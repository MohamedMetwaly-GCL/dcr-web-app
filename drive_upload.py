"""drive_upload.py - Google Drive resumable upload helpers (REST + requests, no googleapiclient).

Auth modes (env DRIVE_UPLOAD_AUTH):
  service_account (default) : GOOGLE_CREDENTIALS_JSON or service_account.json, scope 'drive'
  oauth                     : GOOGLE_OAUTH_CLIENT_ID / GOOGLE_OAUTH_CLIENT_SECRET / GOOGLE_OAUTH_REFRESH_TOKEN
                              Files are then owned by, and counted against, the personal account.
The read-only sync in drive_service.py is NOT affected by anything here.

All HTTP calls accept an injectable `http` object (default: requests) so they are testable offline.
Secrets are never logged or included in exceptions.
"""
import json
import os
import time

import requests

WRITE_SCOPES = ["https://www.googleapis.com/auth/drive"]
UPLOAD_URL = "https://www.googleapis.com/upload/drive/v3/files"
FILES_URL = "https://www.googleapis.com/drive/v3/files"
TOKEN_URI = "https://oauth2.googleapis.com/token"
CHUNK_UNIT = 256 * 1024                      # Drive requires chunk sizes in multiples of 256 KiB
DEFAULT_CHUNK = 8 * 1024 * 1024
FILE_FIELDS = "id,name,size,mimeType,webViewLink,owners(emailAddress),parents"
_cached_creds = None
_shared_session = None

def get_shared_http():
    global _shared_session
    if _shared_session is None:
        _shared_session = requests.Session()
        adapter = requests.adapters.HTTPAdapter(pool_connections=10, pool_maxsize=20, max_retries=2)
        _shared_session.mount("https://", adapter)
        _shared_session.mount("http://", adapter)
    return _shared_session

QUOTA_HINT = ("Drive rejected the upload for storage-quota reasons. A Service Account has no storage of its own "
              "when the target folder is in a personal 'My Drive'. Use OAuth (DRIVE_UPLOAD_AUTH=oauth) or a Shared Drive.")


class DriveUploadError(Exception):
    def __init__(self, message, status=None, reason=None, hint=None):
        super().__init__(message)
        self.status, self.reason, self.hint = status, reason, hint


def classify_drive_error(status, body_text):
    """Return (reason, message, hint) from a Drive error response."""
    reason, message = "", ""
    try:
        err = (json.loads(body_text) or {}).get("error", {})
        message = str(err.get("message", ""))[:300]
        errs = err.get("errors") or []
        reason = (errs[0].get("reason") if errs else "") or str(err.get("status", ""))
    except Exception:
        message = str(body_text)[:200]
    hint = None
    low = (reason + " " + message).lower()
    if "storagequota" in low or "quota" in low and "storage" in low:
        hint = QUOTA_HINT
    elif status == 404:
        hint = "Folder not found or not shared with this account."
    elif status == 403:
        hint = "The authenticated account lacks Editor access to the target folder."
    elif status == 401:
        hint = "Credentials are invalid or expired (for OAuth: refresh token revoked/expired)."
    return reason, message, hint


# ---------------- credentials ----------------
def get_upload_credentials(mode=None, env=None):
    global _cached_creds
    env = os.environ if env is None else env
    has_oauth = bool(env.get("GOOGLE_OAUTH_CLIENT_ID") and env.get("GOOGLE_OAUTH_CLIENT_SECRET") and env.get("GOOGLE_OAUTH_REFRESH_TOKEN"))
    mode = (mode or env.get("DRIVE_UPLOAD_AUTH") or ("oauth" if has_oauth else "service_account")).strip().lower()
    if _cached_creds is not None and getattr(_cached_creds, "_auth_mode", None) == mode:
        if getattr(_cached_creds, "valid", False):
            return _cached_creds
    if mode == "oauth":
        from google.oauth2.credentials import Credentials
        cid, sec, rt = (env.get("GOOGLE_OAUTH_CLIENT_ID"), env.get("GOOGLE_OAUTH_CLIENT_SECRET"),
                        env.get("GOOGLE_OAUTH_REFRESH_TOKEN"))
        if not (cid and sec and rt):
            raise DriveUploadError("OAuth selected but GOOGLE_OAUTH_CLIENT_ID / _CLIENT_SECRET / _REFRESH_TOKEN are not all set",
                                   reason="oauth_not_configured")
        c = Credentials(None, refresh_token=rt, token_uri=TOKEN_URI, client_id=cid,
                        client_secret=sec, scopes=WRITE_SCOPES)
        c._auth_mode = "oauth"
        _cached_creds = c
        return c
    if mode == "service_account":
        from google.oauth2 import service_account
        raw = env.get("GOOGLE_CREDENTIALS_JSON")
        if raw:
            c = service_account.Credentials.from_service_account_info(json.loads(raw), scopes=WRITE_SCOPES)
            c._auth_mode = "service_account"
            _cached_creds = c
            return c
        if os.path.exists("service_account.json"):
            c = service_account.Credentials.from_service_account_file("service_account.json", scopes=WRITE_SCOPES)
            c._auth_mode = "service_account"
            _cached_creds = c
            return c
        raise DriveUploadError("No service account credentials found", reason="service_account_not_configured")
    raise DriveUploadError("Unknown DRIVE_UPLOAD_AUTH mode: %r" % mode, reason="bad_mode")


def access_token(creds):
    if not getattr(creds, "valid", False):
        from google.auth.transport.requests import Request
        try:
            creds.refresh(Request())
        except Exception as e:
            raise DriveUploadError("Could not obtain access token (%s)" % type(e).__name__, status=401,
                                   reason="token_refresh_failed",
                                   hint="For OAuth: refresh token may be revoked/expired, or the consent screen is still in 'Testing' (7-day tokens).")
    return creds.token


def _auth(token, extra=None):
    h = {"Authorization": "Bearer " + token}
    h.update(extra or {})
    return h


def _raise_for(resp, what):
    reason, message, hint = classify_drive_error(resp.status_code, resp.text)
    raise DriveUploadError("%s failed: HTTP %s %s" % (what, resp.status_code, message or reason),
                           status=resp.status_code, reason=reason, hint=hint)


# ---------------- Drive calls ----------------
def get_file(token, file_id, fields="id,name,mimeType,trashed,capabilities(canAddChildren)", http=requests):
    r = http.get("%s/%s" % (FILES_URL, file_id), params={"fields": fields, "supportsAllDrives": "true"},
                 headers=_auth(token), timeout=30)
    if r.status_code != 200:
        _raise_for(r, "files.get")
    return r.json()


def delete_file(token, file_id, http=requests):
    r = http.delete("%s/%s" % (FILES_URL, file_id), params={"supportsAllDrives": "true"},
                    headers=_auth(token), timeout=30)
    if r.status_code not in (200, 204):
        _raise_for(r, "files.delete")
    return True


def trash_file(token, file_id, http=requests):
    """Move a file or folder to Google Drive Trash (recoverable for 30 days)."""
    if not file_id:
        return False
    patch_fn = getattr(http, "patch", None)
    if patch_fn is not None:
        r = patch_fn("%s/%s" % (FILES_URL, file_id),
                     params={"fields": "id,trashed", "supportsAllDrives": "true"},
                     headers=_auth(token, {"Content-Type": "application/json; charset=UTF-8"}),
                     data=json.dumps({"trashed": True}), timeout=30)
        if r.status_code not in (200, 204):
            _raise_for(r, "files.trash")
        return True
    del_fn = getattr(http, "delete", None)
    if del_fn is not None:
        try:
            r = del_fn("%s/%s" % (FILES_URL, file_id), params={"supportsAllDrives": "true"},
                       headers=_auth(token), timeout=30)
            return r.status_code in (200, 204)
        except Exception:
            return False
    return False


def create_resumable_session(token, folder_id, file_name, mime_type, size, origin=None, http=requests):
    """Start a resumable upload. Returns the session URL (valid ~1 week).

    `origin` must be the browser origin when the browser will upload directly (CORS)."""
    if not folder_id:
        raise DriveUploadError("folder_id is required", reason="missing_folder")
    headers = _auth(token, {"Content-Type": "application/json; charset=UTF-8",
                            "X-Upload-Content-Type": mime_type or "application/octet-stream",
                            "X-Upload-Content-Length": str(int(size))})
    if origin:
        headers["Origin"] = origin
    r = http.post(UPLOAD_URL,
                  params={"uploadType": "resumable", "supportsAllDrives": "true", "fields": FILE_FIELDS},
                  headers=headers, data=json.dumps({"name": file_name, "parents": [folder_id]}), timeout=30)
    if r.status_code != 200 or not r.headers.get("Location"):
        _raise_for(r, "create upload session")
    return r.headers["Location"]


def _query_progress(session_url, total, http):
    r = http.put(session_url, headers={"Content-Range": "bytes */%d" % total, "Content-Length": "0"}, timeout=60)
    if r.status_code in (200, 201):
        return total, r.json()
    if r.status_code == 308:
        rng = r.headers.get("Range")             # e.g. 'bytes=0-262143'
        return (int(rng.split("-")[1]) + 1 if rng else 0), None
    _raise_for(r, "upload status query")


def upload_stream_resumable(session_url, fileobj, total, chunk_size=DEFAULT_CHUNK, max_retries=5,
                            http=requests, sleep=time.sleep, progress=None):
    """Upload `total` bytes from a file-like object, one chunk in memory at a time.

    Retries network errors / 5xx with exponential backoff, resuming from Drive's confirmed offset."""
    if chunk_size % CHUNK_UNIT:
        raise DriveUploadError("chunk_size must be a multiple of 256 KiB", reason="bad_chunk")
    if total == 0:
        r = http.put(session_url, headers={"Content-Length": "0"}, timeout=60)
        if r.status_code in (200, 201):
            return r.json()
        _raise_for(r, "upload")
    offset = 0
    while offset < total:
        fileobj.seek(offset)
        chunk = fileobj.read(min(chunk_size, total - offset))
        end = offset + len(chunk) - 1
        headers = {"Content-Length": str(len(chunk)), "Content-Range": "bytes %d-%d/%d" % (offset, end, total)}
        attempt = 0
        while True:
            try:
                r = http.put(session_url, data=chunk, headers=headers, timeout=300)
                if r.status_code in (200, 201):
                    if progress: progress(total, total)
                    return r.json()
                if r.status_code == 308:
                    rng = r.headers.get("Range")
                    offset = int(rng.split("-")[1]) + 1 if rng else 0
                    if progress: progress(offset, total)
                    break
                if r.status_code >= 500 or r.status_code == 429:
                    raise requests.exceptions.ConnectionError("HTTP %d" % r.status_code)
                _raise_for(r, "upload chunk")
            except (requests.exceptions.ConnectionError, requests.exceptions.Timeout):
                # Network problem: back off, then ask Drive how much it really received.
                # The status query itself may fail too; every failure counts as an attempt.
                done = None
                while True:
                    attempt += 1
                    if attempt > max_retries:
                        raise DriveUploadError("Upload failed after %d retries (network)" % max_retries, reason="network")
                    sleep(min(2 ** attempt, 30))
                    try:
                        offset, done = _query_progress(session_url, total, http)
                        break
                    except (requests.exceptions.ConnectionError, requests.exceptions.Timeout):
                        continue
                if done is not None:
                    return done
                break                            # re-read chunk from the confirmed offset
    raise DriveUploadError("Upload ended without completion response", reason="incomplete")



class _RestExec:
    def __init__(self, fn):
        self._fn = fn

    def execute(self):
        return self._fn()


class RestDriveAdapter:
    """Minimal Drive v3 adapter over requests, matching .files().list/get/create().execute()
    expected by drive_routing.py. Works with both OAuth and Service Account tokens."""
    def __init__(self, token, http=None):
        self.token = token
        self.http = http if http is not None else get_shared_http()

    def files(self):
        outer = self

        class _Files:
            def list(self, q, fields="nextPageToken, files(id, name)", pageSize=1000, pageToken=None, **kw):
                def _run():
                    params = {"q": q, "fields": fields, "pageSize": pageSize,
                              "supportsAllDrives": "true", "includeItemsFromAllDrives": "true"}
                    if pageToken:
                        params["pageToken"] = pageToken
                    r = outer.http.get(FILES_URL, params=params, headers=_auth(outer.token), timeout=30)
                    if r.status_code != 200:
                        _raise_for(r, "files.list")
                    return r.json()
                return _RestExec(_run)

            def get(self, fileId, fields="id, mimeType, trashed", **kw):
                return _RestExec(lambda: get_file(outer.token, fileId, fields=fields, http=outer.http))

            def create(self, body, fields="id, name", **kw):
                def _run():
                    r = outer.http.post(FILES_URL, params={"fields": fields, "supportsAllDrives": "true"},
                                        headers=_auth(outer.token, {"Content-Type": "application/json; charset=UTF-8"}),
                                        data=json.dumps(body), timeout=30)
                    if r.status_code not in (200, 201):
                        _raise_for(r, "files.create")
                    return r.json()
                return _RestExec(_run)
        return _Files()


def forward_chunk(session_url, chunk_bytes, content_range, http=requests):
    """Forward a single chunk (or status probe when chunk_bytes is empty) to a Drive session URL."""
    headers = {"Content-Length": str(len(chunk_bytes)), "Content-Range": content_range}
    r = http.put(session_url, data=chunk_bytes, headers=headers, timeout=300)
    if r.status_code in (200, 201):
        return {"status": r.status_code, "done": True, "file": r.json()}
    if r.status_code == 308:
        rng = r.headers.get("Range") or ""
        offset = (int(rng.split("-")[1]) + 1) if "-" in rng else 0
        return {"status": 308, "done": False, "range": rng, "offset": offset}
    _raise_for(r, "upload chunk")
