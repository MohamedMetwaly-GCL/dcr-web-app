"""Offline tests for drive_upload.py with a scripted fake HTTP layer (no network). Run: python test_drive_upload.py"""
import io
import json
import requests
import drive_upload as U


class Resp:
    def __init__(self, status=200, body=None, headers=None):
        self.status_code, self.headers = status, headers or {}
        self.text = json.dumps(body) if isinstance(body, (dict, list)) else (body or "")
        self._b = body

    def json(self):
        return self._b


class FakeHTTP:
    def __init__(self, script):
        self.script, self.calls = list(script), []

    def _next(self, method, url, **kw):
        self.calls.append((method, url, kw))
        item = self.script.pop(0)
        if isinstance(item, Exception):
            raise item
        return item

    post = lambda self, url, **kw: self._next("POST", url, **kw)
    put = lambda self, url, **kw: self._next("PUT", url, **kw)
    get = lambda self, url, **kw: self._next("GET", url, **kw)
    delete = lambda self, url, **kw: self._next("DELETE", url, **kw)


QERR = {"error": {"code": 403, "message": "Service Accounts do not have storage quota.",
                  "errors": [{"reason": "storageQuotaExceeded"}]}}

# ---- error classification
r, m, h = U.classify_drive_error(403, json.dumps(QERR))
assert r == "storageQuotaExceeded" and h == U.QUOTA_HINT
assert "Editor" in U.classify_drive_error(403, json.dumps({"error": {"message": "x", "errors": [{"reason": "forbidden"}]}}))[2]
assert U.classify_drive_error(404, "{}")[2] and U.classify_drive_error(401, "not json")[2]

# ---- create session: request shape + errors
h = FakeHTTP([Resp(200, {}, {"Location": "https://upload/session1"})])
url = U.create_resumable_session("TOK", "FOLDER1", "a.pdf", "application/pdf", 1234, origin="https://dcr.app", http=h)
method, u, kw = h.calls[0]
assert url == "https://upload/session1" and u == U.UPLOAD_URL
assert kw["params"]["uploadType"] == "resumable" and kw["params"]["supportsAllDrives"] == "true"
assert json.loads(kw["data"]) == {"name": "a.pdf", "parents": ["FOLDER1"]}          # exactly one parent
assert kw["headers"]["Authorization"] == "Bearer TOK" and kw["headers"]["X-Upload-Content-Length"] == "1234"
assert kw["headers"]["X-Upload-Content-Type"] == "application/pdf" and kw["headers"]["Origin"] == "https://dcr.app"
try:
    U.create_resumable_session("T", "F", "a", "x", 1, http=FakeHTTP([Resp(403, QERR)])); raise SystemExit("expected")
except U.DriveUploadError as e:
    assert e.reason == "storageQuotaExceeded" and e.hint == U.QUOTA_HINT and e.status == 403
    assert "TOK" not in str(e)
try:
    U.create_resumable_session("T", "", "a", "x", 1, http=FakeHTTP([])); raise SystemExit("expected")
except U.DriveUploadError as e:
    assert e.reason == "missing_folder"

# ---- chunked upload (3 chunks of 256 KiB) with progress; only one chunk read at a time
class SpyIO(io.BytesIO):
    max_read = 0
    def read(self, n=-1):
        SpyIO.max_read = max(SpyIO.max_read, n)
        return super().read(n)

C = U.CHUNK_UNIT
data = bytes(range(256)) * (C * 3 // 256)
h = FakeHTTP([Resp(308, headers={"Range": "bytes=0-%d" % (C - 1)}),
              Resp(308, headers={"Range": "bytes=0-%d" % (2 * C - 1)}),
              Resp(200, {"id": "FID"})])
prog = []
res = U.upload_stream_resumable("https://s", SpyIO(data), len(data), chunk_size=C, http=h, progress=lambda d, t: prog.append(d))
assert res == {"id": "FID"} and prog == [C, 2 * C, 3 * C]
assert [c[2]["headers"]["Content-Range"] for c in h.calls] == [
    "bytes 0-%d/%d" % (C - 1, 3 * C), "bytes %d-%d/%d" % (C, 2 * C - 1, 3 * C), "bytes %d-%d/%d" % (2 * C, 3 * C - 1, 3 * C)]
assert SpyIO.max_read == C                      # never more than one chunk buffered
print("chunked upload OK")

# ---- network drop mid-upload -> status query -> resume from Drive's confirmed offset
h = FakeHTTP([Resp(308, headers={"Range": "bytes=0-%d" % (C - 1)}),
              requests.exceptions.ConnectionError("boom"),                       # chunk 2 fails
              Resp(308, headers={"Range": "bytes=0-%d" % (C + 1024 - 1)}),       # Drive got part of it
              Resp(308, headers={"Range": "bytes=0-%d" % (2 * C - 1)}),          # re-sent from C+1024
              Resp(200, {"id": "FID2"})])
sleeps = []
res = U.upload_stream_resumable("https://s", io.BytesIO(data), len(data), chunk_size=C, http=h, sleep=sleeps.append)
assert res == {"id": "FID2"} and sleeps == [2]
ranges = [c[2]["headers"]["Content-Range"] for c in h.calls]
assert ranges[2] == "bytes */%d" % len(data) and ranges[3].startswith("bytes %d-" % (C + 1024))
print("resume after network drop OK")

# ---- upload finished although the response was lost
h = FakeHTTP([requests.exceptions.Timeout("t"), Resp(200, {"id": "FID3"})])
assert U.upload_stream_resumable("https://s", io.BytesIO(b"abc"), 3, chunk_size=C, http=h, sleep=lambda s: None) == {"id": "FID3"}

# ---- retries exhausted, 5xx treated as retryable, 4xx not
h = FakeHTTP([requests.exceptions.ConnectionError("x")] * 3)   # upload fails, then status queries keep failing
try:
    U.upload_stream_resumable("https://s", io.BytesIO(b"abc"), 3, chunk_size=C, http=h, max_retries=2, sleep=lambda s: None)
    raise SystemExit("expected")
except U.DriveUploadError as e:
    assert e.reason == "network" and not h.script
h = FakeHTTP([Resp(503, "x"), Resp(308, headers={}), Resp(200, {"id": "OK"})])
assert U.upload_stream_resumable("https://s", io.BytesIO(b"abc"), 3, chunk_size=C, http=h, sleep=lambda s: None)["id"] == "OK"
try:
    U.upload_stream_resumable("https://s", io.BytesIO(b"abc"), 3, chunk_size=C, http=FakeHTTP([Resp(403, QERR)])); raise SystemExit("expected")
except U.DriveUploadError as e:
    assert e.status == 403 and e.reason == "storageQuotaExceeded"

# ---- zero byte / bad chunk size
assert U.upload_stream_resumable("https://s", io.BytesIO(b""), 0, http=FakeHTTP([Resp(200, {"id": "E"})]))["id"] == "E"
try:
    U.upload_stream_resumable("https://s", io.BytesIO(b"x"), 1, chunk_size=1000); raise SystemExit("expected")
except U.DriveUploadError as e:
    assert e.reason == "bad_chunk"

# ---- get / delete
h = FakeHTTP([Resp(200, {"id": "F", "name": "n"}), Resp(204, ""), Resp(404, {"error": {"message": "nf"}})])
assert U.get_file("T", "F", http=h)["id"] == "F" and h.calls[0][2]["params"]["supportsAllDrives"] == "true"
assert U.delete_file("T", "F", http=h) is True
try:
    U.delete_file("T", "F", http=h); raise SystemExit("expected")
except U.DriveUploadError as e:
    assert e.status == 404

# ---- credentials selection (no network, no real secrets)
for env, reason in (({"DRIVE_UPLOAD_AUTH": "oauth"}, "oauth_not_configured"),
                    ({"DRIVE_UPLOAD_AUTH": "nonsense"}, "bad_mode")):
    try:
        U.get_upload_credentials(env=env); raise SystemExit("expected")
    except U.DriveUploadError as e:
        assert e.reason == reason
c = U.get_upload_credentials(env={"DRIVE_UPLOAD_AUTH": "oauth", "GOOGLE_OAUTH_CLIENT_ID": "id",
                                  "GOOGLE_OAUTH_CLIENT_SECRET": "sec", "GOOGLE_OAUTH_REFRESH_TOKEN": "rt"})
assert c.refresh_token == "rt" and c.scopes == U.WRITE_SCOPES
assert "rt" not in repr(U.DriveUploadError("x"))
print("ALL DRIVE UPLOAD TESTS PASSED")


# ---- RestDriveAdapter + forward_chunk (Step 10)
h = FakeHTTP([
    Resp(200, {"files": [{"id": "F1", "name": "01 - SD"}]}),
    Resp(200, {"id": "F1", "mimeType": "application/vnd.google-apps.folder", "trashed": False}),
    Resp(201, {"id": "F2", "name": "Submit"}),
    Resp(308, {}, {"Range": "bytes=0-262143"}),
    Resp(200, {"id": "UP_DONE"}),
])
ad = U.RestDriveAdapter("TOK", http=h)
assert ad.files().list(q="x").execute()["files"][0]["id"] == "F1"
assert ad.files().get(fileId="F1").execute()["trashed"] is False
assert ad.files().create(body={"name": "Submit"}).execute()["id"] == "F2"
r1 = U.forward_chunk("https://s", b"x" * 10, "bytes 0-9/20", http=h)
assert r1["status"] == 308 and r1["done"] is False and r1["offset"] == 262144
r2 = U.forward_chunk("https://s", b"x" * 10, "bytes 10-19/20", http=h)
assert r2["status"] == 200 and r2["done"] is True and r2["file"]["id"] == "UP_DONE"
print("RestDriveAdapter + forward_chunk OK")
