"""blueprints/drive_viewer.py - In-app document / folder viewer for DCR attachments.

Endpoints (all require a logged-in user who can view the record's project):
  GET /api/drive/view-info/<record_id>/<slot>               -- what is attached (file or folder listing) + permissions
  GET /api/drive/stream/<record_id>/<slot>/<file_id>        -- stream one file (inline preview, or ?dl=1 download)
  GET /api/drive/folder-zip/<record_id>/<slot>              -- download a whole attached folder as one ZIP

<slot> is 'submission' | 'approval' | 'legacy' (the old single fileLocation link).
A file can only be streamed if it IS the record's attachment or lives directly inside the attached folder,
so users can never reach arbitrary Drive files through this gateway. Drive URLs are not exposed to
non-admin users; downloads require edit permission (or admin).
"""
import logging
import re
import tempfile
import zipfile
from urllib.parse import quote

from flask import Blueprint, Response, jsonify, request, send_file

import attachments
import db
import drive_upload
from auth import can_edit, can_view_project, current_user, is_superadmin

logger = logging.getLogger(__name__)
drive_viewer_bp = Blueprint("drive_viewer", __name__)

MAX_ZIP_BYTES = 1024 * 1024 * 1024
MAX_FOLDER_FILES = 500
FOLDER_MIME = "application/vnd.google-apps.folder"
FILE_FIELDS = "id,name,size,mimeType,parents,trashed"

_IMAGE_EXTS = {"png", "jpg", "jpeg", "gif", "webp", "bmp"}
_INLINE_MIME = {"application/pdf", "text/plain", "image/png", "image/jpeg", "image/gif", "image/webp", "image/bmp"}


def _hooks():
    import blueprints.drive_uploads as du
    return du._get_token(), du._get_http()


def preview_kind(name, mime):
    ext = str(name or "").rsplit(".", 1)[-1].lower() if "." in str(name or "") else ""
    mime = str(mime or "").lower()
    if mime == "application/pdf" or ext == "pdf" or mime.startswith("application/vnd.google-apps."):
        return "pdf"
    if (mime.startswith("image/") and mime not in ("image/svg+xml", "image/tiff")) or ext in _IMAGE_EXTS:
        return "image"
    if mime == "text/plain" or ext == "txt":
        return "text"
    return "other"


def _is_workspace_admin(u, pid):
    try:
        proj = db.get_project(pid) or {}
        return bool(is_superadmin(u) or db.is_user_workspace_admin(u.get("username"), proj.get("workspace_id") or "ENG"))
    except Exception:
        return False


def _context(record_id, slot):
    """Return (ctx, None) or (None, (response, status))."""
    u = current_user()
    if not u:
        return None, (jsonify(error="LOGIN_REQUIRED"), 403)
    slot = str(slot or "").strip().lower()
    if slot not in attachments.SLOTS and slot != "legacy":
        return None, (jsonify(error="Invalid slot"), 400)
    rec = db.get_record_by_id(str(record_id or "").strip())
    if not rec:
        return None, (jsonify(error="Record not found"), 404)
    pid, dt_id = rec.get("_project_id") or "", rec.get("_dt_id") or ""
    if not can_view_project(pid, u):
        try:
            db.log_action(u.get("username", ""), "ACCESS DENIED", project_id=pid, dt_id=dt_id,
                          detail="Unauthorized document view attempt")
        except Exception:
            pass
        return None, (jsonify(error="Forbidden"), 403)
    if not is_superadmin(u) and not db.is_user_workspace_admin(u.get("username")):
        allowed = db.get_user_project_doc_types(u["username"], pid)
        if allowed is not None:
            allowed_set = {str(x).upper() for x in allowed}
            dt_obj = db.get_doc_type(pid, dt_id) or {}
            if str(dt_id).upper() not in allowed_set and str(dt_obj.get("code") or "").upper() not in allowed_set:
                return None, (jsonify(error="Document type access denied"), 403)

    att = rec.get("attachments") if isinstance(rec.get("attachments"), dict) else {}
    if slot == "legacy":
        entry = {"url": attachments._clean_link(rec.get("fileLocation")), "legacy": True}
    else:
        entry = att.get(slot) if isinstance(att.get(slot), dict) else None
    if not entry or not (entry.get("file_id") or entry.get("folder_id") or entry.get("url")):
        return None, (jsonify(error="Nothing is attached in this slot"), 404)

    url = str(entry.get("url") or "")
    if entry.get("is_folder") or "/folders/" in url:
        m = re.search(r"/folders/([a-zA-Z0-9_-]+)", url)
        target_id = entry.get("folder_id") or entry.get("file_id") or (m.group(1) if m else "")
        kind = "folder"
    else:
        target_id = entry.get("file_id") or attachments.extract_file_id(url)
        kind = "file"
    if not target_id:
        return None, (jsonify(error="Attachment has no Drive id", fallback_url=url if is_superadmin(u) else ""), 404)

    can_dl = bool(can_edit(pid) or _is_workspace_admin(u, pid))
    ctx = dict(user=u, rec=rec, pid=pid, dt_id=dt_id, slot=slot, entry=entry, kind=kind, target_id=target_id,
               can_download=can_dl, can_open_drive=_is_workspace_admin(u, pid), url=url)
    return ctx, None


def _list_folder(token, http, folder_id):
    svc = drive_upload.RestDriveAdapter(token, http=http)
    out, page = [], None
    while len(out) < MAX_FOLDER_FILES:
        res = svc.files().list(q="'%s' in parents and trashed=false" % folder_id,
                               fields="nextPageToken, files(id,name,size,mimeType)", pageToken=page).execute()
        for f in res.get("files", []):
            if f.get("mimeType") != FOLDER_MIME:
                out.append(f)
        page = res.get("nextPageToken")
        if not page:
            break
    out.sort(key=lambda f: str(f.get("name", "")).lower())
    return out[:MAX_FOLDER_FILES]


def _file_dict(f, rid, slot, can_dl):
    kind = preview_kind(f.get("name"), f.get("mimeType"))
    base = "/api/drive/stream/%s/%s/%s" % (quote(rid, safe=""), slot, quote(str(f["id"]), safe=""))
    return {"id": f["id"], "name": f.get("name") or "file", "size": int(f.get("size") or 0),
            "mime": f.get("mimeType") or "", "preview": kind,
            "stream_url": base if kind != "other" else "",
            "download_url": (base + "?dl=1") if can_dl else ""}


def _log(ctx, action, detail):
    try:
        rec = ctx["rec"]
        db.log_action(ctx["user"]["username"], action, project_id=ctx["pid"], dt_id=ctx["dt_id"],
                      record_id=rec.get("_id") or rec.get("id"), doc_no=rec.get("docNo", ""), detail=detail)
    except Exception:
        pass


@drive_viewer_bp.route("/api/drive/view-info/<record_id>/<slot>")
def api_view_info(record_id, slot):
    ctx, err = _context(record_id, slot)
    if err:
        return err
    rid = ctx["rec"].get("_id") or record_id
    try:
        token, http = _hooks()
        if ctx["kind"] == "folder":
            files = _list_folder(token, http, ctx["target_id"])
            if not files and isinstance(ctx["entry"].get("files"), list):
                files = [
                    {
                        "id": f.get("id") or f"ITEM_{idx + 1}",
                        "name": f.get("name") or f"File_{idx + 1}",
                        "size": int(f.get("size") or 0),
                        "mimeType": f.get("mimeType") or "",
                    }
                    for idx, f in enumerate(ctx["entry"]["files"])
                ]
        else:
            meta = drive_upload.get_file(token, ctx["target_id"], fields=FILE_FIELDS, http=http)
            files = [meta]
    except drive_upload.DriveUploadError as e:
        return jsonify(ok=False, error=str(e), hint=e.hint,
                       fallback_url=ctx["url"] if (ctx["can_open_drive"] or ctx["entry"].get("legacy")) else ""), 502
    entry = ctx["entry"]
    name = entry.get("folder_name") or entry.get("file_name") or (files[0].get("name") if files else "Document")
    _log(ctx, "DRIVE_LINK_CLICK", "Opened in-app viewer (%s / %s)" % (ctx["slot"], name))
    return jsonify(
        ok=True, kind=ctx["kind"], slot=ctx["slot"], name=name,
        files=[_file_dict(f, rid, ctx["slot"], ctx["can_download"]) for f in files],
        total_size=sum(int(f.get("size") or 0) for f in files),
        can_download=ctx["can_download"], can_open_drive=ctx["can_open_drive"],
        drive_url=ctx["url"] if ctx["can_open_drive"] else "",
        zip_url=("/api/drive/folder-zip/%s/%s" % (quote(rid, safe=""), ctx["slot"]))
        if (ctx["kind"] == "folder" and ctx["can_download"]) else "")


def _authorize_file(ctx, token, http, file_id):
    """Return Drive metadata if file_id is the attachment itself or sits directly inside the attached folder."""
    meta = drive_upload.get_file(token, file_id, fields=FILE_FIELDS, http=http)
    if meta.get("trashed"):
        return None
    if ctx["kind"] == "file":
        return meta if file_id == ctx["target_id"] else None
    return meta if ctx["target_id"] in (meta.get("parents") or []) else None


def _safe_disposition(kind, name):
    fallback = re.sub(r'[^A-Za-z0-9._ -]+', "_", name) or "file"
    return "%s; filename=\"%s\"; filename*=UTF-8''%s" % (kind, fallback, quote(name))


def _drive_media(token, http, meta):
    mime = str(meta.get("mimeType") or "")
    if mime.startswith("application/vnd.google-apps."):
        url, params = "%s/%s/export" % (drive_upload.FILES_URL, meta["id"]), {"mimeType": "application/pdf"}
    else:
        url, params = "%s/%s" % (drive_upload.FILES_URL, meta["id"]), {"alt": "media", "supportsAllDrives": "true"}
    return http.get(url, params=params, headers={"Authorization": "Bearer " + token}, timeout=120, stream=True)


@drive_viewer_bp.route("/api/drive/stream/<record_id>/<slot>/<file_id>")
def api_stream(record_id, slot, file_id):
    ctx, err = _context(record_id, slot)
    if err:
        return err
    download = request.args.get("dl") == "1"
    if download and not ctx["can_download"]:
        return jsonify(error="Download permission required"), 403
    try:
        token, http = _hooks()
        meta = _authorize_file(ctx, token, http, str(file_id))
        if not meta:
            return jsonify(error="File does not belong to this attachment"), 403
        resp = _drive_media(token, http, meta)
    except drive_upload.DriveUploadError as e:
        return jsonify(error=str(e), hint=e.hint), 502
    if resp.status_code != 200:
        return jsonify(error="Drive returned HTTP %s" % resp.status_code), 502

    name = meta.get("name") or "file"
    mime = "application/pdf" if str(meta.get("mimeType") or "").startswith("application/vnd.google-apps.") \
        else str(meta.get("mimeType") or "application/octet-stream").lower()
    inline_ok = (not download) and mime in _INLINE_MIME
    if not inline_ok and not download:
        return jsonify(error="No inline preview for this file type"), 415
    if not request.headers.get("Range") or str(request.headers.get("Range")).startswith("bytes=0-"):
        _log(ctx, "FILE_DOWNLOAD" if download else "DRIVE_LINK_CLICK",
             "%s '%s' via in-app gateway" % ("Downloaded" if download else "Previewed", name))

    chunks = resp.iter_content(1024 * 1024) if hasattr(resp, "iter_content") else iter([resp.content])
    out = Response(chunks, mimetype=mime if inline_ok else "application/octet-stream")
    out.headers["Content-Disposition"] = _safe_disposition("inline" if inline_ok else "attachment", name)
    out.headers["X-Content-Type-Options"] = "nosniff"
    out.headers["Cache-Control"] = "private, no-store"
    clen = resp.headers.get("Content-Length") if hasattr(resp, "headers") else None
    if clen:
        out.headers["Content-Length"] = clen
    return out


@drive_viewer_bp.route("/api/drive/folder-zip/<record_id>/<slot>")
def api_folder_zip(record_id, slot):
    ctx, err = _context(record_id, slot)
    if err:
        return err
    if ctx["kind"] != "folder":
        return jsonify(error="Attachment is not a folder"), 400
    if not ctx["can_download"]:
        return jsonify(error="Download permission required"), 403
    try:
        token, http = _hooks()
        files = _list_folder(token, http, ctx["target_id"])
        if sum(int(f.get("size") or 0) for f in files) > MAX_ZIP_BYTES:
            return jsonify(error="Folder is larger than 1 GB - download files one by one"), 413
        tmp = tempfile.TemporaryFile()
        used = set()
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_STORED, allowZip64=True) as zf:
            for f in files:
                resp = _drive_media(token, http, f)
                if resp.status_code != 200:
                    continue
                arc = f.get("name") or f["id"]
                if arc in used:
                    arc = "%s_%s" % (f["id"][:6], arc)
                used.add(arc)
                with zf.open(arc, "w", force_zip64=True) as dest:
                    it = resp.iter_content(1024 * 1024) if hasattr(resp, "iter_content") else iter([resp.content])
                    for chunk in it:
                        dest.write(chunk)
        tmp.seek(0)
    except drive_upload.DriveUploadError as e:
        return jsonify(error=str(e), hint=e.hint), 502
    zname = re.sub(r'[\\/:*?"<>|]+', "_", ctx["entry"].get("folder_name") or ctx["entry"].get("file_name") or "documents")
    _log(ctx, "FILE_DOWNLOAD", "Downloaded folder '%s' as ZIP (%d files)" % (zname, len(files)))
    return send_file(tmp, mimetype="application/zip", as_attachment=True, download_name=zname + ".zip")
