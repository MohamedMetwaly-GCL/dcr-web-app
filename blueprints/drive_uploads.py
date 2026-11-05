"""blueprints/drive_uploads.py - Direct Resumable Upload & Dual Attachment APIs for DCR.

Endpoints:
  POST /api/drive/upload-session   -- authorize upload, resolve Drive folder, create resumable session
  POST /api/drive/upload-chunk     -- optional chunk-proxy fallback (<=16 MB per chunk, zero full buffering)
  POST /api/drive/upload-complete  -- verify uploaded file on Drive, attach to record, write Audit Log
  DELETE /api/drive/attachment/<record_id>/<slot> -- unlink attachment and move old Drive file/folder to Trash
"""
import logging
import os
import re
from flask import Blueprint, jsonify, request
from itsdangerous import BadSignature, SignatureExpired, URLSafeTimedSerializer

import attachments
import db
import drive_routing
import drive_upload
from auth import can_edit, current_user, is_superadmin
from config import SECRET_KEY
from security import DANGEROUS_EXTS_INSIDE_ARCHIVE

logger = logging.getLogger(__name__)
drive_uploads_bp = Blueprint("drive_uploads", __name__)

MAX_UPLOAD_BYTES = 1024 * 1024 * 1024          # 1 GB max per file/archive
MAX_PROXY_CHUNK_BYTES = 16 * 1024 * 1024       # 16 MB max per proxied chunk
TICKET_MAX_AGE = 86400                         # 24 hours
TICKET_SALT = "dcr-drive-upload-v1"

# Injectable hooks for offline unit/integration testing
_token_provider = None
_http_client = None


def _get_token():
    if _token_provider is not None:
        return _token_provider()
    creds = drive_upload.get_upload_credentials()
    return drive_upload.access_token(creds)


def _get_http():
    return _http_client if _http_client is not None else drive_upload.requests


def _serializer():
    return URLSafeTimedSerializer(SECRET_KEY, salt=TICKET_SALT)


def sign_upload_ticket(payload):
    return _serializer().dumps(payload)


def verify_upload_ticket(token_str, max_age=TICKET_MAX_AGE):
    return _serializer().loads(token_str, max_age=max_age)


def sanitize_upload_filename(raw_name):
    """Return (clean_name, error_reason). Rejects empty names and dangerous executable extensions."""
    name = str(raw_name or "").replace("\\", "/").split("/")[-1].strip()
    name = re.sub(r"[\x00-\x1f\x7f]+", "", name).strip(". ")
    if not name or len(name) > 240:
        return "", "invalid_file_name"
    _, ext = os.path.splitext(name.lower())
    if ext in DANGEROUS_EXTS_INSIDE_ARCHIVE:
        return "", "forbidden_file_extension"
    return name, ""


def _trash_old_attachment(entry, exclude_id="", files_only=False):
    """Best-effort move of an old attachment's Drive file(s) or folder to Google Drive Trash."""
    if not isinstance(entry, dict) or entry.get("legacy"):
        return
    try:
        http = _get_http()
        # Only call real Drive PATCH when http supports .patch (avoids consuming scripted FakeHTTP queues in unit tests)
        if getattr(http, "patch", None) is None:
            return
        token = _get_token()
        if files_only and isinstance(entry.get("files"), list):
            for f in entry["files"]:
                fid = str(f.get("id") or "").strip()
                if fid and fid != exclude_id:
                    try:
                        drive_upload.trash_file(token, fid, http=http)
                    except Exception:
                        pass
            return
        target_id = str(entry.get("folder_id") or entry.get("file_id") or "").strip()
        if target_id and target_id != exclude_id:
            drive_upload.trash_file(token, target_id, http=http)
    except Exception as e:
        logger.warning("Could not move old Drive attachment to trash: %s", e)


def _check_write_access(pid, dt_id):
    """Verify login, project write permission, and doc-type permission.
    Returns (user_dict, None) on success or (None, (response, status_code)) on failure.
    """
    u = current_user()
    if not u:
        return None, (jsonify(error="LOGIN_REQUIRED"), 403)
    if not pid or not dt_id:
        return None, (jsonify(error="project_id and dt_id are required"), 400)
    if not can_edit(pid):
        try:
            db.log_action(u.get("username", ""), "ACCESS DENIED", project_id=pid, dt_id=dt_id,
                          detail=f"Unauthorized upload attempt on project '{pid}'")
        except Exception:
            pass
        return None, (jsonify(error="Forbidden"), 403)
    if not is_superadmin(u) and not db.is_user_workspace_admin(u.get("username")):
        allowed_dts = db.get_user_project_doc_types(u["username"], pid)
        if allowed_dts is not None:
            allowed_set = {str(x).upper() for x in allowed_dts}
            dt_obj = db.get_doc_type(pid, dt_id)
            dt_code = str(dt_obj.get("code") or "").upper() if dt_obj else ""
            if str(dt_id).upper() not in allowed_set and dt_code not in allowed_set:
                return None, (jsonify(error="Document type access denied"), 403)
    return u, None


@drive_uploads_bp.route("/api/drive/upload-session", methods=["POST"])
def api_create_upload_session():
    body = request.get_json(silent=True) or {}
    pid = str(body.get("project_id") or "").strip()
    dt_id = str(body.get("dt_id") or "").strip()
    record_id = str(body.get("record_id") or "").strip()
    slot = str(body.get("slot") or attachments.SLOT_SUBMISSION).strip().lower()
    subfolder_name = str(body.get("subfolder_name") or "").strip()
    link_as_folder = bool(body.get("link_as_folder", False))
    mime_type = str(body.get("mime_type") or "application/octet-stream").strip()
    create_folders = bool(body.get("create_folders", True))

    u, err = _check_write_access(pid, dt_id)
    if err:
        return err

    if slot not in attachments.SLOTS:
        return jsonify(error=f"Invalid slot '{slot}'. Must be 'submission' or 'approval'."), 400

    proj = db.get_project(pid)
    if not proj:
        return jsonify(error="Project not found"), 404
    ws_id = str(proj.get("workspace_id") or "ENG").strip().upper()
    if slot == attachments.SLOT_APPROVAL and ws_id != "ENG":
        return jsonify(error="Approval slot is only available in the Engineering & Projects workspace"), 400
    dt_meta = db.get_doc_type(pid, dt_id) or {}
    has_cycle = attachments.doc_type_has_reply_cycle(dt_meta)
    if slot == attachments.SLOT_APPROVAL and not has_cycle:
        return jsonify(error="This document type has no consultant reply cycle"), 400

    if record_id:
        rec = db.get_record_by_id(record_id)
        if not rec or rec.get("_project_id") != pid or rec.get("_dt_id") != dt_id:
            return jsonify(error="Record not found in this project / document type"), 404

    clean_name, name_err = sanitize_upload_filename(body.get("file_name"))
    if name_err:
        if name_err == "forbidden_file_extension":
            try:
                db.log_action(u.get("username", ""), "FILE_REJECTED", project_id=pid, dt_id=dt_id,
                              record_id=record_id or None, detail=f"Blocked dangerous file extension: {body.get('file_name')}")
            except Exception:
                pass
        return jsonify(error=name_err), 400

    try:
        file_size = int(body.get("file_size"))
    except (TypeError, ValueError):
        return jsonify(error="file_size must be a positive integer"), 400
    if file_size <= 0 or file_size > MAX_UPLOAD_BYTES:
        return jsonify(error=f"file_size must be between 1 byte and {MAX_UPLOAD_BYTES} bytes (1 GB)"), 400

    try:
        token = _get_token()
        http = _get_http()
        service = drive_upload.RestDriveAdapter(token, http=http)
        dest = drive_routing.resolve_upload_destination(
            pid, dt_id, slot, service, db, create_if_missing=create_folders,
            subfolder_name=subfolder_name or None, flat=not has_cycle)
        session_url = drive_upload.create_resumable_session(
            token, dest["folder_id"], clean_name, mime_type, file_size,
            origin=request.headers.get("Origin"), http=http)
    except drive_routing.RoutingError as e:
        return jsonify(error=e.reason, candidates=e.candidates), 400
    except drive_upload.DriveUploadError as e:
        return jsonify(error=str(e), reason=e.reason, hint=e.hint), (e.status if e.status in (400, 401, 403, 404) else 502)

    ticket = sign_upload_ticket({
        "pid": pid,
        "dt_id": dt_id,
        "record_id": record_id,
        "slot": slot,
        "folder_id": dest["folder_id"],
        "subfolder_id": dest.get("subfolder_id", ""),
        "subfolder_name": dest.get("subfolder_name", ""),
        "link_as_folder": link_as_folder,
        "file_name": clean_name,
        "file_size": file_size,
        "mime_type": mime_type,
        "user": u["username"],
        "session_url": session_url,
    })
    return jsonify(
        ok=True,
        session_url=session_url,
        upload_ticket=ticket,
        folder_id=dest["folder_id"],
        doc_type_folder_id=dest["doc_type_folder_id"],
        folder_source=dest["source"],
        chunk_size=drive_upload.DEFAULT_CHUNK,
    )


@drive_uploads_bp.route("/api/drive/upload-chunk", methods=["POST", "PUT"])
def api_upload_chunk():
    """Chunk-proxy fallback (<=16 MB per call). Streams a single chunk to the Drive session URL."""
    ticket_str = request.headers.get("X-Upload-Ticket") or request.args.get("upload_ticket") or ""
    content_range = request.headers.get("X-Content-Range") or request.headers.get("Content-Range") or ""
    if not ticket_str or not content_range:
        return jsonify(error="X-Upload-Ticket and Content-Range headers are required"), 400
    try:
        t = verify_upload_ticket(ticket_str)
    except SignatureExpired:
        return jsonify(error="upload_ticket expired"), 400
    except BadSignature:
        return jsonify(error="invalid upload_ticket"), 400

    u, err = _check_write_access(t["pid"], t["dt_id"])
    if err:
        return err
    if u["username"] != t.get("user") and not is_superadmin(u):
        return jsonify(error="Ticket belongs to a different user"), 403

    if request.content_length and request.content_length > MAX_PROXY_CHUNK_BYTES:
        return jsonify(error="Chunk exceeds maximum proxy chunk size (16 MB)"), 413
    chunk_bytes = request.get_data(cache=False)
    if len(chunk_bytes) > MAX_PROXY_CHUNK_BYTES:
        return jsonify(error="Chunk exceeds maximum proxy chunk size (16 MB)"), 413

    try:
        res = drive_upload.forward_chunk(t["session_url"], chunk_bytes, content_range, http=_get_http())
        return jsonify(ok=True, **res)
    except drive_upload.DriveUploadError as e:
        return jsonify(error=str(e), reason=e.reason, hint=e.hint), (e.status or 502)


@drive_uploads_bp.route("/api/drive/upload-complete", methods=["POST"])
def api_upload_complete():
    body = request.get_json(silent=True) or {}
    ticket_str = str(body.get("upload_ticket") or "").strip()
    file_id = str(body.get("file_id") or "").strip()
    replace_existing = bool(body.get("replace_existing", False))
    rel_subfolder = str(body.get("rel_subfolder") or "").strip()
    if not ticket_str or not file_id:
        return jsonify(error="upload_ticket and file_id are required"), 400

    try:
        t = verify_upload_ticket(ticket_str)
    except SignatureExpired:
        return jsonify(error="upload_ticket expired"), 400
    except BadSignature:
        return jsonify(error="invalid upload_ticket"), 400

    pid, dt_id, slot = t["pid"], t["dt_id"], t["slot"]
    record_id = str(body.get("record_id") or t.get("record_id") or "").strip()
    if not record_id:
        return jsonify(error="record_id is required to attach the uploaded file"), 400

    u, err = _check_write_access(pid, dt_id)
    if err:
        return err
    if u["username"] != t.get("user") and not is_superadmin(u):
        return jsonify(error="Ticket belongs to a different user"), 403

    rec = db.get_record_by_id(record_id)
    if not rec or rec.get("_project_id") != pid or rec.get("_dt_id") != dt_id:
        return jsonify(error="Record not found in this project / document type"), 404

    try:
        token = _get_token()
        remote = drive_upload.get_file(token, file_id, fields=drive_upload.FILE_FIELDS + ",trashed", http=_get_http())
    except drive_upload.DriveUploadError as e:
        return jsonify(error=str(e), reason=e.reason, hint=e.hint), (e.status if e.status in (400, 401, 403, 404) else 502)

    if remote.get("trashed"):
        return jsonify(error="Uploaded file is in trash on Drive"), 400
    parents = remote.get("parents") or []
    if parents and t["folder_id"] not in parents:
        return jsonify(error="File parent folder does not match authorized destination folder"), 400

    view_url = remote.get("webViewLink") or attachments.drive_view_url(file_id)
    extra = {"url": view_url}
    if slot == attachments.SLOT_APPROVAL and rec.get("status"):
        extra["consultant_status"] = str(rec.get("status"))

    # Folder / Multi-file attachment handling
    if t.get("link_as_folder") and (t.get("subfolder_id") or t.get("folder_id")):
        fid = t.get("subfolder_id") or t["folder_id"]
        folder_url = f"https://drive.google.com/drive/folders/{fid}"
        extra["is_folder"] = True
        extra["folder_id"] = fid
        extra["folder_name"] = t.get("subfolder_name") or t["file_name"]
        extra["folder_url"] = folder_url
        extra["url"] = folder_url
        view_url = folder_url

    meta = attachments.build_attachment_meta(
        file_id=file_id,
        name=remote.get("name") or t["file_name"],
        size=int(remote.get("size") or t["file_size"] or 0),
        mime_type=remote.get("mimeType") or t["mime_type"],
        uploaded_by=u["username"],
        **extra,
    )

    existing_att = rec.get("attachments") if isinstance(rec.get("attachments"), dict) else {}
    existing_slot = existing_att.get(slot) if isinstance(existing_att.get(slot), dict) else None
    if extra.get("is_folder"):
        if replace_existing and existing_slot:
            # First file of a replacement folder upload: move previous files inside that folder to Drive Trash
            _trash_old_attachment(existing_slot, exclude_id=file_id, files_only=True)
            files_list = []
        elif existing_slot and existing_slot.get("is_folder") and (existing_slot.get("folder_id") == extra.get("folder_id")):
            files_list = list(existing_slot.get("files") or [])
        else:
            files_list = []
        files_list = [f for f in files_list if f.get("id") != file_id]
        f_item = {
            "id": file_id,
            "name": meta["file_name"],
            "size": meta["file_size"],
            "mimeType": meta["mime_type"],
        }
        if rel_subfolder:
            f_item["subfolder"] = rel_subfolder
        files_list.append(f_item)
        meta["files"] = files_list
        meta["file_count"] = len(files_list)
        meta["file_size"] = sum(int(f.get("size") or 0) for f in files_list)
        meta["file_name"] = (existing_slot or {}).get("folder_name") or extra.get("folder_name") or meta["file_name"]
    else:
        # Single-file upload replacing an existing attachment in the same slot: move the old file/folder to Trash
        if existing_slot:
            _trash_old_attachment(existing_slot, exclude_id=file_id)

    # Preserve legacy single fileLocation as submission if attaching an approval to an old record
    old_loc = attachments._clean_link(rec.get("fileLocation"))
    if slot == attachments.SLOT_APPROVAL and not existing_att.get(attachments.SLOT_SUBMISSION) and old_loc and old_loc != view_url:
        db.set_record_attachment(record_id, attachments.SLOT_SUBMISSION, {"url": old_loc, "legacy": True})

    db.set_record_attachment(record_id, slot, meta)

    # If the record had no fileLocation yet (or if replacing the slot that owned fileLocation), keep fileLocation synced
    prev_slot_url = attachments._clean_link((existing_slot or {}).get("url"))
    if not old_loc or (prev_slot_url and old_loc == prev_slot_url):
        db.q("""UPDATE records SET data = data || jsonb_build_object('fileLocation', %s::text), updated_at = NOW()
                WHERE id = %s RETURNING id""", (view_url, record_id))

    try:
        db.log_action(
            u["username"], "FILE_UPLOAD",
            project_id=pid, dt_id=dt_id, record_id=record_id,
            doc_no=rec.get("docNo", ""),
            field_name=f"attachments.{slot}",
            new_value=meta["file_name"],
            detail=f"Uploaded {slot} file '{meta['file_name']}' ({meta['file_size']} bytes) to Drive [{file_id}]",
        )
    except Exception:
        pass

    return jsonify(
        ok=True,
        record_id=record_id,
        slot=slot,
        attachment=meta,
        attachments=db.get_record_attachments(record_id),
    )


@drive_uploads_bp.route("/api/drive/attachment/<record_id>/<slot>", methods=["DELETE"])
def api_unlink_attachment(record_id, slot):
    """Unlink an attachment slot ('submission' or 'approval') from a record and move the old Drive file/folder to Trash."""
    record_id = str(record_id or "").strip()
    slot = str(slot or "").strip().lower()
    if slot not in attachments.SLOTS:
        return jsonify(error=f"Invalid slot '{slot}'. Must be 'submission' or 'approval'."), 400

    rec = db.get_record_by_id(record_id)
    if not rec:
        return jsonify(error="Record not found"), 404
    pid = rec.get("_project_id") or ""
    dt_id = rec.get("_dt_id") or ""

    u, err = _check_write_access(pid, dt_id)
    if err:
        return err

    existing = (rec.get("attachments") or {}).get(slot) if isinstance(rec.get("attachments"), dict) else None
    old_name = (existing or {}).get("file_name") or (existing or {}).get("url") or slot
    old_url = attachments._clean_link((existing or {}).get("url"))

    if existing:
        _trash_old_attachment(existing)

    db.clear_record_attachment(record_id, slot)
    updated_att = db.get_record_attachments(record_id)

    # Keep legacy fileLocation in sync if it was pointing to the removed attachment
    cur_loc = attachments._clean_link(rec.get("fileLocation"))
    next_loc = cur_loc
    if old_url and cur_loc == old_url:
        other_slot = attachments.SLOT_SUBMISSION if slot == attachments.SLOT_APPROVAL else attachments.SLOT_APPROVAL
        other_entry = (updated_att or {}).get(other_slot) if isinstance(updated_att, dict) else None
        next_loc = attachments._clean_link((other_entry or {}).get("url")) if isinstance(other_entry, dict) else ""
        db.q("""UPDATE records SET data = data || jsonb_build_object('fileLocation', %s::text), updated_at = NOW()
                WHERE id = %s RETURNING id""", (next_loc, record_id))

    try:
        db.log_action(
            u["username"], "EDIT",
            project_id=pid, dt_id=dt_id, record_id=record_id,
            doc_no=rec.get("docNo", ""),
            field_name=f"attachments.{slot}",
            old_value=str(old_name),
            new_value="",
            detail=f"Removed {slot} attachment '{old_name}' and moved to Google Drive Trash",
        )
    except Exception:
        pass

    return jsonify(
        ok=True,
        record_id=record_id,
        slot=slot,
        attachments=updated_att,
        fileLocation=next_loc,
    )
