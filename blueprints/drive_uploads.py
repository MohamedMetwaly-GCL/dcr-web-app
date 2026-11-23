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
from auth import can_edit, current_user, is_superadmin, can_access_doc_type
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


IGNORED_JUNK_NAMES = {
    "desktop.ini",
    "thumbs.db",
    ".ds_store",
    "ehthumbs.db",
    "ehthumbs_vista.db",
    "icon\r",
}
IGNORED_JUNK_PREFIXES = ("~$", ".smbdelete")
IGNORED_JUNK_EXTS = {".tmp", ".bak", ".lnk"}


def is_ignored_junk_filename(raw_name):
    """Return True when filename is OS metadata or temporary Office/lock junk."""
    base = str(raw_name or "").replace("\\", "/").split("/")[-1].strip()
    low = base.lower()
    if not low:
        return False
    if low in IGNORED_JUNK_NAMES:
        return True
    if any(low.startswith(p) for p in IGNORED_JUNK_PREFIXES):
        return True
    _, ext = os.path.splitext(low)
    return ext in IGNORED_JUNK_EXTS


def sanitize_upload_filename(raw_name):
    """Return (clean_name, error_reason). Rejects empty names, OS/Office junk, and dangerous extensions."""
    raw_base = str(raw_name or "").replace("\\", "/").split("/")[-1].strip()
    if is_ignored_junk_filename(raw_base):
        return "", "ignored_system_file"
    name = re.sub(r"[\x00-\x1f\x7f]+", "", raw_base).strip(". ")
    if not name or len(name) > 240:
        return "", "invalid_file_name"
    if is_ignored_junk_filename(name):
        return "", "ignored_system_file"
    _, ext = os.path.splitext(name.lower())
    if ext in DANGEROUS_EXTS_INSIDE_ARCHIVE:
        return "", "forbidden_file_extension"
    return name, ""


def _collect_protected_ids(entry):
    """Collect Drive file/folder IDs belonging to an attachment entry or its archived history."""
    ids = set()
    if not isinstance(entry, dict):
        return ids
    for key in ("file_id", "folder_id"):
        val = str(entry.get(key) or "").strip()
        if val:
            ids.add(val)
    for f in (entry.get("files") or []):
        if isinstance(f, dict):
            fid = str(f.get("id") or f.get("file_id") or "").strip()
            if fid:
                ids.add(fid)
    for h in (entry.get("history") or []):
        ids.update(_collect_protected_ids(h))
    return ids


def _trash_old_attachment(entry, exclude_id="", files_only=False, exclude_ids=None, trash_child_folders_under=None, exclude_folder_ids=None):
    """Best-effort move of an old attachment's Drive file(s) or folder to Google Drive Trash."""
    if not isinstance(entry, dict) or entry.get("legacy"):
        return
    protected = set(exclude_ids or ())
    if exclude_id:
        protected.add(str(exclude_id).strip())
    # Never trash Drive objects still referenced by archived revision history unless caller explicitly strips history
    for h in (entry.get("history") or []):
        protected.update(_collect_protected_ids(h))
    try:
        http = _get_http()
        # Only call real Drive PATCH when http supports .patch (avoids consuming scripted FakeHTTP queues in unit tests)
        if getattr(http, "patch", None) is None:
            return
        token = _get_token()
        if files_only and isinstance(entry.get("files"), list):
            has_nested = any(isinstance(f, dict) and f.get("subfolder") for f in entry["files"])
            for f in entry["files"]:
                if not isinstance(f, dict):
                    continue
                fid = str(f.get("id") or f.get("file_id") or "").strip()
                if fid and fid not in protected:
                    try:
                        drive_upload.trash_file(token, fid, http=http)
                    except Exception:
                        pass
            if trash_child_folders_under and has_nested:
                keep_folders = set(str(x).strip() for x in (exclude_folder_ids or ()) if x)
                keep_folders.update(protected)
                try:
                    service = drive_upload.RestDriveAdapter(token, http=http)
                    for child in drive_routing.list_child_folders(service, trash_child_folders_under):
                        cid = str(child.get("id") or "").strip()
                        if cid and cid not in keep_folders:
                            try:
                                drive_upload.trash_file(token, cid, http=http)
                            except Exception:
                                pass
                    drive_routing.invalidate_subfolder_caches(trash_child_folders_under)
                except Exception:
                    pass
            return
        target_id = str(entry.get("folder_id") or entry.get("file_id") or "").strip()
        if target_id and target_id not in protected:
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
    if not can_access_doc_type(pid, dt_id, u):
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
    rel_subfolder = str(body.get("rel_subfolder") or "").strip()
    link_as_folder = bool(body.get("link_as_folder", False))
    mime_type = str(body.get("mime_type") or "application/octet-stream").strip()
    create_folders = bool(body.get("create_folders", True))
    rev_mode = str(body.get("rev_mode") or "overwrite").strip().lower()
    if rev_mode not in ("overwrite", "new_revision"):
        rev_mode = "overwrite"
    target_rev = body.get("part_rev")

    u, err = _check_write_access(pid, dt_id)
    if err:
        return err

    if slot not in attachments.SLOTS:
        return jsonify(error=f"Invalid slot '{slot}'. Must be one of {attachments.SLOTS}."), 400

    proj = db.get_project(pid)
    if not proj:
        return jsonify(error="Project not found"), 404
    ws_id = str(proj.get("workspace_id") or "ENG").strip().upper()
    if slot in attachments.NOC_SLOTS and ws_id != "ENG":
        return jsonify(error="NOC slots are only available in the Engineering & Projects workspace"), 400
    dt_meta = db.get_doc_type(pid, dt_id) or {}
    if ws_id == "HSE":
        has_cycle = attachments.doc_type_allows_approval_slot(dt_meta, ws_id)
        if slot == attachments.SLOT_APPROVAL and not has_cycle:
            return jsonify(error="Close-Out attachment slot is only available for PTW, OBS, SVP, INC, and JSA in HSE"), 400
    else:
        if slot == attachments.SLOT_APPROVAL and ws_id != "ENG":
            return jsonify(error="Approval / NOC slots are only available in the Engineering & Projects workspace"), 400
        has_cycle = attachments.doc_type_has_reply_cycle(dt_meta)
        if (slot == attachments.SLOT_APPROVAL or slot in attachments.NOC_SLOTS) and not has_cycle:
            return jsonify(error="This document type has no consultant reply cycle"), 400

    rec = None
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
        direction = str(body.get("direction") or (rec.get("direction") if rec else "") or "").strip()
        from_party = str(body.get("from_party") or ((rec.get("fromParty") or rec.get("from")) if rec else "") or "").strip()
        to_party = str(body.get("to_party") or ((rec.get("toParty") or rec.get("to")) if rec else "") or "").strip()
        party = str(body.get("party") or "").strip()
        doc_no = str(body.get("doc_no") or (rec.get("docNo") if rec else "") or clean_name or "").strip()
        raw_subtype = body.get("subtype") or (rec.get("legacySubtype") if rec else "")
        from hse_defaults import normalize_ptw_subtype_value
        subtype = normalize_ptw_subtype_value(raw_subtype)
        dest_folder_id = str(body.get("destination_folder_id") or "").strip()
        dest = drive_routing.resolve_upload_destination(
            pid, dt_id, slot, service, db, create_if_missing=create_folders,
            subfolder_name=subfolder_name or None, flat=not has_cycle,
            direction=direction, destination_folder_id=dest_folder_id,
            party=party, from_party=from_party, to_party=to_party, doc_no=doc_no,
            subtype=subtype)
        root_package_folder_id = dest.get("subfolder_id") or dest["folder_id"]
        rel_parts = drive_routing.normalize_rel_subfolder(rel_subfolder)
        leaf_folder_id = dest["folder_id"]
        ancestor_folder_ids = []
        if rel_parts:
            leaf_folder_id, ancestor_folder_ids = drive_routing.resolve_nested_subfolder_with_ancestors(
                service, root_package_folder_id, rel_subfolder, create_if_missing=create_folders
            )
        session_url = drive_upload.create_resumable_session(
            token, leaf_folder_id, clean_name, mime_type, file_size,
            origin=request.headers.get("Origin"), http=http)
    except drive_routing.RoutingError as e:
        msg = e.reason
        if e.reason == "project_has_no_drive_folder":
            msg = "لم يتم تحديد معرّف مجلد Google Drive لهذا المشروع (يرجى إدخال Drive Root Folder ID في إعدادات المشروع أولاً)"
        return jsonify(error=msg, reason=e.reason, candidates=e.candidates), 400
    except drive_upload.DriveUploadError as e:
        return jsonify(error=str(e), reason=e.reason, hint=e.hint), (e.status if e.status in (400, 401, 403, 404) else 502)
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify(error=f"Upload initialization error: {str(e)}", hint=str(e)), 500

    clean_rel_sub = "/".join(rel_parts)
    ticket_payload = {
        "pid": pid,
        "dt_id": dt_id,
        "ws_id": ws_id,
        "record_id": record_id,
        "slot": slot,
        "folder_id": leaf_folder_id,
        "subfolder_id": dest.get("subfolder_id", ""),
        "subfolder_name": dest.get("subfolder_name", ""),
        "rel_subfolder": clean_rel_sub,
        "ancestor_folder_ids": ancestor_folder_ids,
        "link_as_folder": link_as_folder,
        "rev_mode": rev_mode,
        "file_name": clean_name,
        "file_size": file_size,
        "mime_type": mime_type,
        "user": u["username"],
        "session_url": session_url,
    }
    if target_rev is not None:
        ticket_payload["target_rev"] = attachments.parse_rev_number(target_rev, 0)
    ticket = sign_upload_ticket(ticket_payload)
    return jsonify(
        ok=True,
        session_url=session_url,
        upload_ticket=ticket,
        folder_id=leaf_folder_id,
        package_folder_id=root_package_folder_id,
        doc_type_folder_id=dest["doc_type_folder_id"],
        folder_source=dest["source"],
        rel_subfolder=clean_rel_sub,
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

    if not rel_subfolder:
        rel_subfolder = str(t.get("rel_subfolder") or "").strip()
    rev_mode = str(body.get("rev_mode") or t.get("rev_mode") or "overwrite").strip().lower()
    if rev_mode not in ("overwrite", "new_revision"):
        rev_mode = "overwrite"
    req_part_rev = body.get("part_rev") if body.get("part_rev") is not None else t.get("target_rev")

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

    proj = db.get_project(pid) or {}
    ws_id = str(t.get("ws_id") or proj.get("workspace_id") or "ENG").strip().upper()
    is_hse = (ws_id == "HSE")

    view_url = remote.get("webViewLink") or attachments.drive_view_url(file_id)
    extra = {"url": view_url}
    new_record_status = None
    old_record_status = str(rec.get("status") or "").strip()
    if slot == attachments.SLOT_APPROVAL:
        if is_hse:
            # HSE close-out attachments never carry ENG consultant review codes (Code A/B/C/D).
            closeout_status, new_record_status = attachments.resolve_hse_closeout_status(
                old_record_status, body.get("closeout_status")
            )
            extra["closeout_status"] = closeout_status
        elif rec.get("status"):
            extra["consultant_status"] = str(rec.get("status"))
    elif slot == attachments.SLOT_PART_B:
        pb_st = rec.get("partBStatus") or rec.get("part_b_status") or rec.get("Part B Status") or rec.get("status")
        if pb_st:
            extra["consultant_status"] = str(pb_st)
    elif slot == attachments.SLOT_PART_D:
        pd_st = rec.get("partDStatus") or rec.get("part_d_status") or rec.get("Part D Status")
        if pd_st:
            extra["consultant_status"] = str(pd_st)

    # Folder / Multi-file attachment handling
    link_as_folder = bool(body.get("link_as_folder")) if body.get("link_as_folder") is not None else bool(t.get("link_as_folder"))
    if link_as_folder and (t.get("subfolder_id") or t.get("folder_id")):
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
    same_active_folder = bool(
        extra.get("is_folder")
        and not replace_existing
        and existing_slot
        and existing_slot.get("is_folder")
        and existing_slot.get("folder_id") == extra.get("folder_id")
    )
    starting_new_rev = bool(
        rev_mode == "new_revision"
        and existing_slot
        and not existing_slot.get("legacy")
        and not same_active_folder
    )

    if starting_new_rev:
        # Archive current revision without trashing previous Drive objects
        meta = attachments.archive_and_promote_revision(
            existing_slot,
            meta,
            rev=req_part_rev,
            archived_by=u["username"],
        )
        if extra.get("is_folder"):
            f_item = {
                "id": file_id,
                "name": remote.get("name") or t["file_name"],
                "size": int(remote.get("size") or t["file_size"] or 0),
                "mimeType": meta["mime_type"],
            }
            if rel_subfolder:
                f_item["subfolder"] = rel_subfolder
            meta["files"] = [f_item]
            meta["file_count"] = 1
            meta["file_size"] = int(f_item["size"] or 0)
            meta["file_name"] = extra.get("folder_name") or meta["file_name"]
    else:
        if extra.get("is_folder"):
            if replace_existing and existing_slot:
                # First file of a replacement folder upload: trash previous active files & obsolete subfolders
                _trash_old_attachment(
                    existing_slot,
                    exclude_id=file_id,
                    files_only=True,
                    trash_child_folders_under=extra.get("folder_id"),
                    exclude_folder_ids=t.get("ancestor_folder_ids") or (),
                )
                files_list = []
            elif same_active_folder:
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
            if same_active_folder:
                meta["file_name"] = (existing_slot or {}).get("folder_name") or extra.get("folder_name") or meta["file_name"]
            else:
                meta["file_name"] = extra.get("folder_name") or (existing_slot or {}).get("folder_name") or meta["file_name"]
        else:
            # Single-file upload replacing an existing attachment in the same slot (active rev only)
            if existing_slot:
                _trash_old_attachment(existing_slot, exclude_id=file_id)

        # Preserve active revision number and archived history across folder appends / active overwrites
        cur_rev = (
            attachments.parse_rev_number(req_part_rev, 0)
            if (req_part_rev is not None and not same_active_folder)
            else attachments.get_attachment_rev(existing_slot)
        )
        meta["rev"] = cur_rev
        meta["rev_label"] = attachments.format_rev_label(cur_rev)
        if isinstance(existing_slot, dict) and isinstance(existing_slot.get("history"), list) and existing_slot.get("history"):
            meta["history"] = existing_slot["history"]

    # Preserve legacy single fileLocation as submission (or part_a for NOC) if attaching a later slot to an old record
    old_loc = attachments._clean_link(rec.get("fileLocation"))
    if slot == attachments.SLOT_APPROVAL and not existing_att.get(attachments.SLOT_SUBMISSION) and old_loc and old_loc != view_url:
        db.set_record_attachment(record_id, attachments.SLOT_SUBMISSION, {"url": old_loc, "legacy": True})
    elif slot in (attachments.SLOT_PART_B, attachments.SLOT_PART_C, attachments.SLOT_PART_D):
        has_any_noc = any(existing_att.get(ns) for ns in attachments.NOC_SLOTS)
        if not has_any_noc and old_loc and old_loc != view_url:
            db.set_record_attachment(record_id, attachments.SLOT_PART_A, {"url": old_loc, "legacy": True})

    db.set_record_attachment(record_id, slot, meta)

    # If the record had no fileLocation yet (or if replacing the slot that owned fileLocation), keep fileLocation synced
    prev_slot_url = attachments._clean_link((existing_slot or {}).get("url"))
    if not old_loc or (prev_slot_url and old_loc == prev_slot_url):
        db.q("""UPDATE records SET data = data || jsonb_build_object('fileLocation', %s::text), updated_at = NOW()
                WHERE id = %s RETURNING id""", (view_url, record_id))

    if new_record_status and new_record_status != old_record_status:
        db.q("""UPDATE records SET data = data || jsonb_build_object('status', %s::text), updated_at = NOW()
                WHERE id = %s RETURNING id""", (new_record_status, record_id))
        db.invalidate_dashboard_cache(pid)
        try:
            db.log_action(
                u["username"], "EDIT",
                project_id=pid, dt_id=dt_id, record_id=record_id,
                doc_no=rec.get("docNo", ""),
                field_name="status",
                old_value=old_record_status,
                new_value=new_record_status,
                detail=f"HSE close-out attachment updated status from '{old_record_status}' to '{new_record_status}'",
            )
        except Exception:
            pass

    try:
        rev_tag = f" [{meta.get('rev_label')}]" if slot in attachments.NOC_SLOTS or meta.get("rev", 0) > 0 else ""
        db.log_action(
            u["username"], "FILE_UPLOAD",
            project_id=pid, dt_id=dt_id, record_id=record_id,
            doc_no=rec.get("docNo", ""),
            field_name=f"attachments.{slot}",
            new_value=meta["file_name"],
            detail=f"Uploaded {slot}{rev_tag} file '{meta['file_name']}' ({meta['file_size']} bytes) to Drive [{file_id}]",
        )
    except Exception:
        pass

    res_payload = dict(
        ok=True,
        record_id=record_id,
        slot=slot,
        attachment=meta,
        attachments=db.get_record_attachments(record_id),
    )
    if new_record_status:
        res_payload["status"] = new_record_status
    return jsonify(res_payload)


@drive_uploads_bp.route("/api/drive/attachment/<record_id>/<slot>", methods=["DELETE"])
def api_unlink_attachment(record_id, slot):
    """Unlink an attachment slot (or specific ?rev=<n>) from a record and move its Drive object to Trash."""
    record_id = str(record_id or "").strip()
    slot = str(slot or "").strip().lower()
    if slot not in attachments.SLOTS:
        return jsonify(error=f"Invalid slot '{slot}'. Must be one of {attachments.SLOTS}."), 400

    rec = db.get_record_by_id(record_id)
    if not rec:
        return jsonify(error="Record not found"), 404
    pid = rec.get("_project_id") or ""
    dt_id = rec.get("_dt_id") or ""

    u, err = _check_write_access(pid, dt_id)
    if err:
        return err

    existing = (rec.get("attachments") or {}).get(slot) if isinstance(rec.get("attachments"), dict) else None
    req_rev_raw = request.args.get("rev")
    delete_all = request.args.get("all") == "1"
    active_rev = attachments.get_attachment_rev(existing) if existing else 0
    old_name = (existing or {}).get("file_name") or (existing or {}).get("url") or slot
    old_url = attachments._clean_link((existing or {}).get("url"))

    if req_rev_raw is not None and str(req_rev_raw).strip() != "" and isinstance(existing, dict):
        req_rev = attachments.parse_rev_number(req_rev_raw, -1)
        if req_rev != active_rev:
            raw_hist = list(existing.get("history") or []) if isinstance(existing.get("history"), list) else []
            target_snap = next(
                (h for h in raw_hist if isinstance(h, dict) and attachments.get_attachment_rev(h) == req_rev),
                None,
            )
            if not target_snap:
                return jsonify(error=f"Revision {attachments.format_rev_label(req_rev)} not found"), 404
            active_protected = _collect_protected_ids(attachments.strip_entry_history(existing))
            _trash_old_attachment(target_snap, exclude_ids=active_protected)
            new_hist = [h for h in raw_hist if not (isinstance(h, dict) and attachments.get_attachment_rev(h) == req_rev)]
            updated_slot = dict(existing)
            if new_hist:
                updated_slot["history"] = new_hist
            else:
                updated_slot.pop("history", None)
            db.set_record_attachment(record_id, slot, updated_slot)
            updated_att = db.get_record_attachments(record_id)
            try:
                db.log_action(
                    u["username"], "EDIT",
                    project_id=pid, dt_id=dt_id, record_id=record_id,
                    doc_no=rec.get("docNo", ""),
                    field_name=f"attachments.{slot}",
                    old_value=f"{target_snap.get('rev_label') or ''} {target_snap.get('file_name') or ''}".strip(),
                    new_value="",
                    detail=f"Removed archived {slot} revision REV{req_rev:02d} and moved to Google Drive Trash",
                )
            except Exception:
                pass
            return jsonify(ok=True, record_id=record_id, slot=slot, attachments=updated_att, fileLocation=rec.get("fileLocation", ""))

    cur_loc = attachments._clean_link(rec.get("fileLocation"))
    next_loc = cur_loc
    if existing and not delete_all and isinstance(existing.get("history"), list) and existing.get("history"):
        promoted, removed = attachments.pop_active_revision(existing)
        if removed:
            _trash_old_attachment(removed, exclude_ids=_collect_protected_ids(promoted))
            old_name = removed.get("file_name") or removed.get("url") or old_name
        db.set_record_attachment(record_id, slot, promoted)
        updated_att = db.get_record_attachments(record_id)
        if old_url and cur_loc == old_url:
            next_loc = attachments._clean_link((promoted or {}).get("url"))
            db.q("""UPDATE records SET data = data || jsonb_build_object('fileLocation', %s::text), updated_at = NOW()
                    WHERE id = %s RETURNING id""", (next_loc, record_id))
    else:
        if existing:
            # If deleting everything (?all=1) and history exists, strip history so _trash_old_attachment removes active and optionally archived items
            if delete_all and isinstance(existing.get("history"), list):
                for h in existing.get("history"):
                    if isinstance(h, dict):
                        _trash_old_attachment(h)
                stripped = attachments.strip_entry_history(existing)
                _trash_old_attachment(stripped)
            else:
                _trash_old_attachment(existing)

        db.clear_record_attachment(record_id, slot)
        updated_att = db.get_record_attachments(record_id)

        if old_url and cur_loc == old_url:
            next_loc = ""
            if isinstance(updated_att, dict):
                for cand_slot in ("part_d", "part_c", "part_b", "part_a", "approval", "submission"):
                    if cand_slot == slot:
                        continue
                    cand_entry = updated_att.get(cand_slot)
                    if isinstance(cand_entry, dict):
                        cand_url = attachments._clean_link(cand_entry.get("url"))
                        if cand_url:
                            next_loc = cand_url
                            break
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
