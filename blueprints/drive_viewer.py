"""blueprints/drive_viewer.py - In-app document / folder viewer for DCR attachments.

Endpoints (all require a logged-in user who can view the record's project):
  GET /api/drive/view-info/<record_id>/<slot>               -- what is attached (file or folder listing) + permissions
  GET /api/drive/stream/<record_id>/<slot>/<file_id>        -- stream one file (inline preview, or ?dl=1 download)
  GET /api/drive/office-preview/<record_id>/<slot>/<file_id>-- structured JSON preview for Excel (.xlsx/.csv) & Word (.docx)
  GET /api/drive/folder-zip/<record_id>/<slot>              -- download a whole attached folder as one ZIP

<slot> is 'submission' | 'approval' | 'legacy' (the old single fileLocation link).
A file can only be streamed if it IS the record's attachment or lives inside the attached folder,
so users can never reach arbitrary Drive files through this gateway. Drive URLs are not exposed to
non-admin users; downloads require edit permission (or admin).
"""
import base64
import csv
import io
import logging
import re
import tempfile
import xml.etree.ElementTree as ET
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
MAX_OFFICE_PREVIEW_BYTES = 35 * 1024 * 1024
MAX_FOLDER_FILES = 500
FOLDER_MIME = "application/vnd.google-apps.folder"
FILE_FIELDS = "id,name,size,mimeType,parents,trashed"

_IMAGE_EXTS = {"png", "jpg", "jpeg", "gif", "webp", "bmp"}
_EXCEL_EXTS = {"xlsx", "xlsm", "csv"}
_WORD_EXTS = {"docx"}
_INLINE_MIME = {
    "application/pdf",
    "text/plain",
    "image/png",
    "image/jpeg",
    "image/gif",
    "image/webp",
    "image/bmp",
}


def _hooks():
    import blueprints.drive_uploads as du
    return du._get_token(), du._get_http()


def preview_kind(name, mime):
    ext = str(name or "").rsplit(".", 1)[-1].lower() if "." in str(name or "") else ""
    mime = str(mime or "").lower()
    if ext in _EXCEL_EXTS or mime in (
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "application/vnd.ms-excel.sheet.macroenabled.12",
        "text/csv",
    ):
        return "excel"
    if ext in _WORD_EXTS or mime == "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
        return "word"
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
        # Fallback 1: If approval slot requested, but record only has single submission slot, allow viewing
        if not entry and slot == "approval" and att.get("submission") and not att.get("approval"):
            entry = att.get("submission")
        # Fallback 2: Check any legacy/drive column on record
        if not entry and slot in ("submission", "approval"):
            fallback_url = rec.get("fileLocation") or rec.get("driveLink") or rec.get("drive_link") or rec.get("link")
            if fallback_url:
                entry = {"url": attachments._clean_link(fallback_url), "legacy": True}
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
        return None, (jsonify(error="Attachment has no Drive id",
                              fallback_url=url if (is_superadmin(u) or entry.get("legacy")) else ""), 404)

    can_dl = bool(can_edit(pid) or _is_workspace_admin(u, pid))
    ctx = dict(user=u, rec=rec, pid=pid, dt_id=dt_id, slot=slot, entry=entry, kind=kind, target_id=target_id,
               can_download=can_dl, can_open_drive=_is_workspace_admin(u, pid), url=url)
    return ctx, None


def _list_folder(token, http, folder_id, doc_no="", stored_files=None):
    svc = drive_upload.RestDriveAdapter(token, http=http)
    out, subfolders, page = [], [], None
    while len(out) < MAX_FOLDER_FILES:
        res = svc.files().list(q="'%s' in parents and trashed=false" % folder_id,
                               fields="nextPageToken, files(id,name,size,mimeType)", pageToken=page).execute()
        for f in res.get("files", []):
            if f.get("mimeType") == FOLDER_MIME:
                subfolders.append(f)
            else:
                out.append(f)
        page = res.get("nextPageToken")
        if not page:
            break

    # Also list files inside 1-level subfolders on Google Drive if any exist
    for sf in subfolders[:10]:
        sf_id = sf.get("id")
        sf_name = sf.get("name") or ""
        if not sf_id or len(out) >= MAX_FOLDER_FILES:
            continue
        try:
            sub_res = svc.files().list(q="'%s' in parents and trashed=false" % sf_id,
                                       fields="files(id,name,size,mimeType)").execute()
            for cf in sub_res.get("files", []):
                if cf.get("mimeType") != FOLDER_MIME and len(out) < MAX_FOLDER_FILES:
                    cf["subfolder"] = sf_name
                    out.append(cf)
        except Exception:
            pass

    # Merge any stored relative subfolder metadata from DCR attachment entry
    if isinstance(stored_files, list) and stored_files:
        by_id = {str(sf.get("id") or ""): sf.get("subfolder") or "" for sf in stored_files if isinstance(sf, dict)}
        by_name = {str(sf.get("name") or "").lower(): sf.get("subfolder") or "" for sf in stored_files if isinstance(sf, dict)}
        for f in out:
            if not f.get("subfolder"):
                sub = by_id.get(str(f.get("id") or "")) or by_name.get(str(f.get("name") or "").lower()) or ""
                if sub:
                    f["subfolder"] = sub

    doc_clean = str(doc_no or "").strip().lower()
    doc_base = re.sub(r"\s+rev\d+$", "", doc_clean, flags=re.IGNORECASE).strip()

    def _sort_key(f):
        fname = str(f.get("name") or "").lower()
        sub = str(f.get("subfolder") or "").lower()
        matches_doc = bool(doc_clean and (doc_clean in fname or (doc_base and doc_base in fname)))
        # 1) Files matching the record's Document No. appear FIRST at the very top of the list
        # 2) Root-level files appear before subfolder files
        # 3) Alphabetical by subfolder then file name
        return (0 if matches_doc else 1, 0 if not sub else 1, sub, fname)

    out.sort(key=_sort_key)
    return out[:MAX_FOLDER_FILES]


def _file_dict(f, rid, slot, can_dl):
    kind = preview_kind(f.get("name"), f.get("mimeType"))
    fid_q = quote(str(f["id"]), safe="")
    rid_q = quote(rid, safe="")
    base = "/api/drive/stream/%s/%s/%s" % (rid_q, slot, fid_q)
    office_url = ("/api/drive/office-preview/%s/%s/%s" % (rid_q, slot, fid_q)) if kind in ("excel", "word") else ""
    drive_embed_url = "https://drive.google.com/file/d/%s/preview" % fid_q
    return {
        "id": f["id"],
        "name": f.get("name") or "file",
        "size": int(f.get("size") or 0),
        "mime": f.get("mimeType") or "",
        "subfolder": f.get("subfolder") or "",
        "preview": kind,
        "stream_url": base if kind in ("pdf", "image", "text") else "",
        "office_url": office_url,
        "drive_embed_url": drive_embed_url,
        "download_url": (base + "?dl=1") if can_dl else "",
    }


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
    doc_no = ctx["rec"].get("docNo") or ctx["rec"].get("nocNo") or ""
    stored_files = ctx["entry"].get("files") if isinstance(ctx["entry"].get("files"), list) else None
    try:
        token, http = _hooks()
        if ctx["kind"] == "folder":
            files = _list_folder(token, http, ctx["target_id"], doc_no=doc_no, stored_files=stored_files)
            if not files and stored_files:
                files = [
                    {
                        "id": f.get("id") or f"ITEM_{idx + 1}",
                        "name": f.get("name") or f"File_{idx + 1}",
                        "size": int(f.get("size") or 0),
                        "mimeType": f.get("mimeType") or "",
                        "subfolder": f.get("subfolder") or "",
                    }
                    for idx, f in enumerate(stored_files)
                ]
            # Auto-sync discovered subfolders to database attachment entry if DB lacked subfolders
            if files and any(f.get("subfolder") for f in files) and stored_files is not None:
                if not any(sf.get("subfolder") for sf in stored_files if isinstance(sf, dict)):
                    try:
                        synced_entry = dict(ctx["entry"])
                        synced_entry["files"] = [
                            {
                                "id": f.get("id") or "",
                                "name": f.get("name") or "",
                                "size": int(f.get("size") or 0),
                                "mimeType": f.get("mimeType") or "",
                                "subfolder": f.get("subfolder") or "",
                            }
                            for f in files
                        ]
                        db.set_record_attachment(rid, ctx["slot"], synced_entry)
                    except Exception:
                        pass
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


@drive_viewer_bp.route("/api/drive/rescan-folder/<record_id>/<slot>", methods=["POST"])
def api_rescan_folder(record_id, slot):
    """Force re-query Google Drive for a folder attachment, discover subfolders and all files, and update DB."""
    ctx, err = _context(record_id, slot)
    if err:
        return err
    if ctx["kind"] != "folder":
        return jsonify(ok=False, error="Attachment is not a folder"), 400
    rid = ctx["rec"].get("_id") or record_id
    doc_no = ctx["rec"].get("docNo") or ctx["rec"].get("nocNo") or ""
    try:
        token, http = _hooks()
        files = _list_folder(token, http, ctx["target_id"], doc_no=doc_no, stored_files=None)
        entry = dict(ctx["entry"])
        entry["files"] = [
            {
                "id": f.get("id") or "",
                "name": f.get("name") or "",
                "size": int(f.get("size") or 0),
                "mimeType": f.get("mimeType") or "",
                "subfolder": f.get("subfolder") or "",
            }
            for f in files
        ]
        entry["file_count"] = len(files)
        entry["file_size"] = sum(int(f.get("size") or 0) for f in files)
        db.set_record_attachment(rid, ctx["slot"], entry)
        _log(ctx, "DRIVE_FOLDER_SYNC", f"Re-scanned Drive folder '{entry.get('folder_name') or rid}' ({len(files)} files found)")
        return jsonify(ok=True, files=entry["files"], file_count=entry["file_count"], file_size=entry["file_size"])
    except Exception as e:
        return jsonify(ok=False, error=str(e)), 500


def _authorize_file(ctx, token, http, file_id):
    """Return Drive metadata if file_id is the attachment itself or sits inside the attached folder (or its 1-level subfolder)."""
    meta = drive_upload.get_file(token, file_id, fields=FILE_FIELDS, http=http)
    if meta.get("trashed"):
        return None
    if ctx["kind"] == "file":
        return meta if file_id == ctx["target_id"] else None
    parents = meta.get("parents") or []
    if ctx["target_id"] in parents:
        return meta
    # Check if file_id was explicitly recorded in ctx["entry"]["files"]
    stored = ctx["entry"].get("files") if isinstance(ctx["entry"].get("files"), list) else []
    if any(str(f.get("id") or "") == str(file_id) for f in stored if isinstance(f, dict)):
        return meta
    # Check 1-level subfolder parent on Drive
    for pid in parents[:2]:
        try:
            pmeta = drive_upload.get_file(token, pid, fields="id,parents,trashed", http=http)
            if not pmeta.get("trashed") and ctx["target_id"] in (pmeta.get("parents") or []):
                return meta
        except Exception:
            pass
    return None


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


def _read_media_bytes(resp, max_bytes=MAX_OFFICE_PREVIEW_BYTES):
    if hasattr(resp, "iter_content"):
        buf = bytearray()
        for chunk in resp.iter_content(65536):
            if chunk:
                buf.extend(chunk)
                if len(buf) > max_bytes:
                    raise ValueError("File exceeds maximum in-app preview size (35 MB)")
        return bytes(buf)
    data = resp.content or b""
    if len(data) > max_bytes:
        raise ValueError("File exceeds maximum in-app preview size (35 MB)")
    return data


def _col_letters_to_idx(col_str):
    idx = 0
    for ch in col_str:
        if "A" <= ch <= "Z":
            idx = idx * 26 + (ord(ch) - ord("A") + 1)
    return max(0, idx - 1)


def _parse_xlsx_bytes(raw_bytes, max_rows=1500, max_cols=80):
    """Parse .xlsx bytes into [{'name': sheet_name, 'rows': [[cell, ...], ...]}] using openpyxl or stdlib XML."""
    try:
        import openpyxl
        wb = openpyxl.load_workbook(io.BytesIO(raw_bytes), data_only=True, read_only=True)
        sheets = []
        for sname in wb.sheetnames[:15]:
            ws = wb[sname]
            rows = []
            for r_idx, row in enumerate(ws.iter_rows(values_only=True)):
                if r_idx >= max_rows:
                    break
                vals = ["" if v is None else str(v) for v in row[:max_cols]]
                while vals and vals[-1] == "":
                    vals.pop()
                rows.append(vals)
            while rows and not any(rows[-1]):
                rows.pop()
            max_w = max((len(r) for r in rows), default=0)
            norm_rows = [r + [""] * (max_w - len(r)) for r in rows]
            sheets.append({"name": sname, "rows": norm_rows})
        wb.close()
        if sheets:
            return sheets
    except Exception:
        pass

    # Pure-stdlib XML fallback for .xlsx
    sheets = []
    with zipfile.ZipFile(io.BytesIO(raw_bytes), "r") as zf:
        shared_strings = []
        if "xl/sharedStrings.xml" in zf.namelist():
            ss_root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
            for si in ss_root.iter():
                if si.tag.endswith("}si") or si.tag == "si":
                    texts = [t.text or "" for t in si.iter() if t.tag.endswith("}t") or t.tag == "t"]
                    shared_strings.append("".join(texts))

        sheet_Specs = []
        if "xl/workbook.xml" in zf.namelist():
            wb_root = ET.fromstring(zf.read("xl/workbook.xml"))
            for idx, el in enumerate(wb_root.iter()):
                if el.tag.endswith("}sheet") or el.tag == "sheet":
                    sname = el.attrib.get("name") or f"Sheet{idx + 1}"
                    sheet_Specs.append((sname, f"xl/worksheets/sheet{len(sheet_Specs) + 1}.xml"))
        if not sheet_Specs:
            for name in sorted(zf.namelist()):
                if name.startswith("xl/worksheets/sheet") and name.endswith(".xml"):
                    sheet_Specs.append((f"Sheet{len(sheet_Specs) + 1}", name))

        for sname, spath in sheet_Specs[:15]:
            if spath not in zf.namelist():
                continue
            ws_root = ET.fromstring(zf.read(spath))
            rows = []
            for row_el in ws_root.iter():
                if not (row_el.tag.endswith("}row") or row_el.tag == "row"):
                    continue
                if len(rows) >= max_rows:
                    break
                row_vals = []
                for c_el in row_el:
                    if not (c_el.tag.endswith("}c") or c_el.tag == "c"):
                        continue
                    ref = c_el.attrib.get("r", "")
                    col_letters = "".join(ch for ch in ref if ch.isalpha()).upper()
                    c_idx = _col_letters_to_idx(col_letters) if col_letters else len(row_vals)
                    if c_idx >= max_cols:
                        continue
                    while len(row_vals) <= c_idx:
                        row_vals.append("")
                    ctype = c_el.attrib.get("t", "")
                    val_str = ""
                    if ctype == "inlineStr":
                        texts = [t.text or "" for t in c_el.iter() if t.tag.endswith("}t") or t.tag == "t"]
                        val_str = "".join(texts)
                    else:
                        v_el = next((ch for ch in c_el if ch.tag.endswith("}v") or ch.tag == "v"), None)
                        raw_v = v_el.text if (v_el is not None and v_el.text is not None) else ""
                        if ctype == "s":
                            try:
                                val_str = shared_strings[int(raw_v)]
                            except Exception:
                                val_str = raw_v
                        else:
                            val_str = raw_v
                    row_vals[c_idx] = val_str
                while row_vals and row_vals[-1] == "":
                    row_vals.pop()
                rows.append(row_vals)
            while rows and not any(rows[-1]):
                rows.pop()
            max_w = max((len(r) for r in rows), default=0)
            norm_rows = [r + [""] * (max_w - len(r)) for r in rows]
            sheets.append({"name": sname, "rows": norm_rows})
    return sheets


def _parse_csv_bytes(raw_bytes, max_rows=1500, max_cols=80):
    text = raw_bytes.decode("utf-8-sig", errors="replace")
    reader = csv.reader(io.StringIO(text))
    rows = []
    for idx, r in enumerate(reader):
        if idx >= max_rows:
            break
        rows.append([str(c) for c in r[:max_cols]])
    max_w = max((len(r) for r in rows), default=0)
    norm_rows = [r + [""] * (max_w - len(r)) for r in rows]
    return [{"name": "CSV Data", "rows": norm_rows}]


def _parse_docx_bytes(raw_bytes):
    """Parse .docx bytes into structured blocks [{'type': 'h1'|'h2'|'h3'|'p'|'li'|'table'|'image', ...}] for safe rendering."""
    blocks = []
    with zipfile.ZipFile(io.BytesIO(raw_bytes), "r") as zf:
        if "word/document.xml" not in zf.namelist():
            raise ValueError("Invalid .docx archive (missing word/document.xml)")
        root = ET.fromstring(zf.read("word/document.xml"))

        # Pre-load relationship map for media images
        rels_map = {}
        if "word/_rels/document.xml.rels" in zf.namelist():
            try:
                rels_root = ET.fromstring(zf.read("word/_rels/document.xml.rels"))
                for rel in rels_root:
                    r_id = rel.attrib.get("Id")
                    target = rel.attrib.get("Target")
                    if r_id and target:
                        if not target.startswith("word/"):
                            target = "word/" + target.lstrip("/")
                        rels_map[r_id] = target
            except Exception:
                pass

        img_cache = {}
        for r_id, img_path in rels_map.items():
            if img_path in zf.namelist():
                try:
                    idata = zf.read(img_path)
                    if len(idata) <= 5 * 1024 * 1024:
                        ext = img_path.rsplit(".", 1)[-1].lower()
                        mime = "image/png" if ext == "png" else ("image/jpeg" if ext in ("jpg", "jpeg") else ("image/gif" if ext == "gif" else "image/png"))
                        b64 = base64.b64encode(idata).decode("ascii")
                        img_cache[r_id] = f"data:{mime};base64,{b64}"
                except Exception:
                    pass

    def _local(tag):
        return tag.rsplit("}", 1)[-1] if "}" in tag else tag

    def _para_to_block(p_el):
        style_val = ""
        is_list = False
        align = ""
        for ch in p_el:
            if _local(ch.tag) == "pPr":
                for pr in ch:
                    ln = _local(pr.tag)
                    if ln == "pStyle":
                        for k, v in pr.attrib.items():
                            if _local(k) == "val":
                                style_val = str(v or "")
                    elif ln == "numPr":
                        is_list = True
                    elif ln == "jc":
                        for k, v in pr.attrib.items():
                            if _local(k) == "val":
                                align = str(v or "").lower()

        images_in_p = []
        for el in p_el.iter():
            if _local(el.tag) == "blip":
                for k, v in el.attrib.items():
                    if _local(k) == "embed" and v in img_cache:
                        images_in_p.append(img_cache[v])

        runs = []
        for r_el in p_el.iter():
            if _local(r_el.tag) == "r":
                bold = False
                italic = False
                underline = False
                for r_ch in r_el:
                    if _local(r_ch.tag) == "rPr":
                        for pr in r_ch:
                            pln = _local(pr.tag)
                            if pln == "b":
                                bold = True
                            elif pln == "i":
                                italic = True
                            elif pln == "u":
                                underline = True
                t_parts = []
                for r_ch in r_el:
                    rl = _local(r_ch.tag)
                    if rl == "t" and r_ch.text:
                        t_parts.append(r_ch.text)
                    elif rl == "tab":
                        t_parts.append("\t")
                    elif rl in ("br", "cr"):
                        t_parts.append("\n")
                txt = "".join(t_parts)
                if txt:
                    runs.append({"text": txt, "b": bold, "i": italic, "u": underline})
        full_text = "".join(r["text"] for r in runs).strip()
        if not full_text:
            if images_in_p:
                return {"type": "image", "src": images_in_p[0], "align": align or "center"}
            return None
        st_low = style_val.lower()
        btype = "p"
        if "title" in st_low or "heading1" in st_low or st_low == "h1":
            btype = "h1"
        elif "heading2" in st_low or "subtitle" in st_low or st_low == "h2":
            btype = "h2"
        elif "heading" in st_low or st_low == "h3":
            btype = "h3"
        elif is_list or "list" in st_low:
            btype = "li"
        return {"type": btype, "text": full_text, "runs": runs, "align": align, "images": images_in_p}

    body_el = next((ch for ch in root if _local(ch.tag) == "body"), root)
    for child in body_el:
        cname = _local(child.tag)
        if cname == "p":
            blk = _para_to_block(child)
            if blk:
                if blk.get("images"):
                    for img in blk["images"]:
                        blocks.append({"type": "image", "src": img, "align": blk.get("align") or "center"})
                blocks.append(blk)
        elif cname == "tbl":
            t_rows = []
            for tr in child:
                if _local(tr.tag) != "tr":
                    continue
                row_cells = []
                for tc in tr:
                    if _local(tc.tag) != "tc":
                        continue
                    cell_texts = []
                    for p in tc.iter():
                        if _local(p.tag) == "p":
                            pb = _para_to_block(p)
                            if pb and pb.get("text"):
                                cell_texts.append(pb["text"])
                    row_cells.append("\n".join(cell_texts))
                if row_cells:
                    t_rows.append(row_cells)
            if t_rows:
                blocks.append({"type": "table", "rows": t_rows})
    return blocks


@drive_viewer_bp.route("/api/drive/office-preview/<record_id>/<slot>/<file_id>")
def api_office_preview(record_id, slot, file_id):
    ctx, err = _context(record_id, slot)
    if err:
        return err
    try:
        token, http = _hooks()
        meta = _authorize_file(ctx, token, http, str(file_id))
        if not meta:
            return jsonify(ok=False, error="File does not belong to this attachment"), 403
        name = meta.get("name") or "file"
        kind = preview_kind(name, meta.get("mimeType"))
        if kind not in ("excel", "word"):
            return jsonify(ok=False, error="Not a supported Office file for inline preview"), 415
        resp = _drive_media(token, http, meta)
    except drive_upload.DriveUploadError as e:
        return jsonify(ok=False, error=str(e), hint=e.hint), 502
    if resp.status_code != 200:
        return jsonify(ok=False, error="Drive returned HTTP %s" % resp.status_code), 502

    try:
        raw_bytes = _read_media_bytes(resp)
        if kind == "excel":
            ext = name.rsplit(".", 1)[-1].lower() if "." in name else ""
            sheets = _parse_csv_bytes(raw_bytes) if ext == "csv" else _parse_xlsx_bytes(raw_bytes)
            _log(ctx, "DRIVE_LINK_CLICK", "Previewed Excel '%s' via in-app gateway" % name)
            return jsonify(ok=True, kind="excel", name=name, sheets=sheets)
        else:
            blocks = _parse_docx_bytes(raw_bytes)
            _log(ctx, "DRIVE_LINK_CLICK", "Previewed Word '%s' via in-app gateway" % name)
            return jsonify(ok=True, kind="word", name=name, blocks=blocks)
    except Exception as e:
        logger.warning("Office preview parse failed for %s: %s", file_id, e)
        return jsonify(ok=False, error="Could not parse Office file for preview: %s" % str(e)), 422


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
        doc_no = ctx["rec"].get("docNo") or ctx["rec"].get("nocNo") or ""
        stored_files = ctx["entry"].get("files") if isinstance(ctx["entry"].get("files"), list) else None
        files = _list_folder(token, http, ctx["target_id"], doc_no=doc_no, stored_files=stored_files)
        if sum(int(f.get("size") or 0) for f in files) > MAX_ZIP_BYTES:
            return jsonify(error="Folder is larger than 1 GB - download files one by one"), 413
        tmp = tempfile.TemporaryFile()
        used = set()
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_STORED, allowZip64=True) as zf:
            for f in files:
                resp = _drive_media(token, http, f)
                if resp.status_code != 200:
                    continue
                sub = str(f.get("subfolder") or "").strip("/\\ ")
                fname = f.get("name") or f["id"]
                arc = f"{sub}/{fname}" if sub else fname
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
