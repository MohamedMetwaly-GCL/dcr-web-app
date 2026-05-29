"""blueprints/exporting.py - DCR Export/Import API routes.

Handles export/import endpoints:
  GET  /api/export_all/<pid>
  GET  /api/export/<pid>/<dt_id>
  helper: _build_pdf_for_dt(pid, dt_id, proj, buf=None)
  GET  /api/export_pdf/<pid>/<dt_id>
  GET  /api/export_pdf_all/<pid>
  POST /api/import/<pid>/<dt_id>

Step 13 of the incremental refactor.
Logic is identical to the original app.py routes/helper — only the decorator
changed from @app.route to @exporting_bp.route.
"""
import base64
import datetime
import io
import logging
import re
import uuid

from flask import Blueprint, jsonify, request, send_file

import db
from auth import current_user, can_edit
from utils import compute_expected_reply, compute_duration, is_overdue, format_date, extract_rev

exporting_bp = Blueprint("exporting", __name__)
logger = logging.getLogger(__name__)


def _normalize_sheet_name(name):
    s = str(name or "").strip().lower()
    s = re.sub(r"^\s*\d+\s*[\.\-)\]]*\s*", "", s)
    s = re.sub(r"[(){}\[\],.:;_]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip(" -_/")
    return s


def _sheet_aliases(dt):
    vals = {
        str(dt.get("id", "")).strip().lower(),
        str(dt.get("code", "")).strip().lower(),
        str(dt.get("name", "")).strip().lower(),
        _normalize_sheet_name(dt.get("id", "")),
        _normalize_sheet_name(dt.get("code", "")),
        _normalize_sheet_name(dt.get("name", "")),
    }
    code = str(dt.get("code", "")).strip().lower()
    name = _normalize_sheet_name(dt.get("name", ""))
    if code == "ltr" or "letter" in name:
        vals.update({"letter", "letters"})
    if code == "pr" or "requisition" in name:
        vals.update({"purchase requisition", "purchase requisitions"})
    return {v for v in vals if v}


def _match_sheet_to_dt(sheet_name, dts):
    raw = str(sheet_name or "").strip()
    low = raw.lower()
    norm = _normalize_sheet_name(raw)
    no_paren = _normalize_sheet_name(re.sub(r"\([^)]*\)", " ", raw))
    no_prefix = _normalize_sheet_name(re.sub(r"^\s*\d+\s*[\.\-)\]]*\s*", "", raw))
    paren_code = ""
    m = re.search(r"\(([A-Za-z0-9_-]+)\)", raw)
    if m:
        paren_code = m.group(1).strip().lower()

    def exact_match(field):
        for dt in dts:
            if str(dt.get(field, "")).strip().lower() == low:
                return dt
        return None

    for field in ("id", "code", "name"):
        dt = exact_match(field)
        if dt:
            return dt

    for dt in dts:
        aliases = _sheet_aliases(dt)
        if norm in aliases or no_paren in aliases or no_prefix in aliases:
            return dt

    if paren_code:
        for field in ("id", "code"):
            for dt in dts:
                if str(dt.get(field, "")).strip().lower() == paren_code:
                    return dt

    return None


def _has_meaningful_values(row_data):
    for v in row_data.values():
        if v is None:
            continue
        if isinstance(v, str):
            if v.strip():
                return True
        elif str(v).strip():
            return True
    return False


def _normalize_match_value(value):
    return str(value or "").strip().lower()


def _save_import_row(pid, dt_id, row_data):
    doc_no = str(row_data.get("docNo", "") or "").strip()
    existing = db.get_record_by_doc_no(pid, dt_id, doc_no) if _normalize_match_value(doc_no) else None
    if existing:
        merged = db.merge_record_data(existing, row_data)
        db.save_record(pid, dt_id, existing["_id"], merged)
        return "updated"
    db.save_record(pid, dt_id, str(uuid.uuid4()), row_data)
    return "created"


def _import_excel_worksheet(pid, dt_id, ws, cols):
    import datetime as _dt

    col_map = {c["label"]: c["col_key"] for c in cols}
    header = None
    imported = 0
    created = 0
    updated = 0
    skipped_blank = 0
    skipped_invalid = 0
    warnings = []
    for row_idx, row in enumerate(ws.iter_rows(values_only=True), start=1):
        vals = []
        for v in row:
            if v is None: vals.append("")
            elif isinstance(v, (_dt.datetime, _dt.date)):
                vals.append((v.date() if isinstance(v, _dt.datetime) else v).strftime("%Y-%m-%d"))
            else: vals.append(str(v).strip())
        if not any(vals):
            skipped_blank += 1
            continue
        if header is None:
            if any(v in col_map or v in ("Sr.","Document No.") for v in vals):
                header = [col_map.get(v.strip(), v.strip()) for v in vals]
            continue
        row_data = {header[i]: v for i, v in enumerate(vals)
                    if i < len(header) and header[i] and header[i] not in ("Sr.","sr","")}
        if not _has_meaningful_values(row_data):
            skipped_blank += 1
            continue
        try:
            action = _save_import_row(pid, dt_id, row_data)
            imported += 1
            if action == "updated":
                updated += 1
            else:
                created += 1
        except Exception as e:
            skipped_invalid += 1
            logger.warning("import_row_skipped pid=%s dt_id=%s row=%s error=%s",
                           pid, dt_id, row_idx, e)
            if len(warnings) < 20:
                warnings.append({"row": row_idx, "error": str(e)})
    return imported, created, updated, header is not None, skipped_blank, skipped_invalid, warnings


def _is_pr_dt(dt):
    if not dt: return False
    code = str(dt.get("code","")).strip().upper()
    name = str(dt.get("name","")).strip().lower()
    dt_id = str(dt.get("id","")).strip().upper()
    return dt_id == "PR" or code == "PR" or "requisition" in name or "purchase request" in name


def _pr_details_key(cols):
    def _norm(v):
        return re.sub(r"[^a-z0-9]+", "", str(v or "").strip().lower())
    for c in cols:
        key = str(c.get("col_key", ""))
        label = str(c.get("label", "")).strip()
        if label == "PR Details" or key in ("prDetails", "pr_details"):
            return c.get("col_key")
    for c in cols:
        nk = _norm(c.get("col_key", ""))
        nl = _norm(c.get("label", ""))
        if nk == "prdetails" or nl == "prdetails":
            return c.get("col_key")
    for c in cols:
        nk = _norm(c.get("col_key", ""))
        nl = _norm(c.get("label", ""))
        if "pr" in nk and "detail" in nk:
            return c.get("col_key")
        if "pr" in nl and "detail" in nl:
            return c.get("col_key")
    return None


def _pr_items_text(items):
    if not items: return ""
    lines = []
    for it in items:
        row_type = str(it.get("row_type","item") or "item").strip().lower()
        name = str(it.get("item_name","") or "").strip()
        if row_type == "header":
            if name:
                lines.append(name)
            continue
        unit = str(it.get("unit","") or "").strip()
        qty  = it.get("quantity")
        remarks = str(it.get("remarks","") or "").strip()
        if qty is not None and qty != "":
            qv = str(qty)
        else:
            qv = ""
        parts = [p for p in [name, unit, qv] if p]
        line = " | ".join(parts) if parts else ""
        if remarks:
            line = (line + " — " if line else "") + remarks
        if line:
            lines.append(line)
    return "\n".join(lines)


def _resolve_pr_details_value(row, pr_items_map, pr_details_key):
    manual = str(row.get(pr_details_key, "") or "").strip() if pr_details_key else ""
    if manual:
        return manual
    return _pr_items_text(pr_items_map.get(row.get("_id"), []))


def _safe_excel_name_part(val, fallback):
    text = re.sub(r'[\\/:*?"<>|]+', "_", str(val or "").strip())
    text = re.sub(r"\s+", "_", text).strip("._")
    return text or fallback


def _pick_first(row, keys):
    for key in keys:
        val = str(row.get(key, "") or "").strip()
        if val:
            return val
    return ""


def _is_floor_field(col_key, label=""):
    key = str(col_key or "").strip().lower()
    lbl = str(label or "").strip().lower()
    return key == "floor" or lbl in ("floor", "floors")


def _is_item_ref_field(col_key, label=""):
    key = str(col_key or "").strip().lower()
    lbl = re.sub(r"[_./-]+", " ", str(label or "").strip().lower())
    return key == "itemref" or ("item ref" in lbl and "dwg" in lbl)


def _format_multiline_display_value(col_key, label, value):
    text = str(value or "")
    if not text:
        return ""
    if _is_floor_field(col_key, label):
        return "\n".join([p.strip() for p in text.split(",") if p.strip()])
    if _is_item_ref_field(col_key, label):
        return text.replace("\r\n", "\n").replace("\r", "\n")
    return text


_REV_TOKEN_RE = re.compile(r"\bREV\s*[-_ ]?(\d+)\b", re.IGNORECASE)


def normalize_doc_base(doc_no):
    text = str(doc_no or "").strip()
    if not text:
        return ""
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r"\s+", " ", text)
    text = re.sub(r"(?i)(?:[\s/_-]+)?REV\s*[-_ ]?\d+\s*$", "", text).strip()
    text = re.sub(r"(?i)\s+REV\s*[-_ ]?\d+\b", " ", text).strip()
    return re.sub(r"\s+", " ", text) or str(doc_no or "").strip()


def extract_rev_no(doc_no):
    match = _REV_TOKEN_RE.search(str(doc_no or ""))
    if not match:
        return 0
    try:
        return int(match.group(1))
    except ValueError:
        return 0


def _natural_sort_key(value):
    parts = re.split(r"(\d+)", str(value or ""))
    return [int(part) if part.isdigit() else part.lower() for part in parts]


def _doc_revision_sort_key(row):
    doc_no = row.get("docNo", "") if isinstance(row, dict) else ""
    base = normalize_doc_base(doc_no)
    return (
        _natural_sort_key(base),
        extract_rev_no(doc_no),
        _natural_sort_key(doc_no),
        str(row.get("_id", "")) if isinstance(row, dict) else "",
    )


def _excel_column_width(col_key, label):
    key = str(col_key or "").strip()
    key_l = key.lower()
    label_l = re.sub(r"[_./-]+", " ", str(label or "").strip().lower())
    hay = f"{key_l} {label_l}"

    if key == "_sr":
        return 6
    if key_l in ("docno", "doc_no") or "document no" in hay or "letter ref" in hay:
        return 24
    if "title" in hay or "subject" in hay or "description" in hay:
        return 46
    if "pr detail" in hay or key_l in ("prdetails", "pr_details"):
        return 50
    if "remarks" in hay or "comment" in hay or "note" in hay:
        return 34
    if "ms ref" in hay or "dwg" in hay or "item ref" in hay or "reference" in hay:
        return 26
    if "file" in hay or "location" in hay or "link" in hay:
        return 22
    if "prepared" in hay or "company" in hay or "client" in hay or "consultant" in hay or "contractor" in hay:
        return 24
    if label_l in ("from", "to") or key_l in ("from", "to"):
        return 18
    if "discipline" in hay:
        return 16
    if "sub trade" in hay or "trade" in hay:
        return 18
    if "status" in hay:
        return 24
    if "date" in hay:
        return 14
    if "duration" in hay:
        return 10
    if "floor" in hay:
        return 18
    if "direction" in hay or "revision" in hay:
        return 12
    return 16


def _field_width_role(col):
    key = str(col.get("col_key", "") or "").strip()
    key_l = key.lower()
    label = str(col.get("label", "") or "").strip()
    label_l = re.sub(r"[_./-]+", " ", label.lower())
    hay = f"{key_l} {label_l}"

    if key == "_sr":
        return "serial"
    if key_l in ("docno", "doc_no") or "document no" in hay or "letter ref" in hay:
        return "document_no"
    if "pr detail" in hay or key_l in ("prdetails", "pr_details"):
        return "very_long"
    if "title" in hay or "subject" in hay:
        return "very_long"
    if "status" in hay:
        return "status"
    if "item description" in hay or "description" in hay or "scope" in hay:
        return "long"
    if "remarks" in hay or "comment" in hay or "note" in hay:
        return "long"
    if "file" in hay or "location" in hay or "link" in hay:
        return "file_link"
    if "ms ref" in hay or "dwg" in hay or "item ref" in hay or "parent letter" in hay or "reference" in hay:
        return "technical_ref"
    if "prepared" in hay or "engineer" in hay or "person" in hay:
        return "medium_wide"
    if label_l in ("from", "to") or key_l in ("from", "to"):
        return "medium"
    if "company" in hay or "client" in hay or "consultant" in hay or "contractor" in hay or "party" in hay:
        return "medium"
    if "discipline" in hay:
        return "medium"
    if "sub trade" in hay or "trade" in hay or "brand" in hay or "floor" in hay or "level" in hay:
        return "medium"
    if "duration" in hay:
        return "compact"
    if "date" in hay or "issued" in hay or "submitted" in hay or "received" in hay or "reply" in hay:
        return "date"
    if "direction" in hay or "revision" in hay or re.search(r"\brev\b", hay) or "code" in hay:
        return "compact"
    if "qty" in hay or "quantity" in hay or "unit" in hay:
        return "compact"
    return "default"


def _excel_width_from_role(role):
    return {
        "serial": 5.5,
        "document_no": 30,
        "very_long": 58,
        "long": 42,
        "technical_ref": 31,
        "status": 27,
        "file_link": 17,
        "medium_wide": 24,
        "medium": 18,
        "date": 14,
        "compact": 9.5,
        "default": 16,
    }.get(role, 16)


def _excel_width_from_web_px(px, role):
    try:
        px = float(px)
    except (TypeError, ValueError):
        return None
    if px <= 0:
        return None
    min_by_role = {
        "serial": 5,
        "document_no": 28,
        "very_long": 48,
        "long": 28,
        "technical_ref": 26,
        "status": 25,
        "file_link": 15,
        "medium_wide": 22,
        "medium": 15,
        "date": 13,
        "compact": 8.5,
        "default": 13,
    }
    max_by_role = {
        "serial": 8,
        "document_no": 36,
        "very_long": 72,
        "long": 54,
        "technical_ref": 40,
        "status": 32,
        "file_link": 22,
        "medium_wide": 34,
        "medium": 26,
        "date": 18,
        "compact": 14,
        "default": 24,
    }
    width = px / 7.2
    return max(min_by_role.get(role, 12), min(max_by_role.get(role, 28), width))


def _excel_sheet_column_widths(cols, dt_name=None, web_widths=None):
    """Build the Excel width profile from the actual ordered visible register columns."""
    web_widths = web_widths or {}
    all_cols = [{"col_key":"_sr","label":"Sr."}] + [
        {"col_key":c["col_key"],"label":c["label"]} for c in cols
    ]
    roles = [_field_width_role(c) for c in all_cols]
    widths = []
    for col, role in zip(all_cols, roles):
        live_width = _excel_width_from_web_px(web_widths.get(col["col_key"]), role)
        widths.append(live_width if live_width is not None else _excel_width_from_role(role))

    # Keep dense sheets readable without flattening every register into one generic profile.
    dt_hint = str(dt_name or "").lower()
    visible_keys = [str(c.get("col_key", "") or "").lower() for c in cols]
    visible_labels = [str(c.get("label", "") or "").lower() for c in cols]
    is_letters = (
        "letter" in dt_hint
        or any(k in {"from", "to", "fromparty", "toparty", "direction"} for k in visible_keys)
        and any("subject" in l or "letter" in l for l in visible_labels)
    )
    is_pr = (
        "purchase" in dt_hint
        or "requisition" in dt_hint
        or any("pr detail" in l or k in {"prdetails", "pr_details"} for k, l in zip(visible_keys, visible_labels))
    )

    if is_letters:
        for i, col in enumerate(all_cols):
            key = str(col.get("col_key", "") or "").lower()
            label = str(col.get("label", "") or "").lower()
            if "subject" in label:
                widths[i] = max(widths[i], 58)
            elif "letter ref" in label or key in {"letterref", "letter_ref", "docno"}:
                widths[i] = max(widths[i], 31)
            elif key in {"from", "to", "fromparty", "toparty"}:
                widths[i] = max(widths[i], 22)
            elif "parent letter" in label:
                widths[i] = max(widths[i], 30)
            elif "direction" in label or key == "direction":
                widths[i] = min(widths[i], 12)
    if is_pr:
        for i, col in enumerate(all_cols):
            label = str(col.get("label", "") or "").lower()
            key = str(col.get("col_key", "") or "").lower()
            if "pr detail" in label or key in {"prdetails", "pr_details"}:
                widths[i] = max(widths[i], 60)
            elif "ms ref" in label:
                widths[i] = max(widths[i], 28)

    return widths


def _excel_col_hay(col_key, label):
    key = str(col_key or "").strip().lower()
    label_l = re.sub(r"[_./-]+", " ", str(label or "").strip().lower())
    return re.sub(r"\s+", " ", f"{key} {label_l}").strip()


def _is_excel_date_col(col_key, label):
    hay = _excel_col_hay(col_key, label)
    key = str(col_key or "").strip().lower()
    date_tokens = (
        "date", "issued", "submitted", "submission", "received", "reply",
        "expected", "actual", "rec from", "rec by", "received from", "received by", "sent"
    )
    if _is_excel_duration_col(col_key, label) or key == "_sr" or "prepared by" in hay or "reply status" in hay:
        return False
    return any(token in hay for token in date_tokens) or bool(re.search(r"\brec\b", hay))


def _is_excel_duration_col(col_key, label):
    key = str(col_key or "").strip().lower()
    hay = _excel_col_hay(col_key, label)
    return key == "duration" or "duration" in hay or re.search(r"\bdur\.?\b", hay) is not None


def _looks_like_excel_date_value(value):
    text = str(value or "").strip()
    if not text:
        return False
    text = text.replace("T", " ")
    text = re.sub(r"\s+\d{2}:\d{2}(?::\d{2})?(?:\.\d+)?$", "", text)
    return bool(re.fullmatch(r"\d{4}-\d{2}-\d{2}|\d{2}[-/.]\d{2}[-/.]\d{4}|[0-3]?\d[-/ ][A-Za-z]{3,9}[-/ ]\d{4}", text))


def _parse_excel_date_value(value):
    if value in (None, ""):
        return None
    if isinstance(value, datetime.datetime):
        return value.date()
    if isinstance(value, datetime.date):
        return value
    text = str(value or "").strip()
    if not text:
        return None
    text = text.replace("T", " ")
    text = re.sub(r"\s+\d{2}:\d{2}(?::\d{2})?(?:\.\d+)?$", "", text)
    text = re.sub(r"\s+", " ", text)
    for fmt in (
        "%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y", "%d.%m.%Y",
        "%d/%b/%Y", "%d-%b-%Y", "%d %b %Y", "%d/%B/%Y", "%d-%B-%Y",
    ):
        try:
            return datetime.datetime.strptime(text, fmt).date()
        except ValueError:
            pass
    return None


def _excel_cell_value(col_key, label, value):
    if _is_excel_date_col(col_key, label) or _looks_like_excel_date_value(value):
        parsed = _parse_excel_date_value(value)
        if parsed:
            return parsed
    return value


def _excel_center_col(col_key, label):
    key = str(col_key or "").strip()
    label_l = str(label or "").strip().lower()
    return (
        key in {"_sr", "issuedDate", "expectedReplyDate", "actualReplyDate"}
        or _is_excel_duration_col(col_key, label)
        or "date" in key.lower()
        or "date" in label_l
        or key == "status"
        or "status" in label_l
    )


def _excel_row_height(values, cols=None, widths=None, base=19):
    max_lines = 1
    cols = cols or []
    widths = widths or []
    for idx, value in enumerate(values):
        text = str(value or "")
        if not text:
            continue
        col = cols[idx] if idx < len(cols) else {}
        width = widths[idx] if idx < len(widths) else 18
        hay = _excel_col_hay(col.get("col_key", ""), col.get("label", ""))
        explicit = text.count("\n") + 1
        chars_per_line = max(14, int(width * 1.25))
        wrapped = max(1, (max((len(part) for part in text.split("\n")), default=0) // chars_per_line) + 1)
        line_cap = 2 if ("status" in hay or "document no" in hay or "letter ref" in hay) else 3
        if "title" in hay or "subject" in hay or "remarks" in hay or "description" in hay or "pr detail" in hay:
            line_cap = 4
        max_lines = max(max_lines, min(line_cap, max(explicit, wrapped)))
    return max(base, min(34, 13 * max_lines + 4))


def _excel_wrap_cell(col_key, label):
    return not (str(col_key or "") == "_sr" or _is_excel_duration_col(col_key, label) or _is_excel_date_col(col_key, label))


def _write_register_excel_sheet(ws, proj, dt, cols, records, pr_items_map=None, pr_details_key=None, web_widths=None):
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    pr_items_map = pr_items_map or {}
    PRIMARY = "1A3A5C"; PL = "2563A8"; WHITE = "FFFFFF"
    ALT = "F8FAFC"; OV = "FFF5F5"; MUTED = "9CA3AF"

    STATUS_XL = {
        "A - Approved":              ("BBF7D0","166534"),
        "B - Approved As Noted":     ("DCFCE7","14532D"),
        "B,C - Approved & Resubmit": ("FED7AA","7C2D12"),
        "C - Revise & Resubmit":     ("FCE7F3","831843"),
        "D - Review not Required":   ("FECACA","7F1D1D"),
        "Under Review":              ("FEF9C3","713F12"),
        "Cancelled":                 ("EF4444","FFFFFF"),
        "Open":                      ("FED7AA","7C2D12"),
        "Closed":                    ("BFDBFE","1E3A5F"),
        "Replied":                   ("D1FAE5","064E3B"),
        "Pending":                   ("E0E7FF","312E81"),
    }

    def fill(c): return PatternFill("solid", fgColor=c)
    def thin(): s=Side(style="thin",color="DDE3ED"); return Border(left=s,right=s,top=s,bottom=s)

    ws.sheet_view.showGridLines = False
    all_cols = [{"col_key":"_sr","label":"Sr."}] + [{"col_key":c["col_key"],"label":c["label"]} for c in cols]
    nc = len(all_cols)

    def mcell(row, val, bg, fg="FFFFFF", bold=False, sz=11, halign="center"):
        c = ws.cell(row=row, column=1, value=val)
        ws.merge_cells(f"A{row}:{get_column_letter(nc)}{row}")
        c.font = Font(bold=bold, color=fg, size=sz, name="Arial")
        c.fill = fill(bg)
        c.alignment = Alignment(horizontal=halign, vertical="center")
        return c

    dt_name = (dt.get("name") if dt else "") or (dt.get("id") if dt else "") or "REGISTER"
    pid = (proj or {}).get("id")
    dt_id = (dt or {}).get("id")
    expected_reply_rule = db.get_expected_reply_rule(pid, dt_id) if pid and dt_id else None
    mcell(1, f"DOCUMENT CONTROL REGISTER  -  {str(dt_name).upper()}",
          PRIMARY, bold=True, sz=13)
    ws.row_dimensions[1].height = 30

    info = "   |   ".join(f"{k}: {v}" for k,v in [
        ("Project",proj.get("name","")),("Code",proj.get("code","")),
        ("Client",proj.get("client","")),("Consultant",proj.get("mainConsultant","")),
        ("Exported",datetime.datetime.now().strftime("%d-%m-%Y %H:%M"))] if v)
    mcell(2, info, PL, sz=9)
    ws.row_dimensions[2].height = 16
    ws.row_dimensions[3].height = 3

    ws.freeze_panes = "A5"
    ws.print_title_rows = "4:4"
    ws.sheet_view.zoomScale = 90
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.oddFooter.right.text = "Page &[Page] of &[Pages]"
    ws.oddFooter.left.text = "Generated from DCR System"

    col_widths = _excel_sheet_column_widths(cols, dt_name=dt_name, web_widths=web_widths)
    for ci, col in enumerate(all_cols, 1):
        ws.column_dimensions[get_column_letter(ci)].width = col_widths[ci - 1]
        c = ws.cell(row=4, column=ci, value=col["label"])
        c.font = Font(bold=True,color=WHITE,size=10,name="Arial")
        c.fill = fill(PRIMARY)
        c.alignment = Alignment(horizontal="center",vertical="center",wrap_text=True)
        c.border = thin()
    ws.row_dimensions[4].height = 22

    has_exp_col = any(c["col_key"]=="expectedReplyDate" for c in all_cols)
    sr = 1
    if not records:
        c = ws.cell(row=5, column=1, value="No records in this register")
        c.font = Font(italic=True, size=10, name="Arial", color=MUTED)
        ws.merge_cells(f"A5:{get_column_letter(nc)}5")
        c.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[5].height = 24
        total_row = 6
    else:
        for ri, row in enumerate(records):
            rn   = 5 + ri
            is_rev = extract_rev(row.get("docNo","")) > 0
            status_val = row.get("status")
            action_val = row.get("action")
            ov     = is_overdue(row.get("issuedDate"), row.get("docNo"), row.get("actualReplyDate"), has_exp_col, expected_reply_rule, status=status_val, action=action_val)
            bg     = OV if ov else (ALT if sr%2==0 else WHITE)
            row_values = []

            for ci, col in enumerate(all_cols, 1):
                key = col["col_key"]
                if key=="_sr":                  val = "" if is_rev else str(sr)
                elif key=="expectedReplyDate":  val = format_date(compute_expected_reply(row.get("issuedDate"),row.get("docNo"), expected_reply_rule, status=status_val, action=action_val))
                elif _is_excel_duration_col(key, col["label"]):
                    dur_val = compute_duration(row.get("issuedDate"), row.get("actualReplyDate"), expected_reply_rule, status=status_val, action=action_val)
                    val = str(dur_val) if dur_val is not None else ""
                elif key=="issuedDate":         val = format_date(row.get(key,""))
                elif key=="actualReplyDate":    val = format_date(row.get(key,""))
                elif pr_details_key and key == pr_details_key:
                    val = _resolve_pr_details_value(row, pr_items_map, pr_details_key) or str(row.get(key,"") or "")
                else:                           val = str(row.get(key,"") or "")
                val = _format_multiline_display_value(key, col["label"], val)
                row_values.append(val)
                cell_value = _excel_cell_value(key, col["label"], val)

                if key == "fileLocation" and val and val.startswith("http"):
                    c = ws.cell(row=rn, column=ci)
                    c.value = "View"
                    c.hyperlink = val
                    c.font = Font(size=9,name="Arial",color="2563A8",underline="single")
                else:
                    c = ws.cell(row=rn, column=ci, value=cell_value)
                    if isinstance(cell_value, datetime.date):
                        c.number_format = "DD-MM-YYYY"
                    if _is_excel_duration_col(key, col["label"]) and val == "0":
                        c.value = 0
                    if key=="status" and val:
                        bg2, fg2 = STATUS_XL.get(val, ("F3F4F6","374151"))
                        c.fill = fill(bg2); c.font = Font(bold=True,size=9,name="Arial",color=fg2)
                    elif key=="docNo":
                        c.fill = fill(bg)
                        c.font = Font(size=10,name="Consolas",bold=not is_rev,
                                      color=MUTED if is_rev else PRIMARY)
                    else:
                        c.fill = fill(bg)
                        c.font = Font(size=10,name="Arial",
                                      color=MUTED if is_rev else ("991B1B" if ov else "1E2A3A"))
                c.border    = thin()
                c.alignment = Alignment(
                    vertical="center",
                    wrap_text=_excel_wrap_cell(key, col["label"]),
                    horizontal="center" if _excel_center_col(key, col["label"]) else "left",
                )
            ws.row_dimensions[rn].height = _excel_row_height(row_values, all_cols, col_widths, base=19)
            if not is_rev: sr += 1
        total_row = 5 + len(records)

    real = sum(1 for r in records if extract_rev(r.get("docNo",""))==0)
    mcell(total_row, f"TOTAL: {real} documents  |  {len(records)} submissions", PRIMARY, sz=10, halign="left")
    ws.row_dimensions[total_row].height = 22
    if total_row >= 4:
        ws.auto_filter.ref = f"A4:{get_column_letter(nc)}{max(4, total_row - 1)}"


def _build_pr_register_excel(proj, dt, records, pr_items_map, pr_details_key):
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "PR Summary"
    raw_ws = wb.create_sheet("PR Items Raw")
    ws.sheet_view.showGridLines = False
    raw_ws.sheet_view.showGridLines = False

    PRIMARY = "1A3A5C"
    ACCENT = "2563A8"
    LIGHT = "F0F4F8"
    SECTION = "E8EEF6"
    LABEL = "E2E8F0"
    SUBTLE = "F8FAFC"
    WHITE = "FFFFFF"
    MUTED = "64748B"
    TEXT = "1E2A3A"

    thin_side = Side(style="thin", color="DDE3ED")
    thin = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    top_rule = Border(top=thin_side)

    def fill(color):
        return PatternFill("solid", fgColor=color)

    def style_cell(cell, *, font=None, fill_color=None, border=None, align=None):
        if font: cell.font = font
        if fill_color: cell.fill = fill(fill_color)
        if border: cell.border = border
        if align: cell.alignment = align

    for col, width in {"A": 6, "B": 22, "C": 13, "D": 18, "E": 64, "F": 14}.items():
        ws.column_dimensions[col].width = width
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.print_options.horizontalCentered = False
    ws.oddFooter.right.text = "Page &[Page] of &[Pages]"
    ws.oddFooter.left.text = "Generated from DCR System"

    ws.merge_cells("A1:F1")
    c = ws["A1"]
    c.value = "PURCHASE REQUISITION REGISTER"
    style_cell(
        c,
        font=Font(name="Arial", size=16, bold=True, color=PRIMARY),
        fill_color=LIGHT,
        border=thin,
        align=Alignment(horizontal="center", vertical="center"),
    )
    ws.row_dimensions[1].height = 30

    ws.merge_cells("A2:F2")
    c = ws["A2"]
    c.value = "Document Control Register"
    style_cell(
        c,
        font=Font(name="Arial", size=10, italic=True, color=MUTED),
        fill_color=LIGHT,
        border=thin,
        align=Alignment(horizontal="center", vertical="center"),
    )
    ws.row_dimensions[2].height = 18
    ws.merge_cells("A3:F3")
    c = ws["A3"]
    c.value = ""
    c.fill = fill(WHITE)
    c.border = Border(bottom=thin_side)
    ws.row_dimensions[3].height = 6

    ws.merge_cells("A4:F4")
    c = ws["A4"]
    c.value = "PROJECT / PR INFORMATION"
    style_cell(
        c,
        font=Font(name="Arial", size=11, bold=True, color=WHITE),
        fill_color=ACCENT,
        border=thin,
        align=Alignment(horizontal="left", vertical="center"),
    )
    ws.row_dimensions[4].height = 20

    meta_fields = [
        ("Project Name", proj.get("name", "")),
        ("Project Code", proj.get("code", "")),
        ("Register", str(dt.get("name", "") or dt.get("id", "PR")).strip()),
        ("Exported", datetime.datetime.now().strftime("%d-%m-%Y %H:%M")),
        ("Total PRs", str(len(records))),
    ]
    meta_fields = [(k, v) for k, v in meta_fields if str(v or "").strip()]
    row_no = 5
    for i in range(0, len(meta_fields), 2):
        left = meta_fields[i]
        right = meta_fields[i + 1] if i + 1 < len(meta_fields) else None
        ws[f"A{row_no}"] = left[0]
        style_cell(
            ws[f"A{row_no}"],
            font=Font(name="Arial", size=10, bold=True, color=PRIMARY),
            fill_color=LABEL,
            border=thin,
            align=Alignment(vertical="center"),
        )
        ws.merge_cells(f"B{row_no}:C{row_no}")
        ws[f"B{row_no}"] = left[1]
        style_cell(
            ws[f"B{row_no}"],
            font=Font(name="Arial", size=10, color=TEXT),
            fill_color=WHITE,
            border=thin,
            align=Alignment(vertical="center", wrap_text=True),
        )
        for cell in (f"C{row_no}",):
            ws[cell].border = thin
        if right:
            ws[f"D{row_no}"] = right[0]
            style_cell(
                ws[f"D{row_no}"],
                font=Font(name="Arial", size=10, bold=True, color=PRIMARY),
                fill_color=LABEL,
                border=thin,
                align=Alignment(vertical="center"),
            )
            ws[f"E{row_no}"] = right[1]
            style_cell(
                ws[f"E{row_no}"],
                font=Font(name="Arial", size=10, color=TEXT),
                fill_color=WHITE,
                border=thin,
                align=Alignment(vertical="center", wrap_text=True),
            )
        else:
            ws[f"D{row_no}"].fill = fill(WHITE)
            ws[f"D{row_no}"].border = thin
            ws[f"E{row_no}"].fill = fill(WHITE)
            ws[f"E{row_no}"].border = thin
        ws.row_dimensions[row_no].height = 20
        row_no += 1

    ws.merge_cells(f"A{row_no}:F{row_no}")
    c = ws[f"A{row_no}"]
    c.value = "REQUISITION SUMMARY"
    style_cell(
        c,
        font=Font(name="Arial", size=11, bold=True, color=WHITE),
        fill_color=PRIMARY,
        border=thin,
        align=Alignment(horizontal="left", vertical="center"),
    )
    ws.row_dimensions[row_no].height = 20
    row_no += 1

    summary_text = (
        f"Total requisitions: {len(records)}\n"
        f"Project: {proj.get('name', '')} ({proj.get('code', '')})\n"
        f"Generated from DCR System on {datetime.datetime.now().strftime('%d-%m-%Y %H:%M')}"
    )
    detail_lines = max(3, min(8, len(str(summary_text).splitlines()) + 1))
    ws.merge_cells(f"A{row_no}:F{row_no}")
    c = ws[f"A{row_no}"]
    c.value = summary_text
    style_cell(
        c,
        font=Font(name="Arial", size=10, color=TEXT),
        fill_color=SUBTLE,
        border=thin,
        align=Alignment(vertical="top", wrap_text=True),
    )
    ws.row_dimensions[row_no].height = detail_lines * 16
    row_no += 2

    table_header_row = row_no
    headers = ["No.", "PR Number", "PR Date", "Discipline / Trade", "PR Title", "Prepared By"]
    for idx, label in enumerate(headers, start=1):
        c = ws.cell(row=table_header_row, column=idx, value=label)
        style_cell(
            c,
            font=Font(name="Arial", size=10, bold=True, color=WHITE),
            fill_color=PRIMARY,
            border=thin,
            align=Alignment(horizontal="center", vertical="center", wrap_text=True),
        )
    ws.row_dimensions[table_header_row].height = 22
    ws.freeze_panes = f"A{table_header_row + 1}"
    ws.print_title_rows = f"{table_header_row}:{table_header_row}"
    ws.sheet_view.zoomScale = 90

    item_no = 1
    data_row = table_header_row + 1
    if records:
        for row in records:
            pr_number = _pick_first(row, ("docNo", "prNo", "prNumber")) or f"PR-{str(row.get('_id',''))[:8]}"
            pr_date = format_date(_pick_first(row, ("issuedDate", "prDate", "date")))
            discipline = _pick_first(row, ("discipline",))
            trade = _pick_first(row, ("trade",))
            disc_trade = " / ".join([v for v in [discipline, trade] if v]) or "Unspecified"
            pr_title = _pick_first(row, ("title", "subject", "description", "docTitle")) or _resolve_pr_details_value(row, pr_items_map, pr_details_key)
            prepared_by = _pick_first(row, ("preparedBy", "prepared_by", "requestedBy", "requester", "requested_by"))
            values = [item_no, pr_number, pr_date, disc_trade, pr_title, prepared_by]
            row_fill = WHITE if item_no % 2 else SUBTLE
            for col, val in enumerate(values, start=1):
                cell = ws.cell(row=data_row, column=col, value=val)
                style_cell(
                    cell,
                    font=Font(name="Arial", size=10, color=TEXT),
                    fill_color=row_fill,
                    border=thin,
                    align=Alignment(
                        horizontal="center" if col in (1, 3) else "left",
                        vertical="top",
                        wrap_text=True,
                    ),
                )
            ws.row_dimensions[data_row].height = 34 if len(str(pr_title or "")) > 80 else 22
            item_no += 1
            data_row += 1
    else:
        ws.merge_cells(start_row=data_row, start_column=1, end_row=data_row, end_column=6)
        c = ws.cell(row=data_row, column=1, value="No requisitions in this register")
        c.font = Font(name="Arial", size=10, italic=True, color=MUTED)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = thin
        ws.row_dimensions[data_row].height = 20
        data_row += 1

    sig_row = data_row + 2
    signatures = [("Prepared By", "A", "B"), ("Reviewed By", "C", "D"), ("Approved By", "E", "F")]
    ws.merge_cells(f"A{sig_row-1}:F{sig_row-1}")
    c = ws[f"A{sig_row-1}"]
    c.value = "SIGNATURES"
    style_cell(
        c,
        font=Font(name="Arial", size=11, bold=True, color=WHITE),
        fill_color=ACCENT,
        border=thin,
        align=Alignment(horizontal="left", vertical="center"),
    )
    ws.row_dimensions[sig_row-1].height = 20
    for label, start_col, end_col in signatures:
        ws.merge_cells(f"{start_col}{sig_row}:{end_col}{sig_row}")
        top = ws[f"{start_col}{sig_row}"]
        top.value = ""
        top.border = top_rule
        top.fill = fill(WHITE)
        top.alignment = Alignment(horizontal="center")
        ws.merge_cells(f"{start_col}{sig_row+1}:{end_col}{sig_row+1}")
        lbl = ws[f"{start_col}{sig_row+1}"]
        lbl.value = label
        lbl.font = Font(name="Arial", size=9, bold=True, color=MUTED)
        lbl.alignment = Alignment(horizontal="center")
        lbl.fill = fill(SUBTLE)
    ws.row_dimensions[sig_row].height = 22
    ws.row_dimensions[sig_row + 1].height = 18

    footer_row = sig_row + 3
    ws.merge_cells(f"A{footer_row}:F{footer_row}")
    c = ws[f"A{footer_row}"]
    c.value = f"Generated from DCR System | Export date: {datetime.datetime.now().strftime('%d-%m-%Y %H:%M')}"
    c.font = Font(name="Arial", size=8, italic=True, color=MUTED)
    c.alignment = Alignment(horizontal="right")
    c.border = Border(top=thin_side)

    raw_headers = ["sort_order", "row_type", "description", "unit", "qty", "remarks"]
    raw_ws.merge_cells("A1:F1")
    c = raw_ws["A1"]
    c.value = "PR ITEMS RAW"
    style_cell(
        c,
        font=Font(name="Arial", size=12, bold=True, color=PRIMARY),
        fill_color=LIGHT,
        border=thin,
        align=Alignment(horizontal="center", vertical="center"),
    )
    raw_ws.row_dimensions[1].height = 22
    for idx, label in enumerate(raw_headers, start=1):
        c = raw_ws.cell(row=2, column=idx, value=label)
        c.font = Font(name="Arial", size=10, bold=True)
        c.fill = fill(LABEL)
        c.border = thin
        c.alignment = Alignment(horizontal="center")
    for col, width in {1: 12, 2: 12, 3: 52, 4: 12, 5: 10, 6: 24}.items():
        raw_ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = width
    raw_ws.freeze_panes = "A3"
    raw_row = 3
    for row in records:
        pr_number = _pick_first(row, ("docNo", "prNo", "prNumber")) or f"PR-{str(row.get('_id',''))[:8]}"
        items = pr_items_map.get(row.get("_id"), [])
        raw_ws.merge_cells(start_row=raw_row, start_column=1, end_row=raw_row, end_column=6)
        c = raw_ws.cell(row=raw_row, column=1, value=pr_number)
        c.border = thin
        c.alignment = Alignment(horizontal="left", vertical="center")
        c.font = Font(name="Arial", size=10, bold=True, color=PRIMARY)
        c.fill = fill(LABEL)
        for col in range(2, 7):
            raw_ws.cell(row=raw_row, column=col).border = thin
            raw_ws.cell(row=raw_row, column=col).fill = fill(LABEL)
        raw_ws.row_dimensions[raw_row].height = 20
        raw_row += 1
        if not items:
            values = ["", "item", "No PR items", "", "", ""]
            for col, val in enumerate(values, start=1):
                c = raw_ws.cell(row=raw_row, column=col, value=val)
                c.border = thin
                c.alignment = Alignment(vertical="top", wrap_text=True)
            raw_row += 1
        else:
            item_sort = 1
            for it in items:
                row_type = str(it.get("row_type", "item") or "item").strip().lower()
                is_header = row_type == "header"
                values = [
                    "" if is_header else item_sort,
                    row_type,
                    str(it.get("item_name", "") or "").strip(),
                    "" if is_header else str(it.get("unit", "") or "").strip(),
                    "" if is_header else it.get("quantity", ""),
                    "" if is_header else str(it.get("remarks", "") or "").strip(),
                ]
                for col, val in enumerate(values, start=1):
                    c = raw_ws.cell(row=raw_row, column=col, value=val)
                    c.border = thin
                    c.alignment = Alignment(vertical="top", wrap_text=True)
                    if is_header:
                        c.font = Font(name="Arial", size=10, bold=True, color=PRIMARY)
                        c.fill = fill(SECTION)
                if is_header:
                    raw_ws.row_dimensions[raw_row].height = 22
                else:
                    item_sort += 1
                raw_row += 1

    if raw_row > 3:
        raw_ws.auto_filter.ref = f"A2:F{raw_row-1}"

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


@exporting_bp.route("/api/export_all/<pid>")
def api_export_all(pid):
    u = current_user()
    if u: db.log_action(u["username"],"EXPORT_EXCEL",pid,detail="Export All")
    import openpyxl

    proj    = db.get_project(pid) or {}
    dts     = db.get_doc_types(pid)
    wb      = openpyxl.Workbook()
    wb.remove(wb.active)  # remove default sheet

    PRIMARY="1A3A5C"; PL="2563A8"; WHITE="FFFFFF"; ALT="F8FAFC"; OV="FFF5F5"; MUTED="9CA3AF"
    STATUS_XL = {
        "A - Approved":              ("BBF7D0","166534"),
        "B - Approved As Noted":     ("DCFCE7","14532D"),
        "B,C - Approved & Resubmit": ("FED7AA","7C2D12"),
        "C - Revise & Resubmit":     ("FCE7F3","831843"),
        "D - Review not Required":   ("FECACA","7F1D1D"),
        "Under Review":              ("FEF9C3","713F12"),
        "Cancelled":                 ("EF4444","FFFFFF"),
        "Open":                      ("FED7AA","7C2D12"),
        "Closed":                    ("BFDBFE","1E3A5F"),
        "Replied":                   ("D1FAE5","064E3B"),
        "Pending":                   ("E0E7FF","312E81"),
    }
    def fill(c): return PatternFill("solid", fgColor=c)
    def thin(): s=Side(style="thin",color="DDE3ED"); return Border(left=s,right=s,top=s,bottom=s)

    for dt in dts:
        cols    = [c for c in db.get_columns(pid, dt["id"]) if c["visible"]]
        records = sorted(db.get_records(pid, dt["id"]), key=_doc_revision_sort_key)
        is_pr = _is_pr_dt(dt)
        pr_details_key = _pr_details_key(cols) if is_pr else None
        pr_items_map = db.get_pr_items_for_records([r.get("_id") for r in records]) if is_pr else {}
        web_widths = db.get_col_widths(pid, dt["id"])
        ws = wb.create_sheet(title=dt["id"][:31])
        _write_register_excel_sheet(ws, proj, dt, cols, records, pr_items_map, pr_details_key, web_widths)

    if not wb.sheetnames:
        ws = wb.create_sheet("Empty"); ws.cell(1,1,"No data")

    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    fname = f"{proj.get('code','DCR')}_All_Registers.xlsx"
    return send_file(buf, as_attachment=True, download_name=fname,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@exporting_bp.route("/api/export/<pid>/<dt_id>")
def api_export(pid, dt_id):
    import openpyxl

    proj    = db.get_project(pid) or {}
    cols    = [c for c in db.get_columns(pid, dt_id) if c["visible"]]
    records = sorted(db.get_records(pid, dt_id), key=_doc_revision_sort_key)
    dts     = db.get_doc_types(pid)
    dt      = next((d for d in dts if d["id"] == dt_id), None)
    is_pr = _is_pr_dt(dt)
    pr_details_key = _pr_details_key(cols) if is_pr else None
    pr_items_map = db.get_pr_items_for_records([r.get("_id") for r in records]) if is_pr else {}
    web_widths = db.get_col_widths(pid, dt_id)

    if is_pr:
        buf = _build_pr_register_excel(proj, dt or {"id": dt_id, "name": dt_id}, records, pr_items_map, pr_details_key)
        fname = f"{_safe_excel_name_part(proj.get('code','DCR'), 'DCR')}_{_safe_excel_name_part(dt_id, 'PR')}_Register.xlsx"
        return send_file(
            buf,
            as_attachment=True,
            download_name=fname,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = dt_id[:31]
    _write_register_excel_sheet(ws, proj, dt or {"id": dt_id, "name": dt_id}, cols, records, pr_items_map, pr_details_key, web_widths)
    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    fname = f"{proj.get('code','DCR')}_{dt_id}_Register.xlsx"
    return send_file(buf, as_attachment=True, download_name=fname,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


def _build_pdf_for_dt(pid, dt_id, proj, buf=None):
    """Build a PDF for one document type. Returns BytesIO."""
    from reportlab.platypus import (SimpleDocTemplate, Table, TableStyle,
                                    Paragraph, Spacer, HRFlowable)
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.lib import colors as rl_colors
    from xml.sax.saxutils import escape as xml_escape

    buf = buf or io.BytesIO()
    cols    = [c for c in db.get_columns(pid, dt_id) if c["visible"]]
    records = db.get_records(pid, dt_id)
    dts     = db.get_doc_types(pid)
    dt      = next((d for d in dts if d["id"] == dt_id), {"name": dt_id, "code": dt_id})
    is_pr = _is_pr_dt(dt)
    pr_details_key = _pr_details_key(cols) if is_pr else None
    pr_items_map = db.get_pr_items_for_records([r.get("_id") for r in records]) if is_pr else {}

    PAGE = landscape(A4)
    doc = SimpleDocTemplate(buf, pagesize=PAGE,
                            leftMargin=10*mm, rightMargin=10*mm,
                            topMargin=12*mm, bottomMargin=12*mm)
    styles = getSampleStyleSheet()
    DARK   = rl_colors.HexColor("#1A3A5C")
    LIGHT  = rl_colors.HexColor("#F8FAFC")
    OV_COL = rl_colors.HexColor("#FEF2F2")
    REV_COL= rl_colors.HexColor("#F0F9FF")
    ALT    = rl_colors.HexColor("#F0F4F8")

    STATUS_PDF = {
        "A - Approved":              rl_colors.HexColor("#BBF7D0"),
        "B - Approved As Noted":     rl_colors.HexColor("#DCFCE7"),
        "B,C - Approved & Resubmit": rl_colors.HexColor("#FED7AA"),
        "C - Revise & Resubmit":     rl_colors.HexColor("#FCE7F3"),
        "D - Review not Required":   rl_colors.HexColor("#FECACA"),
        "Under Review":              rl_colors.HexColor("#FEF9C3"),
        "Cancelled":                 rl_colors.HexColor("#EF4444"),
        "Open":                      rl_colors.HexColor("#FED7AA"),
        "Closed":                    rl_colors.HexColor("#BFDBFE"),
        "Replied":                   rl_colors.HexColor("#D1FAE5"),
        "Pending":                   rl_colors.HexColor("#E0E7FF"),
    }

    pstyle = ParagraphStyle("cell", fontName="Helvetica", fontSize=7, leading=9,
                             wordWrap="LTR", spaceAfter=0, spaceBefore=0)
    hstyle = ParagraphStyle("hdr", fontName="Helvetica-Bold", fontSize=7.5,
                             leading=10, textColor=rl_colors.white)

    hdr_cols = [{"col_key":"_sr","label":"Sr."}] + [{"col_key":c["col_key"],"label":c["label"]} for c in cols]

    # Column widths (points)
    W_MAP = {"_sr":18,"docNo":68,"discipline":55,"trade":55,"title":110,
             "floor":38,"itemRef":48,"issuedDate":46,"expectedReplyDate":48,
             "actualReplyDate":46,"status":80,"duration":34,"remarks":80,"fileLocation":60}
    page_w = PAGE[0] - 20*mm
    col_ws = [W_MAP.get(c["col_key"], 55) for c in hdr_cols]
    scale  = page_w / sum(col_ws)
    col_ws = [w * scale for w in col_ws]

    # Build header row
    hdr_row = [Paragraph(c["label"], hstyle) for c in hdr_cols]
    data    = [hdr_row]
    row_meta = []   # (bg, is_header)


    sr = 1
    has_exp_col_pdf = any(c["col_key"]=="expectedReplyDate" for c in hdr_cols)
    for row in records:
        is_rev = extract_rev(row.get("docNo","")) > 0
        status_val = row.get("status")
        action_val = row.get("action")
        ov     = is_overdue(row.get("issuedDate"), row.get("docNo"), row.get("actualReplyDate"), has_exp_col_pdf, rule=None, status=status_val, action=action_val)
        cells  = []
        for c in hdr_cols:
            key = c["col_key"]
            if key == "_sr":
                val = "" if is_rev else str(sr)
            elif key == "expectedReplyDate":
                val = format_date(compute_expected_reply(row.get("issuedDate"), row.get("docNo"), rule=None, status=status_val, action=action_val))
            elif key == "duration":
                val = str(compute_duration(row.get("issuedDate"), row.get("actualReplyDate"), rule=None, status=status_val, action=action_val) or "")
            elif key in ("issuedDate","actualReplyDate"):
                val = format_date(row.get(key,""))
            elif pr_details_key and key == pr_details_key:
                val = _resolve_pr_details_value(row, pr_items_map, pr_details_key) or str(row.get(key,"") or "")
            else:
                val = str(row.get(key,"") or "")
            val = _format_multiline_display_value(key, c["label"], val)
            cells.append(Paragraph(xml_escape(val).replace("\n", "<br/>"), pstyle))
        data.append(cells)
        if ov:     row_meta.append(OV_COL)
        elif is_rev: row_meta.append(REV_COL)
        elif sr % 2 == 0: row_meta.append(ALT)
        else:      row_meta.append(rl_colors.white)
        if not is_rev: sr += 1

    # Build table style
    ts = [
        ("BACKGROUND", (0,0), (-1,0), DARK),
        ("TEXTCOLOR",  (0,0), (-1,0), rl_colors.white),
        ("FONTNAME",   (0,0), (-1,0), "Helvetica-Bold"),
        ("FONTSIZE",   (0,0), (-1,-1), 7),
        ("ROWBACKGROUND", (0,0), (-1,0), DARK),
        ("VALIGN",     (0,0), (-1,-1), "MIDDLE"),
        ("TOPPADDING", (0,0), (-1,-1), 2),
        ("BOTTOMPADDING",(0,0),(-1,-1), 2),
        ("LEFTPADDING",(0,0),(-1,-1), 2),
        ("RIGHTPADDING",(0,0),(-1,-1), 2),
        ("GRID",       (0,0), (-1,-1), 0.3, rl_colors.HexColor("#CBD5E1")),
    ]
    # Row backgrounds
    for i, bg in enumerate(row_meta, start=1):
        ts.append(("BACKGROUND", (0,i), (-1,i), bg))
    # Status column color
    status_idx = next((i for i,c in enumerate(hdr_cols) if c["col_key"]=="status"), None)
    if status_idx is not None:
        for i, row in enumerate(records, start=1):
            sv = row.get("status","")
            bg = STATUS_PDF.get(sv)
            if bg: ts.append(("BACKGROUND", (status_idx,i), (status_idx,i), bg))

    tbl = Table(data, colWidths=col_ws, repeatRows=1)
    tbl.setStyle(TableStyle(ts))

    proj_name = proj.get("name","") if isinstance(proj, dict) else ""
    title_st  = ParagraphStyle("t", fontName="Helvetica-Bold", fontSize=12,
                                textColor=DARK, spaceAfter=4)
    sub_st    = ParagraphStyle("s", fontName="Helvetica", fontSize=9,
                                textColor=rl_colors.HexColor("#64748B"), spaceAfter=8)

    story = [
        Paragraph(f"{dt['name'].upper()} — {proj_name}", title_st),
        Paragraph(f"Project: {proj.get('code','')}  |  Exported: {datetime.datetime.now().strftime('%d-%m-%Y %H:%M')}", sub_st),
        HRFlowable(width="100%", thickness=1, color=DARK, spaceAfter=6),
        tbl,
    ]
    doc.build(story)
    buf.seek(0)
    return buf


@exporting_bp.route("/api/export_pdf/<pid>/<dt_id>")
def api_export_pdf(pid, dt_id):
    proj = db.get_project(pid) or {}
    buf  = _build_pdf_for_dt(pid, dt_id, proj)
    fname = f"{proj.get('code','DCR')}_{dt_id}_Register.pdf"
    return send_file(buf, as_attachment=True, download_name=fname,
                     mimetype="application/pdf")


@exporting_bp.route("/api/export_pdf_all/<pid>")
def api_export_pdf_all(pid):
    from reportlab.platypus import SimpleDocTemplate, PageBreak
    proj = db.get_project(pid) or {}
    dts  = db.get_doc_types(pid)

    # Build each DT as separate PDF bytes then merge
    try:
        from pypdf import PdfWriter
        writer = PdfWriter()
        for dt in dts:
            if not db.count_records(pid, dt["id"]): continue
            single_buf = _build_pdf_for_dt(pid, dt["id"], proj)
            from pypdf import PdfReader
            reader = PdfReader(single_buf)
            for page in reader.pages:
                writer.add_page(page)
        out = io.BytesIO()
        writer.write(out); out.seek(0)
    except ImportError:
        # Fallback: just export first DT or give error
        out = _build_pdf_for_dt(pid, dts[0]["id"], proj) if dts else io.BytesIO()

    fname = f"{proj.get('code','DCR')}_All_Registers.pdf"
    return send_file(out, as_attachment=True, download_name=fname,
                     mimetype="application/pdf")


@exporting_bp.route("/api/import/<pid>/<dt_id>", methods=["POST"])
def api_import(pid, dt_id):
    if not can_edit(pid): return jsonify(error="LOGIN_REQUIRED"), 403
    data = request.get_json(silent=True) or {}
    b64  = data.get("file_b64","")
    ext  = data.get("ext","csv")
    if "," in b64: b64 = b64.split(",",1)[1]

    cols    = db.get_columns(pid, dt_id)
    col_map = {c["label"]: c["col_key"] for c in cols}
    imported = 0
    created = 0
    updated = 0
    skipped_blank = 0
    skipped_invalid = 0
    warnings = []

    try:
        if ext in ("xlsx","xls"):
            import openpyxl
            wb  = openpyxl.load_workbook(io.BytesIO(base64.b64decode(b64)), data_only=True)
            ws  = wb.active
            imported, created, updated, _, skipped_blank, skipped_invalid, warnings = _import_excel_worksheet(pid, dt_id, ws, cols)
        else:
            import csv, datetime as _dt
            text   = base64.b64decode(b64).decode("utf-8","ignore")
            header = None
            date_cols = {c["col_key"] for c in cols if c.get("col_type") in ("date","auto_date")}
            for row_idx, line in enumerate(csv.reader(io.StringIO(text)), start=1):
                if not header:
                    if any(c in line for c in ["Document No.","Sr.","docNo"]):
                        header = [col_map.get(h.strip(),h.strip()) for h in line]
                    continue
                if not any(line):
                    skipped_blank += 1
                    continue
                row_data = {}
                for i,val in enumerate(line):
                    if i<len(header) and header[i] and header[i] not in ("sr","Sr.","Sr",""):
                        v = val.strip()
                        if header[i] in date_cols and v:
                            for fmt in ("%d/%b/%Y","%d-%b-%Y","%Y-%m-%d","%d/%m/%Y"):
                                try: v = _dt.datetime.strptime(v,fmt).strftime("%Y-%m-%d"); break
                                except: pass
                        row_data[header[i]] = v
                if not _has_meaningful_values(row_data):
                    skipped_blank += 1
                    continue
                try:
                    action = _save_import_row(pid, dt_id, row_data)
                    imported += 1
                    if action == "updated":
                        updated += 1
                    else:
                        created += 1
                except Exception as e:
                    skipped_invalid += 1
                    logger.warning("import_csv_row_skipped pid=%s dt_id=%s row=%s error=%s",
                                   pid, dt_id, row_idx, e)
                    if len(warnings) < 20:
                        warnings.append({"row": row_idx, "error": str(e)})
    except Exception as e:
        logger.error("import_failed pid=%s dt_id=%s ext=%s error=%s", pid, dt_id, ext, e)
        return jsonify(ok=False, error=str(e)), 500

    return jsonify(
        ok=True,
        imported=imported,
        created=created,
        updated=updated,
        skipped_blank=skipped_blank,
        skipped_invalid=skipped_invalid,
        warnings=warnings,
    )


@exporting_bp.route("/api/import_project/<pid>", methods=["POST"])
def api_import_project(pid):
    if not can_edit(pid): return jsonify(error="LOGIN_REQUIRED"), 403
    data = request.get_json(silent=True) or {}
    b64  = data.get("file_b64","")
    ext  = data.get("ext","xlsx")
    if "," in b64: b64 = b64.split(",",1)[1]
    if ext not in ("xlsx","xls"):
        return jsonify(ok=False, error="Only Excel workbooks are supported"), 400

    try:
        import openpyxl

        wb = openpyxl.load_workbook(io.BytesIO(base64.b64decode(b64)), data_only=True)
        dts = db.get_doc_types(pid)
        imported_total = 0
        created_total = 0
        updated_total = 0
        skipped_blank_total = 0
        skipped_invalid_total = 0
        matched_sheets = []
        skipped_sheets = []
        errors = []
        warnings = []

        for ws in wb.worksheets:
            dt = _match_sheet_to_dt(ws.title, dts)
            if not dt:
                logger.warning("import_project_sheet_skipped pid=%s sheet=%s reason=no_matching_document_type",
                               pid, ws.title)
                skipped_sheets.append({"sheet": ws.title, "reason": "No matching document type"})
                continue
            try:
                cols = db.get_columns(pid, dt["id"])
                imported, created, updated, has_header, skipped_blank, skipped_invalid, sheet_warnings = _import_excel_worksheet(pid, dt["id"], ws, cols)
                if not has_header:
                    logger.warning("import_project_sheet_skipped pid=%s dt_id=%s sheet=%s reason=no_valid_header",
                                   pid, dt["id"], ws.title)
                    skipped_sheets.append({"sheet": ws.title, "reason": "No valid header row"})
                    continue
                imported_total += imported
                created_total += created
                updated_total += updated
                skipped_blank_total += skipped_blank
                skipped_invalid_total += skipped_invalid
                matched_sheets.append({
                    "sheet": ws.title,
                    "dt_id": dt["id"],
                    "dt_name": dt["name"],
                    "imported": imported,
                    "created": created,
                    "updated": updated,
                    "skipped_blank": skipped_blank,
                    "skipped_invalid": skipped_invalid,
                })
                if sheet_warnings and len(warnings) < 20:
                    for w in sheet_warnings:
                        warnings.append({"sheet": ws.title, **w})
                        if len(warnings) >= 20:
                            break
            except Exception as e:
                logger.error("import_project_sheet_failed pid=%s dt_id=%s sheet=%s error=%s",
                             pid, dt.get("id",""), ws.title, e)
                errors.append({"sheet": ws.title, "error": str(e)})

        return jsonify(
            ok=True,
            project_id=pid,
            imported_total=imported_total,
            created_total=created_total,
            updated_total=updated_total,
            skipped_blank_total=skipped_blank_total,
            skipped_invalid_total=skipped_invalid_total,
            matched_sheets=matched_sheets,
            skipped_sheets=skipped_sheets,
            errors=errors,
            warnings=warnings,
        )
    except Exception as e:
        logger.error("import_project_failed pid=%s ext=%s error=%s", pid, ext, e)
        return jsonify(ok=False, error=str(e)), 500


