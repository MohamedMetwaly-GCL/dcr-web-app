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


def _import_excel_worksheet(pid, dt_id, ws, cols):
    import datetime as _dt

    col_map = {c["label"]: c["col_key"] for c in cols}
    header = None
    imported = 0
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
            db.save_record(pid, dt_id, str(uuid.uuid4()), row_data)
            imported += 1
        except Exception as e:
            skipped_invalid += 1
            logger.warning("import_row_skipped pid=%s dt_id=%s row=%s error=%s",
                           pid, dt_id, row_idx, e)
            if len(warnings) < 20:
                warnings.append({"row": row_idx, "error": str(e)})
    return imported, header is not None, skipped_blank, skipped_invalid, warnings


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


def _build_pr_excel(record, proj, pr_items, pr_details):
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.worksheet.table import Table, TableStyleInfo

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

    for col, width in {"A": 6, "B": 46, "C": 12, "D": 10, "E": 24}.items():
        ws.column_dimensions[col].width = width
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.print_options.horizontalCentered = False
    ws.oddFooter.right.text = "Page &[Page] of &[Pages]"
    ws.oddFooter.left.text = "Generated from DCR System"

    ws.merge_cells("A1:E1")
    c = ws["A1"]
    c.value = "PURCHASE REQUISITION"
    style_cell(
        c,
        font=Font(name="Arial", size=16, bold=True, color=PRIMARY),
        fill_color=LIGHT,
        border=thin,
        align=Alignment(horizontal="center", vertical="center"),
    )
    ws.row_dimensions[1].height = 30

    ws.merge_cells("A2:E2")
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
    ws.merge_cells("A3:E3")
    c = ws["A3"]
    c.value = ""
    c.fill = fill(WHITE)
    c.border = Border(bottom=thin_side)
    ws.row_dimensions[3].height = 6

    pr_number = _pick_first(record, ("docNo", "prNo", "prNumber")) or f"PR-{str(record.get('_id',''))[:8]}"
    pr_date = format_date(_pick_first(record, ("issuedDate", "prDate", "date")))
    discipline = _pick_first(record, ("discipline",))
    trade = _pick_first(record, ("trade",))
    disc_trade = " / ".join([v for v in [discipline, trade] if v])
    requested_by = _pick_first(record, ("requestedBy", "requester", "preparedBy", "prepared_by", "requested_by"))

    meta_fields = [
        ("Project Name", proj.get("name", "")),
        ("Project Code", proj.get("code", "")),
        ("PR Number", pr_number),
        ("PR Date", pr_date),
        ("Discipline / Trade", disc_trade),
        ("Requested By", requested_by),
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
            ws[f"D{row_no}"].border = thin
            ws[f"E{row_no}"].border = thin
        ws.row_dimensions[row_no].height = 20
        row_no += 1

    ws.merge_cells(f"A{row_no}:E{row_no}")
    c = ws[f"A{row_no}"]
    c.value = "PR DETAILS"
    style_cell(
        c,
        font=Font(name="Arial", size=11, bold=True, color=WHITE),
        fill_color=PRIMARY,
        border=thin,
        align=Alignment(horizontal="left", vertical="center"),
    )
    ws.row_dimensions[row_no].height = 20
    row_no += 1

    detail_lines = max(3, min(8, len(str(pr_details or "").splitlines()) + 1))
    ws.merge_cells(f"A{row_no}:E{row_no}")
    c = ws[f"A{row_no}"]
    c.value = pr_details or ""
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
    headers = ["No.", "Item / Description", "Unit", "Qty", "Remarks"]
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

    item_no = 1
    data_row = table_header_row + 1
    if pr_items:
        for it in pr_items:
            row_type = str(it.get("row_type", "item") or "item").strip().lower()
            if row_type == "header":
                ws.merge_cells(start_row=data_row, start_column=1, end_row=data_row, end_column=5)
                cell = ws.cell(row=data_row, column=1, value=str(it.get("item_name", "") or "").strip())
                style_cell(
                    cell,
                    font=Font(name="Arial", size=11, bold=True, color=PRIMARY),
                    fill_color=SECTION,
                    border=thin,
                    align=Alignment(vertical="center", wrap_text=True),
                )
                for col in range(2, 6):
                    ws.cell(row=data_row, column=col).border = thin
                    ws.cell(row=data_row, column=col).fill = fill(SECTION)
                ws.row_dimensions[data_row].height = 24
            else:
                values = [
                    item_no,
                    str(it.get("item_name", "") or "").strip(),
                    str(it.get("unit", "") or "").strip(),
                    it.get("quantity", ""),
                    str(it.get("remarks", "") or "").strip(),
                ]
                for col, val in enumerate(values, start=1):
                    cell = ws.cell(row=data_row, column=col, value=val)
                    style_cell(
                        cell,
                        font=Font(name="Arial", size=10, color=TEXT),
                        fill_color=WHITE if item_no % 2 else SUBTLE,
                        border=thin,
                        align=Alignment(
                            horizontal="center" if col in (1, 3, 4) else "left",
                            vertical="top",
                            wrap_text=True,
                        ),
                    )
                ws.row_dimensions[data_row].height = 20
                item_no += 1
            data_row += 1
    else:
        ws.merge_cells(start_row=data_row, start_column=1, end_row=data_row, end_column=5)
        c = ws.cell(row=data_row, column=1, value="No PR items")
        c.font = Font(name="Arial", size=10, italic=True, color=MUTED)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = thin
        ws.row_dimensions[data_row].height = 20
        data_row += 1

    sig_row = data_row + 2
    signatures = [("Prepared By", "A", "B"), ("Reviewed By", "C", "D"), ("Approved By", "E", "E")]
    for label, start_col, end_col in signatures:
        ws.merge_cells(f"{start_col}{sig_row}:{end_col}{sig_row}")
        top = ws[f"{start_col}{sig_row}"]
        top.value = ""
        top.border = top_rule
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
    ws.merge_cells(f"A{footer_row}:E{footer_row}")
    c = ws[f"A{footer_row}"]
    c.value = f"Generated from DCR System | Export date: {datetime.datetime.now().strftime('%d-%m-%Y %H:%M')}"
    c.font = Font(name="Arial", size=8, italic=True, color=MUTED)
    c.alignment = Alignment(horizontal="right")

    raw_headers = ["record_id", "sort_order", "row_type", "description", "unit", "qty", "remarks"]
    for idx, label in enumerate(raw_headers, start=1):
        c = raw_ws.cell(row=1, column=idx, value=label)
        c.font = Font(name="Arial", size=10, bold=True)
        c.fill = fill(LABEL)
        c.border = thin
        c.alignment = Alignment(horizontal="center")
    for col, width in {1: 40, 2: 12, 3: 12, 4: 46, 5: 12, 6: 10, 7: 24}.items():
        raw_ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = width
    raw_ws.freeze_panes = "A2"
    raw_ws.auto_filter.ref = f"A1:G{max(2, len(pr_items) + 1)}"

    for idx, it in enumerate(pr_items, start=2):
        row_type = str(it.get("row_type", "item") or "item").strip().lower()
        values = [
            record.get("_id", ""),
            it.get("sort_order", idx - 2),
            row_type,
            str(it.get("item_name", "") or "").strip(),
            "" if row_type == "header" else str(it.get("unit", "") or "").strip(),
            "" if row_type == "header" else it.get("quantity", ""),
            "" if row_type == "header" else str(it.get("remarks", "") or "").strip(),
        ]
        for col, val in enumerate(values, start=1):
            c = raw_ws.cell(row=idx, column=col, value=val)
            c.border = thin
            c.alignment = Alignment(vertical="top", wrap_text=True)

    if pr_items:
        raw_tbl = Table(displayName="PRItemsRaw", ref=f"A1:G{len(pr_items)+1}")
        raw_tbl.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium2",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False,
        )
        raw_ws.add_table(raw_tbl)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


@exporting_bp.route("/api/export_pr/<record_id>")
def api_export_pr(record_id):
    full_row = db.get_record_by_id(record_id)
    if not full_row:
        return jsonify(error="Not found"), 404

    pid = full_row.get("_project_id", "")
    dt_id = full_row.get("_dt_id", "")
    dts = db.get_doc_types(pid)
    dt = next((d for d in dts if d["id"] == dt_id), None)
    if not _is_pr_dt(dt):
        return jsonify(error="Not a PR document"), 400

    proj = db.get_project(pid) or {}
    cols = [c for c in db.get_columns(pid, dt_id) if c["visible"]]
    pr_details_key = _pr_details_key(cols)
    pr_items = db.get_pr_items(record_id)
    pr_details = _resolve_pr_details_value(full_row, {record_id: pr_items}, pr_details_key)

    if current_user():
        db.log_action(current_user()["username"], "EXPORT_EXCEL", pid, dt_id, record_id, full_row.get("docNo", ""), detail="Export PR workbook")

    pr_number = _pick_first(full_row, ("docNo", "prNo", "prNumber")) or f"PR-{record_id[:8]}"
    fname = f"PR_{_safe_excel_name_part(proj.get('code','DCR'), 'DCR')}_{_safe_excel_name_part(pr_number, 'PR_Record')}.xlsx"
    buf = _build_pr_excel(full_row, proj, pr_items, pr_details)
    return send_file(buf, as_attachment=True, download_name=fname,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@exporting_bp.route("/api/export_all/<pid>")
def api_export_all(pid):
    u = current_user()
    if u: db.log_action(u["username"],"EXPORT_EXCEL",pid,detail="Export All")
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

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
        records = db.get_records(pid, dt["id"])
        is_pr = _is_pr_dt(dt)
        pr_details_key = _pr_details_key(cols) if is_pr else None
        pr_items_map = db.get_pr_items_for_records([r.get("_id") for r in records]) if is_pr else {}
        # always include tab (even if empty)

        ws = wb.create_sheet(title=dt["id"][:31])
        ws.sheet_view.showGridLines = False
        all_cols = [{"col_key":"_sr","label":"Sr."}] + [{"col_key":c["col_key"],"label":c["label"]} for c in cols]
        nc = len(all_cols)

        from openpyxl.utils import get_column_letter as gcl
        def mcell(row, val, bg, fg="FFFFFF", bold=False, sz=11):
            c = ws.cell(row=row, column=1, value=val)
            ws.merge_cells(f"A{row}:{gcl(nc)}{row}")
            c.font = Font(bold=bold, color=fg, size=sz, name="Arial")
            c.fill = fill(bg)
            c.alignment = Alignment(horizontal="left", vertical="center")
            return c

        mcell(1, f"{dt['name'].upper()} — {proj.get('name','')} ({proj.get('code','')})",
              PRIMARY, bold=True, sz=12)
        ws.row_dimensions[1].height = 30
        ws.row_dimensions[2].height = 4

        COL_W = {"_sr":5,"docNo":22,"discipline":14,"trade":14,"title":38,"floor":12,
                 "itemRef":16,"issuedDate":13,"expectedReplyDate":14,"actualReplyDate":13,
                 "status":24,"duration":10,"remarks":28,"fileLocation":18}
        CENTER = {"_sr","duration","issuedDate","expectedReplyDate","actualReplyDate"}

        for ci, col in enumerate(all_cols, 1):
            ws.column_dimensions[gcl(ci)].width = COL_W.get(col["col_key"],13)
            c = ws.cell(row=3, column=ci, value=col["label"])
            c.font = Font(bold=True,color=WHITE,size=10,name="Arial")
            c.fill = fill(PRIMARY)
            c.alignment = Alignment(horizontal="center",vertical="center",wrap_text=True)
            c.border = thin()
        ws.row_dimensions[3].height = 22

        sr = 1
        if not records:
            # Empty tab — write a "no records" row
            c = ws.cell(row=4, column=1, value="No records in this register")
            c.font = Font(italic=True, size=10, name="Arial", color="9CA3AF")
            ws.merge_cells(f"A4:{gcl(nc)}4")
            c.alignment = Alignment(horizontal="center", vertical="center")
            ws.row_dimensions[4].height = 24
        has_exp_col = any(c["col_key"]=="expectedReplyDate" for c in all_cols)
        for ri, row in enumerate(records):
            rn = 4 + ri
            is_rev = extract_rev(row.get("docNo","")) > 0
            ov     = is_overdue(row.get("issuedDate"), row.get("docNo"), row.get("actualReplyDate"), has_exp_col)
            bg     = OV if ov else (ALT if sr%2==0 else WHITE)
            ws.row_dimensions[rn].height = 18
            for ci, col in enumerate(all_cols, 1):
                key = col["col_key"]
                if key=="_sr":                   val = "" if is_rev else str(sr)
                elif key=="expectedReplyDate":   val = format_date(compute_expected_reply(row.get("issuedDate"),row.get("docNo")))
                elif key=="duration":
                    dur_val = compute_duration(row.get("issuedDate"),row.get("actualReplyDate"))
                    val = str(dur_val) if dur_val is not None else ""
                elif key in ("issuedDate","actualReplyDate"): val = format_date(row.get(key,""))
                elif pr_details_key and key == pr_details_key:
                    val = _resolve_pr_details_value(row, pr_items_map, pr_details_key) or str(row.get(key,"") or "")
                else:                            val = str(row.get(key,"") or "")
                if key == "fileLocation" and val and val.startswith("http"):
                    c = ws.cell(row=rn, column=ci)
                    c.value = "View"
                    c.hyperlink = val
                    c.font = Font(size=9,name="Arial",color="2563A8",underline="single")
                else:
                    c = ws.cell(row=rn, column=ci, value=val)
                    if key == "duration" and val == "0":
                        c.value = 0
                c.border = thin()
                c.alignment = Alignment(vertical="center", wrap_text=True,
                                        horizontal="center" if key in CENTER else "left")
                if key=="status" and val:
                    bg2, fg2 = STATUS_XL.get(val, ("F3F4F6","374151"))
                    c.fill = fill(bg2); c.font = Font(bold=True,size=9,name="Arial",color=fg2)
                elif key != "fileLocation" or not (val and val.startswith("http")):
                    c.fill = fill(bg)
                    c.font = Font(size=10,name="Arial",
                                  color=MUTED if is_rev else ("991B1B" if ov else "1E2A3A"))
            if not is_rev: sr += 1

    if not wb.sheetnames:
        ws = wb.create_sheet("Empty"); ws.cell(1,1,"No data")

    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    fname = f"{proj.get('code','DCR')}_All_Registers.xlsx"
    return send_file(buf, as_attachment=True, download_name=fname,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@exporting_bp.route("/api/export/<pid>/<dt_id>")
def api_export(pid, dt_id):
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    proj    = db.get_project(pid) or {}
    cols    = [c for c in db.get_columns(pid, dt_id) if c["visible"]]
    records = db.get_records(pid, dt_id)
    dts     = db.get_doc_types(pid)
    dt      = next((d for d in dts if d["id"] == dt_id), None)
    is_pr = _is_pr_dt(dt)
    pr_details_key = _pr_details_key(cols) if is_pr else None
    pr_items_map = db.get_pr_items_for_records([r.get("_id") for r in records]) if is_pr else {}

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

    wb = openpyxl.Workbook(); ws = wb.active
    ws.title = dt_id; ws.sheet_view.showGridLines = False

    all_cols = [{"col_key":"_sr","label":"Sr."}] + [{"col_key":c["col_key"],"label":c["label"]} for c in cols]
    nc = len(all_cols)

    def mcell(row, val, bg, fg="FFFFFF", bold=False, sz=11, halign="center"):
        c = ws.cell(row=row, column=1, value=val)
        ws.merge_cells(f"A{row}:{get_column_letter(nc)}{row}")
        c.font = Font(bold=bold, color=fg, size=sz, name="Arial")
        c.fill = fill(bg)
        c.alignment = Alignment(horizontal=halign, vertical="center")
        return c

    mcell(1, f"DOCUMENT CONTROL REGISTER  —  {(dt['name'] if dt else dt_id).upper()}",
          PRIMARY, bold=True, sz=13)
    ws.row_dimensions[1].height = 36

    info = "   |   ".join(f"{k}: {v}" for k,v in [
        ("Project",proj.get("name","")),("Code",proj.get("code","")),
        ("Client",proj.get("client","")),("Consultant",proj.get("mainConsultant","")),
        ("Exported",datetime.datetime.now().strftime("%d/%b/%Y %H:%M"))] if v)
    mcell(2, info, PL, sz=9); ws.row_dimensions[2].height=18
    ws.row_dimensions[3].height = 4

    COL_W = {"_sr":5,"docNo":22,"discipline":14,"trade":14,"title":38,"floor":12,
             "itemRef":16,"issuedDate":13,"expectedReplyDate":14,"actualReplyDate":13,
             "status":24,"duration":10,"remarks":28,"fileLocation":18}
    CENTER = {"_sr","duration","issuedDate","expectedReplyDate","actualReplyDate"}

    for ci, col in enumerate(all_cols, 1):
        ws.column_dimensions[get_column_letter(ci)].width = COL_W.get(col["col_key"],13)
        c = ws.cell(row=4, column=ci, value=col["label"])
        c.font = Font(bold=True,color=WHITE,size=10,name="Arial")
        c.fill = fill(PRIMARY)
        c.alignment = Alignment(horizontal="center",vertical="center",wrap_text=True)
        c.border = thin()
    ws.row_dimensions[4].height = 24

    has_exp_col_s = any(c["col_key"]=="expectedReplyDate" for c in cols)
    sr = 1
    for ri, row in enumerate(records):
        rn   = 5 + ri
        is_rev = extract_rev(row.get("docNo","")) > 0
        ov     = is_overdue(row.get("issuedDate"), row.get("docNo"), row.get("actualReplyDate"), has_exp_col_s)
        bg     = OV if ov else (ALT if sr%2==0 else WHITE)
        ws.row_dimensions[rn].height = 20

        for ci, col in enumerate(all_cols, 1):
            key = col["col_key"]
            if key=="_sr":                  val = "" if is_rev else str(sr)
            elif key=="expectedReplyDate":  val = format_date(compute_expected_reply(row.get("issuedDate"),row.get("docNo")))
            elif key=="duration":
                dur_val = compute_duration(row.get("issuedDate"),row.get("actualReplyDate"))
                val = str(dur_val) if dur_val is not None else ""
            elif key=="issuedDate":         val = format_date(row.get(key,""))
            elif key=="actualReplyDate":    val = format_date(row.get(key,""))
            elif pr_details_key and key == pr_details_key:
                val = _resolve_pr_details_value(row, pr_items_map, pr_details_key) or str(row.get(key,"") or "")
            else:                           val = str(row.get(key,"") or "")

            # fileLocation → hyperlink "View"
            if key == "fileLocation" and val and val.startswith("http"):
                c = ws.cell(row=rn, column=ci)
                c.value = "View"
                c.hyperlink = val
                c.font = Font(size=9,name="Arial",color="2563A8",underline="single")
            else:
                c = ws.cell(row=rn, column=ci, value=val)
                # Duration 0 should show as number not empty
                if key == "duration" and val == "0":
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
            c.alignment = Alignment(vertical="center", wrap_text=True,
                                    horizontal="center" if key in CENTER else "left")
        if not is_rev: sr += 1

    tot = 5 + len(records)
    real = sum(1 for r in records if extract_rev(r.get("docNo",""))==0)
    mcell(tot, f"TOTAL: {real} documents  |  {len(records)} submissions", PRIMARY, sz=10, halign="left")
    ws.row_dimensions[tot].height = 22

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
        ov     = is_overdue(row.get("issuedDate"), row.get("docNo"), row.get("actualReplyDate"), has_exp_col_pdf)
        cells  = []
        for c in hdr_cols:
            key = c["col_key"]
            if key == "_sr":
                val = "" if is_rev else str(sr)
            elif key == "expectedReplyDate":
                val = format_date(compute_expected_reply(row.get("issuedDate"), row.get("docNo")))
            elif key == "duration":
                val = str(compute_duration(row.get("issuedDate"), row.get("actualReplyDate")) or "")
            elif key in ("issuedDate","actualReplyDate"):
                val = format_date(row.get(key,""))
            elif pr_details_key and key == pr_details_key:
                val = _resolve_pr_details_value(row, pr_items_map, pr_details_key) or str(row.get(key,"") or "")
            else:
                val = str(row.get(key,"") or "")
            cells.append(Paragraph(val, pstyle))
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
    skipped_blank = 0
    skipped_invalid = 0
    warnings = []

    try:
        if ext in ("xlsx","xls"):
            import openpyxl
            wb  = openpyxl.load_workbook(io.BytesIO(base64.b64decode(b64)), data_only=True)
            ws  = wb.active
            imported, _, skipped_blank, skipped_invalid, warnings = _import_excel_worksheet(pid, dt_id, ws, cols)
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
                    db.save_record(pid, dt_id, str(uuid.uuid4()), row_data)
                    imported += 1
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
                imported, has_header, skipped_blank, skipped_invalid, sheet_warnings = _import_excel_worksheet(pid, dt["id"], ws, cols)
                if not has_header:
                    logger.warning("import_project_sheet_skipped pid=%s dt_id=%s sheet=%s reason=no_valid_header",
                                   pid, dt["id"], ws.title)
                    skipped_sheets.append({"sheet": ws.title, "reason": "No valid header row"})
                    continue
                imported_total += imported
                skipped_blank_total += skipped_blank
                skipped_invalid_total += skipped_invalid
                matched_sheets.append({
                    "sheet": ws.title,
                    "dt_id": dt["id"],
                    "dt_name": dt["name"],
                    "imported": imported,
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
