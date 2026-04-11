"""blueprints/records.py - DCR Records read-only API routes.

Handles the read-only records endpoint:
  GET /api/records/<pid>/<dt_id> -- list records for one doc type

Step 7A of the incremental refactor.
Logic is identical to the original app.py route — only the decorator
changed from @app.route to @records_bp.route.
"""
import logging
import uuid

from flask import Blueprint, jsonify, request

import db
from auth import current_user, can_edit
from utils import compute_expected_reply, compute_duration, is_overdue, format_date, extract_rev

records_bp = Blueprint("records", __name__)
logger = logging.getLogger(__name__)

def _is_pr_doc_type(pid, dt_id):
    if str(dt_id or "").upper() == "PR":
        return True
    dts = db.get_doc_types(pid)
    dt = next((d for d in dts if d["id"] == dt_id), None)
    if not dt:
        return False
    code = str(dt.get("code","")).strip().upper()
    name = str(dt.get("name","")).strip().lower()
    return code == "PR" or "requisition" in name or "purchase request" in name


@records_bp.route("/api/records/<pid>/<dt_id>")
def api_records(pid, dt_id):
    search  = request.args.get("search","")
    records = db.get_records(pid, dt_id, search=search)
    cols    = db.get_columns(pid, dt_id)
    pr_items_map = db.get_pr_items_for_records([r.get("_id") for r in records]) if _is_pr_doc_type(pid, dt_id) else {}
    date_col_keys = {c["col_key"] for c in cols if c.get("col_type") in ("date","auto_date")}
    has_exp_reply = any(c["col_key"] == "expectedReplyDate" for c in cols)
    has_status    = any(c["col_key"] == "status" for c in cols)
    for row in records:
        if has_exp_reply:
            issued_date = row.get("issuedDate")
            doc_no = row.get("docNo")
            exp = None
            if issued_date and doc_no:
                try:
                    exp = compute_expected_reply(issued_date, doc_no)
                except Exception as e:
                    logger.warning("expected_reply_calc_failed pid=%s dt_id=%s record_id=%s error=%s",
                                   pid, dt_id, row.get("_id",""), e)
            row["_expectedReplyDate"] = format_date(exp) if exp else ""
        else:
            row["_expectedReplyDate"] = ""
        issued   = row.get("issuedDate","")
        actual   = row.get("actualReplyDate","")
        dur = compute_duration(issued, actual)
        row["_duration"] = str(dur) if dur is not None else ""
        row["_overdue"]    = is_overdue(row.get("issuedDate"), row.get("docNo"), row.get("actualReplyDate"), has_exp_reply)
        row["_isRev"]      = extract_rev(row.get("docNo","")) > 0
        # Format ALL date columns (any col_type=date)
        for dk in date_col_keys:
            if dk in row and row[dk]:
                row["_fmt_" + dk] = format_date(row[dk])
        # Standard aliases
        row["_issuedFmt"]  = format_date(row.get("issuedDate",""))
        row["_replyFmt"]   = format_date(row.get("actualReplyDate",""))
    return jsonify(records=records, columns=cols, count=db.count_records(pid, dt_id), pr_items_map=pr_items_map)


@records_bp.route("/api/records/<pid>/<dt_id>", methods=["POST"])
def api_save_record(pid, dt_id):
    if not can_edit(pid): return jsonify(error="LOGIN_REQUIRED"), 403
    u      = current_user()
    uname  = u["username"] if u else "unknown"
    data   = request.get_json(silent=True) or {}
    rec_id = data.pop("_id", None)
    clean  = {k:v for k,v in data.items() if not k.startswith("_")}
    if rec_id:
        full_row = db.get_record_by_id(rec_id) or {}
        # Extract only stored document fields for diff (exclude _ meta keys)
        old_data = {k: v for k, v in full_row.items() if not k.startswith("_")}
        db.save_record(pid, dt_id, rec_id, clean)
        for field, new_val in clean.items():
            old_val = old_data.get(field,"")
            if str(old_val).strip() != str(new_val or "").strip():
                db.log_action(uname,"EDIT",pid,dt_id,rec_id,
                    clean.get("docNo",""),field,
                    str(old_val)[:200],str(new_val)[:200])
    else:
        rec_id = str(uuid.uuid4())
        db.save_record(pid, dt_id, rec_id, clean)
        db.log_action(uname,"ADD",pid,dt_id,rec_id,
            clean.get("docNo",""),detail="New document added")
    return jsonify(ok=True, id=rec_id)


@records_bp.route("/api/records/<rec_id>", methods=["DELETE"])
def api_delete_record(rec_id):
    u = current_user()
    if not u:
        return jsonify(error="LOGIN_REQUIRED"), 403
    full_row = db.get_record_by_id(rec_id)
    if not full_row:
        return jsonify(error="Not found"), 404
    pid   = full_row.get("_project_id", "")
    dt_id = full_row.get("_dt_id", "")
    doc_no = full_row.get("docNo", "")
    if not can_edit(pid):
        return jsonify(error="Forbidden"), 403
    db.delete_record(rec_id)
    db.log_action(
        u["username"], "DELETE",
        pid or None, dt_id or None, rec_id,
        doc_no, detail=f"Deleted: {doc_no}"
    )
    return jsonify(ok=True)


@records_bp.route("/api/records/bulk_delete", methods=["POST"])
def api_bulk_delete_records():
    u = current_user()
    if not u:
        return jsonify(error="LOGIN_REQUIRED"), 403
    data = request.get_json(silent=True) or {}
    ids = data.get("ids") if isinstance(data, dict) else None
    if not isinstance(ids, list):
        return jsonify(ok=False, error="Invalid ids"), 400
    ids = [str(i).strip() for i in ids if str(i).strip()]
    if not ids:
        return jsonify(ok=True, deleted=0)

    recs = db.get_records_meta(ids)
    if not recs:
        return jsonify(ok=True, deleted=0)

    blocked = next((r for r in recs if not can_edit(r.get("project_id", ""))), None)
    if blocked:
        return jsonify(error="Forbidden"), 403

    deleted = db.delete_records_bulk([r["id"] for r in recs])
    by_scope = {}
    for r in recs:
        key = (r.get("project_id") or None, r.get("dt_id") or None)
        by_scope.setdefault(key, []).append(r.get("doc_no") or r.get("id"))
    for (pid, dt_id), doc_nos in by_scope.items():
        db.log_action(
            u["username"], "DELETE",
            pid, dt_id, None, "",
            detail=f"Bulk deleted {len(doc_nos)} record(s)"
        )
    return jsonify(ok=True, deleted=deleted)


@records_bp.route("/api/pr_items/<record_id>")
def api_get_pr_items(record_id):
    full_row = db.get_record_by_id(record_id)
    if not full_row:
        return jsonify(error="Not found"), 404
    pid   = full_row.get("_project_id", "")
    dt_id = full_row.get("_dt_id", "")
    if not _is_pr_doc_type(pid, dt_id):
        return jsonify(error="Not a PR document"), 400
    items = db.get_pr_items(record_id)
    return jsonify(ok=True, items=items)


@records_bp.route("/api/pr_items/<record_id>", methods=["POST","PUT"])
def api_save_pr_items(record_id):
    full_row = db.get_record_by_id(record_id)
    if not full_row:
        return jsonify(error="Not found"), 404
    pid   = full_row.get("_project_id", "")
    dt_id = full_row.get("_dt_id", "")
    if not _is_pr_doc_type(pid, dt_id):
        return jsonify(error="Not a PR document"), 400
    if not can_edit(pid): return jsonify(error="LOGIN_REQUIRED"), 403
    data = request.get_json(silent=True) or {}
    items = data.get("items") if isinstance(data, dict) else data
    if not isinstance(items, list):
        return jsonify(ok=False, error="Invalid items"), 400
    clean = []
    from decimal import Decimal, InvalidOperation
    for it in items:
        if not isinstance(it, dict):
            continue
        row_type = str(it.get("row_type", "item")).strip().lower()
        if row_type not in ("header", "item"):
            row_type = "item"
        item_name = str(it.get("item_name","")).strip()
        unit = str(it.get("unit","")).strip() if it.get("unit") is not None else ""
        remarks = str(it.get("remarks","")).strip() if it.get("remarks") is not None else ""
        if row_type == "header":
            if not item_name:
                continue
            clean.append({
                "row_type": "header",
                "item_name": item_name,
                "unit": None,
                "quantity": None,
                "remarks": None,
            })
            continue
        qty_raw = it.get("quantity", "")
        qty = None
        if qty_raw not in (None, ""):
            try:
                qty = Decimal(str(qty_raw).strip())
            except (InvalidOperation, ValueError):
                qty = None
        if not item_name and not unit and not remarks and qty is None:
            continue
        if not item_name:
            continue
        clean.append({
            "row_type": "item",
            "item_name": item_name,
            "unit": unit or None,
            "quantity": qty,
            "remarks": remarks or None,
        })
    saved = db.save_pr_items(record_id, clean)
    return jsonify(ok=True, saved=saved)
