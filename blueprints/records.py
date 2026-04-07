"""blueprints/records.py - DCR Records read-only API routes.

Handles the read-only records endpoint:
  GET /api/records/<pid>/<dt_id> -- list records for one doc type

Step 7A of the incremental refactor.
Logic is identical to the original app.py route — only the decorator
changed from @app.route to @records_bp.route.
"""
import uuid

from flask import Blueprint, jsonify, request

import db
from auth import current_user, can_edit
from utils import compute_expected_reply, compute_duration, is_overdue, format_date, extract_rev

records_bp = Blueprint("records", __name__)


@records_bp.route("/api/records/<pid>/<dt_id>")
def api_records(pid, dt_id):
    search  = request.args.get("search","")
    records = db.get_records(pid, dt_id, search=search)
    cols    = db.get_columns(pid, dt_id)
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
                    print(f"[WARN] expectedReply calc failed for rec {row.get('_id','')}: {e}")
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
    return jsonify(records=records, columns=cols, count=db.count_records(pid, dt_id))


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
