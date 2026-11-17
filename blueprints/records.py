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
from auth import (
    current_user,
    can_edit,
    can_view_project,
    get_allowed_project_ids,
    get_user_doc_type_scope,
    is_superadmin,
    can_access_doc_type,
    _resolve_project,
)
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

def _is_ltr_doc_type(pid, dt_id):
    if str(dt_id or "").upper() == "LTR":
        return True
    dts = db.get_doc_types(pid)
    dt = next((d for d in dts if d["id"] == dt_id), None)
    if not dt:
        return False
    code = str(dt.get("code","")).strip().upper()
    name = str(dt.get("name","")).strip().lower()
    return code == "LTR" or "letter" in name or "correspondence" in name

def _is_mhr_doc_type(pid, dt_id):
    if str(dt_id or "").strip().upper() == "MHR":
        return True
    dt = db.get_doc_type(pid, dt_id)
    if not dt:
        return False
    return str(dt.get("code") or "").strip().upper() == "MHR"

def _norm_ltr_text(v):
    return "".join(ch for ch in str(v or "").strip().lower() if ch.isalnum())

def _ltr_field_key(pid, dt_id, role):
    cols = db.get_columns(pid, dt_id)
    wanted = {
        "docNo": {"docno", "letterref"},
        "title": {"title", "subject"},
        "description": {"description"},
        "fileLocation": {"filelocation", "attachment"},
        "remarks": {"remarks"},
        "direction": {"direction"},
        "fromParty": {"fromparty", "from"},
        "toParty": {"toparty", "to"},
        "issuedDate": {"issueddate"},
        "receivedDate": {"receiveddate"},
        "parentLetterId": {"parentletterid"},
        "parentLetterRef": {"parentletterref", "responseref", "parentletter"},
        "status": {"status"},
    }.get(role, set())
    for c in cols:
        key = _norm_ltr_text(c.get("col_key"))
        label = _norm_ltr_text(c.get("label"))
        if key in wanted or label in wanted:
            return c.get("col_key")
    return role


@records_bp.route("/api/records/search/global")
def api_search_global():
    u = current_user()
    if not u:
        return jsonify(error="LOGIN_REQUIRED"), 403

    q_str = str(request.args.get("q", "")).strip()
    dt_id = str(request.args.get("dt_id", "")).strip()
    project_filter = str(request.args.get("project_id", "")).strip()
    status_filter = str(request.args.get("status", "")).strip()
    date_from = str(request.args.get("date_from", "")).strip()
    date_to = str(request.args.get("date_to", "")).strip()
    try:
        limit = min(int(request.args.get("limit", 300)), 500)
        limit = max(1, limit)
    except (ValueError, TypeError):
        limit = 300

    ws_filter = str(request.args.get("workspace_id", "")).strip() or None
    user_projs = db.get_projects_for_user(u.get("username", ""), role=u.get("role"), workspace_id=ws_filter)

    allowed_ids = []
    seen_pids = set()
    for p in user_projs:
        canon_pid = str(p.get("id") or "").strip()
        if canon_pid and canon_pid not in seen_pids:
            seen_pids.add(canon_pid)
            allowed_ids.append(canon_pid)

    is_sa = str(u.get("role", "")).strip().lower() == "superadmin"
    user_scope = "all" if is_sa and not ws_filter else "filtered"

    if not allowed_ids or len(q_str) < 2:
        return jsonify(query=q_str, total_found=0, user_scope=user_scope, results=[])

    if project_filter and project_filter.upper() != "ALL":
        proj_filter = _resolve_project(project_filter)
        canon_filter = proj_filter["id"] if proj_filter else None
        if not canon_filter or canon_filter not in seen_pids:
            return jsonify(query=q_str, total_found=0, user_scope=user_scope, results=[])
        target_pids = [canon_filter]
    else:
        target_pids = allowed_ids

    dt_scope = get_user_doc_type_scope(u, target_pids)
    scope_branches = []
    scope_params = []
    if dt_scope is None:
        scope_branches.append("p.id = ANY(%s)")
        scope_params.append(target_pids)
    else:
        unrestricted_pids = []
        for pid in target_pids:
            allowed_dts = dt_scope.get(pid)
            if allowed_dts is None:
                unrestricted_pids.append(pid)
            elif allowed_dts:
                scope_branches.append("(p.id = %s AND UPPER(r.dt_id) = ANY(%s))")
                scope_params.extend([pid, sorted(allowed_dts)])
        if unrestricted_pids:
            scope_branches.insert(0, "p.id = ANY(%s)")
            scope_params.insert(0, unrestricted_pids)

    if not scope_branches:
        return jsonify(query=q_str, total_found=0, user_scope=user_scope, results=[])

    like_param = f"%{q_str}%"
    where_clauses = [
        "(" + " OR ".join(scope_branches) + ")",
        """(
            r.data->>'docNo' ILIKE %s
            OR r.data->>'title' ILIKE %s
            OR r.data->>'nocSubject' ILIKE %s
            OR r.data->>'nocDescription' ILIKE %s
            OR r.data::text ILIKE %s
        )"""
    ]
    params = scope_params + [like_param, like_param, like_param, like_param, like_param]

    if dt_id and dt_id.upper() != "ALL":
        where_clauses.append("UPPER(r.dt_id) = UPPER(%s)")
        params.append(dt_id)

    if status_filter and status_filter.upper() != "ALL":
        if "REVISE" in status_filter.upper():
            st_pattern = "%Revise%"
        elif "UNDER REVIEW" in status_filter.upper():
            st_pattern = "%Review%"
        elif "REJECT" in status_filter.upper():
            st_pattern = "%Reject%"
        elif "APPROV" in status_filter.upper():
            st_pattern = "%Approv%"
        else:
            st_pattern = f"%{status_filter}%"
        where_clauses.append("(COALESCE(r.data->>'status', r.data->>'partBStatus', r.data->>'partDStatus', '') ILIKE %s)")
        params.append(st_pattern)

    safe_issued_date = r"""(CASE 
        WHEN (r.data->>'issuedDate') ~ '^\d{4}-\d{2}-\d{2}' THEN SUBSTRING(r.data->>'issuedDate', 1, 10)::date
        WHEN (r.data->>'partAIssueDate') ~ '^\d{4}-\d{2}-\d{2}' THEN SUBSTRING(r.data->>'partAIssueDate', 1, 10)::date
        ELSE NULL 
    END)"""

    import re
    if date_from and re.match(r"^\d{4}-\d{2}-\d{2}$", date_from):
        where_clauses.append(f"(r.created_at::date >= %s::date OR {safe_issued_date} >= %s::date)")
        params.extend([date_from, date_from])

    if date_to and re.match(r"^\d{4}-\d{2}-\d{2}$", date_to):
        where_clauses.append(f"(r.created_at::date <= %s::date OR {safe_issued_date} <= %s::date)")
        params.extend([date_to, date_to])

    where_sql = " AND ".join(where_clauses)
    sql = f"""
        SELECT 
            r.id,
            r.project_id,
            COALESCE(p.code, '') AS project_code,
            COALESCE(p.name, '') AS project_name,
            r.dt_id,
            COALESCE(r.data->>'docNo', r.data->>'nocNo', r.data->>'woNo', r.data->>'refNo', '') AS doc_no,
            COALESCE(r.data->>'title', r.data->>'nocSubject', r.data->>'subject', r.data->>'description', '') AS title,
            COALESCE(r.data->>'status', r.data->>'partBStatus', r.data->>'partDStatus', '') AS status,
            COALESCE(r.data->>'driveLink', r.data->>'fileLocation', '') AS drive_link,
            COALESCE(r.data->>'issuedDate', r.data->>'partAIssueDate', r.data->>'actualReply', r.data->>'actualReplyDate', r.data->>'date', '') AS doc_date,
            r.created_at,
            COUNT(*) OVER() AS total_count
        FROM records r
        JOIN projects p ON r.project_id = p.id
        WHERE {where_sql}
        ORDER BY r.created_at DESC NULLS LAST
        LIMIT %s;
    """
    params.append(limit)

    rows = db.q(sql, params)
    total_found = int(rows[0]["total_count"]) if rows and rows[0].get("total_count") is not None else len(rows)

    seen_ids = set()
    results = []
    for row in rows:
        rec_id = row.get("id")
        if rec_id in seen_ids:
            continue
        seen_ids.add(rec_id)
        results.append({
            "id": rec_id,
            "project_id": row.get("project_id") or "",
            "project_code": row.get("project_code") or "",
            "project_name": row.get("project_name") or "",
            "dt_id": row.get("dt_id") or "",
            "doc_no": row.get("doc_no") or "",
            "title": row.get("title") or "",
            "status": row.get("status") or "",
            "drive_link": row.get("drive_link") or "",
            "doc_date": row.get("doc_date") or "",
            "issued_date": row.get("doc_date") or ""
        })

    if q_str and len(q_str) >= 2:
        try:
            db.log_action(
                username=u.get("username", ""),
                action="SEARCH",
                project_id=project_filter or None,
                dt_id=dt_id or None,
                detail=f"Global query: '{q_str[:80]}' | Found: {total_found}"
            )
        except Exception:
            pass

    return jsonify(
        query=q_str,
        total_found=total_found,
        user_scope=user_scope,
        results=results
    )


@records_bp.route("/api/records/<pid>/<dt_id>")
def api_records(pid, dt_id):
    u = current_user()
    if not u:
        return jsonify(error="LOGIN_REQUIRED"), 403
    if not can_view_project(pid, u):
        return jsonify(error="Forbidden"), 403
    if not can_access_doc_type(pid, dt_id, u):
        return jsonify(error="Document type access denied"), 403
    search  = request.args.get("search","")
    is_pr   = _is_pr_doc_type(pid, dt_id)
    records = db.get_records(pid, dt_id, search=search, search_pr_items=is_pr)
    cols    = db.get_columns(pid, dt_id)
    pr_items_map = db.get_pr_items_for_records([r.get("_id") for r in records]) if _is_pr_doc_type(pid, dt_id) else {}
    date_col_keys = {c["col_key"] for c in cols if c.get("col_type") in ("date","auto_date")}
    has_exp_reply = any(c["col_key"] == "expectedReplyDate" for c in cols)
    has_status    = any(c["col_key"] == "status" for c in cols)
    expected_reply_rule = db.get_expected_reply_rule(pid, dt_id)
    status_meta = db.get_status_meta_map(pid)
    for row in records:
        if has_exp_reply:
            issued_date = row.get("issuedDate")
            doc_no = row.get("docNo")
            exp = None
            if issued_date and doc_no:
                try:
                    exp = compute_expected_reply(issued_date, doc_no, expected_reply_rule, row.get("status"), row.get("action"), row=row)
                except Exception as e:
                    logger.warning("expected_reply_calc_failed pid=%s dt_id=%s record_id=%s error=%s",
                                   pid, dt_id, row.get("_id",""), e)
            row["_expectedReplyDate"] = format_date(exp) if exp else ""
        else:
            row["_expectedReplyDate"] = ""
        issued   = row.get("issuedDate","")
        actual   = row.get("actualReplyDate","")
        status_val = row.get("status")
        action_val = row.get("action")
        dur = compute_duration(issued, actual, expected_reply_rule, status_val, action_val)
        row["_duration"] = str(dur) if dur is not None else ""
        
        meta = db.resolve_status_meta(status_val, status_meta) if has_status else "pending"
        row["_overdue"]    = (meta == "pending") and is_overdue(row.get("issuedDate"), row.get("docNo"), row.get("actualReplyDate"), has_exp_reply, expected_reply_rule, status_val, action_val, row=row)
        row["_isRev"]      = extract_rev(row.get("docNo","")) > 0
        # Format ALL date columns (any col_type=date)
        for dk in date_col_keys:
            if dk in row and row[dk]:
                row["_fmt_" + dk] = format_date(row[dk])
        # Standard aliases
        row["_issuedFmt"]  = format_date(row.get("issuedDate",""))
        row["_replyFmt"]   = format_date(row.get("actualReplyDate",""))
    return jsonify(records=records, columns=cols, count=db.count_records(pid, dt_id), pr_items_map=pr_items_map)


@records_bp.route("/api/letters/parent-options/<pid>")
def api_letter_parent_options(pid):
    if not can_view_project(pid):
        return jsonify(error="Forbidden"), 403
    ltr_dts = [d for d in (db.get_doc_types(pid) or []) if _is_ltr_doc_type(pid, d.get("id"))]
    if ltr_dts:
        if not any(can_access_doc_type(pid, d.get("id")) for d in ltr_dts):
            return jsonify(error="Document type access denied"), 403
    elif not can_access_doc_type(pid, "LTR"):
        return jsonify(error="Document type access denied"), 403
    exclude_id = str(request.args.get("record_id","") or "").strip() or None
    return jsonify(options=db.get_letter_parent_options(pid, exclude_id))


@records_bp.route("/api/letters/thread/<pid>/<record_id>")
def api_letter_thread(pid, record_id):
    if not can_view_project(pid):
        return jsonify(error="Forbidden"), 403
    resolved = _resolve_project(pid)
    canon_pid = resolved["id"] if resolved else pid
    full_row = db.get_record_by_id(record_id)
    if not full_row:
        return jsonify(error="Not found"), 404
    if full_row.get("_project_id") != canon_pid:
        return jsonify(error="Invalid project"), 400
    if not can_access_doc_type(canon_pid, full_row.get("_dt_id", "")):
        return jsonify(error="Document type access denied"), 403
    if not _is_ltr_doc_type(canon_pid, full_row.get("_dt_id", "")):
        return jsonify(error="Not an LTR record"), 400
    thread = db.get_letter_thread(canon_pid, record_id)
    if not thread:
        return jsonify(error="Thread not found"), 404
    return jsonify(thread)


@records_bp.route("/api/letters/timeline/<pid>/<record_id>")
def api_letter_timeline(pid, record_id):
    if not can_view_project(pid):
        return jsonify(error="Forbidden"), 403
    resolved = _resolve_project(pid)
    canon_pid = resolved["id"] if resolved else pid
    full_row = db.get_record_by_id(record_id)
    if not full_row:
        return jsonify(error="Not found"), 404
    if full_row.get("_project_id") != canon_pid:
        return jsonify(error="Invalid project"), 400
    if not can_access_doc_type(canon_pid, full_row.get("_dt_id", "")):
        return jsonify(error="Document type access denied"), 403
    if not _is_ltr_doc_type(canon_pid, full_row.get("_dt_id", "")):
        return jsonify(error="Not an LTR record"), 400
    timeline = db.get_letter_timeline(canon_pid, record_id)
    if not timeline:
        return jsonify(error="Timeline not found"), 404
    return jsonify(timeline)


def _reject_record_ownership(u, pid, dt_id, rec_id, existing):
    """404 + audit entry for an update whose record id is owned by another project / document type.
    The response never reveals where the record lives; the audit detail (admin-only) does."""
    stored = ""
    if existing:
        stored = " (stored in project '%s' / type '%s')" % (existing.get("_project_id"), existing.get("_dt_id"))
    try:
        db.log_action(u.get("username", "") if u else "unknown", "ACCESS DENIED",
                      project_id=pid, dt_id=dt_id, record_id=rec_id,
                      detail="Record ownership mismatch on update attempt%s" % stored)
    except Exception:
        pass
    return jsonify(error="Record not found in this project / document type"), 404


@records_bp.route("/api/records/<pid>/<dt_id>", methods=["POST"])
def api_save_record(pid, dt_id):
    if not can_edit(pid):
        u = current_user()
        if u:
            try:
                db.log_action(u.get("username", ""), "ACCESS DENIED", project_id=pid, dt_id=dt_id,
                              detail=f"Unauthorized write attempt on project '{pid}'")
            except Exception:
                pass
        return jsonify(error="LOGIN_REQUIRED"), 403
    u      = current_user()
    if not can_access_doc_type(pid, dt_id, u):
        return jsonify(error="Document type access denied"), 403
    uname  = u["username"] if u else "unknown"
    data   = request.get_json(silent=True) or {}
    rec_id = data.pop("_id", None)
    if rec_id is not None and not isinstance(rec_id, str):
        return jsonify(error="Invalid record id"), 400
    # Finding B: an existing record may only be updated inside the project / document type that
    # stores it. Authorization above covers the *supplied* pid/dt_id, so the stored ownership must
    # match them; mismatches are rejected, never silently moved or overwritten.
    existing = db.get_record_by_id(rec_id) if rec_id else None
    if existing and (existing.get("_project_id") != pid or existing.get("_dt_id") != dt_id):
        return _reject_record_ownership(u, pid, dt_id, rec_id, existing)
    clean  = {k:v for k,v in data.items() if not k.startswith("_") and k != "attachments"}
    if _is_mhr_doc_type(pid, dt_id):
        try:
            year_val, month_val = db.normalize_mhr_period(clean)
            clean["year"] = year_val
            clean["month"] = month_val
            if str(dt_id or "").strip().upper() != "MHR":
                db.ensure_unique_mhr_period(pid, dt_id, rec_id or "", clean)
        except db.InvalidMhrPeriodError as e:
            return jsonify(error=str(e)), 400
        except db.DuplicateMhrPeriodError as e:
            return jsonify(error=str(e), code="DUPLICATE_MHR_PERIOD"), 409
    if _is_ltr_doc_type(pid, dt_id):
        direction_key = _ltr_field_key(pid, dt_id, "direction")
        issued_key = _ltr_field_key(pid, dt_id, "issuedDate")
        received_key = _ltr_field_key(pid, dt_id, "receivedDate")
        parent_id_key = _ltr_field_key(pid, dt_id, "parentLetterId")
        parent_ref_key = _ltr_field_key(pid, dt_id, "parentLetterRef")
        direction = str(clean.get(direction_key,"")).strip().lower()
        issued = str(clean.get(issued_key,"")).strip()
        received = str(clean.get(received_key,"")).strip()
        parent_id = str(clean.get(parent_id_key,"")).strip()
        if direction == "sent" and not issued:
            return jsonify(error="Issue Date is required for Sent letters"), 400
        if direction == "received" and not received:
            return jsonify(error="Received Date is required for Received letters"), 400
        if rec_id and parent_id and parent_id == rec_id:
            return jsonify(error="A letter cannot reference itself as parent"), 400
        if parent_id:
            parent = db.get_record_by_id(parent_id)
            if not parent or parent.get("_project_id") != pid or parent.get("_dt_id") != dt_id:
                return jsonify(error="Selected parent letter is invalid"), 400
            clean[parent_ref_key] = str(parent.get("docNo","") or "").strip()
        else:
            clean[parent_ref_key] = ""
    if rec_id:
        full_row = existing or {}
        # Extract only stored document fields for diff (exclude _ meta keys)
        old_data = {k: v for k, v in full_row.items() if not k.startswith("_")}
        try:
            db.save_record(pid, dt_id, rec_id, clean)
        except db.InvalidMhrPeriodError as e:
            return jsonify(error=str(e)), 400
        except db.DuplicateMhrPeriodError as e:
            return jsonify(error=str(e), code="DUPLICATE_MHR_PERIOD"), 409
        except db.RecordOwnershipError:
            # Lost a race with a concurrent create of the same id elsewhere: the atomic SQL guard refused.
            return _reject_record_ownership(u, pid, dt_id, rec_id, None)
        for field, new_val in clean.items():
            old_val = old_data.get(field,"")
            if str(old_val).strip() != str(new_val or "").strip():
                db.log_action(uname,"EDIT",pid,dt_id,rec_id,
                    clean.get("docNo",""),field,
                    str(old_val)[:200],str(new_val)[:200])
    else:
        rec_id = str(uuid.uuid4())
        try:
            db.save_record(pid, dt_id, rec_id, clean)
        except db.InvalidMhrPeriodError as e:
            return jsonify(error=str(e)), 400
        except db.DuplicateMhrPeriodError as e:
            return jsonify(error=str(e), code="DUPLICATE_MHR_PERIOD"), 409
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
    doc_no = (
        full_row.get("docNo")
        or full_row.get("nocNo")
        or full_row.get("prNo")
        or full_row.get("prNumber")
        or full_row.get("letterRef")
        or ""
    )
    if not can_edit(pid):
        try:
            db.log_action(u.get("username", ""), "ACCESS DENIED", project_id=pid, dt_id=dt_id,
                          doc_no=doc_no, detail=f"Unauthorized delete attempt on {doc_no or rec_id}")
        except Exception:
            pass
        return jsonify(error="Forbidden"), 403
    if not can_access_doc_type(pid, dt_id, u):
        return jsonify(error="Document type access denied"), 403
    db.delete_record(rec_id)
    db.log_action(
        u["username"], "DELETE",
        pid or None, dt_id or None, rec_id,
        doc_no, detail=f"Deleted: {doc_no}" if doc_no else f"Deleted record {rec_id}"
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

    blocked = next((r for r in recs if not can_edit(r.get("project_id", "")) or not can_access_doc_type(r.get("project_id", ""), r.get("dt_id", ""), u)), None)
    if blocked:
        return jsonify(error="Forbidden"), 403

    deleted = db.delete_records_bulk([r["id"] for r in recs])
    for r in recs:
        pid = r.get("project_id")
        dt_id = r.get("dt_id")
        rec_id = r.get("id")
        doc_no = (r.get("doc_no") or "").strip()
        detail = f"Deleted: {doc_no}" if doc_no else f"Deleted record {rec_id}"
        db.log_action(
            u["username"], "DELETE",
            pid, dt_id, rec_id, doc_no,
            detail=detail
        )
    return jsonify(ok=True, deleted=deleted)


@records_bp.route("/api/pr_items/<record_id>")
def api_get_pr_items(record_id):
    full_row = db.get_record_by_id(record_id)
    if not full_row:
        return jsonify(error="Not found"), 404
    pid   = full_row.get("_project_id", "")
    dt_id = full_row.get("_dt_id", "")
    if not can_view_project(pid):
        return jsonify(error="Forbidden"), 403
    if not can_access_doc_type(pid, dt_id):
        return jsonify(error="Document type access denied"), 403
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
    if not can_edit(pid): return jsonify(error="LOGIN_REQUIRED"), 403
    if not can_access_doc_type(pid, dt_id):
        return jsonify(error="Document type access denied"), 403
    if not _is_pr_doc_type(pid, dt_id):
        return jsonify(error="Not a PR document"), 400
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
        po_ref = str(it.get("po_ref", "")).strip()
        
        po_qty_raw = it.get("po_qty", "")
        po_qty = None
        if po_qty_raw not in (None, ""):
            try: po_qty = Decimal(str(po_qty_raw).strip())
            except (InvalidOperation, ValueError): po_qty = None
            
        del_qty_raw = it.get("delivered_qty", "")
        del_qty = None
        if del_qty_raw not in (None, ""):
            try: del_qty = Decimal(str(del_qty_raw).strip())
            except (InvalidOperation, ValueError): del_qty = None

        if not item_name and not unit and not remarks and qty is None and not po_ref and po_qty is None and del_qty is None:
            continue
        if not item_name:
            continue
        clean.append({
            "row_type": "item",
            "item_name": item_name,
            "unit": unit or None,
            "quantity": qty,
            "po_ref": po_ref or None,
            "po_qty": po_qty,
            "delivered_qty": del_qty,
            "remarks": remarks or None,
        })
    saved = db.save_pr_items(record_id, clean)
    return jsonify(ok=True, saved=saved)
