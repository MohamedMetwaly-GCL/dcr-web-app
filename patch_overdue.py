import re

# 1. db.py
with open(r'D:\DCR\db.py', 'r', encoding='utf-8') as f:
    code = f.read()

# get_dashboard_stats
old_dash = """                dt_rule = dt_rules.get(dt["id"])
                is_ov = is_overdue(d.get("issuedDate"), doc_no, d.get("actualReplyDate"), _has_both, rule=dt_rule, status=status_val, action=action_val)"""

new_dash = """                dt_rule = dt_rules.get(dt["id"])
                is_ov = is_pe and is_overdue(d.get("issuedDate"), doc_no, d.get("actualReplyDate"), _has_both, rule=dt_rule, status=status_val, action=action_val)"""

code = code.replace(old_dash, new_dash)

# get_overdue_records
old_ov = """    dt_rows = q(f"SELECT id, code, name FROM doc_types {dt_where}", dt_params)
    dt_with_exp = {dt_id for dt_id in dt_with_exp
                   if not _is_non_workflow_dt(
                       next((d["code"] for d in dt_rows if d["id"] == dt_id), ""),
                       next((d["name"] for d in dt_rows if d["id"] == dt_id), "")
                   )}
    result = []
    for row in rows:"""

new_ov = """    dt_rows = q(f"SELECT id, code, name FROM doc_types {dt_where}", dt_params)
    dt_with_exp = {dt_id for dt_id in dt_with_exp
                   if not _is_non_workflow_dt(
                       next((d["code"] for d in dt_rows if d["id"] == dt_id), ""),
                       next((d["name"] for d in dt_rows if d["id"] == dt_id), "")
                   )}
    
    list_where, list_params = _project_scope_clause("project_id", pid, project_ids)
    meta_rows = q(f"SELECT project_id, item_value, meta FROM dropdown_lists {list_where} AND list_name LIKE %s AND meta IS NOT NULL", list_params + ["status%"])
    meta_map = {}
    for r in meta_rows:
        meta_map.setdefault(r["project_id"], {})[r["item_value"]] = r["meta"]

    result = []
    for row in rows:"""

code = code.replace(old_ov, new_ov)

old_ov2 = """        status_val = d.get("status")
        action_val = d.get("action")
        dt_rule = dt_rules.get(row["dt_id"])
        if is_overdue(issued, doc_no, None, True, rule=dt_rule, status=status_val, action=action_val):"""

new_ov2 = """        status_val = d.get("status")
        action_val = d.get("action")
        dt_rule = dt_rules.get(row["dt_id"])
        
        meta = resolve_status_meta(status_val, meta_map.get(row["project_id"], {}))
        if meta != "pending":
            continue

        if is_overdue(issued, doc_no, None, True, rule=dt_rule, status=status_val, action=action_val):"""

code = code.replace(old_ov2, new_ov2)

with open(r'D:\DCR\db.py', 'w', encoding='utf-8') as f:
    f.write(code)

# 2. blueprints/records.py
with open(r'D:\DCR\blueprints\records.py', 'r', encoding='utf-8') as f:
    code = f.read()

old_rec = """    has_status    = any(c["col_key"] == "status" for c in cols)
    expected_reply_rule = db.get_expected_reply_rule(pid, dt_id)
    for row in records:"""

new_rec = """    has_status    = any(c["col_key"] == "status" for c in cols)
    expected_reply_rule = db.get_expected_reply_rule(pid, dt_id)
    status_meta = db.get_status_meta_map(pid)
    for row in records:"""

code = code.replace(old_rec, new_rec)

old_rec2 = """        dur = compute_duration(issued, actual, expected_reply_rule, status_val, action_val)
        row["_duration"] = str(dur) if dur is not None else ""
        row["_overdue"]    = is_overdue(row.get("issuedDate"), row.get("docNo"), row.get("actualReplyDate"), has_exp_reply, expected_reply_rule, status_val, action_val)"""

new_rec2 = """        dur = compute_duration(issued, actual, expected_reply_rule, status_val, action_val)
        row["_duration"] = str(dur) if dur is not None else ""
        
        meta = db.resolve_status_meta(status_val, status_meta) if has_status else "pending"
        row["_overdue"]    = (meta == "pending") and is_overdue(row.get("issuedDate"), row.get("docNo"), row.get("actualReplyDate"), has_exp_reply, expected_reply_rule, status_val, action_val)"""

code = code.replace(old_rec2, new_rec2)

with open(r'D:\DCR\blueprints\records.py', 'w', encoding='utf-8') as f:
    f.write(code)


# 3. blueprints/exporting.py
with open(r'D:\DCR\blueprints\exporting.py', 'r', encoding='utf-8') as f:
    code = f.read()

old_exp = """    dt_rule_map = {}
    
    excel_rows = []"""

new_exp = """    dt_rule_map = {}
    status_meta = db.get_status_meta_map(pid)
    
    excel_rows = []"""

code = code.replace(old_exp, new_exp)

old_exp2 = """            dur    = compute_duration(row.get("issuedDate"), row.get("actualReplyDate"), expected_reply_rule, status=status_val, action=action_val)
            ov     = is_overdue(row.get("issuedDate"), row.get("docNo"), row.get("actualReplyDate"), has_exp_col, expected_reply_rule, status=status_val, action=action_val)"""

new_exp2 = """            dur    = compute_duration(row.get("issuedDate"), row.get("actualReplyDate"), expected_reply_rule, status=status_val, action=action_val)
            meta = db.resolve_status_meta(status_val, status_meta) if has_status_col else "pending"
            ov     = (meta == "pending") and is_overdue(row.get("issuedDate"), row.get("docNo"), row.get("actualReplyDate"), has_exp_col, expected_reply_rule, status=status_val, action=action_val)"""

code = code.replace(old_exp2, new_exp2)

old_exp3 = """    dt_rule_map = {}
    
    pdf_rows = []"""

new_exp3 = """    dt_rule_map = {}
    status_meta = db.get_status_meta_map(pid)
    
    pdf_rows = []"""

code = code.replace(old_exp3, new_exp3)

old_exp4 = """        dur    = compute_duration(row.get("issuedDate"), row.get("actualReplyDate"), rule=None, status=status_val, action=action_val)
        ov     = is_overdue(row.get("issuedDate"), row.get("docNo"), row.get("actualReplyDate"), has_exp_col_pdf, rule=None, status=status_val, action=action_val)"""

new_exp4 = """        dur    = compute_duration(row.get("issuedDate"), row.get("actualReplyDate"), rule=None, status=status_val, action=action_val)
        meta = db.resolve_status_meta(status_val, status_meta) if has_status_col_pdf else "pending"
        ov     = (meta == "pending") and is_overdue(row.get("issuedDate"), row.get("docNo"), row.get("actualReplyDate"), has_exp_col_pdf, rule=None, status=status_val, action=action_val)"""

code = code.replace(old_exp4, new_exp4)

with open(r'D:\DCR\blueprints\exporting.py', 'w', encoding='utf-8') as f:
    f.write(code)

print("SUCCESS")
