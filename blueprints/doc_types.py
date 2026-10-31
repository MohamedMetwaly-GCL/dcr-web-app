"""blueprints/doc_types.py - DCR Doc Types API routes.

Handles all document-type management endpoints:
  GET    /api/doc_types/<pid>              -- list doc types for a project
  POST   /api/doc_types/<pid>             -- add a new doc type
  PATCH  /api/doc_types/<pid>/<dt_id>     -- rename a doc type (code/name)
  POST   /api/doc_types/<pid>/reorder     -- reorder doc types
  DELETE /api/doc_types/<pid>/<dt_id>     -- delete a doc type and its records

Step 5 of the incremental refactor.
Logic is identical to the original app.py routes — only the decorator
changed from @app.route to @doc_types_bp.route.
"""
from flask import Blueprint, request, jsonify

import db
from auth import can_edit

doc_types_bp = Blueprint("doc_types", __name__)


@doc_types_bp.route("/api/doc_types/<pid>")
def api_doc_types(pid):
    from auth import current_user, is_superadmin
    u = current_user()
    dts = db.get_doc_types(pid)
    if request.args.get("all") == "1":
        return jsonify(dts)
    if u and not is_superadmin(u) and not db.is_user_workspace_admin(u.get("username")):
        allowed_dts = db.get_user_project_doc_types(u["username"], pid)
        if allowed_dts is not None:
            allowed_set = {str(x).upper() for x in allowed_dts}
            dts = [d for d in dts if str(d["id"]).upper() in allowed_set or str(d.get("code", "")).upper() in allowed_set]
    return jsonify(dts)

@doc_types_bp.route("/api/doc_types/<pid>", methods=["POST"])
def api_add_doc_type(pid):
    if not can_edit(pid): return jsonify(error="LOGIN_REQUIRED"), 403
    data = request.get_json(silent=True) or {}
    code = data.get("code","").strip().upper()
    name = data.get("name","").strip()
    if not code or not name:
        return jsonify(ok=False, error="Code and name required"), 400
    dup = db.q("""SELECT id FROM doc_types
                  WHERE project_id=%s AND (LOWER(code)=LOWER(%s) OR LOWER(name)=LOWER(%s))
                  LIMIT 1""", (pid, code, name), one=True)
    if dup:
        return jsonify(ok=False, error="Document type code/name already exists"), 400
    override = data.get("expected_reply_override") or {}
    db.add_doc_type(pid, code, name, override)
    return jsonify(ok=True)


@doc_types_bp.route("/api/doc_types/<pid>/<dt_id>", methods=["PATCH"])
def api_rename_doc_type(pid, dt_id):
    if not can_edit(pid): return jsonify(error="LOGIN_REQUIRED"), 403
    data = request.get_json(silent=True) or {}
    new_code = data.get("code", "").strip().upper()
    new_name = data.get("name", "").strip()
    if not new_code or not new_name:
        return jsonify(ok=False, error="Code and name required"), 400
    existing = db.q("SELECT id FROM doc_types WHERE id=%s AND project_id=%s", (dt_id, pid), one=True)
    if not existing:
        return jsonify(ok=False, error="Document type not found"), 404
    dup = db.q("""SELECT id FROM doc_types
                  WHERE project_id=%s AND id<>%s AND (LOWER(code)=LOWER(%s) OR LOWER(name)=LOWER(%s))
                  LIMIT 1""", (pid, dt_id, new_code, new_name), one=True)
    if dup:
        return jsonify(ok=False, error="Document type code/name already exists"), 400
    db.exe("UPDATE doc_types SET code=%s, name=%s WHERE id=%s AND project_id=%s", (new_code, new_name, dt_id, pid))
    if "expected_reply_override" in data:
        db.save_doc_type_expected_reply_override(pid, dt_id, data.get("expected_reply_override") or {})
    return jsonify(ok=True)

@doc_types_bp.route("/api/doc_types/<pid>/reorder", methods=["POST"])
def api_reorder_doc_types(pid):
    if not can_edit(pid): return jsonify(error="LOGIN_REQUIRED"), 403
    data = request.get_json(silent=True) or {}
    order = data.get("order", [])
    for i, dt_id in enumerate(order):
        db.exe("UPDATE doc_types SET sort_order=%s WHERE id=%s AND project_id=%s", (i, dt_id, pid))
    return jsonify(ok=True)

@doc_types_bp.route("/api/doc_types/<pid>/<dt_id>", methods=["DELETE"])
def api_delete_doc_type(pid, dt_id):
    if not can_edit(pid): return jsonify(error="LOGIN_REQUIRED"), 403
    db.delete_doc_type(pid, dt_id)
    return jsonify(ok=True)
