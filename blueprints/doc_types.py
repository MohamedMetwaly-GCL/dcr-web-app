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
from auth import (
    current_user,
    can_edit,
    can_view_project,
    is_project_admin,
    get_allowed_doc_type_filter,
    filter_allowed_doc_types,
    can_access_doc_type,
)

doc_types_bp = Blueprint("doc_types", __name__)


def _can_manage_doc_types(pid, user, dt_id=None):
    """Allow project schema/doc-type administration for project admins and unrestricted editors."""
    if not user or not can_edit(pid, user):
        return False
    if is_project_admin(pid, user):
        return True
    # Restricted users (non-None allowed_doc_types) cannot add, rename, reorder, or delete doc types.
    if get_allowed_doc_type_filter(pid, user) is not None:
        return False
    if dt_id and not can_access_doc_type(pid, dt_id, user):
        return False
    return True


@doc_types_bp.route("/api/doc_types/<pid>")
def api_doc_types(pid):
    u = current_user()
    if not u or not can_view_project(pid, u):
        return jsonify(error="Forbidden"), 403
    dts = db.get_doc_types(pid)
    if request.args.get("all") == "1" and is_project_admin(pid, u):
        return jsonify(dts)
    return jsonify(filter_allowed_doc_types(pid, dts, u))

@doc_types_bp.route("/api/doc_types/<pid>", methods=["POST"])
def api_add_doc_type(pid):
    u = current_user()
    if not u:
        return jsonify(error="LOGIN_REQUIRED"), 403
    if not can_edit(pid, u):
        return jsonify(error="LOGIN_REQUIRED"), 403
    if not _can_manage_doc_types(pid, u):
        return jsonify(error="Document type access denied"), 403
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
    u = current_user()
    if not u or not can_edit(pid, u):
        return jsonify(error="LOGIN_REQUIRED"), 403
    if not _can_manage_doc_types(pid, u, dt_id):
        return jsonify(error="Document type access denied"), 403
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
    u = current_user()
    if not u or not can_edit(pid, u):
        return jsonify(error="LOGIN_REQUIRED"), 403
    if not _can_manage_doc_types(pid, u):
        return jsonify(error="Document type access denied"), 403
    data = request.get_json(silent=True) or {}
    order = data.get("order", [])
    for i, dt_id in enumerate(order):
        db.exe("UPDATE doc_types SET sort_order=%s WHERE id=%s AND project_id=%s", (i, dt_id, pid))
    return jsonify(ok=True)

@doc_types_bp.route("/api/doc_types/<pid>/<dt_id>", methods=["DELETE"])
def api_delete_doc_type(pid, dt_id):
    u = current_user()
    if not u or not can_edit(pid, u):
        return jsonify(error="LOGIN_REQUIRED"), 403
    if not _can_manage_doc_types(pid, u, dt_id):
        try:
            db.log_action(u.get("username", ""), "ACCESS DENIED", project_id=pid, dt_id=dt_id,
                          detail="Unauthorized document-type deletion attempt")
        except Exception:
            pass
        return jsonify(error="Document type access denied"), 403
    db.delete_doc_type(pid, dt_id)
    return jsonify(ok=True)
