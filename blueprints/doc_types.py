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

import attachments
import db
import drive_routing
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
    raw_cycle = data.get("has_reply_cycle")
    has_reply_cycle = raw_cycle if isinstance(raw_cycle, bool) else None
    try:
        db.add_doc_type(pid, code, name, override, has_reply_cycle=has_reply_cycle)
    except TypeError:
        db.add_doc_type(pid, code, name, override)
        if isinstance(has_reply_cycle, bool):
            db.save_doc_type_metadata(pid, code, has_reply_cycle=has_reply_cycle)

    drive_prov = {"ok": True, "skipped": True, "reason": "no_drive_folder", "created": 0, "existing": 0, "errors": []}
    try:
        drive_prov = drive_routing.provision_project_drive_structure(pid, only_doc_type_id=code)
        proj = db.get_project(pid) or {}
        ws_id = str(proj.get("workspace_id") or "ENG").strip().upper()
        for other_p in (db.get_projects(workspace_id=ws_id) or []):
            other_pid = str((other_p or {}).get("id") or "").strip()
            if not other_pid or other_pid == pid:
                continue
            if db.project_has_doc_type(other_pid, code):
                drive_routing.provision_project_drive_structure(other_pid, only_doc_type_id=code)
    except Exception:
        pass
    return jsonify(ok=True, drive_provision=drive_prov)


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
    existing = db.q("SELECT id, code, name, data FROM doc_types WHERE id=%s AND project_id=%s", (dt_id, pid), one=True)
    if not existing:
        return jsonify(ok=False, error="Document type not found"), 404
    dup = db.q("""SELECT id FROM doc_types
                  WHERE project_id=%s AND id<>%s AND (LOWER(code)=LOWER(%s) OR LOWER(name)=LOWER(%s))
                  LIMIT 1""", (pid, dt_id, new_code, new_name), one=True)
    if dup:
        return jsonify(ok=False, error="Document type code/name already exists"), 400

    old_dt = db.get_doc_type(pid, dt_id) or existing
    old_code = str(old_dt.get("code") or dt_id).strip()
    old_name = str(old_dt.get("name") or "").strip()
    proj = {}
    try:
        proj = db.get_project(pid) or {}
    except Exception:
        proj = {}
    ws_id = str(proj.get("workspace_id") or "ENG").strip().upper()
    old_has_cycle = attachments.doc_type_has_reply_cycle(old_dt, workspace_id=ws_id)

    db.exe("UPDATE doc_types SET code=%s, name=%s WHERE id=%s AND project_id=%s", (new_code, new_name, dt_id, pid))

    has_override_key = "expected_reply_override" in data
    raw_cycle = data.get("has_reply_cycle")
    has_cycle_bool = raw_cycle if isinstance(raw_cycle, bool) else None
    if has_override_key or has_cycle_bool is not None:
        db.save_doc_type_metadata(
            pid,
            dt_id,
            expected_reply_override=(data.get("expected_reply_override") or {}) if has_override_key else None,
            has_reply_cycle=has_cycle_bool,
        )

    drive_rename = {"ok": True, "skipped": True, "renamed": False}
    drive_prov = None
    try:
        if (old_code, old_name) != (new_code, new_name):
            drive_rename = drive_routing.rename_doc_type_drive_folder(
                pid, dt_id, old_code, old_name, new_code, new_name
            )
        updated_dt = db.get_doc_type(pid, dt_id) or {"id": dt_id, "code": new_code, "name": new_name}
        new_has_cycle = attachments.doc_type_has_reply_cycle(updated_dt, workspace_id=ws_id)
        if (new_has_cycle and not old_has_cycle) or (drive_rename.get("reason") == "folder_not_found"):
            drive_prov = drive_routing.provision_project_drive_structure(pid, only_doc_type_id=dt_id)
    except Exception:
        pass
    return jsonify(ok=True, drive_rename=drive_rename, drive_provision=drive_prov)

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
    proj = None
    dt = None
    try:
        proj = db.get_project(pid)
        dt = db.get_doc_type(pid, dt_id)
    except Exception:
        pass
    db.delete_doc_type(pid, dt_id)
    drive_cleanup = {"ok": True, "skipped": True, "reason": "no_drive_folder"}
    try:
        drive_cleanup = drive_routing.cleanup_doc_type_drive_folder(
            proj,
            dt or {"id": dt_id, "code": dt_id, "name": ""},
        )
    except Exception as err:
        drive_cleanup = {"ok": False, "action": "error", "error": str(err)}
    return jsonify(ok=True, drive_cleanup=drive_cleanup)
