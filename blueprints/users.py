"""blueprints/users.py - DCR Users/Admin API routes with Multi-Workspace RBAC.

Handles users/admin endpoints:
  GET  /api/users                    -- list users (scoped for Dept Admins, global for Superadmin)
  POST /api/users                    -- manage user actions (scoped)
  GET  /api/users/<username>/projects-- user's assigned projects
  GET  /api/workspaces               -- list workspaces accessible to current user
  GET  /api/whoami                   -- current session profile + active workspaces
  POST /api/change_password          -- user changing own password
"""
from flask import Blueprint, jsonify, request

import db
from auth import (
    current_user,
    require_superadmin,
    require_admin,
    require_login,
    is_superadmin,
    get_allowed_project_ids,
)
from security import validate_username

users_bp = Blueprint("users", __name__)


@users_bp.route("/api/users")
@require_admin
def api_users():
    u = current_user()
    if not u:
        return jsonify(error="LOGIN_REQUIRED"), 403
    users = db.get_users_for_admin(u["username"])
    all_projects = db.get_all_user_projects()
    for usr in users:
        usr["projects"] = all_projects.get(usr["username"], [])
    return jsonify(users)


@users_bp.route("/api/users", methods=["POST"])
@require_admin
def api_users_action():
    data   = request.get_json(silent=True) or {}
    action = data.get("action")
    u = current_user() or {"username": "admin"}
    is_super = is_superadmin(u)
    admin_ws = db.get_admin_workspaces_for_user(u["username"]) if not is_super else []

    if action == "add":
        uname = data.get("username", "").strip().lower()
        pw    = data.get("password", "")
        role  = data.get("role", "viewer")
        email = data.get("email", "")
        ws_id = data.get("workspace_id") or data.get("primary_workspace") or "ENG"

        if not uname or not pw:
            return jsonify(ok=False, error="Username and password required"), 400
        is_valid, err_msg, _ = validate_username(uname)
        if not is_valid:
            return jsonify(ok=False, error=err_msg), 400

        if not is_super:
            if role == "superadmin":
                return jsonify(ok=False, error="Cannot create superadmin"), 403
            if ws_id not in admin_ws:
                ws_id = admin_ws[0] if admin_ws else "ENG"

        db.add_user(uname, pw, role, email, primary_workspace=ws_id)
        db.log_action(u["username"], "USER_CREATE",
                      detail=f"Created user '{uname}' (Role: {role}, Dept: {ws_id}, Email: {email or '—'})")
        return jsonify(ok=True)

    if action == "delete":
        uname = data.get("username", "")
        if uname == "admin":
            return jsonify(ok=False, error="Cannot delete super admin"), 400
        target_user = db.get_user(uname)
        if not target_user:
            return jsonify(ok=False, error="User not found"), 404
        if target_user.get("role") == "superadmin":
            return jsonify(ok=False, error="Cannot delete super admin"), 403
        if not is_super:
            target_workspaces = [w["id"] for w in db.get_user_workspaces(uname)]
            if not any(w in admin_ws for w in target_workspaces):
                return jsonify(ok=False, error="Cannot delete user from another department"), 403

        db.delete_user(uname)
        db.log_action(u["username"], "USER_DELETE", detail=f"Deleted user '{uname}'")
        return jsonify(ok=True)

    if action == "change_password":
        uname = data.get("username", "")
        target_user = db.get_user(uname)
        if not target_user:
            return jsonify(ok=False, error="User not found"), 404
        if not is_super:
            if target_user.get("role") == "superadmin":
                return jsonify(ok=False, error="Cannot change password of super admin"), 403
            target_workspaces = [w["id"] for w in db.get_user_workspaces(uname)]
            if not any(w in admin_ws for w in target_workspaces):
                return jsonify(ok=False, error="Cannot modify user from another department"), 403

        db.change_pw(uname, data.get("password", ""))
        db.log_action(u["username"], "PASSWORD_CHANGE", detail=f"Admin changed password for user '{uname}'")
        return jsonify(ok=True)

    if action == "update_role":
        uname = data.get("username", "").strip().lower()
        role = data.get("role", "viewer")
        target_user = db.get_user(uname)
        if not target_user:
            return jsonify(ok=False, error="User not found"), 404
        if uname == "admin" and role != "superadmin":
            return jsonify(ok=False, error="Cannot demote protected super admin"), 400
        if not is_super:
            if role == "superadmin":
                return jsonify(ok=False, error="Cannot promote to superadmin"), 403
            if target_user.get("role") == "superadmin":
                return jsonify(ok=False, error="Cannot change role of super admin"), 403
            target_workspaces = [w["id"] for w in db.get_user_workspaces(uname)]
            if not any(w in admin_ws for w in target_workspaces):
                return jsonify(ok=False, error="Cannot modify user from another department"), 403

        if not db.set_user_role(uname, role):
            return jsonify(ok=False, error="Invalid role"), 400
        db.log_action(u["username"], "ROLE_UPDATE", detail=f"Updated role of '{uname}' to '{role}'")
        return jsonify(ok=True)

    if action == "update_email":
        uname = data.get("username", "").strip().lower()
        target_user = db.get_user(uname)
        if not target_user:
            return jsonify(ok=False, error="User not found"), 404
        if not is_super:
            if target_user.get("role") == "superadmin":
                return jsonify(ok=False, error="Cannot modify super admin"), 403
            target_workspaces = [w["id"] for w in db.get_user_workspaces(uname)]
            if not any(w in admin_ws for w in target_workspaces):
                return jsonify(ok=False, error="Cannot modify user from another department"), 403

        db.set_user_email(uname, data.get("email", ""))
        db.log_action(u["username"], "USER_EDIT", detail=f"Updated email for '{uname}' to '{data.get('email')}'")
        return jsonify(ok=True)

    if action == "assign":
        uname = data.get("username", "")
        pid = data.get("project_id", "")
        is_dc = data.get("is_dc", False)
        if not is_super:
            proj = db.get_project(pid)
            if not proj or proj.get("workspace_id", "ENG") not in admin_ws:
                return jsonify(ok=False, error="Cannot assign project from another department"), 403

        db.assign_project(uname, pid, is_dc)
        db.log_action(u["username"], "PROJECT_ASSIGN", project_id=pid,
                      detail=f"Assigned project '{pid}' to '{uname}' (DC: {is_dc})")
        return jsonify(ok=True)

    if action == "unassign":
        uname = data.get("username", "")
        pid = data.get("project_id", "")
        if not is_super:
            proj = db.get_project(pid)
            if not proj or proj.get("workspace_id", "ENG") not in admin_ws:
                return jsonify(ok=False, error="Cannot unassign project from another department"), 403

        db.unassign_project(uname, pid)
        db.log_action(u["username"], "PROJECT_UNASSIGN", project_id=pid,
                      detail=f"Unassigned project '{pid}' from '{uname}'")
        return jsonify(ok=True)

    if action == "update_workspaces":
        uname = data.get("username", "").strip().lower()
        ws_ids = data.get("workspaces", [])
        if not isinstance(ws_ids, list) or not ws_ids:
            return jsonify(ok=False, error="User must belong to at least one department"), 400
        if not is_super:
            return jsonify(ok=False, error="Superadmin only"), 403
        target_user = db.get_user(uname)
        if not target_user:
            return jsonify(ok=False, error="User not found"), 404
        if target_user.get("role") == "superadmin" or uname == "admin":
            return jsonify(ok=False, error="Cannot modify superadmin workspaces"), 400

        db.set_user_workspaces(uname, ws_ids, target_user.get("role", "viewer"))
        db.log_action(u["username"], "WORKSPACE_UPDATE", detail=f"Updated departments for '{uname}' to {ws_ids}")
        return jsonify(ok=True)

    if action in ("assign_workspace", "update_workspace", "update_department"):
        uname = data.get("username", "").strip().lower()
        ws_id = data.get("workspace_id") or data.get("department", "")
        target_user = db.get_user(uname)
        if not target_user:
            return jsonify(ok=False, error="User not found"), 404
        ws_role = data.get("role") or target_user.get("role", "viewer")
        if not is_super:
            if ws_id not in admin_ws:
                return jsonify(ok=False, error="Cannot assign workspace outside your authority"), 403
            if ws_role == "superadmin":
                return jsonify(ok=False, error="Cannot grant superadmin"), 403
            target_workspaces = [w["id"] for w in db.get_user_workspaces(uname)]
            if not any(w in admin_ws for w in target_workspaces):
                return jsonify(ok=False, error="Cannot modify user from another department"), 403
        if hasattr(db, "set_user_primary_workspace"):
            db.set_user_primary_workspace(uname, ws_id)
        db.assign_user_workspace(uname, ws_id, ws_role)
        db.log_action(u["username"], "WORKSPACE_ASSIGN", detail=f"Assigned workspace '{ws_id}' to '{uname}' (Role: {ws_role})")
        return jsonify(ok=True)

    return jsonify(ok=False, error="Unknown action"), 400


@users_bp.route("/api/users/<username>/projects")
@require_admin
def api_user_projects(username):
    u = current_user()
    if not is_superadmin(u):
        allowed_users = [usr["username"] for usr in db.get_users_for_admin(u["username"])]
        if username not in allowed_users:
            return jsonify(error="Forbidden"), 403
    return jsonify(db.get_user_projects(username))


@users_bp.route("/api/workspaces")
@require_login
def api_workspaces():
    u = current_user()
    if not u:
        return jsonify(error="LOGIN_REQUIRED"), 403
    workspaces = db.get_user_workspaces(u["username"])
    return jsonify(workspaces)


@users_bp.route("/api/whoami")
def api_whoami():
    u = current_user()
    if not u:
        return jsonify(username="guest", role="guest", is_superadmin=False, projects=[], dc_projects=[], workspaces=[])
    projs = get_allowed_project_ids(u)
    user_ws = db.get_user_workspaces(u["username"])
    is_super = is_superadmin(u)
    if u["role"] in ("superadmin", "admin"):
        return jsonify(username=u["username"], role=u["role"], is_superadmin=is_super, projects=projs, dc_projects=[], workspaces=user_ws)
    user_projects = db.get_user_projects(u["username"])
    dc_projects = [p["project_id"] for p in user_projects if p.get("is_dc")]
    return jsonify(username=u["username"], role=u["role"], is_superadmin=is_super, projects=projs, dc_projects=dc_projects, workspaces=user_ws)


@users_bp.route("/api/change_password", methods=["POST"])
def api_change_own_pw():
    u = current_user()
    if not u:
        return jsonify(error="LOGIN_REQUIRED"), 403
    data = request.get_json(silent=True) or {}
    pw   = data.get("password", "")
    if len(pw) < 4:
        return jsonify(ok=False, error="Min 4 characters"), 400
    db.change_pw(u["username"], pw)
    db.log_action(u["username"], "PASSWORD_CHANGE", detail="User changed own password")
    return jsonify(ok=True)
