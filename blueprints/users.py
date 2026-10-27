"""blueprints/users.py - DCR Users/Admin API routes.

Handles users/admin endpoints:
  GET  /api/users
  POST /api/users
  GET  /api/users/<username>/projects
  GET  /api/whoami
  POST /api/change_password

Step 12 of the incremental refactor.
Logic is identical to the original app.py routes — only the decorator
changed from @app.route to @users_bp.route.
"""
from flask import Blueprint, jsonify, request

import db
from auth import current_user, require_superadmin, get_allowed_project_ids
from security import validate_username

users_bp = Blueprint("users", __name__)


@users_bp.route("/api/users")
@require_superadmin
def api_users():
    users = db.get_all_users()
    all_projects = db.get_all_user_projects()
    all_depts = db.get_all_user_departments()
    for u in users:
        u["projects"] = all_projects.get(u["username"], [])
        u["departments"] = all_depts.get(u["username"], [])
        u["primary_dept"] = u.get("primary_dept") or "ENG"
        u["confidentiality_clearance"] = u.get("confidentiality_clearance") or "internal"
    return jsonify(users)


@users_bp.route("/api/users", methods=["POST"])
@require_superadmin
def api_users_action():
    data   = request.get_json(silent=True) or {}
    action = data.get("action")
    u = current_user() or {"username": "admin"}
    if action == "add":
        uname = data.get("username","").strip().lower()
        pw    = data.get("password","")
        role  = data.get("role","viewer")
        email = data.get("email","")
        if not uname or not pw:
            return jsonify(ok=False, error="Username and password required"), 400
        is_valid, err_msg, _ = validate_username(uname)
        if not is_valid:
            return jsonify(ok=False, error=err_msg), 400
        pdept = str(data.get("primary_dept") or "ENG").strip().upper()
        clearance = str(data.get("confidentiality_clearance") or "internal").strip().lower()
        db.add_user(uname, pw, role, email, primary_dept=pdept, confidentiality_clearance=clearance)
        if pdept:
            try:
                db.assign_user_department(uname, pdept, "editor" if role in ("editor", "admin", "superadmin") else "viewer")
            except Exception:
                pass
        db.log_action(u["username"], "USER_CREATE", detail=f"Created user '{uname}' (Role: {role}, Dept: {pdept}, Clearance: {clearance})")
        return jsonify(ok=True)
    if action == "delete":
        uname = data.get("username","")
        if uname == "admin":
            return jsonify(ok=False, error="Cannot delete super admin"), 400
        db.delete_user(uname)
        db.log_action(u["username"], "USER_DELETE", detail=f"Deleted user '{uname}'")
        return jsonify(ok=True)
    if action == "change_password":
        db.change_pw(data.get("username",""), data.get("password",""))
        db.log_action(u["username"], "PASSWORD_CHANGE", detail=f"Admin changed password for user '{data.get('username')}'")
        return jsonify(ok=True)
    if action == "update_role":
        uname = data.get("username","").strip().lower()
        role = data.get("role","viewer")
        if uname == "admin" and role != "superadmin":
            return jsonify(ok=False, error="Cannot demote protected super admin"), 400
        if not db.set_user_role(uname, role):
            return jsonify(ok=False, error="Invalid role"), 400
        db.log_action(u["username"], "ROLE_UPDATE", detail=f"Updated role of '{uname}' to '{role}'")
        return jsonify(ok=True)
    if action == "update_email":
        db.set_user_email(data.get("username","").strip().lower(), data.get("email",""))
        db.log_action(u["username"], "USER_EDIT", detail=f"Updated email for '{data.get('username')}' to '{data.get('email')}'")
        return jsonify(ok=True)
    if action == "assign":
        db.assign_project(data.get("username",""), data.get("project_id",""), data.get("is_dc", False))
        db.log_action(u["username"], "PROJECT_ASSIGN", project_id=data.get("project_id"),
                      detail=f"Assigned project '{data.get('project_id')}' to '{data.get('username')}' (DC: {data.get('is_dc', False)})")
        return jsonify(ok=True)
    if action == "unassign":
        db.unassign_project(data.get("username",""), data.get("project_id",""))
        db.log_action(u["username"], "PROJECT_UNASSIGN", project_id=data.get("project_id"),
                      detail=f"Unassigned project '{data.get('project_id')}' from '{data.get('username')}'")
        return jsonify(ok=True)
    if action == "assign_dept":
        uname = data.get("username","").strip().lower()
        dept_id = str(data.get("department_id","")).strip().upper()
        dept_role = str(data.get("dept_role","viewer")).strip().lower()
        db.assign_user_department(uname, dept_id, dept_role)
        db.log_action(u["username"], "DEPT_ASSIGN",
                      detail=f"Assigned department '{dept_id}' to '{uname}' (Dept Role: {dept_role})")
        return jsonify(ok=True)
    if action == "unassign_dept":
        uname = data.get("username","").strip().lower()
        dept_id = str(data.get("department_id","")).strip().upper()
        db.unassign_user_department(uname, dept_id)
        db.log_action(u["username"], "DEPT_UNASSIGN",
                      detail=f"Unassigned department '{dept_id}' from '{uname}'")
        return jsonify(ok=True)
    if action == "update_clearance":
        uname = data.get("username","").strip().lower()
        clearance = str(data.get("confidentiality_clearance","internal")).strip().lower()
        pdept = data.get("primary_dept")
        db.set_user_clearance(uname, clearance, pdept)
        db.log_action(u["username"], "USER_EDIT",
                      detail=f"Updated clearance for '{uname}' to '{clearance}'" + (f" (Primary Dept: {pdept})" if pdept else ""))
        return jsonify(ok=True)
    return jsonify(ok=False, error="Unknown action"), 400


@users_bp.route("/api/users/<username>/projects")
@require_superadmin
def api_user_projects(username):
    return jsonify(db.get_user_projects(username))


@users_bp.route("/api/whoami")
def api_whoami():
    u = current_user()
    if not u: return jsonify(username="guest", role="guest", projects=[], dc_projects=[])
    projs = get_allowed_project_ids(u)
    # Admins/superadmins have global access — dc_projects is not applicable
    if u["role"] in ("superadmin", "admin"):
        return jsonify(username=u["username"], role=u["role"], projects=projs, dc_projects=[])
    # For regular users: single query to get assigned projects with is_dc flag
    user_projects = db.get_user_projects(u["username"])
    dc_projects = [p["project_id"] for p in user_projects if p.get("is_dc")]
    return jsonify(username=u["username"], role=u["role"], projects=projs, dc_projects=dc_projects)


@users_bp.route("/api/change_password", methods=["POST"])
def api_change_own_pw():
    u = current_user()
    if not u: return jsonify(error="LOGIN_REQUIRED"), 403
    data = request.get_json(silent=True) or {}
    pw   = data.get("password","")
    if len(pw) < 4: return jsonify(ok=False, error="Min 4 characters"), 400
    db.change_pw(u["username"], pw)
    db.log_action(u["username"], "PASSWORD_CHANGE", detail="User changed own password")
    return jsonify(ok=True)
