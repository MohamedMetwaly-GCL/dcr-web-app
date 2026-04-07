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
from auth import current_user, require_superadmin

users_bp = Blueprint("users", __name__)


@users_bp.route("/api/users")
@require_superadmin
def api_users():
    return jsonify(db.get_all_users())


@users_bp.route("/api/users", methods=["POST"])
@require_superadmin
def api_users_action():
    data   = request.get_json(silent=True) or {}
    action = data.get("action")
    if action == "add":
        uname = data.get("username","").strip().lower()
        pw    = data.get("password","")
        role  = data.get("role","viewer")
        if not uname or not pw:
            return jsonify(ok=False, error="Username and password required"), 400
        db.add_user(uname, pw, role)
        return jsonify(ok=True)
    if action == "delete":
        uname = data.get("username","")
        if uname == "admin":
            return jsonify(ok=False, error="Cannot delete super admin"), 400
        db.delete_user(uname)
        return jsonify(ok=True)
    if action == "change_password":
        db.change_pw(data.get("username",""), data.get("password",""))
        return jsonify(ok=True)
    if action == "assign":
        db.assign_project(data.get("username",""), data.get("project_id",""))
        return jsonify(ok=True)
    if action == "unassign":
        db.unassign_project(data.get("username",""), data.get("project_id",""))
        return jsonify(ok=True)
    return jsonify(ok=False, error="Unknown action"), 400


@users_bp.route("/api/users/<username>/projects")
@require_superadmin
def api_user_projects(username):
    return jsonify(db.get_user_projects(username))


@users_bp.route("/api/whoami")
def api_whoami():
    u = current_user()
    if not u: return jsonify(username="guest", role="guest", projects=[])
    projs = [] if u["role"] == "superadmin" else db.get_user_projects(u["username"])
    return jsonify(username=u["username"], role=u["role"], projects=projs)


@users_bp.route("/api/change_password", methods=["POST"])
def api_change_own_pw():
    u = current_user()
    if not u: return jsonify(error="LOGIN_REQUIRED"), 403
    data = request.get_json(silent=True) or {}
    pw   = data.get("password","")
    if len(pw) < 4: return jsonify(ok=False, error="Min 4 characters"), 400
    db.change_pw(u["username"], pw)
    return jsonify(ok=True)
