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
import re

import db
import drive_routing
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


def _target_workspace_ids(target_user):
    uname = str((target_user or {}).get("username") or "").strip().lower()
    if not uname:
        return set()
    rows = db.get_user_workspaces(uname) or []
    ws_ids = {str(w.get("id") or "").strip().upper() for w in rows if w.get("id")}
    if not ws_ids:
        primary = str((target_user or {}).get("primary_workspace") or "").strip().upper()
        if primary:
            ws_ids.add(primary)
    return ws_ids


def _is_peer_or_higher_admin(target_user):
    if not target_user:
        return False
    uname = str(target_user.get("username") or "").strip().lower()
    global_role = str(target_user.get("role") or "").strip().lower()
    if uname == "admin" or global_role in ("superadmin", "admin"):
        return True
    admin_ws = db.get_admin_workspaces_for_user(uname) or []
    return bool(admin_ws)


def _can_manage_target_user(actor, target_user, admin_ws, require_exclusive_workspace=False):
    """Verify actor may manage target_user under scoped Department Admin RBAC."""
    if is_superadmin(actor):
        return True, None
    actor_uname = str((actor or {}).get("username") or "").strip().lower()
    target_uname = str((target_user or {}).get("username") or "").strip().lower()
    if not target_uname:
        return False, (jsonify(ok=False, error="User not found"), 404)

    admin_ws_set = {str(w).strip().upper() for w in (admin_ws or []) if str(w).strip()}
    if not admin_ws_set:
        return False, (jsonify(ok=False, error="Forbidden"), 403)

    if target_uname == actor_uname:
        return False, (jsonify(ok=False, error="Cannot modify or delete your own admin account"), 403)

    target_global_role = str(target_user.get("role") or "").strip().lower()
    if target_global_role == "superadmin" or target_uname == "admin":
        return False, (jsonify(ok=False, error="Cannot modify superadmin account"), 403)

    if _is_peer_or_higher_admin(target_user):
        return False, (jsonify(ok=False, error="Cannot modify peer admin account"), 403)

    target_ws = _target_workspace_ids(target_user)
    if not target_ws or not (target_ws & admin_ws_set):
        return False, (jsonify(ok=False, error="Cannot modify user from another department"), 403)

    if require_exclusive_workspace and not (target_ws <= admin_ws_set):
        return False, (jsonify(ok=False, error="Cannot perform global account changes on a multi-department user"), 403)

    return True, None


@users_bp.route("/api/users")
@require_admin
def api_users():
    u = current_user()
    if not u:
        return jsonify(error="LOGIN_REQUIRED"), 403
    users = db.get_users_for_admin(u["username"])
    all_projects = db.get_all_user_projects()
    allowed_pids = None
    if not is_superadmin(u):
        admin_ws = [str(w).strip().upper() for w in (db.get_admin_workspaces_for_user(u["username"]) or [])]
        allowed_pids = {p["id"] for ws in admin_ws for p in (db.get_projects(workspace_id=ws) or [])}
    for usr in users:
        projs = all_projects.get(usr["username"], [])
        if allowed_pids is not None:
            projs = [p for p in projs if p.get("project_id") in allowed_pids]
        usr["projects"] = projs
    return jsonify(users)


@users_bp.route("/api/users", methods=["POST"])
@require_admin
def api_users_action():
    data   = request.get_json(silent=True) or {}
    action = data.get("action")
    u = current_user() or {"username": "admin"}
    is_super = is_superadmin(u)
    admin_ws = db.get_admin_workspaces_for_user(u["username"]) if not is_super else []
    admin_ws_upper = [str(w).strip().upper() for w in (admin_ws or []) if str(w).strip()]

    if action == "add":
        uname = str(data.get("username") or "").strip().lower()
        pw    = data.get("password", "")
        role  = str(data.get("role") or "viewer").strip().lower()
        email = data.get("email", "")
        raw_ws = str(data.get("workspace_id") or data.get("primary_workspace") or "").strip().upper()

        if not uname or not pw:
            return jsonify(ok=False, error="Username and password required"), 400
        is_valid, err_msg, _ = validate_username(uname)
        if not is_valid:
            return jsonify(ok=False, error=err_msg), 400
        if len(str(pw)) < 4:
            return jsonify(ok=False, error="Min 4 characters"), 400

        if not is_super:
            if role == "superadmin":
                return jsonify(ok=False, error="Cannot create superadmin"), 403
            if role not in ("viewer", "editor"):
                return jsonify(ok=False, error="Department admins can only create viewer or editor users"), 403
            if raw_ws and raw_ws not in admin_ws_upper:
                return jsonify(ok=False, error="Cannot create user in another department"), 403
            if not admin_ws_upper:
                return jsonify(ok=False, error="Forbidden"), 403
            ws_id = raw_ws or admin_ws_upper[0]
        else:
            if role not in ("viewer", "editor", "admin", "superadmin"):
                return jsonify(ok=False, error="Invalid role"), 400
            ws_id = raw_ws or "ENG"

        if db.get_user(uname):
            return jsonify(ok=False, error="User already exists"), 409

        can_dl = bool(data.get("can_download", False))
        db.add_user(uname, pw, role, email, primary_workspace=ws_id, can_download=can_dl)
        db.log_action(u["username"], "USER_CREATE",
                      detail=f"Created user '{uname}' (Role: {role}, Dept: {ws_id}, Email: {email or '—'})")
        return jsonify(ok=True)

    if action == "delete":
        uname = str(data.get("username") or "").strip().lower()
        if uname == "admin":
            return jsonify(ok=False, error="Cannot delete super admin"), 400
        target_user = db.get_user(uname)
        if not target_user:
            return jsonify(ok=False, error="User not found"), 404
        if target_user.get("role") == "superadmin":
            return jsonify(ok=False, error="Cannot delete super admin"), 403
        allowed, err_resp = _can_manage_target_user(u, target_user, admin_ws_upper, require_exclusive_workspace=True)
        if not allowed:
            return err_resp

        db.delete_user(uname)
        db.log_action(u["username"], "USER_DELETE", detail=f"Deleted user '{uname}'")
        return jsonify(ok=True)

    if action == "toggle_download_permission":
        uname = str(data.get("username") or "").strip().lower()
        can_dl = bool(data.get("can_download", False))
        target_user = db.get_user(uname)
        if not target_user:
            return jsonify(ok=False, error="User not found"), 404
        if is_super and (target_user.get("role") == "superadmin" or uname == "admin"):
            return jsonify(ok=False, error="Superadmin permissions cannot be modified"), 400
        allowed, err_resp = _can_manage_target_user(u, target_user, admin_ws_upper, require_exclusive_workspace=True)
        if not allowed:
            return err_resp

        db.set_user_can_download(uname, can_dl)
        status_txt = "Allowed" if can_dl else "Restricted"
        db.log_action(u["username"], "DOWNLOAD_PERM_UPDATE",
                      detail=f"Set file download permission for '{uname}' to {status_txt}")
        return jsonify(ok=True, can_download=can_dl)

    if action == "change_password":
        uname = str(data.get("username") or "").strip().lower()
        pw = data.get("password", "")
        target_user = db.get_user(uname)
        if not target_user:
            return jsonify(ok=False, error="User not found"), 404
        allowed, err_resp = _can_manage_target_user(u, target_user, admin_ws_upper, require_exclusive_workspace=True)
        if not allowed:
            return err_resp
        if len(str(pw)) < 4:
            return jsonify(ok=False, error="Min 4 characters"), 400

        db.change_pw(uname, pw)
        db.log_action(u["username"], "PASSWORD_CHANGE", detail=f"Admin changed password for user '{uname}'")
        return jsonify(ok=True)

    if action == "update_role":
        uname = str(data.get("username") or "").strip().lower()
        role = str(data.get("role") or "viewer").strip().lower()
        target_user = db.get_user(uname)
        if not target_user:
            return jsonify(ok=False, error="User not found"), 404
        if uname == "admin" and role != "superadmin":
            return jsonify(ok=False, error="Cannot demote protected super admin"), 400
        if not is_super:
            if role == "superadmin":
                return jsonify(ok=False, error="Cannot promote to superadmin"), 403
            if role not in ("viewer", "editor"):
                return jsonify(ok=False, error="Department admins can only assign viewer or editor roles"), 403
            allowed, err_resp = _can_manage_target_user(u, target_user, admin_ws_upper, require_exclusive_workspace=True)
            if not allowed:
                return err_resp

        if not db.set_user_role(uname, role):
            return jsonify(ok=False, error="Invalid role"), 400
        if not is_super and role in ("viewer", "editor"):
            primary_ws = str(target_user.get("primary_workspace") or "").strip().upper()
            target_ws = primary_ws if primary_ws in admin_ws_upper else (admin_ws_upper[0] if admin_ws_upper else "")
            if target_ws:
                db.set_user_workspace_role(uname, target_ws, role)
        db.log_action(u["username"], "ROLE_UPDATE", detail=f"Updated role of '{uname}' to '{role}'")
        return jsonify(ok=True)

    if action == "update_email":
        uname = str(data.get("username") or "").strip().lower()
        target_user = db.get_user(uname)
        if not target_user:
            return jsonify(ok=False, error="User not found"), 404
        allowed, err_resp = _can_manage_target_user(u, target_user, admin_ws_upper, require_exclusive_workspace=True)
        if not allowed:
            return err_resp

        db.set_user_email(uname, data.get("email", ""))
        db.log_action(u["username"], "USER_EDIT", detail=f"Updated email for '{uname}' to '{data.get('email')}'")
        return jsonify(ok=True)

    if action == "assign":
        uname = str(data.get("username") or "").strip().lower()
        pid = data.get("project_id", "")
        is_dc = data.get("is_dc", False)
        if not is_super:
            proj = db.get_project(pid)
            proj_ws = str(proj.get("workspace_id") or "ENG").upper() if proj else "ENG"
            if not proj or proj_ws not in admin_ws_upper:
                return jsonify(ok=False, error="Cannot assign project from another department"), 403
            target_user = db.get_user(uname)
            if not target_user:
                return jsonify(ok=False, error="User not found"), 404
            allowed, err_resp = _can_manage_target_user(u, target_user, admin_ws_upper, require_exclusive_workspace=False)
            if not allowed:
                return err_resp

        db.assign_project(uname, pid, is_dc)
        if "doc_types" in data:
            db.set_user_project_doc_types(uname, pid, data.get("doc_types"))
        db.log_action(u["username"], "PROJECT_ASSIGN", project_id=pid,
                      detail=f"Assigned project '{pid}' to '{uname}' (DC: {is_dc})")
        return jsonify(ok=True)

    if action == "unassign":
        uname = str(data.get("username") or "").strip().lower()
        pid = data.get("project_id", "")
        if not is_super:
            proj = db.get_project(pid)
            proj_ws = str(proj.get("workspace_id") or "ENG").upper() if proj else "ENG"
            if not proj or proj_ws not in admin_ws_upper:
                return jsonify(ok=False, error="Cannot unassign project from another department"), 403
            target_user = db.get_user(uname)
            if not target_user:
                return jsonify(ok=False, error="User not found"), 404
            allowed, err_resp = _can_manage_target_user(u, target_user, admin_ws_upper, require_exclusive_workspace=False)
            if not allowed:
                return err_resp

        db.unassign_project(uname, pid)
        db.log_action(u["username"], "PROJECT_UNASSIGN", project_id=pid,
                      detail=f"Unassigned project '{pid}' from '{uname}'")
        return jsonify(ok=True)

    if action == "update_workspace_role":
        uname = str(data.get("username") or "").strip().lower()
        ws_id = str(data.get("workspace_id") or "").strip().upper()
        ws_role = str(data.get("role") or "viewer").strip().lower()
        if not uname or not ws_id:
            return jsonify(ok=False, error="Username and workspace_id required"), 400
        if ws_role not in ("viewer", "editor", "admin"):
            return jsonify(ok=False, error="Invalid role"), 400
        target_user = db.get_user(uname)
        if not target_user:
            return jsonify(ok=False, error="User not found"), 404
        if not is_super:
            if ws_id not in admin_ws_upper:
                return jsonify(ok=False, error="Cannot modify department outside your authority"), 403
            if ws_role in ("admin", "superadmin"):
                return jsonify(ok=False, error="Cannot grant admin privileges"), 403
            allowed, err_resp = _can_manage_target_user(u, target_user, admin_ws_upper, require_exclusive_workspace=False)
            if not allowed:
                return err_resp
        elif target_user.get("role") == "superadmin" or uname == "admin":
            return jsonify(ok=False, error="Cannot modify superadmin"), 400

        db.set_user_workspace_role(uname, ws_id, ws_role)
        primary_ws = str(target_user.get("primary_workspace") or "ENG").strip().upper()
        target_ws = _target_workspace_ids(target_user)
        should_sync_global = (
            ws_id == primary_ws
            and ws_role in ("viewer", "editor")
            and str(target_user.get("role") or "").strip().lower() != "superadmin"
            and (is_super or (target_ws <= set(admin_ws_upper)))
        )
        if should_sync_global:
            db.set_user_role(uname, ws_role)
        db.log_action(u["username"], "WORKSPACE_ROLE_UPDATE",
                      detail=f"Updated role of '{uname}' in '{ws_id}' to '{ws_role}'")
        return jsonify(ok=True)

    if action == "set_doc_type_permissions":
        uname = str(data.get("username") or "").strip().lower()
        pid = str(data.get("project_id") or "").strip()
        doc_types = data.get("doc_types")
        if not uname or not pid:
            return jsonify(ok=False, error="Username and project_id required"), 400
        if not is_super:
            proj = db.get_project(pid)
            proj_ws = str(proj.get("workspace_id") or "ENG").upper() if proj else "ENG"
            if not proj or proj_ws not in admin_ws_upper:
                return jsonify(ok=False, error="Cannot modify project outside your department"), 403
        target_user = db.get_user(uname)
        if not target_user:
            return jsonify(ok=False, error="User not found"), 404
        if not is_super:
            allowed, err_resp = _can_manage_target_user(u, target_user, admin_ws_upper, require_exclusive_workspace=False)
            if not allowed:
                return err_resp

        db.set_user_project_doc_types(uname, pid, doc_types)
        dt_desc = "ALL" if doc_types is None else (", ".join(str(d) for d in doc_types) if doc_types else "NONE")
        db.log_action(u["username"], "DOC_TYPE_PERM_UPDATE", project_id=pid,
                      detail=f"Updated allowed doc types for '{uname}' on '{pid}': {dt_desc}")
        return jsonify(ok=True)

    if action == "update_workspaces":
        uname = str(data.get("username") or "").strip().lower()
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
        uname = str(data.get("username") or "").strip().lower()
        ws_id = str(data.get("workspace_id") or data.get("department") or "").strip().upper()
        target_user = db.get_user(uname)
        if not target_user:
            return jsonify(ok=False, error="User not found"), 404
        ws_role = str(data.get("role") or target_user.get("role", "viewer")).strip().lower()
        if not is_super:
            if ws_id not in admin_ws_upper:
                return jsonify(ok=False, error="Cannot assign workspace outside your authority"), 403
            if ws_role == "superadmin":
                return jsonify(ok=False, error="Cannot grant superadmin"), 403
            if ws_role not in ("viewer", "editor"):
                return jsonify(ok=False, error="Cannot grant admin privileges"), 403
            allowed, err_resp = _can_manage_target_user(u, target_user, admin_ws_upper, require_exclusive_workspace=True)
            if not allowed:
                return err_resp
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
        admin_ws = [str(w).strip().upper() for w in (db.get_admin_workspaces_for_user(u["username"]) or [])]
        allowed_pids = {p["id"] for ws in admin_ws for p in (db.get_projects(workspace_id=ws) or [])}
        return jsonify([p for p in db.get_user_projects(username) if p.get("project_id") in allowed_pids])
    return jsonify(db.get_user_projects(username))


@users_bp.route("/api/workspaces", methods=["GET", "POST"])
@require_login
def api_workspaces():
    u = current_user()
    if not u:
        return jsonify(error="LOGIN_REQUIRED"), 403

    if request.method == "POST":
        if not is_superadmin(u):
            return jsonify(error="Only Super Admin can manage departments"), 403
        data = request.get_json(silent=True) or {}
        action = data.get("action", "add")
        
        if action in ("add", "save"):
            code = str(data.get("code") or data.get("id") or "").strip().upper()
            raw_id = str(data.get("id") or code).strip().upper()
            ws_id = re.sub(r'[^A-Z0-9_]', '', raw_id) or raw_id
            name = str(data.get("name") or "").strip()
            name_ar = str(data.get("name_ar") or "").strip()
            description = str(data.get("description") or "").strip()
            icon = str(data.get("icon") or "📁").strip()
            ws_type = str(data.get("type") or "general").strip()
            sort_order = 0
            if "sort_order" in data and data["sort_order"] is not None:
                try: sort_order = int(data["sort_order"])
                except: sort_order = 0
            elif action == "add":
                try:
                    max_order = db.q("SELECT COALESCE(MAX(sort_order), 0) as m FROM workspaces", one=True)
                    if max_order and isinstance(max_order, dict) and max_order.get("m"):
                        sort_order = int(max_order["m"]) + 1
                except Exception:
                    sort_order = 0
            
            if not ws_id or not name:
                return jsonify(error="Department Code/ID and Name are required"), 400
                
            db.create_workspace(ws_id, code, name, name_ar, icon, ws_type, sort_order, description=description)
            drive_provisioned = []
            try:
                ws_projects = db.get_projects(workspace_id=ws_id) or []
                for p in ws_projects:
                    p_id = str((p or {}).get("id") or "").strip()
                    if not p_id:
                        continue
                    if not (db.get_doc_types(p_id) or []):
                        db._seed_project(p_id, ws_id)
                    prov = drive_routing.provision_project_drive_structure(p_id)
                    if not prov.get("skipped"):
                        drive_provisioned.append(prov)
            except Exception:
                pass
            db.log_action(u["username"], "WORKSPACE_CREATE" if action == "add" else "WORKSPACE_UPDATE", detail=f"Managed workspace {ws_id} ({name})")
            return jsonify(ok=True, drive_provisioned=drive_provisioned)
            
        elif action == "update":
            ws_id = str(data.get("id") or "").strip().upper()
            if not ws_id:
                return jsonify(error="Department ID is required"), 400
            code = str(data.get("code") or "").strip().upper() if data.get("code") else None
            name = str(data.get("name") or "").strip() if data.get("name") else None
            name_ar = str(data.get("name_ar") or "").strip() if data.get("name_ar") is not None else None
            description = str(data.get("description") or "").strip() if "description" in data else None
            icon = str(data.get("icon") or "").strip() if data.get("icon") else None
            sort_order = None
            if "sort_order" in data and data["sort_order"] is not None:
                try: sort_order = int(data["sort_order"])
                except: pass
            is_active = data.get("is_active")
            db.update_workspace(ws_id, code=code, name=name, name_ar=name_ar, icon=icon, sort_order=sort_order, is_active=is_active, description=description)
            db.log_action(u["username"], "WORKSPACE_UPDATE", detail=f"Updated workspace {ws_id} ({name or ''})")
            return jsonify(ok=True)
            
        return jsonify(error="Unknown action"), 400

    if is_superadmin(u):
        workspaces = db.get_workspaces(active_only=False)
    else:
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
