"""auth.py - DCR authentication and authorisation helpers.

Contains the four auth primitives used throughout app.py:

  current_user()        -- read the session cookie, return user dict or None
  require_login         -- decorator: reject unauthenticated requests with 403
  require_superadmin    -- decorator: reject non-superadmin requests with 403
  can_edit(project_id)  -- return True if the current user may write to a project

Design notes
------------
- No Flask app object is imported here, so there is zero circular-import risk.
  Flask's `request` and `jsonify` are imported directly from the flask package,
  which is always safe to do outside the application factory.
- `db` is imported at module level. db.py has no dependency on auth.py or app.py,
  so the import graph is a clean DAG: app → auth → db.
- This module has no side effects on import (no global state, no DB calls).

Step 2 of the incremental refactor. Only auth logic lives here.
"""
from functools import wraps

from flask import request, jsonify

import db


def current_user():
    """Return the current user dict {username, role} or None if not logged in.

    Reads the dcr_token cookie and validates it against the sessions table.
    Called on every request that needs to know who the caller is.
    """
    try:
        token = request.cookies.get("dcr_token")
        if not token:
            return None
        return db.get_session(token)
    except RuntimeError:
        return None


def require_login(fn):
    """Decorator: allow only authenticated users.

    Returns HTTP 403 + {error: "LOGIN_REQUIRED"} for unauthenticated callers.
    The frontend JS (apiFetch) watches for LOGIN_REQUIRED and redirects to /login.
    """
    @wraps(fn)
    def wrapper(*a, **kw):
        if not current_user():
            return jsonify(error="LOGIN_REQUIRED"), 403
        return fn(*a, **kw)
    return wrapper


def _resolve_project(project_id):
    """Look up a project by canonical id (or unambiguous code/name alias) and return its normalized dict or None."""
    pid = str(project_id or "").strip()
    if not pid:
        return None
    return db.resolve_project_alias(pid)


def get_allowed_project_ids(user=None, workspace_id=None):
    """Return the canonical project ids the given/current user may view."""
    u = user or current_user()
    if not u:
        return []
    if is_superadmin(u):
        return [p["id"] for p in db.get_projects(workspace_id=workspace_id)]

    # Check if user is admin of the requested workspace:
    if workspace_id:
        if db.is_user_workspace_admin(u.get("username"), workspace_id):
            return [p["id"] for p in db.get_projects(workspace_id=workspace_id)]
    else:
        # Check if user is admin in any workspace
        admin_ws = db.get_admin_workspaces_for_user(u.get("username")) if hasattr(db, "get_admin_workspaces_for_user") else []
        if admin_ws:
            all_projs = []
            seen_admin = set()
            for aws in admin_ws:
                for p in db.get_projects(workspace_id=aws):
                    if p["id"] not in seen_admin:
                        seen_admin.add(p["id"])
                        all_projs.append(p["id"])
            user_projs = db.get_user_projects(u.get("username"))
            assigned_raw = [p["project_id"] for p in user_projs if p.get("project_id")]
            if assigned_raw:
                resolved_ids = db.resolve_project_identifiers(assigned_raw)
                if resolved_ids:
                    for p in db.get_projects(resolved_ids):
                        if p["id"] not in seen_admin:
                            seen_admin.add(p["id"])
                            all_projs.append(p["id"])
            return all_projs

    user_projs = db.get_user_projects(u.get("username"))
    assigned_raw = [p["project_id"] for p in user_projs if p.get("project_id")]
    if not assigned_raw:
        return []
    resolved_ids = db.resolve_project_identifiers(assigned_raw, workspace_id=workspace_id)
    if not resolved_ids:
        return []
    matched_projs = db.get_projects(resolved_ids, workspace_id=workspace_id)
    if not matched_projs:
        return []
    resolved = []
    seen = set()
    for p in matched_projs:
        pid = p.get("id")
        if pid and pid not in seen:
            seen.add(pid)
            resolved.append(pid)
    return resolved


def has_project_access(username, project_id, role=None):
    """Return True if the user may view the project."""
    if not username or not project_id:
        return False
    user_role = str(role or "").strip().lower()
    if not user_role:
        u = db.get_user(username)
        user_role = str((u or {}).get("role") or "").strip().lower()
    if user_role == "superadmin":
        return True
    proj = _resolve_project(project_id)
    if not proj:
        return False
    canon_id = proj["id"]
    ws_id = proj.get("workspace_id") or "ENG"
    if db.is_user_workspace_admin(username, ws_id):
        return True
    user_projects = db.get_user_projects(username)
    assigned_raw = [p["project_id"] for p in user_projects if p.get("project_id")]
    if not assigned_raw:
        return False
    canonical_assigned = set(db.resolve_project_identifiers(assigned_raw))
    return canon_id in canonical_assigned


def can_view_project(project_id, user=None):
    """Return True if the current/given user may view project_id."""
    u = user or current_user()
    if not u or not project_id:
        return False
    if is_superadmin(u):
        return True
    return has_project_access(u.get("username"), project_id, u.get("role"))


def is_project_admin(project_id, user=None):
    """Return True if the current/given user is superadmin or admin of project_id's workspace."""
    u = user or current_user()
    if not u or not project_id:
        return False
    if is_superadmin(u):
        return True
    proj = _resolve_project(project_id)
    if not proj:
        return False
    ws_id = proj.get("workspace_id") or "ENG"
    return db.is_user_workspace_admin(u.get("username"), ws_id)


def get_allowed_doc_type_filter(project_id, user=None):
    """Return None when user can access all doc types in project_id, else list of allowed IDs/codes.

    Fails closed (`[]`) when unauthenticated, project is missing/ambiguous, or the user holds
    no canonical/unambiguous assignment row for the project.
    """
    u = user or current_user()
    if not u or not project_id:
        return []
    proj = _resolve_project(project_id)
    if not proj:
        return []
    canon_pid = proj["id"]
    if is_project_admin(canon_pid, u):
        return None
    return db.get_user_project_doc_types(u.get("username"), canon_pid)


def can_access_doc_type(project_id, dt_id, user=None):
    """Return True if user may access dt_id within project_id."""
    u = user or current_user()
    if not u or not project_id or not dt_id:
        return False
    proj = _resolve_project(project_id)
    canon_pid = proj["id"] if proj else project_id
    if not can_view_project(canon_pid, u):
        return False
    allowed_dts = get_allowed_doc_type_filter(canon_pid, user=u)
    if allowed_dts is None:
        return True
    allowed_set = {str(x).strip().upper() for x in allowed_dts if str(x).strip()}
    if not allowed_set:
        return False
    dt_obj = db.get_doc_type(canon_pid, dt_id)
    dt_code = str(dt_obj.get("code") or "").strip().upper() if dt_obj else ""
    return str(dt_id).strip().upper() in allowed_set or (bool(dt_code) and dt_code in allowed_set)


def filter_allowed_doc_types(project_id, doc_types, user=None):
    """Filter a list of doc_type dicts to those permitted for user in project_id."""
    u = user or current_user()
    if not u:
        return []
    allowed_dts = get_allowed_doc_type_filter(project_id, user=u)
    if allowed_dts is None:
        return list(doc_types or [])
    allowed_set = {str(x).strip().upper() for x in allowed_dts if str(x).strip()}
    if not allowed_set:
        return []
    return [
        d for d in (doc_types or [])
        if str(d.get("id", "")).strip().upper() in allowed_set
        or str(d.get("code", "")).strip().upper() in allowed_set
    ]


def get_user_doc_type_scope(user=None, project_ids=None):
    """Return None if user has unrestricted doc-type access across all project_ids,
    else a dict {pid: set_of_uppercase_allowed_ids_or_None} for filtering aggregates."""
    u = user or current_user()
    if not u:
        return {}
    if is_superadmin(u):
        return None
    scope = {}
    has_restriction = False
    for raw_pid in (project_ids or []):
        pid = str(raw_pid or "").strip()
        if not pid:
            continue
        allowed = get_allowed_doc_type_filter(pid, user=u)
        if allowed is None:
            scope[pid] = None
            continue
        has_restriction = True
        allowed_set = {str(x).strip().upper() for x in allowed if str(x).strip()}
        expanded = set(allowed_set)
        if allowed_set:
            for dt in (db.get_doc_types(pid) or []):
                dt_id = str(dt.get("id", "")).strip().upper()
                dt_code = str(dt.get("code", "")).strip().upper()
                if dt_id in allowed_set or dt_code in allowed_set:
                    if dt_id:
                        expanded.add(dt_id)
                    if dt_code:
                        expanded.add(dt_code)
        scope[pid] = expanded
    return scope if has_restriction else None


def is_superadmin(user=None):
    """Return True if the current or specified user is a superadmin."""
    u = user or current_user()
    if not u:
        return False
    return str(u.get("role", "")).strip().lower() == "superadmin"


def is_workspace_admin(username, workspace_id=None):
    """Return True if the user is superadmin or admin of the given workspace (or any workspace if None)."""
    if not username:
        return False
    return db.is_user_workspace_admin(username, workspace_id)


def has_workspace_access(username, workspace_id):
    """Return True if user can access this workspace."""
    if not username or not workspace_id:
        return False
    u = db.get_user(username)
    if not u:
        return False
    if str(u.get("role", "")).strip().lower() == "superadmin":
        return True
    workspaces = db.get_user_workspaces(username)
    return any(w.get("id") == workspace_id for w in workspaces)


def require_superadmin(fn):
    """Decorator: allow only users with role == 'superadmin'."""
    @wraps(fn)
    def wrapper(*a, **kw):
        u = current_user()
        if not u or u.get("role") != "superadmin":
            if u:
                try:
                    db.log_action(u.get("username", ""), "ACCESS DENIED",
                                  detail=f"Unauthorized superadmin attempt on {request.path}")
                except Exception:
                    pass
            return jsonify(error="Forbidden"), 403
        return fn(*a, **kw)
    return wrapper


def require_admin(fn):
    """Decorator: allow superadmin or any department admin."""
    @wraps(fn)
    def wrapper(*a, **kw):
        u = current_user()
        if not u:
            return jsonify(error="LOGIN_REQUIRED"), 403
        if not is_superadmin(u) and not db.is_user_workspace_admin(u.get("username")):
            try:
                db.log_action(u.get("username", ""), "ACCESS DENIED",
                              detail=f"Unauthorized admin attempt on {request.path}")
            except Exception:
                pass
            return jsonify(error="Forbidden"), 403
        return fn(*a, **kw)
    return wrapper


def can_edit(project_id, user=None):
    """Return True if the current/given user has write access to project_id."""
    u = user or current_user()
    if not u or not project_id:
        return False
    if is_superadmin(u):
        return True
    proj = _resolve_project(project_id)
    if not proj:
        return False
    ws_id = proj.get("workspace_id") or "ENG"

    # Workspace Admin check
    if db.is_user_workspace_admin(u.get("username"), ws_id):
        return True

    # User must have access to the project
    if not has_project_access(u.get("username"), project_id, u.get("role")):
        return False

    # Check user's role in this specific workspace
    ws_role = db.get_user_workspace_role(u.get("username"), ws_id) if hasattr(db, "get_user_workspace_role") else (u.get("role") or "viewer")
    return ws_role in ("editor", "admin", "superadmin")
