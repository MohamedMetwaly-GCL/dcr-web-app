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
    token = request.cookies.get("dcr_token")
    if not token:
        return None
    return db.get_session(token)


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


def get_allowed_project_ids(user=None):
    """Return the project ids the given/current user may view."""
    u = user or current_user()
    if not u:
        return []
    if u["role"] in ("superadmin", "admin"):
        return [p["id"] for p in db.get_projects()]
    return [p["project_id"] for p in db.get_user_projects(u["username"])]


def has_project_access(username, project_id, role=None):
    """Return True if the user may view the project."""
    if not username or not project_id:
        return False
    user_role = str(role or "").strip().lower()
    if user_role in ("superadmin", "admin"):
        return True
    return project_id in [p["project_id"] for p in db.get_user_projects(username)]


def can_view_project(project_id, user=None):
    """Return True if the current/given user may view project_id."""
    u = user or current_user()
    if not u or not project_id:
        return False
    if u["role"] in ("superadmin", "admin"):
        return True
    return has_project_access(u["username"], project_id, u.get("role"))


def require_superadmin(fn):
    """Decorator: allow only users with role == 'superadmin'."""
    @wraps(fn)
    def wrapper(*a, **kw):
        u = current_user()
        if not u or u["role"] != "superadmin":
            return jsonify(error="Forbidden"), 403
        return fn(*a, **kw)
    return wrapper


def can_edit(project_id):
    """Return True if the current user has write access to project_id."""
    u = current_user()
    if not u:
        return False
    if u["role"] in ("superadmin", "admin"):
        return True
    if u["role"] == "editor":
        return project_id in [p["project_id"] for p in db.get_user_projects(u["username"])]
    return False
