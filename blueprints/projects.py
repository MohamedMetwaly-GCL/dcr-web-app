"""blueprints/projects.py - DCR Projects API routes.

Handles all project-level CRUD endpoints:
  GET  /api/projects              -- list all projects with can_edit flag
  GET  /api/project/<pid>         -- get one project
  POST /api/project/<pid>         -- save/update one project
  POST /api/projects/create       -- create new project (superadmin only)
  POST /api/projects/delete/<pid> -- delete project and all its data (superadmin only)

Step 4 of the incremental refactor.
Logic is identical to the original app.py routes — only the decorator
changed from @app.route to @projects_bp.route.
"""
from flask import Blueprint, jsonify, request

import db
from auth import current_user, require_superadmin, can_edit

projects_bp = Blueprint("projects", __name__)


@projects_bp.route("/api/projects")
def api_projects():
    u = current_user()
    projects = db.get_projects()
    out = []
    for p in projects:
        try: data = p.get("data") or {}
        except: data = {}
        out.append({
            "id": p["id"], "name": p["name"], "code": p["code"],
            "client": data.get("client","") if isinstance(data,dict) else "",
            "can_edit": can_edit(p["id"])
        })
    return jsonify(out)

@projects_bp.route("/api/project/<pid>")
def api_get_project(pid):
    p = db.get_project(pid)
    if not p: return jsonify(error="Not found"), 404
    return jsonify(p)

@projects_bp.route("/api/project/<pid>", methods=["POST"])
def api_save_project(pid):
    if not can_edit(pid):
        return jsonify(error="LOGIN_REQUIRED"), 403
    data = request.get_json(silent=True) or {}
    name = data.pop("name", pid) or pid
    code = data.pop("code", pid) or pid
    data.pop("id", None)
    # _labels is valid custom data, keep it
    db.save_project(pid, name, code, data)
    return jsonify(ok=True)

@projects_bp.route("/api/projects/create", methods=["POST"])
@require_superadmin
def api_create_project():
    data = request.get_json(silent=True) or {}
    pid  = data.get("id","").strip().upper()
    name = data.get("name","").strip()
    code = data.get("code","").strip().upper()
    if not pid or not name or not code:
        return jsonify(ok=False, error="ID, name and code required"), 400
    u = current_user()
    db.create_project(pid, name, code, creator=u["username"])
    return jsonify(ok=True)

@projects_bp.route("/api/projects/delete/<pid>", methods=["POST"])
@require_superadmin
def api_delete_project(pid):
    db.delete_project(pid)
    return jsonify(ok=True)
