"""blueprints/projects.py - DCR Projects API routes with Multi-Workspace support.

Handles all project-level CRUD endpoints:
  GET  /api/projects              -- list projects with can_edit flag (supports ?workspace_id=)
  GET  /api/project/<pid>         -- get one project
  POST /api/project/<pid>         -- save/update one project
  POST /api/projects/create       -- create new project (Superadmin or Dept Admin)
  POST /api/projects/reorder      -- reorder projects
  POST /api/projects/delete/<pid> -- delete project and all its data (Superadmin or Dept Admin)
"""
from flask import Blueprint, jsonify, request

import db
from auth import current_user, require_superadmin, require_admin, is_superadmin, can_edit, can_view_project, is_project_admin, _resolve_project

projects_bp = Blueprint("projects", __name__)


@projects_bp.route("/api/projects")
def api_projects():
    u = current_user()
    if not u:
        return jsonify(error="LOGIN_REQUIRED"), 403
    ws_id = request.args.get("workspace_id")
    projects = db.get_projects_for_user(u["username"], u.get("role"), workspace_id=ws_id)
    out = []
    for p in projects:
        try: data = p.get("data") or {}
        except: data = {}
        out.append({
            "id": p["id"], "name": p["name"], "code": p["code"],
            "workspace_id": p.get("workspace_id", "ENG"),
            "client": data.get("client","") if isinstance(data,dict) else "",
            "can_edit": can_edit(p["id"]),
            "can_admin": is_project_admin(p["id"], u),
        })
    return jsonify(out)


@projects_bp.route("/api/project/<pid>")
def api_get_project(pid):
    if not can_view_project(pid):
        return jsonify(error="Forbidden"), 403
    p = db.get_project(pid)
    if not p: return jsonify(error="Not found"), 404
    return jsonify(p)


@projects_bp.route("/api/project/<pid>", methods=["POST"])
def api_save_project(pid):
    if not can_edit(pid):
        return jsonify(error="LOGIN_REQUIRED"), 403
    u = current_user()
    proj = db.get_project(pid)
    if not proj:
        return jsonify(ok=False, error="Project not found"), 404
    stored_ws = str(proj.get("workspace_id") or "ENG").strip().upper()
    data = request.get_json(silent=True) or {}
    name = data.pop("name", pid) or pid
    code = data.pop("code", pid) or pid
    data.pop("id", None)
    ws_id = data.pop("workspace_id", None)
    if ws_id is not None and str(ws_id).strip():
        req_ws = str(ws_id).strip().upper()
        if req_ws != stored_ws:
            try:
                db.log_action(u.get("username", "") if u else "unknown", "ACCESS DENIED",
                              project_id=pid,
                              detail=f"Rejected project workspace reassignment attempt to '{req_ws}'")
            except Exception:
                pass
            return jsonify(ok=False, error="Cannot change project workspace"), 403
    try:
        db.save_project(pid, name, code, data, workspace_id=stored_ws)
    except db.ProjectOwnershipError:
        return jsonify(ok=False, error="Project workspace ownership conflict"), 409
    return jsonify(ok=True)


@projects_bp.route("/api/projects/create", methods=["POST"])
@require_admin
def api_create_project():
    data = request.get_json(silent=True) or {}
    pid  = data.get("id","").strip().upper()
    name = data.get("name","").strip()
    code = data.get("code","").strip().upper()
    ws_id = str(data.get("workspace_id") or "ENG").strip().upper()
    client = data.get("client", "").strip()
    if not pid or not name or not code:
        return jsonify(ok=False, error="ID, name and code required"), 400
    u = current_user()
    if not is_superadmin(u):
        admin_ws = [str(w).strip().upper() for w in (db.get_admin_workspaces_for_user(u["username"]) or [])]
        if ws_id not in admin_ws:
            return jsonify(ok=False, error="Cannot create project in another workspace"), 403
    try:
        db.create_project(pid, name, code, creator=u["username"], workspace_id=ws_id, client=client)
    except db.ProjectOwnershipError:
        try:
            db.log_action(u.get("username", ""), "ACCESS DENIED", project_id=pid,
                          detail=f"Rejected duplicate/colliding project creation attempt in workspace '{ws_id}'")
        except Exception:
            pass
        # Generic message: never reveal which workspace owns the colliding ID.
        return jsonify(ok=False, error="Project ID already exists"), 409
    return jsonify(ok=True)


@projects_bp.route("/api/projects/reorder", methods=["POST"])
def api_reorder_projects():
    u = current_user()
    if not u:
        return jsonify(error="LOGIN_REQUIRED"), 403
    if not is_superadmin(u) and not db.is_user_workspace_admin(u.get("username")):
        return jsonify(error="Forbidden"), 403
    data = request.get_json(silent=True) or {}
    order = data.get("order")
    if not isinstance(order, list):
        return jsonify(ok=False, error="Invalid project order"), 400
    clean_order = []
    seen = set()
    submitted_projs = []
    for raw_id in order:
        val = str(raw_id or "").strip()
        if not val:
            continue
        proj = _resolve_project(val)
        if not proj:
            return jsonify(ok=False, error="Project not found"), 404
        canon_id = proj["id"]
        if not is_project_admin(canon_id, u):
            try:
                db.log_action(u.get("username", ""), "ACCESS DENIED", project_id=canon_id,
                              detail="Unauthorized project reorder attempt")
            except Exception:
                pass
            return jsonify(error="Forbidden"), 403
        if canon_id not in seen:
            seen.add(canon_id)
            clean_order.append(canon_id)
            submitted_projs.append(proj)

    raw_ws = str(data.get("workspace_id") or "").strip().upper()
    if raw_ws and raw_ws != "ALL":
        scoped_visible = db.get_projects_for_user(
            u.get("username"), role=u.get("role"), workspace_id=raw_ws
        )
    else:
        all_visible = db.get_projects_for_user(u.get("username"), role=u.get("role"))
        if raw_ws == "ALL":
            scoped_visible = all_visible
        else:
            # Scope to workspaces represented in clean_order plus any cross-workspace project
            # with an already-persisted explicit dashboard_order so interleaved read-only cards stay pinned.
            submitted_wss = {
                str(p.get("workspace_id") or "ENG").strip().upper() for p in submitted_projs
            }
            scoped_visible = [
                p for p in all_visible
                if str(p.get("workspace_id") or "ENG").strip().upper() in submitted_wss
                or db._extract_dashboard_order(p.get("data")) is not None
            ]

    # Only truly non-admin (read-only) projects act as fixed interleaved anchors; omitted
    # admin projects must never be treated as immovable barriers at key 999999.
    visible_projects = [
        p for p in scoped_visible
        if p.get("id") in seen or not is_project_admin(p.get("id"), u)
    ]

    merged_order = db.reorder_projects(clean_order, visible_projects=visible_projects)
    return jsonify(ok=True, order=merged_order)


@projects_bp.route("/api/projects/delete/<pid>", methods=["POST"])
@require_admin
def api_delete_project(pid):
    u = current_user()
    if not is_superadmin(u):
        proj = db.get_project(pid)
        if not proj:
            return jsonify(ok=False, error="Project not found"), 404
        admin_ws = [str(w).strip().upper() for w in (db.get_admin_workspaces_for_user(u["username"]) or [])]
        if str(proj.get("workspace_id") or "ENG").strip().upper() not in admin_ws:
            return jsonify(ok=False, error="Cannot delete project from another workspace"), 403
    db.delete_project(pid)
    return jsonify(ok=True)
