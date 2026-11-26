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
import drive_routing
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
    try:
        p = dict(p)
        p["master_drive_root_active"] = bool(db.get_master_drive_root_id())
    except Exception:
        pass
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
    old_drv_id = drive_routing.extract_drive_folder_id(proj.get("drive_folder_id"))
    old_name = str(proj.get("name") or "")
    old_code = str(proj.get("code") or "")
    was_auto_drv = bool(proj.get("auto_provisioned_drive"))
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
    if "drive_folder_id" in data:
        data["drive_folder_id"] = drive_routing.extract_drive_folder_id(data.get("drive_folder_id"))
    new_drv_id = drive_routing.extract_drive_folder_id(
        data.get("drive_folder_id") if "drive_folder_id" in data else old_drv_id
    )
    try:
        db.save_project(pid, name, code, data, workspace_id=stored_ws)
    except db.ProjectOwnershipError:
        return jsonify(ok=False, error="Project workspace ownership conflict"), 409
    if old_drv_id and old_drv_id != new_drv_id:
        db.clear_project_drive_caches(pid)
    drive_rename = {"ok": True, "skipped": True, "renamed": False}
    if was_auto_drv and old_drv_id and new_drv_id == old_drv_id and (str(name) != old_name or str(code) != old_code):
        try:
            drive_rename = drive_routing.rename_project_drive_folder(None, pid, name, new_project_code=code)
        except Exception:
            drive_rename = {"ok": False, "skipped": False, "renamed": False}
    drive_prov = {"ok": True, "skipped": True, "reason": "no_drive_folder", "created": 0, "existing": 0, "errors": []}
    has_master_root = bool(db.get_master_drive_root_id()) if hasattr(db, "get_master_drive_root_id") else False
    if new_drv_id or has_master_root:
        try:
            drive_prov = drive_routing.provision_project_drive_structure(pid)
        except Exception as err:
            drive_prov = {"ok": False, "skipped": False, "created": 0, "existing": 0, "errors": [str(err)]}
    return jsonify(ok=True, drive_provision=drive_prov, drive_rename=drive_rename)


api_update_project = api_save_project


@projects_bp.route("/api/projects/create", methods=["POST"])
@require_admin
def api_create_project():
    data = request.get_json(silent=True) or {}
    pid  = data.get("id","").strip().upper()
    name = data.get("name","").strip()
    code = data.get("code","").strip().upper()
    ws_id = str(data.get("workspace_id") or "ENG").strip().upper()
    client = data.get("client", "").strip()
    drv_id = drive_routing.extract_drive_folder_id(data.get("drive_folder_id"))
    if not pid or not name or not code:
        return jsonify(ok=False, error="ID, name and code required"), 400
    u = current_user()
    if not is_superadmin(u):
        admin_ws = [str(w).strip().upper() for w in (db.get_admin_workspaces_for_user(u["username"]) or [])]
        if ws_id not in admin_ws:
            return jsonify(ok=False, error="Cannot create project in another workspace"), 403
    create_kwargs = {"creator": u["username"], "workspace_id": ws_id, "client": client}
    if drv_id:
        create_kwargs["drive_folder_id"] = drv_id
    try:
        db.create_project(pid, name, code, **create_kwargs)
    except TypeError:
        db.create_project(pid, name, code, creator=u["username"], workspace_id=ws_id, client=client)
        if drv_id:
            db.update_project_settings(pid, {"drive_folder_id": drv_id})
    except db.ProjectOwnershipError:
        try:
            db.log_action(u.get("username", ""), "ACCESS DENIED", project_id=pid,
                          detail=f"Rejected duplicate/colliding project creation attempt in workspace '{ws_id}'")
        except Exception:
            pass
        # Generic message: never reveal which workspace owns the colliding ID.
        return jsonify(ok=False, error="Project ID already exists"), 409
    drive_prov = {"ok": True, "skipped": True, "reason": "no_drive_folder", "created": 0, "existing": 0, "errors": []}
    has_master_root = bool(db.get_master_drive_root_id()) if hasattr(db, "get_master_drive_root_id") else False
    if drv_id or has_master_root:
        try:
            drive_prov = drive_routing.provision_project_drive_structure(pid)
        except Exception as err:
            drive_prov = {"ok": False, "skipped": False, "created": 0, "existing": 0, "errors": [str(err)]}
    return jsonify(ok=True, drive_provision=drive_prov)


@projects_bp.route("/api/projects/<pid>/init-drive-folders", methods=["POST"])
def api_init_project_drive_folders(pid):
    u = current_user()
    if not u:
        return jsonify(ok=False, error="LOGIN_REQUIRED"), 403
    if not can_edit(pid, u):
        try:
            db.log_action(u.get("username", ""), "ACCESS DENIED", project_id=pid,
                          detail="Unauthorized Drive folder initialization attempt")
        except Exception:
            pass
        return jsonify(ok=False, error="Forbidden"), 403

    proj = db.get_project(pid)
    if not proj:
        return jsonify(ok=False, error="Project not found"), 404

    body = request.get_json(silent=True) or {}
    req_fid = drive_routing.extract_drive_folder_id(body.get("drive_folder_id") or body.get("folder_id"))
    old_fid = drive_routing.extract_drive_folder_id(proj.get("drive_folder_id"))
    if req_fid and req_fid != old_fid:
        db.update_project_settings(pid, {"drive_folder_id": req_fid})
        if old_fid:
            db.clear_project_drive_caches(pid)
        old_fid = req_fid

    has_master_root = bool(db.get_master_drive_root_id()) if hasattr(db, "get_master_drive_root_id") else False
    if not old_fid and not has_master_root:
        return jsonify(
            ok=False,
            error="Project has no Drive Folder ID configured",
            reason="no_drive_folder",
            created=0,
            existing=0,
            errors=["no_drive_folder"],
        ), 400

    res = drive_routing.provision_project_drive_structure(pid)
    if res.get("skipped") and res.get("reason") == "drive_not_configured":
        return jsonify(
            ok=False,
            error="Google Drive credentials are not configured",
            reason="drive_not_configured",
            created=0,
            existing=0,
            errors=["drive_not_configured"],
        ), 503

    try:
        db.log_action(
            u.get("username", ""),
            "DRIVE_FOLDERS_INIT",
            project_id=pid,
            detail=f"Initialized project Drive folders (created={res.get('created', 0)}, existing={res.get('existing', 0)}, errors={len(res.get('errors') or [])})",
        )
    except Exception:
        pass

    return jsonify(
        ok=bool(res.get("ok", True)),
        project_id=pid,
        drive_folder_id=res.get("drive_folder_id") or old_fid,
        created=int(res.get("created") or 0),
        existing=int(res.get("existing") or 0),
        errors=list(res.get("errors") or []),
    )


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
