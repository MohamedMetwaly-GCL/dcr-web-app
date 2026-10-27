"""blueprints/departments.py - Enterprise Organizational Hierarchy (Departments) API routes.

Handles department-level CRUD and project-department mapping:
  GET  /api/departments   -- list accessible departments with project counts
  POST /api/departments   -- create, update/rename, delete, reorder departments & map projects
"""
import re
from flask import Blueprint, jsonify, request

import db
from auth import current_user

departments_bp = Blueprint("departments", __name__)

DEPT_CODE_RE = re.compile(r"^[A-Z0-9_-]{2,20}$")


def _require_admin():
    u = current_user()
    if not u:
        return None, (jsonify(error="LOGIN_REQUIRED"), 403)
    if str(u.get("role", "")).lower() not in ("superadmin", "admin"):
        try:
            db.log_action(u.get("username", ""), "ACCESS DENIED",
                          detail=f"Unauthorized department management attempt on {request.path}")
        except Exception:
            pass
        return None, (jsonify(error="Forbidden"), 403)
    return u, None


@departments_bp.route("/api/departments")
def api_get_departments():
    u = current_user()
    if not u:
        return jsonify(error="LOGIN_REQUIRED"), 403
    depts = db.get_departments_for_user(u["username"], u.get("role"))
    projs = db.get_projects_for_user(u["username"], u.get("role"))
    proj_counts = {}
    for p in projs:
        did = str(p.get("department_id") or "ENG").strip().upper()
        proj_counts[did] = proj_counts.get(did, 0) + 1

    out = []
    for d in depts:
        did = str(d.get("id", "")).strip().upper()
        mod_type = db._normalize_module_type(d.get("module_type"), did, d.get("is_project_based", True))
        out.append({
            "id": did,
            "code": d.get("code") or did,
            "name": d.get("name") or did,
            "name_ar": d.get("name_ar") or "",
            "icon": d.get("icon") or "🏢",
            "description": d.get("description") or "",
            "module_type": mod_type,
            "is_project_based": mod_type == "engineering",
            "sort_order": d.get("sort_order", 0),
            "is_active": bool(d.get("is_active", True)),
            "project_count": proj_counts.get(did, 0),
        })
    return jsonify(out)


@departments_bp.route("/api/departments", methods=["POST"])
def api_departments_action():
    u, err = _require_admin()
    if err:
        return err

    data = request.get_json(silent=True) or {}
    action = str(data.get("action") or "").strip().lower()

    if action in ("create", "update"):
        raw_id = str(data.get("id") or data.get("code") or "").strip().upper()
        raw_code = str(data.get("code") or raw_id).strip().upper()
        name = str(data.get("name") or "").strip()
        name_ar = str(data.get("name_ar") or "").strip()
        icon = str(data.get("icon") or "🏢").strip() or "🏢"
        description = str(data.get("description") or "").strip()
        module_type = str(data.get("module_type") or "").strip().lower()
        is_project_based = bool(data.get("is_project_based", True))
        if module_type:
            is_project_based = (module_type == "engineering")
        sort_order = data.get("sort_order")

        if not raw_id or not DEPT_CODE_RE.match(raw_id):
            return jsonify(ok=False, error="Department ID/Code must be 2–20 uppercase letters, numbers, hyphens, or underscores (e.g. HR, FIN, PROC)."), 400
        if not name:
            return jsonify(ok=False, error="Department name is required."), 400

        if action == "create" and db.get_department(raw_id):
            return jsonify(ok=False, error=f"Department '{raw_id}' already exists."), 400

        dept = db.save_department(
            dept_id=raw_id,
            code=raw_code,
            name=name,
            name_ar=name_ar,
            icon=icon,
            description=description,
            is_project_based=is_project_based,
            sort_order=sort_order,
            module_type=module_type or None,
        )
        act_type = "DEPT_CREATE" if action == "create" else "DEPT_EDIT"
        db.log_action(
            u["username"], act_type,
            detail=f"{'Created' if action == 'create' else 'Updated'} department '{name}' ({raw_code}, module={dept.get('module_type')})"
                   + (f" - {name_ar}" if name_ar else "")
        )
        return jsonify(ok=True, department=dept)

    if action == "delete":
        dept_id = str(data.get("id") or "").strip().upper()
        if not dept_id:
            return jsonify(ok=False, error="Department ID required"), 400
        old_dept = db.get_department(dept_id)
        ok, msg = db.delete_department(dept_id)
        if not ok:
            return jsonify(ok=False, error=msg), 400
        dept_label = f"{old_dept.get('name')} ({dept_id})" if old_dept else dept_id
        db.log_action(u["username"], "DEPT_DELETE", detail=f"Deleted department '{dept_label}'")
        return jsonify(ok=True)

    if action == "reorder":
        order = data.get("order") or []
        if not isinstance(order, list):
            return jsonify(ok=False, error="Invalid order list"), 400
        db.reorder_departments(order)
        return jsonify(ok=True)

    if action == "assign_project":
        pid = str(data.get("project_id") or "").strip()
        dept_id = str(data.get("department_id") or "ENG").strip().upper()
        if not pid or not dept_id:
            return jsonify(ok=False, error="project_id and department_id required"), 400
        db.set_project_department(pid, dept_id)
        db.log_action(
            u["username"], "PROJECT_DEPT_CHANGE", project_id=pid,
            detail=f"Assigned project '{pid}' to department '{dept_id}'"
        )
        return jsonify(ok=True)

    if action == "reset_workspace_template":
        pid = str(data.get("project_id") or "").strip()
        mod_type = str(data.get("module_type") or "corporate_edms").strip().lower()
        if not pid:
            return jsonify(ok=False, error="project_id required"), 400
        ok, msg = db.reset_workspace_to_module_template(pid, mod_type)
        if not ok:
            return jsonify(ok=False, error=msg), 400
        db.log_action(
            u["username"], "DEPT_EDIT", project_id=pid,
            detail=f"Initialized workspace '{pid}' with '{mod_type}' document categories"
        )
        return jsonify(ok=True)

    return jsonify(ok=False, error="Unknown action"), 400
