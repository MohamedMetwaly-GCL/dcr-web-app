"""blueprints/lists.py - DCR Lists API routes.

Handles list endpoints:
  GET    /api/lists/<pid>
  POST   /api/lists/<pid>
  DELETE /api/lists/<pid>
  GET    /api/lists_meta/<pid>
  POST   /api/lists_meta/<pid>

Step 9 of the incremental refactor.
Logic is identical to the original app.py routes — only the decorator
changed from @app.route to @lists_bp.route.
"""
from flask import Blueprint, jsonify, request

import db
from auth import current_user, can_edit, is_project_admin

lists_bp = Blueprint("lists", __name__)


@lists_bp.route("/api/lists/<pid>")
def api_lists(pid):
    return jsonify(db.get_lists(pid))


@lists_bp.route("/api/lists/<pid>", methods=["POST"])
def api_add_list_item(pid):
    if not can_edit(pid): return jsonify(error="LOGIN_REQUIRED"), 403
    data = request.get_json(silent=True) or {}
    list_name = str(data.get("list_name", "") or "").strip()
    dt_id = str(data.get("dt_id", "") or "").strip()
    res = db.add_list_item(pid, list_name, data.get("item", ""), meta=data.get("meta"))
    column_wired = False
    if dt_id and list_name and (list_name.lower().endswith("_subtype") or list_name.lower() == "subtype"):
        column_wired = bool(db.ensure_doc_type_subtype_column(pid, dt_id, list_name))
    return jsonify(ok=True, item=res.get("item", ""), meta=res.get("meta"), list_name=list_name, column_wired=column_wired)


@lists_bp.route("/api/lists/<pid>", methods=["PATCH"])
def api_rename_list_item(pid):
    if not can_edit(pid): return jsonify(error="LOGIN_REQUIRED"), 403
    data = request.get_json(silent=True) or {}
    renamed = db.rename_list_item(
        pid,
        data.get("list_name",""),
        data.get("old_item",""),
        data.get("new_item",""),
        meta=data.get("meta"),
    )
    return jsonify(ok=True, renamed=renamed)


@lists_bp.route("/api/lists/<pid>/reorder", methods=["POST"])
def api_reorder_list_items(pid):
    if not can_edit(pid): return jsonify(error="LOGIN_REQUIRED"), 403
    data = request.get_json(silent=True) or {}
    db.reorder_list_items(pid, data.get("list_name",""), data.get("order", []))
    return jsonify(ok=True)


@lists_bp.route("/api/lists/<pid>", methods=["DELETE"])
def api_remove_list_item(pid):
    if not can_edit(pid): return jsonify(error="LOGIN_REQUIRED"), 403
    data = request.get_json(silent=True) or {}
    db.remove_list_item(pid, data.get("list_name",""), data.get("item",""))
    return jsonify(ok=True)


@lists_bp.route("/api/lists_meta/<pid>")
def api_lists_meta(pid):
    return jsonify(db.get_lists_with_meta(pid))


@lists_bp.route("/api/lists_meta/<pid>", methods=["POST"])
def api_set_list_meta(pid):
    u = current_user()
    if not u or not is_project_admin(pid, u):
        return jsonify(error="Admin only"), 403
    data = request.get_json(silent=True) or {}
    db.set_list_item_meta(pid, data.get("list_name",""),
                         data.get("item_value",""), data.get("meta","pending"))
    return jsonify(ok=True)
