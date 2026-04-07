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
from auth import current_user, can_edit

lists_bp = Blueprint("lists", __name__)


@lists_bp.route("/api/lists/<pid>")
def api_lists(pid):
    return jsonify(db.get_lists(pid))


@lists_bp.route("/api/lists/<pid>", methods=["POST"])
def api_add_list_item(pid):
    if not can_edit(pid): return jsonify(error="LOGIN_REQUIRED"), 403
    data = request.get_json(silent=True) or {}
    db.add_list_item(pid, data.get("list_name",""), data.get("item",""))
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
    if not u or u.get("role") not in ("superadmin","admin"):
        return jsonify(error="Admin only"), 403
    data = request.get_json(silent=True) or {}
    db.set_list_item_meta(pid, data.get("list_name",""),
                         data.get("item_value",""), data.get("meta","pending"))
    return jsonify(ok=True)
