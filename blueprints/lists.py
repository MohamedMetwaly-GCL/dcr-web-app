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
import logging
from flask import Blueprint, jsonify, request

import db
from auth import current_user, can_edit

lists_bp = Blueprint("lists", __name__)
logger = logging.getLogger(__name__)


@lists_bp.route("/api/lists/<pid>")
def api_lists(pid):
    logger.info("api_lists start pid=%s", pid)
    try:
        db.cleanup_orphan_lists(pid)
    except Exception as e:
        logger.exception("api_lists cleanup failed pid=%s error=%s", pid, e)
    try:
        logger.info("api_lists before get_lists pid=%s", pid)
        payload = db.get_lists(pid)
        logger.info("api_lists after get_lists pid=%s list_count=%s", pid, len(payload))
        return jsonify(payload)
    except Exception as e:
        logger.exception("api_lists get_lists failed pid=%s error=%s", pid, e)
        raise


@lists_bp.route("/api/lists/<pid>", methods=["POST"])
def api_add_list_item(pid):
    if not can_edit(pid): return jsonify(error="LOGIN_REQUIRED"), 403
    data = request.get_json(silent=True) or {}
    if not db.is_allowed_list_name(pid, data.get("list_name","")):
        return jsonify(error="Invalid list"), 400
    db.add_list_item(pid, data.get("list_name",""), data.get("item",""))
    return jsonify(ok=True)


@lists_bp.route("/api/lists/<pid>", methods=["PATCH"])
def api_rename_list_item(pid):
    if not can_edit(pid): return jsonify(error="LOGIN_REQUIRED"), 403
    data = request.get_json(silent=True) or {}
    if not db.is_allowed_list_name(pid, data.get("list_name","")):
        return jsonify(error="Invalid list"), 400
    renamed = db.rename_list_item(pid, data.get("list_name",""), data.get("old_item",""), data.get("new_item",""))
    return jsonify(ok=True, renamed=renamed)


@lists_bp.route("/api/lists/<pid>/reorder", methods=["POST"])
def api_reorder_list_items(pid):
    if not can_edit(pid): return jsonify(error="LOGIN_REQUIRED"), 403
    data = request.get_json(silent=True) or {}
    if not db.is_allowed_list_name(pid, data.get("list_name","")):
        return jsonify(error="Invalid list"), 400
    db.reorder_list_items(pid, data.get("list_name",""), data.get("order", []))
    return jsonify(ok=True)


@lists_bp.route("/api/lists/<pid>", methods=["DELETE"])
def api_remove_list_item(pid):
    if not can_edit(pid): return jsonify(error="LOGIN_REQUIRED"), 403
    data = request.get_json(silent=True) or {}
    if not db.is_allowed_list_name(pid, data.get("list_name","")):
        return jsonify(error="Invalid list"), 400
    db.remove_list_item(pid, data.get("list_name",""), data.get("item",""))
    return jsonify(ok=True)


@lists_bp.route("/api/lists_meta/<pid>")
def api_lists_meta(pid):
    logger.info("api_lists_meta start pid=%s", pid)
    try:
        logger.info("api_lists_meta before get_lists_with_meta pid=%s", pid)
        payload = db.get_lists_with_meta(pid)
        logger.info("api_lists_meta after get_lists_with_meta pid=%s list_count=%s", pid, len(payload))
        return jsonify(payload)
    except Exception as e:
        logger.exception("api_lists_meta failed pid=%s error=%s", pid, e)
        raise


@lists_bp.route("/api/lists_meta/<pid>", methods=["POST"])
def api_set_list_meta(pid):
    u = current_user()
    if not u or u.get("role") not in ("superadmin","admin"):
        return jsonify(error="Admin only"), 403
    data = request.get_json(silent=True) or {}
    if not db.is_allowed_list_name(pid, data.get("list_name","")):
        return jsonify(error="Invalid list"), 400
    db.set_list_item_meta(pid, data.get("list_name",""),
                         data.get("item_value",""), data.get("meta","pending"))
    return jsonify(ok=True)
