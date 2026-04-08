"""blueprints/columns.py - DCR Columns API routes.

Handles all column configuration and width endpoints:
  GET  /api/columns/<pid>/<dt_id>                 -- list columns for a doc type
  POST /api/columns/<pid>/<dt_id>                 -- add a custom column
  POST /api/columns/visibility/<int:col_id>        -- show/hide a column
  DEL  /api/columns/<int:col_id>                  -- delete a column
  POST /api/columns/rename/<int:col_id>            -- rename a column label
  GET  /api/col_width/<pid>/<dt_id>               -- get column widths
  POST /api/col_width/<pid>/<dt_id>               -- save column widths
  POST /api/columns/reorder/<pid>/<dt_id>          -- reorder columns (array of ids)
  POST /api/columns/reorder                        -- reorder single column by position

Step 6 of the incremental refactor.
Logic is identical to the original app.py routes — only the decorator
changed from @app.route to @columns_bp.route.
"""
import uuid

from flask import Blueprint, request, jsonify

import db
from auth import can_edit, current_user

columns_bp = Blueprint("columns", __name__)


def _col_project_id(col_id):
    row = db.q("SELECT project_id FROM columns_config WHERE id=%s", (col_id,), one=True)
    return row["project_id"] if row else None


@columns_bp.route("/api/columns/<pid>/<dt_id>")
def api_columns(pid, dt_id):
    return jsonify(db.get_columns(pid, dt_id))

@columns_bp.route("/api/columns/<pid>/<dt_id>", methods=["POST"])
def api_add_column(pid, dt_id):
    if not can_edit(pid): return jsonify(error="LOGIN_REQUIRED"), 403
    data = request.get_json(silent=True) or {}
    key  = "custom_" + data.get("label","").lower().replace(" ","_") + "_" + uuid.uuid4().hex[:6]
    db.add_column(pid, dt_id, key, data.get("label",""),
                  data.get("col_type","text"), data.get("list_name"))
    return jsonify(ok=True)

@columns_bp.route("/api/columns/visibility/<int:col_id>", methods=["POST"])
def api_col_visibility(col_id):
    pid = _col_project_id(col_id)
    if not pid: return jsonify(error="Column not found"), 404
    if not can_edit(pid): return jsonify(error="LOGIN_REQUIRED"), 403
    data = request.get_json(silent=True) or {}
    db.set_col_visible(col_id, data.get("visible", True))
    return jsonify(ok=True)

@columns_bp.route("/api/columns/<int:col_id>", methods=["DELETE"])
def api_delete_col(col_id):
    pid = _col_project_id(col_id)
    if not pid: return jsonify(error="Column not found"), 404
    if not can_edit(pid): return jsonify(error="LOGIN_REQUIRED"), 403
    db.delete_col(col_id)
    return jsonify(ok=True)

@columns_bp.route("/api/columns/rename/<int:col_id>", methods=["POST"])
def api_rename_col(col_id):
    pid = _col_project_id(col_id)
    if not pid: return jsonify(error="Column not found"), 404
    if not can_edit(pid): return jsonify(error="LOGIN_REQUIRED"), 403
    data = request.get_json(silent=True) or {}
    label = data.get("label","").strip()
    if not label: return jsonify(error="Label required"), 400
    db.rename_column(col_id, label)
    return jsonify(ok=True)

@columns_bp.route("/api/col_width/<pid>/<dt_id>", methods=["GET"])
def api_get_col_widths(pid, dt_id):
    return jsonify(db.get_col_widths(pid, dt_id))

@columns_bp.route("/api/col_width/<pid>/<dt_id>", methods=["POST"])
def api_save_col_width(pid, dt_id):
    u = current_user()
    if not u or u.get("role") not in ("superadmin","admin"):
        return jsonify(error="Superadmin only"), 403
    data = request.get_json(silent=True) or {}
    col_key   = data.get("col_key","")
    width_px  = int(data.get("width_px", 120))
    if not col_key: return jsonify(error="col_key required"), 400
    db.save_col_width(pid, dt_id, col_key, max(40, min(600, width_px)))
    return jsonify(ok=True)

@columns_bp.route("/api/columns/reorder/<pid>/<dt_id>", methods=["POST"])
def api_reorder_cols(pid, dt_id):
    if not can_edit(pid): return jsonify(error="LOGIN_REQUIRED"), 403
    data = request.get_json(silent=True) or {}
    db.reorder_columns(pid, dt_id, data.get("order", []))
    return jsonify(ok=True)

@columns_bp.route("/api/columns/reorder", methods=["POST"])
def api_col_reorder():
    data   = request.get_json(silent=True) or {}
    pid    = data.get("pid")
    dt_id  = data.get("dt_id")
    col_id = data.get("col_id")
    new_order = data.get("new_order", 0)
    if not can_edit(pid): return jsonify(error="LOGIN_REQUIRED"), 403
    # Get all columns and rebuild sort_order
    cols = db.get_columns(pid, dt_id)
    # Move col_id to new_order position
    moving = next((c for c in cols if c["id"]==col_id), None)
    if not moving: return jsonify(ok=False, error="Column not found"), 404
    cols = [c for c in cols if c["id"]!=col_id]
    cols.insert(int(new_order), moving)
    for i, c in enumerate(cols):
        db.exe("UPDATE columns_config SET sort_order=%s WHERE id=%s", (i, c["id"]))
    return jsonify(ok=True)
