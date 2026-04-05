"""blueprints/doc_types.py - DCR Doc Types API routes.

Handles all document-type management endpoints:
  GET    /api/doc_types/<pid>              -- list doc types for a project
  POST   /api/doc_types/<pid>             -- add a new doc type
  PATCH  /api/doc_types/<pid>/<dt_id>     -- rename a doc type (code/name)
  POST   /api/doc_types/<pid>/reorder     -- reorder doc types
  DELETE /api/doc_types/<pid>/<dt_id>     -- delete a doc type and its records

Step 5 of the incremental refactor.
Logic is identical to the original app.py routes — only the decorator
changed from @app.route to @doc_types_bp.route.
"""
from flask import Blueprint, request, jsonify

import db
from auth import can_edit

doc_types_bp = Blueprint("doc_types", __name__)


@doc_types_bp.route("/api/doc_types/<pid>")
def api_doc_types(pid):
    return jsonify(db.get_doc_types(pid))

@doc_types_bp.route("/api/doc_types/<pid>", methods=["POST"])
def api_add_doc_type(pid):
    if not can_edit(pid): return jsonify(error="LOGIN_REQUIRED"), 403
    data = request.get_json(silent=True) or {}
    code = data.get("code","").strip().upper()
    name = data.get("name","").strip()
    if not code or not name:
        return jsonify(ok=False, error="Code and name required"), 400
    db.add_doc_type(pid, code, name)
    return jsonify(ok=True)


@doc_types_bp.route("/api/doc_types/<pid>/<dt_id>", methods=["PATCH"])
def api_rename_doc_type(pid, dt_id):
    if not can_edit(pid): return jsonify(error="LOGIN_REQUIRED"), 403
    data = request.get_json(silent=True) or {}
    new_code = data.get("code","").strip().upper()
    new_name = data.get("name","").strip()
    if new_code:
        db.exe("UPDATE doc_types SET code=%s WHERE id=%s AND project_id=%s", (new_code, dt_id, pid))
    if new_name:
        db.exe("UPDATE doc_types SET name=%s WHERE id=%s AND project_id=%s", (new_name, dt_id, pid))
    return jsonify(ok=True)

@doc_types_bp.route("/api/doc_types/<pid>/reorder", methods=["POST"])
def api_reorder_doc_types(pid):
    if not can_edit(pid): return jsonify(error="LOGIN_REQUIRED"), 403
    data = request.get_json(silent=True) or {}
    order = data.get("order", [])
    for i, dt_id in enumerate(order):
        db.exe("UPDATE doc_types SET sort_order=%s WHERE id=%s AND project_id=%s", (i, dt_id, pid))
    return jsonify(ok=True)

@doc_types_bp.route("/api/doc_types/<pid>/<dt_id>", methods=["DELETE"])
def api_delete_doc_type(pid, dt_id):
    if not can_edit(pid): return jsonify(error="LOGIN_REQUIRED"), 403
    db.delete_doc_type(pid, dt_id)
    return jsonify(ok=True)
