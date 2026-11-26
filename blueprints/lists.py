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
import re

from flask import Blueprint, jsonify, request

import db
import drive_routing
from auth import current_user, can_edit, can_access_doc_type, is_project_admin

lists_bp = Blueprint("lists", __name__)


def _is_valid_subtype_list_for_doc_type(pid, dt, list_name):
    ln = str(list_name or "").strip().lower()
    if not ln or not (ln == "subtype" or ln.endswith("_subtype")):
        return False
    if ln == "subtype":
        return True
    proj = db.get_project(pid) or {}
    ws_low = re.sub(r"[^a-z0-9]+", "", str(proj.get("workspace_id") or "ENG").strip().lower()) or "eng"
    tokens = set()
    for raw in (dt.get("code"), dt.get("id")):
        slug = re.sub(r"[^a-z0-9]+", "", str(raw or "").strip().lower())
        if slug:
            tokens.add(slug)
    allowed = set()
    for tok in tokens:
        allowed.add(f"{ws_low}_{tok}_subtype")
        allowed.add(f"{tok}_subtype")
    return ln in allowed


@lists_bp.route("/api/lists/<pid>")
def api_lists(pid):
    return jsonify(db.get_lists(pid))


@lists_bp.route("/api/lists/<pid>", methods=["POST"])
def api_add_list_item(pid):
    if not can_edit(pid): return jsonify(error="LOGIN_REQUIRED"), 403
    data = request.get_json(silent=True) or {}
    list_name = str(data.get("list_name", "") or "").strip()
    dt_id = str(data.get("dt_id", "") or "").strip()
    if dt_id:
        dt = db.get_doc_type(pid, dt_id)
        if not dt:
            return jsonify(error="Document type not found"), 404
        if not can_access_doc_type(pid, dt_id, current_user()):
            return jsonify(error="Document type access denied"), 403
        if not _is_valid_subtype_list_for_doc_type(pid, dt, list_name):
            return jsonify(error="Invalid subtype list_name for document type"), 400
    try:
        res = db.add_list_item(pid, list_name, data.get("item", ""), meta=data.get("meta"))
    except ValueError as e:
        return jsonify(error=str(e)), 400
    column_wired = False
    is_sub_list = bool(list_name and db._is_subtype_list_name(list_name))
    if dt_id and is_sub_list:
        column_wired = bool(db.ensure_doc_type_subtype_column(pid, dt_id, list_name))
    drive_prov = None
    saved_item = str(res.get("item") or "").strip()
    if is_sub_list and saved_item:
        try:
            target_dt_ids = [dt_id] if dt_id else db.resolve_doc_types_for_subtype_list(pid, list_name)
            for t_dt in target_dt_ids:
                drive_prov = drive_routing.provision_project_drive_structure(
                    pid, only_doc_type_id=t_dt, only_subtype=saved_item
                )
        except Exception:
            drive_prov = None
    return jsonify(
        ok=True,
        item=saved_item,
        meta=res.get("meta"),
        list_name=list_name,
        column_wired=column_wired,
        drive_provision=drive_prov,
    )


@lists_bp.route("/api/lists/<pid>", methods=["PATCH"])
def api_rename_list_item(pid):
    if not can_edit(pid): return jsonify(error="LOGIN_REQUIRED"), 403
    data = request.get_json(silent=True) or {}
    list_name = str(data.get("list_name") or "").strip()
    old_item = str(data.get("old_item") or "").strip()
    raw_new_item = str(data.get("new_item") or "").strip()
    dt_id = str(data.get("dt_id") or "").strip()
    try:
        renamed = db.rename_list_item(
            pid,
            list_name,
            old_item,
            raw_new_item,
            meta=data.get("meta"),
        )
    except ValueError as e:
        return jsonify(error=str(e)), 400

    drive_rename = None
    if renamed and db._is_subtype_list_name(list_name):
        clean_new = raw_new_item
        if list_name.lower() == "hse_ptw_subtype":
            from hse_defaults import parse_ptw_list_item_input
            parsed_name, _ = parse_ptw_list_item_input(raw_new_item)
            if parsed_name:
                clean_new = parsed_name
        if old_item and clean_new and old_item != clean_new:
            try:
                target_dt_ids = [dt_id] if dt_id else db.resolve_doc_types_for_subtype_list(pid, list_name)
                for t_dt in target_dt_ids:
                    drive_rename = drive_routing.rename_subtype_drive_folder(pid, t_dt, old_item, clean_new)
            except Exception:
                drive_rename = None
    return jsonify(ok=True, renamed=renamed, drive_rename=drive_rename)


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
