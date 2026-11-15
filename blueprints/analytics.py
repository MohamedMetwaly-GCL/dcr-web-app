"""blueprints/analytics.py - DCR Analytics API routes.

Handles analytics endpoints:
  GET /api/analytics/trend
  GET /api/analytics/aging
  GET /api/analytics/quality
  GET /api/analytics/overdue

Step 8 of the incremental refactor.
Logic is identical to the original app.py routes — only the decorator
changed from @app.route to @analytics_bp.route.
"""
from flask import Blueprint, jsonify, request

import db
from auth import current_user, get_allowed_project_ids, get_user_doc_type_scope

analytics_bp = Blueprint("analytics", __name__)


def _analytics_scope():
    pid = request.args.get("pid")
    ws = request.args.get("workspace_id") or request.args.get("ws")
    u = current_user()
    if not u:
        return None, None, None, None
    allowed_ids = get_allowed_project_ids(u, workspace_id=ws)
    dt_scope = get_user_doc_type_scope(u, allowed_ids)
    return u, pid, allowed_ids, dt_scope


@analytics_bp.route("/api/analytics/trend")
def api_trend():
    u, pid, allowed_ids, dt_scope = _analytics_scope()
    if not u: return jsonify(error="LOGIN_REQUIRED"), 403
    return jsonify(db.get_monthly_trend(pid, project_ids=allowed_ids, doc_type_scope=dt_scope))


@analytics_bp.route("/api/analytics/aging")
def api_aging():
    u, pid, allowed_ids, dt_scope = _analytics_scope()
    if not u: return jsonify(error="LOGIN_REQUIRED"), 403
    return jsonify(db.get_aging_report(pid, project_ids=allowed_ids, doc_type_scope=dt_scope))


@analytics_bp.route("/api/analytics/quality")
def api_quality():
    u, pid, allowed_ids, dt_scope = _analytics_scope()
    if not u: return jsonify(error="LOGIN_REQUIRED"), 403
    return jsonify(db.get_quality_report(pid, project_ids=allowed_ids, doc_type_scope=dt_scope))


@analytics_bp.route("/api/analytics/overdue")
def api_overdue_list():
    u, pid, allowed_ids, dt_scope = _analytics_scope()
    if not u: return jsonify(error="LOGIN_REQUIRED"), 403
    return jsonify(db.get_overdue_records(pid, project_ids=allowed_ids, doc_type_scope=dt_scope))
