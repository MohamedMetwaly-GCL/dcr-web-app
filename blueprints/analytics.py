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

analytics_bp = Blueprint("analytics", __name__)


@analytics_bp.route("/api/analytics/trend")
def api_trend():
    pid = request.args.get("pid")
    return jsonify(db.get_monthly_trend(pid))


@analytics_bp.route("/api/analytics/aging")
def api_aging():
    pid = request.args.get("pid")
    return jsonify(db.get_aging_report(pid))


@analytics_bp.route("/api/analytics/quality")
def api_quality():
    pid = request.args.get("pid")
    return jsonify(db.get_quality_report(pid))


@analytics_bp.route("/api/analytics/overdue")
def api_overdue_list():
    pid = request.args.get("pid")
    return jsonify(db.get_overdue_records(pid))
