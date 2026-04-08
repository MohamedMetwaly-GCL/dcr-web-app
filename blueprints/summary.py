"""blueprints/summary.py - DCR Summary/Misc API routes.

Handles summary and misc endpoints:
  GET /api/overdue_digest
  GET /api/records_range/<pid>/<dt_id>
  GET /api/executive_summary

Step 10 of the incremental refactor.
Logic is identical to the original app.py routes — only the decorator
changed from @app.route to @summary_bp.route.
"""
import datetime

from flask import Blueprint, jsonify, request

import db
from auth import current_user

summary_bp = Blueprint("summary", __name__)


@summary_bp.route("/api/overdue_digest")
def api_overdue_digest():
    """Returns overdue summary for email/display."""
    u = current_user()
    if not u: return jsonify(error="LOGIN_REQUIRED"), 403
    records = db.get_overdue_records()
    return jsonify({
        "total": len(records),
        "records": records[:50],
        "generated_at": datetime.datetime.now().isoformat()
    })


@summary_bp.route("/api/records_range/<pid>/<dt_id>")
def api_records_range(pid, dt_id):
    """Filter records by issued date range."""
    from_date = request.args.get("from")
    to_date   = request.args.get("to")
    records = db.get_records(pid, dt_id)
    if from_date:
        records = [r for r in records if (r.get("issuedDate") or "") >= from_date]
    if to_date:
        records = [r for r in records if (r.get("issuedDate") or "") <= to_date]
    return jsonify(records)


@summary_bp.route("/api/executive_summary")
def api_executive_summary():
    """One-page executive summary data."""
    stats  = db.get_dashboard_stats()
    aging  = db.get_aging_report()
    quality= db.get_quality_report()
    overdue= db.get_overdue_records()
    total_docs = sum(p["total"]    for p in stats)
    total_ap   = sum(p["approved"] for p in stats)
    total_pe   = sum(p["pending"]  for p in stats)
    total_rj   = sum(p["rejected"] for p in stats)
    total_ov   = sum(p["overdue"]  for p in stats)
    return jsonify({
        "summary": {
            "total": total_docs, "approved": total_ap,
            "pending": total_pe, "rejected": total_rj,
            "overdue": total_ov,
            "completion_pct": round(total_ap/total_docs*100) if total_docs else 0,
            "projects": len(stats),
        },
        "projects": [{"name":p["name"],"code":p["code"],"total":p["total"],
                      "approved":p["approved"],"pct":p["pct"],"overdue":p["overdue"]}
                     for p in stats],
        "aging":   aging,
        "quality": quality,
        "top_overdue": overdue[:10],
        "generated_at": datetime.datetime.now().strftime("%d-%m-%Y %H:%M"),
    })


@summary_bp.route("/api/data_quality_summary")
def api_data_quality_summary():
    return jsonify(db.get_data_quality_summary())


@summary_bp.route("/api/action_required_summary")
def api_action_required_summary():
    return jsonify(db.get_action_required_summary())


@summary_bp.route("/api/pr_analytics_summary")
def api_pr_analytics_summary():
    return jsonify(db.get_pr_analytics_summary())
