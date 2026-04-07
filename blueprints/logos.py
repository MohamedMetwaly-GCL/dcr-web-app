"""blueprints/logos.py - DCR Logos API routes.

Handles logo endpoints:
  GET  /api/logo/<pid>/<key>
  POST /api/logo/<pid>/<key>

Step 11 of the incremental refactor.
Logic is identical to the original app.py routes — only the decorator
changed from @app.route to @logos_bp.route.
"""
import base64
import io

from flask import Blueprint, jsonify, request, send_file

import db
from auth import can_edit

logos_bp = Blueprint("logos", __name__)


@logos_bp.route("/api/logo/<pid>/<key>")
def api_get_logo(pid, key):
    data = db.get_logo(pid, key)
    if not data: return "", 404
    raw = base64.b64decode(data)
    return send_file(io.BytesIO(raw), mimetype="image/png")


@logos_bp.route("/api/logo/<pid>/<key>", methods=["POST"])
def api_save_logo(pid, key):
    if not can_edit(pid): return jsonify(error="LOGIN_REQUIRED"), 403
    data = request.get_json(silent=True) or {}
    db.save_logo(pid, key, data.get("data",""))
    return jsonify(ok=True)
