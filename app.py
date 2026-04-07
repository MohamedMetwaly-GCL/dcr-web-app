"""app.py - DCR Flask v1"""
import os, json, uuid, io, base64, datetime
from flask import (Flask, request, session, redirect, url_for,
                   jsonify, make_response, send_file, g)
import db
from utils import (compute_expected_reply, compute_duration, is_overdue,
                   format_date, get_next_doc_no, extract_rev, STATUS_COLORS,
                   REJECTED_STATUSES, invalidate_holidays_cache)

from config import IS_RENDER, SECRET_KEY, SESSION_CONFIG
from html_render import render_login, render_dashboard, render_register
from blueprints.projects import projects_bp
from blueprints.doc_types import doc_types_bp
from blueprints.columns import columns_bp
from blueprints.records import records_bp
from blueprints.analytics import analytics_bp
from blueprints.lists import lists_bp
from blueprints.summary import summary_bp
from blueprints.logos import logos_bp
from blueprints.users import users_bp
from blueprints.exporting import exporting_bp

app = Flask(__name__)
app.secret_key = SECRET_KEY
app.config.update(**SESSION_CONFIG)
app.register_blueprint(projects_bp)
app.register_blueprint(doc_types_bp)
app.register_blueprint(columns_bp)
app.register_blueprint(records_bp)
app.register_blueprint(analytics_bp)
app.register_blueprint(lists_bp)
app.register_blueprint(summary_bp)
app.register_blueprint(logos_bp)
app.register_blueprint(users_bp)
app.register_blueprint(exporting_bp)

# ── Auth helpers (extracted to auth.py — Step 2 refactor) ─────
from auth import current_user, require_login, require_superadmin, can_edit

# ── Pages ─────────────────────────────────────────────────────
@app.route("/ping")
def ping():
    return jsonify(status="ok", time=datetime.datetime.now().isoformat())

@app.route("/health")
def health():
    try:
        db.get_projects()
        return jsonify(status="ok", db="connected")
    except Exception as e:
        return jsonify(status="error", db=str(e)), 503

@app.route("/api/keepalive")
def keepalive():
    """Called periodically to prevent free-tier sleep."""
    return jsonify(status="ok", ts=datetime.datetime.now().isoformat())

@app.route("/login", methods=["GET","POST"])
def login():
    if request.method == "POST":
        data  = request.get_json(silent=True) or {}
        uname = data.get("username","").strip().lower()
        pw    = data.get("password","")
        if db.verify_pw(uname, pw):
            u     = db.get_user(uname)
            token = db.create_session(uname, u["role"])
            db.log_action(uname,"LOGIN",detail=f"Role: {u['role']}")
            resp  = make_response(jsonify(ok=True, role=u["role"], username=uname))
            resp.set_cookie("dcr_token", token, httponly=True,
                            samesite="Lax", secure=IS_RENDER,
                            max_age=db.SESSION_TTL)
            return resp
        db.log_action(uname,"LOGIN_FAIL",detail="Invalid credentials")
        return jsonify(ok=False, error="Invalid username or password"), 401
    return render_login()

@app.route("/logout", methods=["POST"])
def logout():
    token = request.cookies.get("dcr_token")
    if token: db.delete_session(token)
    resp = make_response(redirect("/login"))
    resp.delete_cookie("dcr_token")
    return resp

@app.route("/")
def home():
    return render_dashboard(current_user())

@app.route("/app")
def register():
    pid = request.args.get("p","")
    if not pid: return redirect("/")
    u = current_user()
    proj = db.get_project(pid)
    if not proj: return "Project not found", 404
    return render_register(u, proj)

@app.route("/api/settings/holidays", methods=["GET"])
def api_get_holidays():
    from utils import DEFAULT_HOLIDAYS
    h = db.get_setting("holidays", sorted(DEFAULT_HOLIDAYS))
    return jsonify(holidays=sorted(h) if h else [])

@app.route("/api/settings/holidays", methods=["POST"])
def api_save_holidays():
    u = current_user()
    if not u or u.get("role") not in ("superadmin","admin"):
        return jsonify(error="Superadmin only"), 403
    data = request.get_json(silent=True) or {}
    holidays = data.get("holidays", [])
    if not isinstance(holidays, list): return jsonify(error="Invalid"), 400
    # Validate format
    import datetime
    valid = []
    for h in holidays:
        try: datetime.date.fromisoformat(h); valid.append(h)
        except: pass
    db.save_setting("holidays", sorted(valid))
    from utils import invalidate_holidays_cache
    invalidate_holidays_cache()
    return jsonify(ok=True, count=len(valid))

# ── API: Records ──────────────────────────────────────────────
@app.route("/api/counts/<pid>")
def api_counts(pid):
    dts = db.get_doc_types(pid)
    return jsonify({dt["id"]: db.count_records(pid, dt["id"]) for dt in dts})

@app.route("/api/dashboard_stats")
def api_dashboard_stats():
    stats = db.get_dashboard_stats()
    u = current_user()
    for s in stats:
        s["can_edit"] = can_edit(s["id"])
    return jsonify(stats)

@app.route("/api/next_doc_no/<pid>/<dt_id>")
def api_next_doc_no(pid, dt_id):
    dts    = db.get_doc_types(pid)
    dt     = next((d for d in dts if d["id"] == dt_id), None)
    records = db.get_records(pid, dt_id)
    # Auto-detect prefix from existing records or build from project+dt codes
    if records:
        first_doc = next((r.get("docNo","") for r in records if r.get("docNo","")), "")
        import re as _re
        m = _re.match(r"^(.+?)\s*-\s*\d+", first_doc)
        prefix = m.group(1).strip() if m else (dt["code"] if dt else dt_id)
    else:
        # Build default prefix: PROJECT-DTCODE (e.g. PEM064-DS)
        proj = db.get_project(pid) or {}
        proj_code = proj.get("code", pid).replace("-","")
        dt_code = dt["code"] if dt else dt_id
        prefix = f"{proj_code}-{dt_code}"
    return jsonify(next=get_next_doc_no(prefix, records))

# ── API: Dropdown Lists ───────────────────────────────────────
# ── API: Logos ────────────────────────────────────────────────
# ── API: Users ────────────────────────────────────────────────
# ── Export Excel ──────────────────────────────────────────────


# ── PDF Export ────────────────────────────────────────────────



# ── Phase 2 — Analytics APIs ─────────────────────────────────
# ── Phase 3 — Overdue Digest API ─────────────────────────────
# ── Audit Log API ─────────────────────────────────────────────
@app.route("/api/audit")
def api_audit():
    u = current_user()
    if not u or u.get("role") not in ("superadmin","admin"):
        return jsonify(error="Admin only"), 403
    pid      = request.args.get("pid")
    username = request.args.get("username")
    action   = request.args.get("action")
    offset   = int(request.args.get("offset","0"))
    rows     = db.get_audit_log(project_id=pid, username=username,
                                action=action, limit=100, offset=offset)
    users    = [r["username"] for r in db.get_all_users()]
    actions  = db.get_audit_actions()
    projects = db.get_projects()
    return jsonify(
        rows=[{**dict(r), "ts": str(r["ts"])} for r in rows],
        users=users, actions=actions,
        projects=[{"id":p["id"],"name":p["name"],"code":p["code"]} for p in projects]
    )


if __name__ == "__main__":
    db.init()
    db.cleanup_sessions()
    port = int(os.environ.get("PORT", 5000))
    print(f"[DCR Flask] Running on http://localhost:{port}")
    app.run(host="0.0.0.0", port=port, debug=not IS_RENDER)
