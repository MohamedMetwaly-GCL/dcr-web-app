"""app.py - DCR Flask v1"""
import os, datetime
from flask import Flask, request, redirect, jsonify, make_response
import db
from utils import get_next_doc_no, get_next_plain_doc_no, doc_type_uses_revision

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
from auth import current_user, can_edit, can_view_project, get_allowed_project_ids

PUBLIC_ENDPOINTS = {"login", "logout", "static"}


@app.before_request
def enforce_login_and_project_scope():
    endpoint = request.endpoint or ""
    if endpoint in PUBLIC_ENDPOINTS or request.path.startswith("/static/"):
        return None

    u = current_user()
    if not u:
        if request.path.startswith("/api/"):
            return jsonify(error="LOGIN_REQUIRED"), 403
        return redirect("/login")

    pid = ""
    if request.view_args:
        pid = str(request.view_args.get("pid") or "").strip()
    if not pid:
        pid = str(request.args.get("p") or request.args.get("project_id") or request.args.get("pid") or "").strip()
    if pid and not can_view_project(pid, u):
        if request.path.startswith("/api/"):
            return jsonify(error="Forbidden"), 403
        return "Forbidden", 403
    return None

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
    u = current_user()
    stats = db.get_dashboard_stats(project_ids=get_allowed_project_ids(u))
    for s in stats:
        s["can_edit"] = can_edit(s["id"])
    return jsonify(stats)

@app.route("/api/next_doc_no/<pid>/<dt_id>")
def api_next_doc_no(pid, dt_id):
    dts    = db.get_doc_types(pid)
    dt     = next((d for d in dts if d["id"] == dt_id), None)
    records = db.get_records(pid, dt_id)
    dt_code = str((dt or {}).get("code", dt_id) or dt_id).strip().upper()
    dt_name = str((dt or {}).get("name", "") or "").strip().lower()
    is_ltr = dt_code == "LTR" or "letter" in dt_name or "correspondence" in dt_name
    if is_ltr:
        import re as _re
        doc_nos = [str(r.get("docNo","") or "").strip() for r in records if str(r.get("docNo","") or "").strip()]
        if doc_nos:
            def _ltr_num_parts(doc_no):
                m = _re.match(r"^(.*?)(\d+)$", doc_no)
                if not m:
                    return None
                prefix = m.group(1)
                num_txt = m.group(2)
                return prefix, int(num_txt), len(num_txt)
            parsed = [p for p in (_ltr_num_parts(d) for d in doc_nos) if p]
            if parsed:
                prefix, num, width = max(parsed, key=lambda t: (t[1], t[2], t[0]))
                return jsonify(next=f"{prefix}{str(num + 1).zfill(width)}")
            last_doc = max(doc_nos)
            return jsonify(next=last_doc + "-001")
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
    if not doc_type_uses_revision(dt_code, records):
        return jsonify(next=get_next_plain_doc_no(prefix, records))
    return jsonify(next=get_next_doc_no(prefix, records))

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

# ── API: Google Drive Webhooks ────────────────────────────────
@app.route("/api/webhooks/drive", methods=["POST"])
def api_webhook_drive():
    """
    Receives Push Notifications from Google Drive API.
    Google sends headers like X-Goog-Resource-State, X-Goog-Channel-Id, etc.
    """
    state = request.headers.get("X-Goog-Resource-State")
    if state == "sync":
        return "OK", 200  # Initial setup event
        
    channel_id = request.headers.get("X-Goog-Channel-Id")
    
    # 1. Look up the folder_id associated with this channel_id from your DB/Config
    # folder_id = get_folder_id_for_channel(channel_id)
    folder_id = "YOUR_DRIVE_FOLDER_ID" # Placeholder
    
    # 2. Process the folder in a separate thread so Google gets an immediate 200 OK
    if folder_id:
        from drive_service import process_drive_folder
        import threading
        threading.Thread(target=process_drive_folder, args=(folder_id,)).start()
        
    return "OK", 200

@app.route("/api/drive/sync/<pid>", methods=["POST"])
def api_drive_sync(pid):
    u = current_user()
    if not u or not can_edit(pid):
        return jsonify(error="Forbidden"), 403
        
    data = request.get_json(silent=True) or {}
    folder_id = data.get("folder_id", "").strip()
    if not folder_id:
        return jsonify(error="No folder ID provided"), 400
        
    # Save folder_id to project data
    proj = db.get_project(pid)
    if proj:
        proj["drive_folder_id"] = folder_id
        import json
        db.exe("UPDATE projects SET data=%s WHERE id=%s", (json.dumps(proj), pid))
        
    # Start sync process in background
    from drive_service import process_drive_folder
    import threading
    threading.Thread(target=process_drive_folder, args=(folder_id,)).start()
    
    return jsonify(ok=True)

# ── Distribution Matrix API ──────────────────────────────────────
@app.route("/api/distribution/<pid>", methods=["GET"])
def api_get_distribution(pid):
    """Get the full distribution matrix for a project."""
    try:
        u = current_user()
        if not u: return jsonify(error="LOGIN_REQUIRED"), 403
        if not can_view_project(pid): return jsonify(error="Forbidden"), 403
        data = db.get_distribution(pid)
        return jsonify(data)
    except Exception as e:
        print(f"[Distribution GET Error]: {e}")
        return jsonify(error="Server error"), 500

@app.route("/api/distribution/<pid>", methods=["POST"])
def api_save_distribution(pid):
    """Save/update a distribution row for a project.
    Body: {doc_type_id, event_type, emails: [...]}
    Requires DC role or admin+.
    """
    try:
        u = current_user()
        if not u: return jsonify(error="LOGIN_REQUIRED"), 403
        if not (u["role"] in ("superadmin", "admin") or db.user_is_dc(u["username"], pid)):
            return jsonify(error="Forbidden – DC role required"), 403
        body = request.get_json(silent=True) or {}
        doc_type_id = body.get("doc_type_id", "").strip()
        event_type  = body.get("event_type", "").strip()
        emails      = body.get("emails", [])
        if not doc_type_id or not event_type:
            return jsonify(error="doc_type_id and event_type required"), 400
        if not isinstance(emails, list):
            return jsonify(error="emails must be a list"), 400
        db.upsert_distribution(pid, doc_type_id, event_type, emails)
        return jsonify(ok=True)
    except Exception as e:
        print(f"[Distribution POST Error]: {e}")
        return jsonify(error="Server error"), 500

@app.route("/api/project_users/<pid>", methods=["GET"])
def api_project_users(pid):
    try:
        u = current_user()
        if not u: return jsonify(error="LOGIN_REQUIRED"), 403
        if not can_view_project(pid): return jsonify(error="Forbidden"), 403
        return jsonify(db.get_project_users_full(pid))
    except Exception as e:
        return jsonify(error="Server error"), 500

import hashlib
import os
MAGIC_SECRET = os.environ.get("MAGIC_SECRET", "super_secret_magic_dcr_key")

def _generate_magic_token(pid, dt_id):
    return hashlib.sha256(f"{pid}:{dt_id}:{MAGIC_SECRET}".encode()).hexdigest()[:16]

@app.route("/api/magic/generate/<pid>/<dt_id>", methods=["POST"])
def api_generate_magic_link(pid, dt_id):
    try:
        u = current_user()
        if not u: return jsonify(error="LOGIN_REQUIRED"), 403
        if not can_view_project(pid): return jsonify(error="Forbidden"), 403
        token = _generate_magic_token(pid, dt_id)
        # Generate absolute URL without hardcoding domain (using request.host_url)
        link = f"{request.host_url.rstrip('/')}/magic/{pid}/{dt_id}?token={token}"
        return jsonify(ok=True, link=link)
    except Exception as e:
        return jsonify(error="Server error"), 500

@app.route("/magic/<pid>/<dt_id>", methods=["GET"])
def magic_digest_view(pid, dt_id):
    token = request.args.get("token", "")
    expected = _generate_magic_token(pid, dt_id)
    if token != expected:
        return "Invalid or unauthorized magic link.", 403
    
    # We will build a lightweight HTML template inline or via render_template_string
    # But since DCR uses html_render for SPA, we can return a simple HTML page.
    digest = db.get_daily_digest(pid, [dt_id])
    
    from flask import render_template_string
    template = """
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Daily Digest</title>
        <style>
            body { font-family: 'Inter', sans-serif; background: #f8fafc; color: #0f172a; padding: 20px; margin: 0; }
            .card { background: white; border-radius: 12px; padding: 20px; margin-bottom: 20px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1); }
            h2 { font-size: 18px; margin-top: 0; border-bottom: 2px solid #e2e8f0; padding-bottom: 10px; }
            .h-recv { color: #2563eb; border-color: #bfdbfe; }
            .h-sent { color: #ea580c; border-color: #fed7aa; }
            .h-repl { color: #16a34a; border-color: #bbf7d0; }
            .item { padding: 10px 0; border-bottom: 1px solid #f1f5f9; font-size: 14px; }
            .item:last-child { border: none; padding-bottom: 0; }
            .docno { font-weight: bold; color: #334155; }
            .title { color: #64748b; margin-top: 4px; font-size: 13px; }
            .empty { color: #94a3b8; font-size: 14px; font-style: italic; }
        </style>
    </head>
    <body>
        <h1 style="font-size:20px;text-align:center;color:#0f172a;margin-bottom:24px;">📋 Daily Digest - Today</h1>
        
        <div class="card">
            <h2 class="h-recv">📥 Received Today ({{ digest.received|length }})</h2>
            {% for doc in digest.received %}
            <div class="item">
                <div class="docno">{{ doc.docNo }}</div>
                <div class="title">{{ doc.title }}</div>
            </div>
            {% else %}
            <div class="empty">No documents received today.</div>
            {% endfor %}
        </div>
        
        <div class="card">
            <h2 class="h-sent">📤 Issued/Sent Today ({{ digest.issued|length }})</h2>
            {% for doc in digest.issued %}
            <div class="item">
                <div class="docno">{{ doc.docNo }}</div>
                <div class="title">{{ doc.title }}</div>
            </div>
            {% else %}
            <div class="empty">No documents issued today.</div>
            {% endfor %}
        </div>
        
        <div class="card">
            <h2 class="h-repl">✅ Replied Today ({{ digest.replied|length }})</h2>
            {% for doc in digest.replied %}
            <div class="item">
                <div class="docno">{{ doc.docNo }}</div>
                <div class="title">{{ doc.title }}</div>
                {% if doc.status %}<div style="margin-top:4px;font-size:12px;font-weight:bold;color:#475569;">Status: {{ doc.status }}</div>{% endif %}
            </div>
            {% else %}
            <div class="empty">No documents replied today.</div>
            {% endfor %}
        </div>
    </body>
    </html>
    """
    return render_template_string(template, digest=digest)

@app.route("/api/daily_digest/<pid>", methods=["GET"])
def api_daily_digest(pid):
    try:
        u = current_user()
        if not u: return jsonify(error="LOGIN_REQUIRED"), 403
        
        projects_to_check = []
        if pid == "all":
            from auth import get_allowed_project_ids
            projects_to_check = get_allowed_project_ids(u)
        else:
            if not can_view_project(pid): return jsonify(error="Forbidden"), 403
            projects_to_check = [pid]
            
        combined_digest = {"received": [], "issued": [], "replied": []}
        is_admin = u["role"] in ("superadmin", "admin")
        
        for p_id in projects_to_check:
            dist = db.get_distribution(p_id)
            assigned_dt_ids = []
            for dt_id, events in dist.items():
                users = events.get("access", [])
                if is_admin or u["username"] in users:
                    assigned_dt_ids.append(dt_id)
            if assigned_dt_ids:
                d = db.get_daily_digest(p_id, assigned_dt_ids)
                combined_digest["received"].extend(d["received"])
                combined_digest["issued"].extend(d["issued"])
                combined_digest["replied"].extend(d["replied"])
                
        return jsonify(combined_digest)
    except Exception as e:
        print(f"[Daily Digest Error]: {e}")
        return jsonify(error="Server error"), 500


def drive_polling_job():
    """Background thread to poll Google Drive every 15 minutes as a robust Failsafe/Sync mechanism."""
    import time
    while True:
        time.sleep(15 * 60) # Wait 15 minutes
        try:
            projects = db.get_projects()
            for p in projects:
                folder_id = p.get("data", {}).get("drive_folder_id")
                if folder_id:
                    from drive_service import process_drive_folder
                    process_drive_folder(folder_id)
        except Exception as e:
            print(f"[Drive Polling Error]: {e}")

if __name__ == "__main__":
    db.init()
    db.cleanup_sessions()
    
    # Start the Drive Polling Daemon Thread
    import threading
    threading.Thread(target=drive_polling_job, daemon=True).start()
    
    port = int(os.environ.get("PORT", 5000))
    print(f"[DCR Flask] Running on http://localhost:{port}")
    app.run(host="0.0.0.0", port=port, debug=not IS_RENDER)
