"""app.py - DCR Flask v1"""
import os, datetime
from flask import Flask, request, redirect, jsonify, make_response, send_from_directory
import db
from utils import get_next_doc_no, get_next_plain_doc_no, doc_type_uses_revision

from config import IS_RENDER, SECRET_KEY, SESSION_CONFIG
from html_render import render_login, render_launchpad, render_dashboard, render_register, _head_icon_tags
from i18n import get_lang, LANG_COOKIE_NAME, LANG_COOKIE_MAX_AGE, SUPPORTED_LANGS, html_dir, t, get_client_i18n_json
from security import validate_username, login_limiter, get_client_ip
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
from blueprints.ai import ai_bp
from blueprints.drive_uploads import drive_uploads_bp
from blueprints.drive_viewer import drive_viewer_bp
from blueprints.hse import hse_bp

app = Flask(__name__)
APP_VERSION = "2.4.9"
BUILD_ID = "20261111-security-hardening"
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
app.register_blueprint(ai_bp, url_prefix="/api/ai")
app.register_blueprint(drive_uploads_bp)
app.register_blueprint(drive_viewer_bp)
app.register_blueprint(hse_bp)

# ── Auth helpers (extracted to auth.py — Step 2 refactor) ─────
from auth import (
    current_user,
    can_edit,
    can_view_project,
    get_allowed_project_ids,
    is_superadmin,
    is_project_admin,
    can_access_doc_type,
    filter_allowed_doc_types,
    get_user_doc_type_scope,
    _resolve_project,
)

PUBLIC_ENDPOINTS = {"login", "logout", "static", "favicon", "manifest_webmanifest", "magic_master_view", "api_magic_data", "hse.public_qr_verify"}


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
    if pid and pid != "all" and not can_view_project(pid, u):
        if request.path.startswith("/api/"):
            try:
                db.log_action(u.get("username", "unknown"), "ACCESS DENIED", project_id=pid,
                              detail=f"Unauthorized API access attempt to project '{pid}' on {request.path}")
            except Exception:
                pass
            return jsonify(error="Forbidden"), 403
        return "Forbidden", 403
    return None


@app.after_request
def apply_security_headers(response):
    """Apply safe, non-intrusive security headers to all HTTP responses."""
    response.headers.setdefault("X-Content-Type-Options", "nosniff")
    response.headers.setdefault("X-Frame-Options", "SAMEORIGIN")
    response.headers.setdefault("Referrer-Policy", "strict-origin-when-cross-origin")
    is_https = IS_RENDER or request.is_secure or (request.headers.get("X-Forwarded-Proto", "").lower() == "https")
    if is_https:
        response.headers.setdefault("Strict-Transport-Security", "max-age=31536000; includeSubDomains")
    q_lang = (request.args.get("lang") or "").strip().lower()
    if q_lang in SUPPORTED_LANGS:
        response.set_cookie(
            LANG_COOKIE_NAME,
            q_lang,
            max_age=LANG_COOKIE_MAX_AGE,
            path="/",
            samesite="Lax",
            secure=is_https,
            httponly=False,
        )
    return response


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

@app.route("/favicon.ico")
def favicon():
    return send_from_directory(
        os.path.join(app.root_path, "static"),
        "favicon.ico",
        mimetype="image/x-icon",
    )

@app.route("/manifest.webmanifest")
def manifest_webmanifest():
    return send_from_directory(
        os.path.join(app.root_path, "static"),
        "manifest.json",
        mimetype="application/manifest+json",
    )

@app.route("/api/keepalive")
def keepalive():
    """Called periodically to prevent free-tier sleep."""
    return jsonify(status="ok", ts=datetime.datetime.now().isoformat())

@app.route("/login", methods=["GET","POST"])
def login():
    if request.method == "POST":
        data  = request.get_json(silent=True) or {}
        raw_uname = str(data.get("username","")).strip()
        pw    = str(data.get("password",""))
        ip    = get_client_ip()

        # 1. Brute Force Protection & Rate Limiting check
        is_locked, remaining_sec = login_limiter.is_locked(ip, raw_uname)
        if is_locked:
            mins = max(1, round(remaining_sec / 60))
            db.log_action(raw_uname or "unknown", "BRUTE_FORCE_BLOCKED",
                          detail=f"Blocked login attempt during active lockout ({remaining_sec}s remaining)")
            return jsonify(
                ok=False,
                error=f"Too many failed login attempts. Access is temporarily locked for {mins} minute(s). Please try again later.",
                locked=True,
                remaining_seconds=remaining_sec
            ), 429

        # 2. Username sanitization & injection prevention
        is_valid, err_msg, is_suspicious = validate_username(raw_uname)
        if not is_valid:
            if is_suspicious:
                db.log_action(raw_uname[:50] or "unknown", "SECURITY_ALERT",
                              detail=f"Blocked suspicious/injected username attempt: '{raw_uname[:50]}'")
            else:
                db.log_action(raw_uname[:50] or "unknown", "LOGIN_FAIL", detail="Invalid username format")
            return jsonify(ok=False, error=err_msg), 400

        uname = raw_uname.lower()

        # 3. Verify credentials
        if db.verify_pw(uname, pw):
            login_limiter.record_success(ip, uname)
            u     = db.get_user(uname)
            token = db.create_session(uname, u["role"])
            db.log_action(uname,"LOGIN",detail=f"Role: {u['role']}")
            try:
                user_ws = db.get_user_workspaces(uname)
            except Exception:
                user_ws = []
            is_multi_or_super = (str(u.get("role", "")).lower() == "superadmin") or (len(user_ws) > 1)
            if is_multi_or_super:
                redirect_to = "/hub"
            elif len(user_ws) == 1:
                redirect_to = f"/?ws={user_ws[0]['id']}"
            else:
                redirect_to = "/hub"
            resp  = make_response(jsonify(ok=True, role=u["role"], username=uname, redirect_to=redirect_to))
            is_https = IS_RENDER or request.is_secure or (request.headers.get("X-Forwarded-Proto", "").lower() == "https")
            resp.set_cookie("dcr_token", token, httponly=True,
                            samesite="Lax", secure=is_https,
                            max_age=db.SESSION_TTL)
            return resp

        # Record failed attempt (6 attempts threshold)
        is_locked_now, remaining_attempts, lockout_sec = login_limiter.record_failure(ip, uname)
        if is_locked_now:
            mins = max(1, round(lockout_sec / 60))
            db.log_action(uname, "LOGIN_LOCKOUT",
                          detail=f"Maximum failed attempts (6) reached. Locked out for {mins} minutes.")
            return jsonify(
                ok=False,
                error=f"Too many failed login attempts (6/6). Account / IP is temporarily locked for {mins} minutes.",
                locked=True,
                remaining_seconds=lockout_sec
            ), 429
        else:
            db.log_action(uname, "LOGIN_FAIL",
                          detail=f"Invalid credentials ({remaining_attempts} attempt(s) remaining)")
            return jsonify(
                ok=False,
                error=f"Invalid username or password. ({remaining_attempts} attempt(s) remaining before temporary lockout).",
                remaining_attempts=remaining_attempts
            ), 401
    lang = get_lang(request)
    return render_login(lang=lang)

@app.route("/logout", methods=["POST", "GET"])
def logout():
    u = current_user()
    if u:
        try:
            db.log_action(u.get("username", ""), "LOGOUT", detail="User logged out")
        except Exception:
            pass
    token = request.cookies.get("dcr_token")
    if token: db.delete_session(token)
    resp = make_response(redirect("/login"))
    resp.delete_cookie("dcr_token")
    return resp

@app.route("/hub")
def launchpad_hub():
    u = current_user()
    lang = get_lang(request)
    return render_launchpad(u, lang=lang)

@app.route("/")
def home():
    u = current_user()
    lang = get_lang(request)
    ws_id = request.args.get("ws")
    if not ws_id and u:
        try:
            user_ws = db.get_user_workspaces(u.get("username", ""))
        except Exception:
            user_ws = []
        is_super = (str(u.get("role", "")).lower() == "superadmin")
        if is_super or len(user_ws) > 1:
            return render_launchpad(u, lang=lang)
        elif len(user_ws) == 1:
            ws_id = user_ws[0]["id"]
    return render_dashboard(u, active_ws_id=ws_id, lang=lang)

_recent_project_views = {}

@app.route("/app")
def register():
    pid = request.args.get("p","")
    if not pid: return redirect("/")
    u = current_user()
    proj = db.get_project(pid)
    if not proj: return "Project not found", 404
    try:
        now_ts = datetime.datetime.now().timestamp()
        last_view_ts = _recent_project_views.get((u.get("username"), pid), 0)
        if now_ts - last_view_ts > 60:
            _recent_project_views[(u.get("username"), pid)] = now_ts
            db.log_action(
                username=u.get("username"),
                action="PROJECT VIEW",
                project_id=pid,
                detail=f"Opened project: {proj.get('name', pid)}"
            )
    except Exception as e:
        pass
    ws_id = request.args.get("ws") or proj.get("workspace_id", "ENG")
    lang = get_lang(request)
    return render_register(u, proj, active_ws_id=ws_id, lang=lang)

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
    u = current_user()
    if not u or not can_view_project(pid, u):
        return jsonify(error="Forbidden"), 403
    dts = filter_allowed_doc_types(pid, db.get_doc_types(pid), u)
    return jsonify({dt["id"]: db.count_records(pid, dt["id"]) for dt in dts})

@app.route("/api/dashboard_stats")
def api_dashboard_stats():
    u = current_user()
    ws_id = request.args.get("ws") or request.args.get("workspace_id") or "ENG"
    allowed = get_allowed_project_ids(u, workspace_id=ws_id)
    dt_scope = get_user_doc_type_scope(u, allowed)
    stats = db.get_dashboard_stats(project_ids=allowed, workspace_id=ws_id, doc_type_scope=dt_scope)
    for s in stats:
        s["can_edit"] = can_edit(s["id"])
        s["can_admin"] = is_project_admin(s["id"], u)
    return jsonify(stats)

@app.route("/api/next_doc_no/<pid>/<dt_id>")
def api_next_doc_no(pid, dt_id):
    if not can_access_doc_type(pid, dt_id):
        return jsonify(error="Document type access denied"), 403
    subtype = str(request.args.get("subtype", "") or "").strip()
    dt_meta = db.get_doc_type(pid, dt_id) or {}
    dt_code = str(dt_meta.get("code") or dt_id or "").strip().upper()
    project = db.get_project(pid) or {}
    ws_id = str(project.get("workspace_id") or "ENG").strip().upper()
    records = db.get_records(pid, dt_id)
    if ws_id == "HSE" and dt_code == "PTW":
        from hse_defaults import generate_next_ptw_doc_no, resolve_ptw_project_prefix
        lists_meta = db.get_lists_with_meta(pid).get("hse_ptw_subtype", [])
        meta_map = {
            it["value"]: it.get("meta")
            for it in lists_meta
            if isinstance(it, dict) and it.get("value")
        }
        if subtype:
            res = generate_next_ptw_doc_no(
                pid, subtype, project=project, records=records, subtype_meta_map=meta_map
            )
        else:
            res = {
                "next": "",
                "prefix": resolve_ptw_project_prefix(pid, project=project, records=records),
                "subtype": "",
                "subtype_code": "",
                "seq": 0,
            }
        resp = jsonify(**res)
        resp.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
        return resp
    doc_nos = [str(r.get("docNo", "")).strip() for r in records if str(r.get("docNo", "")).strip()]
    if not doc_nos:
        # First Document: Leave blank for the user to define the pattern
        return jsonify(next="")
        
    import re as _re
    parsed_docs = []
    
    for d in doc_nos:
        # 1. Check if it ends with REV
        has_rev = bool(_re.search(r'\bREV\d+$', d, flags=_re.IGNORECASE))
        # 2. Extract base document
        base_d = _re.sub(r'\s*REV\d+$', '', d, flags=_re.IGNORECASE).strip()
        # 3. Determine prefix and numeric sequence
        m_d = _re.search(r'^(.*?)(\d+)$', base_d)
        if m_d:
            parsed_docs.append({
                "orig": d,
                "prefix": m_d.group(1),
                "num": int(m_d.group(2)),
                "width": len(m_d.group(2)),
                "has_rev": has_rev
            })
            
    if parsed_docs:
        # Find the document with the absolute highest sequence number in this section
        last_doc_meta = max(parsed_docs, key=lambda x: (x["num"], x["width"], x["prefix"]))
        prefix = last_doc_meta["prefix"]
        width = last_doc_meta["width"]
        has_rev = last_doc_meta["has_rev"]
        
        # Max sequence for this specific prefix (will be last_doc_meta["num"] unless tied)
        max_num = max(x["num"] for x in parsed_docs if x["prefix"] == prefix)
        
        next_num = max_num + 1
        next_base = f"{prefix}{str(next_num).zfill(width)}"
        resp = jsonify(next=f"{next_base} REV00" if has_rev else next_base)
        resp.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
        return resp
    else:
        # Fallback if no document ends with digits
        last_doc = doc_nos[-1]
        has_rev = bool(_re.search(r'\bREV\d+$', last_doc, flags=_re.IGNORECASE))
        base_last = _re.sub(r'\s*REV\d+$', '', last_doc, flags=_re.IGNORECASE).strip()
        next_base = f"{base_last}-001"
        resp = jsonify(next=f"{next_base} REV00" if has_rev else next_base)
        resp.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
        return resp

# ── Audit Log API ─────────────────────────────────────────────
# ── Audit Log Event Beacon ────────────────────────────────────
@app.route("/api/audit/log_event", methods=["POST"])
def api_audit_log_event():
    u = current_user()
    if not u:
        return jsonify(error="LOGIN_REQUIRED"), 403
    data = request.get_json(silent=True) or {}
    event_type = str(data.get("event") or "").strip().upper()
    ALLOWED_EVENTS = {"DOC_VIEW", "DRIVE_LINK_CLICK", "TAB_VIEW", "FILE_DOWNLOAD"}
    if event_type not in ALLOWED_EVENTS:
        return jsonify(error="Invalid event"), 400
    pid = data.get("project_id")
    dt_id = data.get("dt_id")
    doc_no = data.get("doc_no")
    detail = data.get("detail")
    try:
        db.log_action(
            username=u.get("username", ""),
            action=event_type,
            project_id=pid,
            dt_id=dt_id,
            doc_no=doc_no,
            detail=str(detail)[:250] if detail else None
        )
    except Exception:
        pass
    return jsonify(ok=True)

@app.route("/api/audit")
def api_audit():
    u = current_user()
    if not u or (not is_superadmin(u) and not db.is_user_workspace_admin(u.get("username"))):
        return jsonify(error="Admin only"), 403
    pid      = request.args.get("pid")
    username = request.args.get("username")
    action   = request.args.get("action")
    offset   = int(request.args.get("offset","0"))
    ws_filter = request.args.get("workspace_id")

    is_super = is_superadmin(u)
    if not is_super:
        try:
            admin_ws = [str(w).strip().upper() for w in (db.get_admin_workspaces_for_user(u["username"]) or [])]
            if ws_filter:
                ws_up = str(ws_filter).strip().upper()
                target_ws = [ws_up] if ws_up in admin_ws else []
            else:
                target_ws = admin_ws
            dept_projs = []
            seen_pids = set()
            for tws in target_ws:
                for p in db.get_projects(workspace_id=tws):
                    if p["id"] not in seen_pids:
                        seen_pids.add(p["id"])
                        dept_projs.append(p["id"])
            dept_users = [usr["username"] for usr in db.get_users_for_admin(u["username"])]
            if (pid and pid not in dept_projs) or not dept_projs:
                return jsonify(rows=[], users=dept_users if target_ws else [], actions=[], projects=[])
            rows = db.get_audit_log(project_id=pid, username=username,
                                    action=action, limit=100, offset=offset,
                                    allowed_project_ids=dept_projs, allowed_usernames=dept_users)
            users = dept_users
            projects = db.get_projects(dept_projs)
        except Exception:
            # Fallback if unconfigured/mock DB in unit tests
            rows = db.get_audit_log(project_id=pid, username=username,
                                    action=action, limit=100, offset=offset)
            users = [r["username"] for r in db.get_all_users()]
            projects = db.get_projects()
    else:
        if ws_filter:
            try:
                ws_projs = [p["id"] for p in db.get_projects(workspace_id=ws_filter)]
                if pid and pid not in ws_projs:
                    return jsonify(rows=[], users=[], actions=[], projects=[])
                rows = db.get_audit_log(project_id=pid, username=username,
                                        action=action, limit=100, offset=offset,
                                        allowed_project_ids=ws_projs)
                users = [r["username"] for r in db.get_all_users()]
                projects = db.get_projects(project_ids=ws_projs)
            except Exception:
                rows = db.get_audit_log(project_id=pid, username=username,
                                        action=action, limit=100, offset=offset)
                users = [r["username"] for r in db.get_all_users()]
                projects = db.get_projects()
        else:
            rows = db.get_audit_log(project_id=pid, username=username,
                                    action=action, limit=100, offset=offset)
            users = [r["username"] for r in db.get_all_users()]
            projects = db.get_projects()

    actions  = db.get_audit_actions()
    if "PROJECT VIEW" not in actions:
        actions = sorted(actions + ["PROJECT VIEW"])
    if "TAB_VIEW" not in actions:
        actions = sorted(actions + ["TAB_VIEW"])
    if "FILE_DOWNLOAD" not in actions:
        actions = sorted(actions + ["FILE_DOWNLOAD"])
    if "DRIVE_LINK_CLICK" not in actions:
        actions = sorted(actions + ["DRIVE_LINK_CLICK"])
    for sec_act in ["SECURITY_ALERT", "BRUTE_FORCE_BLOCKED", "LOGIN_LOCKOUT", "FILE_REJECTED"]:
        if sec_act not in actions:
            actions = sorted(actions + [sec_act])
    return jsonify(
        rows=[{**dict(r), "ts": str(r["ts"])} for r in rows],
        users=users, actions=actions,
        projects=[{"id":p["id"],"name":p["name"],"code":p["code"]} for p in projects]
    )

# ── API: Global Sequence Audit ────────────────────────────────
@app.route("/api/audit-sequence", methods=["GET"])
def api_audit_sequence():
    u = current_user()
    pid = request.args.get("project_id")
    dt_id = request.args.get("dt_id")
    if not u:
        return jsonify(error="LOGIN_REQUIRED"), 403
    if not pid or not dt_id:
        return jsonify(error="Project ID and Document Type ID are required"), 400
    if not is_project_admin(pid, u):
        return jsonify(error="Superadmin/Admin only"), 403
    
    records = db.q("SELECT COALESCE(data->>'docNo', '') AS docno FROM records WHERE project_id=%s AND dt_id=%s", (pid, dt_id))
    
    import re
    nums = []
    for r in records:
        docno = r["docno"]
        if not docno:
            continue
        # Strip revision parts like "rev 1", "REV-A", etc.
        clean_docno = re.sub(r'rev.*$', '', docno, flags=re.IGNORECASE).strip()
        # Find all contiguous integer groups
        matches = re.findall(r'\d+', clean_docno)
        if matches:
            # We assume the last integer group before revision is the sequence number
            nums.append(int(matches[-1]))
    
    if not nums:
        return jsonify(missing=[], min=0, max=0)
    
    nums = set(nums)
    min_n = min(nums)
    max_n = max(nums)
    missing = [n for n in range(min_n, max_n + 1) if n not in nums]
    
    return jsonify(missing=missing, min=min_n, max=max_n)

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
        
    # Start sync process in background (scoped to pid)
    from drive_service import process_drive_folder
    import threading
    threading.Thread(target=process_drive_folder, args=(folder_id,), kwargs={"project_id": pid}).start()
    
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
        if not (is_project_admin(pid, u) or (can_view_project(pid, u) and db.user_is_dc(u["username"], pid))):
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

@app.route("/api/project/<pid>/settings", methods=["POST"])
def update_project_settings_route(pid):
    try:
        u = current_user()
        if not u: return jsonify(error="LOGIN_REQUIRED"), 403
        if not is_project_admin(pid, u):
            return jsonify(error="Forbidden. Only Admins can modify project settings."), 403
            
        body = request.get_json()
        if not body:
            return jsonify(error="Empty payload"), 400
            
        db.update_project_settings(pid, body)
        return jsonify(ok=True)
    except Exception as e:
        print(f"[Project Settings POST Error]: {e}")
        return jsonify(error="Server error"), 500

import hashlib
import os
MAGIC_SECRET = os.environ.get("MAGIC_SECRET", "super_secret_magic_dcr_key")

def _generate_master_magic_token(pid):
    return hashlib.sha256(f"{pid}:master:{MAGIC_SECRET}".encode()).hexdigest()[:16]

@app.route("/api/magic/generate_master/<pid>", methods=["POST"])
def api_generate_master_magic_link(pid):
    try:
        u = current_user()
        if not u: return jsonify(error="LOGIN_REQUIRED"), 403
        if not can_view_project(pid): return jsonify(error="Forbidden"), 403
        token = _generate_master_magic_token(pid)
        link = f"{request.host_url.rstrip('/')}/magic/{pid}?token={token}"
        return jsonify(ok=True, link=link)
    except Exception as e:
        return jsonify(error="Server error"), 500

@app.route("/magic/<pid>", methods=["GET"])
def magic_master_view(pid):
    token = request.args.get("token", "")
    expected = _generate_master_magic_token(pid)
    if token != expected:
        return "Invalid or unauthorized magic link.", 403
    
    # Fetch project details
    projects = db.get_projects()
    project = next((p for p in projects if p["id"] == pid or p.get("code") == pid), None)
    if not project:
        return "Project not found", 404
        
    p_code = project.get("code", "")
    p_name = project.get("name", "")
    lang = get_lang(request)
    dir_attr = html_dir(lang)
    i18n_json = get_client_i18n_json(lang)
    
    # We serve the empty shell HTML. The frontend will fetch data via AJAX.
    # We embed the pid and token in the template.
    template = """
    <!DOCTYPE html>
    <html lang="{{ lang }}" dir="{{ dir_attr }}">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>{{ t_title }}</title>
        {{ icon_tags|safe }}
        <style>
            @font-face {
                font-family: 'Cairo';
                src: url('/static/fonts/Cairo-Regular.ttf') format('truetype');
                font-weight: 400 800;
                font-display: swap;
            }
            body { font-family: 'Inter', sans-serif; background: #f8fafc; color: #0f172a; padding: 20px; margin: 0; }
            html[dir="rtl"] body { font-family: 'Cairo', 'Inter', sans-serif; direction: rtl; text-align: right; }
            .bidi-ltr { direction: ltr !important; unicode-bidi: isolate !important; display: inline-block; }
            .header-bar { display: flex; justify-content: space-between; align-items: center; margin-bottom: 20px; flex-wrap: wrap; gap: 10px; }
            .header-bar h1 { font-size: 20px; margin: 0; color: #0f172a; }
            .proj-subtitle { font-size: 16px; color: #64748b; margin-left: 10px; }
            html[dir="rtl"] .proj-subtitle { margin-left: 0; margin-right: 10px; }
            .controls { display: flex; gap: 10px; align-items: center; flex-wrap: wrap; }
            select, input[type="date"] { padding: 8px 12px; border-radius: 6px; border: 1px solid #cbd5e1; font-size: 14px; outline: none; background: #fff; color: #0f172a; }
            .lang-btn { padding: 8px 12px; border-radius: 6px; border: 1px solid #cbd5e1; background: #fff; color: #0f172a; font-size: 13px; font-weight: 700; cursor: pointer; }
            .lang-btn:hover { background: #f1f5f9; border-color: #94a3b8; }
            .card { background: white; border-radius: 12px; padding: 20px; margin-bottom: 20px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1); }
            .stats-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(120px, 1fr)); gap: 15px; margin-bottom: 20px; }
            .stat-box { background: white; padding: 15px; border-radius: 10px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); text-align: center; border: 1px solid #e2e8f0; }
            .stat-box .val { font-size: 24px; font-weight: bold; color: #0f172a; }
            .stat-box .lbl { font-size: 12px; color: #64748b; margin-top: 4px; text-transform: uppercase; }
            html[dir="rtl"] .stat-box .lbl { text-transform: none; }
            
            h2 { font-size: 18px; margin-top: 0; border-bottom: 2px solid #e2e8f0; padding-bottom: 10px; }
            .h-action { color: #dc2626; border-color: #fecaca; }
            .h-info { color: #2563eb; border-color: #bfdbfe; }
            
            .item { padding: 10px 0; border-bottom: 1px solid #f1f5f9; font-size: 14px; }
            .item:last-child { border: none; padding-bottom: 0; }
            .docno { font-weight: bold; color: #334155; }
            .title { color: #64748b; margin-top: 4px; font-size: 13px; }
            .disc { display: inline-block; padding: 2px 8px; border-radius: 12px; background: #e2e8f0; font-size: 11px; font-weight: 600; color: #475569; margin-top: 4px; }
            .status-tag { font-size: 11px; font-weight: bold; color: #475569; margin-left: 8px; }
            html[dir="rtl"] .status-tag { margin-left: 0; margin-right: 8px; }
            .empty { color: #94a3b8; font-size: 14px; font-style: italic; }
            html[dir="rtl"] .empty { font-style: normal; }
            
            .top-overdue-card { border: 2px solid #fee2e2; }
            .top-overdue-card h2 { color: #b91c1c; border-color: #fecaca; }
            
            #loading { text-align: center; color: #64748b; padding: 40px; display: none; }
            .digest-footer { text-align: center; color: #64748b; font-size: 12px; margin-top: 24px; padding-top: 12px; border-top: 1px solid #e2e8f0; }
        </style>
    </head>
    <body>
        <div class="header-bar">
            <h1>📋 {{ t_title }} <span class="proj-subtitle"><span class="bidi-ltr">{{ p_code }}</span> - {{ p_name }}</span></h1>
            <div class="controls">
                <input type="date" id="datePicker" class="bidi-ltr" title="{{ t_history_date }}">
                <select id="roleFilter">
                    <option value="__ALL__" selected>{{ t_show_all }}</option>
                </select>
                <button type="button" id="langSwitcherBtn" class="lang-btn" onclick="switchDcrLang('{{ switch_target }}')" title="{{ switch_title }}" aria-label="{{ switch_title }}">{{ switch_label }}</button>
            </div>
        </div>
        
        <div class="stats-grid" id="statsContainer"></div>
        
        <div id="loading">{{ t_loading }}</div>
        <div id="content" style="display:none;">

            
            <div class="card">
                <h2 style="color: #059669; border-color: #a7f3d0;">📥 {{ t_sec_received }}</h2>
                <div id="receivedList"></div>
            </div>
            
            <div class="card">
                <h2 style="color: #2563eb; border-color: #bfdbfe;">📤 {{ t_sec_sent }}</h2>
                <div id="sentList"></div>
            </div>
            
            <div class="card">
                <h2 style="color: #7c3aed; border-color: #ddd6fe;">↩️ {{ t_sec_replied }}</h2>
                <div id="repliedList"></div>
            </div>
            
            <div class="card">
                <h2 style="color: #ea580c; border-color: #fdba74;">📩 {{ t_sec_incoming_letters }}</h2>
                <div id="incomingLettersList"></div>
            </div>
            
            <div class="card">
                <h2 style="color: #db2777; border-color: #fbcfe8;">📤 {{ t_sec_outgoing_letters }}</h2>
                <div id="outgoingLettersList"></div>
            </div>
        </div>
        <div class="digest-footer">{{ t_footer }}</div>

        <script>
            window.DCR_LANG = "{{ lang }}";
            window.DCR_DIR = "{{ dir_attr }}";
            window.DCR_I18N = {{ i18n_json|safe }};
            function _t(key, fallback) {
                const dict = window.DCR_I18N || {};
                return Object.prototype.hasOwnProperty.call(dict, key) ? dict[key] : (fallback || key);
            }
            function _esc(s) {
                return String(s ?? '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
            }
            function _tDt(code, fallback) {
                const fb = fallback || code || '';
                if (window.DCR_LANG !== 'ar' || !code) return fb;
                const k = 'dt_name.' + String(code).trim().toUpperCase();
                return _t(k, fb);
            }
            function _tStatus(val) {
                if (!val) return _t('magic.no_status', 'No Status');
                if (window.DCR_LANG !== 'ar') return String(val);
                const low = String(val).trim().toLowerCase();
                const map = {
                    'approved': 'status.approved',
                    'approved with comments': 'status.approved_with_comments',
                    'pending': 'status.pending',
                    'under review': 'status.under_review',
                    'open': 'status.open',
                    'in progress': 'status.in_progress',
                    'in-progress': 'status.in_progress',
                    'closed': 'status.closed',
                    'completed': 'status.completed',
                    'rejected': 'status.rejected',
                    'revise & resubmit': 'status.revise_resubmit',
                    'revise and resubmit': 'status.revise_resubmit',
                    'overdue': 'status.overdue',
                    'expired': 'status.expired',
                    'valid': 'status.valid',
                    'expiring': 'status.expiring',
                    'active': 'status.active',
                    'info / closed': 'status.info_closed',
                    'info/closed': 'status.info_closed',
                    'not closed': 'status.not_closed',
                    'cancelled': 'status.cancelled',
                    'canceled': 'status.cancelled',
                    'a - approved': 'status.a_approved',
                    'b - approved as noted': 'status.b_approved_noted',
                    'b,c - approved & resubmit': 'status.bc_approved_resubmit',
                    'c - revise & resubmit': 'status.c_revise_resubmit',
                    'd - review not required': 'status.d_review_not_required',
                    'replied': 'status.replied',
                    'accepted & to proceed with part c': 'status.noc_accepted_part_c',
                    'information required': 'status.info_required',
                    'needs maintenance': 'status.needs_maintenance',
                    'under refill': 'status.under_refill',
                    'out of service': 'status.out_of_service',
                    'draft': 'status.draft',
                    'revised': 'status.revised'
                };
                const k = map[low];
                return k ? _t(k, val) : String(val);
            }
            function switchDcrLang(nextLang) {
                const next = (String(nextLang || '').trim().toLowerCase() === 'ar') ? 'ar' : 'en';
                try {
                    document.cookie = 'dcr_lang=' + next + '; Path=/; Max-Age=31536000; SameSite=Lax';
                } catch(e) {}
                try {
                    const u = new URL(window.location.href);
                    u.searchParams.set('lang', next);
                    window.location.href = u.toString();
                } catch(e) {
                    window.location.reload();
                }
            }
            const PID = "{{ pid }}";
            const TOKEN = "{{ token }}";
            const datePicker = document.getElementById('datePicker');
            const roleFilter = document.getElementById('roleFilter');
            const loading = document.getElementById('loading');
            const content = document.getElementById('content');
            
            let rawData = null;
            
            // Set default date to today
            const today = new Date().toISOString().split('T')[0];
            datePicker.value = today;
            
            async function fetchData() {
                loading.style.display = 'block';
                content.style.display = 'none';
                
                try {
                    const res = await fetch(`/api/magic/data/${PID}?token=${TOKEN}&date=${datePicker.value}`);
                    if (!res.ok) throw new Error("Failed to fetch");
                    rawData = await res.json();
                    
                    // Populate Roles Dropdown if not already populated
                    if (roleFilter.options.length <= 2 && rawData.matrix) {
                        for (const discipline in rawData.matrix) {
                            if (discipline !== "Project Management") {
                                const opt = document.createElement('option');
                                opt.value = discipline;
                                opt.textContent = discipline;
                                roleFilter.appendChild(opt);
                            }
                        }
                    }
                    
                    renderData();
                } catch (e) {
                    alert(_t('magic.load_error', 'Error loading data.'));
                } finally {
                    loading.style.display = 'none';
                    if (rawData) content.style.display = 'block';
                }
            }
            
            function renderData() {
                if (!rawData) return;
                
                const selectedRole = roleFilter.value;
                const isPM = selectedRole === "__ALL__" || selectedRole === "";
                
                // Filter Documents Function
                function filterList(list, type) {
                    if(!list) return [];
                    const filtered = [];
                    for (const doc of list) {
                        if (type === 'docs' && doc.is_letter) continue;
                        if (type === 'letters' && !doc.is_letter) continue;
                        
                        const dDisc = (doc.discipline || "").toLowerCase().trim();
                        const rDisc = selectedRole.toLowerCase().trim();
                        
                        let show = isPM;
                        if (!show && dDisc && rDisc && (dDisc.includes(rDisc) || rDisc.includes(dDisc))) {
                            show = true;
                        }
                        
                        if (show) filtered.push(doc);
                    }
                    return filtered;
                }
                
                const receivedDocs = filterList(rawData.digest.received, 'docs');
                const sentDocs = filterList(rawData.digest.issued, 'docs');
                const repliedDocs = filterList(rawData.digest.replied, 'docs');
                
                // Letters
                // Deduplicate letters across all lists
                const allLettersRaw = [
                    ...(rawData.digest.received || []),
                    ...(rawData.digest.issued || []),
                    ...(rawData.digest.replied || [])
                ].filter(d => d.is_letter);
                
                const uniqueLettersMap = new Map();
                for (const l of allLettersRaw) {
                    uniqueLettersMap.set(l.id, l);
                }
                const uniqueLetters = filterList(Array.from(uniqueLettersMap.values()), 'letters');
                
                const incomingLetters = [];
                const outgoingLetters = [];
                for (const l of uniqueLetters) {
                    const dir = (l.direction || '').toLowerCase();
                    if (dir === 'received' || dir === 'in') {
                        incomingLetters.push(l);
                    } else {
                        outgoingLetters.push(l);
                    }
                }
                
                // Render Stats
                document.getElementById('statsContainer').innerHTML = `
                    <div class="stat-box"><div class="val bidi-ltr" style="color:#2563eb">${sentDocs.length}</div><div class="lbl">${_esc(_t('magic.kpi_sent_today', 'Sent Today'))}</div></div>
                    <div class="stat-box"><div class="val bidi-ltr" style="color:#7c3aed">${repliedDocs.length}</div><div class="lbl">${_esc(_t('magic.kpi_replied_today', 'Replied Today'))}</div></div>
                    <div class="stat-box"><div class="val bidi-ltr" style="color:#ea580c">${incomingLetters.length}</div><div class="lbl">${_esc(_t('magic.kpi_incoming_letters', 'Incoming Letters'))}</div></div>
                    <div class="stat-box"><div class="val bidi-ltr" style="color:#db2777">${outgoingLetters.length}</div><div class="lbl">${_esc(_t('magic.kpi_outgoing_letters', 'Outgoing Letters'))}</div></div>
                `;
                
                // Render Lists
                const renderList = (arr, elId, emptyMsg, hideStatus=false) => {
                    const el = document.getElementById(elId);
                    if (arr.length === 0) {
                        el.innerHTML = `<div class="empty">${_esc(emptyMsg)}</div>`;
                    } else {
                        el.innerHTML = arr.map(d => {
                            const disc = (d.discipline || '').trim();
                            const showDisc = disc && disc.toLowerCase() !== 'none' && disc.toLowerCase() !== 'no disc' && disc.toLowerCase() !== 'no disc.';
                            const dtCode = (d.dt_code || d.doc_type || '').trim();
                            const dtLabel = dtCode ? _tDt(dtCode, dtCode) : '';
                            return `
                                <div class="item">
                                    <div class="docno"><span class="bidi-ltr">${_esc(d.docNo)}</span>${dtLabel ? ` <span class="disc">${_esc(dtLabel)}</span>` : ''}</div>
                                    <div class="title">${_esc(d.title)}</div>
                                    <div>
                                        ${showDisc ? `<span class="disc">${_esc(disc)}</span>` : ''}
                                        ${!hideStatus ? `<span class="status-tag">${_esc(_tStatus(d.status))}</span>` : ''}
                                    </div>
                                </div>
                            `;
                        }).join('');
                    }
                };
                
                renderList(receivedDocs, 'receivedList', _t('magic.empty_received', 'No documents received today.'));
                renderList(sentDocs, 'sentList', _t('magic.empty_sent', 'No documents sent today.'));
                renderList(repliedDocs, 'repliedList', _t('magic.empty_replied', 'No documents replied today.'));
                renderList(incomingLetters, 'incomingLettersList', _t('magic.empty_incoming_letters', 'No incoming letters today.'), true);
                renderList(outgoingLetters, 'outgoingLettersList', _t('magic.empty_outgoing_letters', 'No outgoing letters today.'), true);
            }
            
            datePicker.addEventListener('change', fetchData);
            roleFilter.addEventListener('change', renderData);
            
            // Initial fetch
            fetchData();
        </script>
    </body>
    </html>
    """
    from flask import render_template_string
    return render_template_string(
        template,
        pid=pid,
        token=token,
        p_code=p_code,
        p_name=p_name,
        lang=lang,
        dir_attr=dir_attr,
        i18n_json=i18n_json,
        switch_label=t("lang.switch_label", lang),
        switch_target=t("lang.switch_target", lang),
        switch_title=t("lang.switch_title", lang),
        t_title=t("magic.title", lang),
        t_history_date=t("magic.history_date_title", lang),
        t_show_all=t("magic.show_all", lang),
        t_loading=t("magic.loading", lang),
        t_sec_received=t("magic.sec_received", lang),
        t_sec_sent=t("magic.sec_sent", lang),
        t_sec_replied=t("magic.sec_replied", lang),
        t_sec_incoming_letters=t("magic.sec_incoming_letters", lang),
        t_sec_outgoing_letters=t("magic.sec_outgoing_letters", lang),
        t_footer=t("magic.footer", lang),
        icon_tags=_head_icon_tags(),
    )

@app.route("/api/magic/data/<pid>", methods=["GET"])
def api_magic_data(pid):
    try:
        token = request.args.get("token", "")
        expected = _generate_master_magic_token(pid)
        if token != expected:
            return jsonify(error="Unauthorized"), 403
            
        date_str = request.args.get("date")
        
        # Get all doc types for the project
        dt_ids = [dt["id"] for dt in db.get_doc_types(pid)]
        
        digest = db.get_daily_digest(pid, dt_ids, target_date=date_str)
        
        # Matrix config
        projects = db.get_projects()
        proj = next((p for p in projects if p["id"] == pid or p.get("code") == pid), None)
        matrix = {}
        if proj and proj.get("data"):
            matrix = proj["data"].get("distribution_matrix", {})
            
        return jsonify({
            "digest": digest,
            "matrix": matrix
        })
    except Exception as e:
        return jsonify(error=str(e)), 500

@app.route("/api/daily_digest/<pid>", methods=["GET"])
def api_daily_digest(pid):
    try:
        u = current_user()
        if not u: return jsonify(error="LOGIN_REQUIRED"), 403
        
        projects_to_check = []
        if pid == "all":
            ws = request.args.get("workspace_id") or request.args.get("ws")
            projects_to_check = get_allowed_project_ids(u, workspace_id=ws)
        else:
            if not can_view_project(pid): return jsonify(error="Forbidden"), 403
            resolved = _resolve_project(pid)
            if not resolved: return jsonify(error="Forbidden"), 403
            projects_to_check = [resolved["id"]]
            
        combined_digest = {"received": [], "issued": [], "replied": []}
        for p_id in projects_to_check:
            try:
                # Filter digest document types strictly to those the current user is allowed to read.
                all_dts = filter_allowed_doc_types(p_id, db.get_doc_types(p_id), u)
                assigned_dt_ids = [dt["id"] for dt in all_dts]
                if assigned_dt_ids:
                    target_date = request.args.get("date")
                    d = db.get_daily_digest(p_id, assigned_dt_ids, target_date)
                    combined_digest["received"].extend(d.get("received", []))
                    combined_digest["issued"].extend(d.get("issued", []))
                    combined_digest["replied"].extend(d.get("replied", []))
            except Exception as inner_e:
                import traceback
                error_msg = traceback.format_exc()
                print(error_msg)
                return jsonify({"success": False, "error": str(inner_e), "trace": error_msg}), 500
                
        return jsonify(combined_digest)
    except Exception as e:
        import traceback
        error_msg = traceback.format_exc()
        print(error_msg)
        return jsonify({"success": False, "error": str(e), "trace": error_msg}), 500


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
                    process_drive_folder(folder_id, project_id=p.get("id"))
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
