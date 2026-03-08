"""app.py - DCR Flask v1"""
import os, json, uuid, io, base64, datetime
from flask import (Flask, request, session, redirect, url_for,
                   jsonify, make_response, send_file, g)
import db
from utils import (compute_expected_reply, compute_duration, is_overdue,
                   format_date, get_next_doc_no, extract_rev, STATUS_COLORS,
                   REJECTED_STATUSES, invalidate_holidays_cache)

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", os.urandom(32))
IS_RENDER = bool(os.environ.get("RENDER"))

# ── Session cookie settings ───────────────────────────────────
app.config.update(
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE="Lax",
    SESSION_COOKIE_SECURE=IS_RENDER,
    PERMANENT_SESSION_LIFETIME=datetime.timedelta(hours=8),
)

# ── Auth helpers ──────────────────────────────────────────────
def current_user():
    token = request.cookies.get("dcr_token")
    if not token: return None
    return db.get_session(token)

def require_login(fn):
    from functools import wraps
    @wraps(fn)
    def wrapper(*a, **kw):
        if not current_user():
            return jsonify(error="LOGIN_REQUIRED"), 403
        return fn(*a, **kw)
    return wrapper

def require_superadmin(fn):
    from functools import wraps
    @wraps(fn)
    def wrapper(*a, **kw):
        u = current_user()
        if not u or u["role"] != "superadmin":
            return jsonify(error="Forbidden"), 403
        return fn(*a, **kw)
    return wrapper

def can_edit(project_id):
    u = current_user()
    if not u: return False
    if u["role"] == "superadmin": return True
    if u["role"] in ("admin","editor"):
        return project_id in db.get_user_projects(u["username"])
    return False

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

@app.route("/login", methods=["GET","POST"])
def login():
    if request.method == "POST":
        data  = request.get_json(silent=True) or {}
        uname = data.get("username","").strip().lower()
        pw    = data.get("password","")
        if db.verify_pw(uname, pw):
            u     = db.get_user(uname)
            token = db.create_session(uname, u["role"])
            resp  = make_response(jsonify(ok=True, role=u["role"], username=uname))
            resp.set_cookie("dcr_token", token, httponly=True,
                            samesite="Lax", secure=IS_RENDER,
                            max_age=db.SESSION_TTL)
            return resp
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

# ── API: Projects ─────────────────────────────────────────────
@app.route("/api/projects")
def api_projects():
    u = current_user()
    projects = db.get_projects()
    out = []
    for p in projects:
        try: data = p.get("data") or {}
        except: data = {}
        out.append({
            "id": p["id"], "name": p["name"], "code": p["code"],
            "client": data.get("client","") if isinstance(data,dict) else "",
            "can_edit": can_edit(p["id"])
        })
    return jsonify(out)

@app.route("/api/project/<pid>")
def api_get_project(pid):
    p = db.get_project(pid)
    if not p: return jsonify(error="Not found"), 404
    return jsonify(p)

@app.route("/api/project/<pid>", methods=["POST"])
def api_save_project(pid):
    if not can_edit(pid):
        return jsonify(error="LOGIN_REQUIRED"), 403
    data = request.get_json(silent=True) or {}
    name = data.pop("name", pid)
    code = data.pop("code", pid)
    data.pop("id", None)
    db.save_project(pid, name, code, data)
    return jsonify(ok=True)

@app.route("/api/projects/create", methods=["POST"])
@require_superadmin
def api_create_project():
    data = request.get_json(silent=True) or {}
    pid  = data.get("id","").strip().upper()
    name = data.get("name","").strip()
    code = data.get("code","").strip().upper()
    if not pid or not name or not code:
        return jsonify(ok=False, error="ID, name and code required"), 400
    u = current_user()
    db.create_project(pid, name, code, creator=u["username"])
    return jsonify(ok=True)

@app.route("/api/projects/delete/<pid>", methods=["POST"])
@require_superadmin
def api_delete_project(pid):
    db.delete_project(pid)
    return jsonify(ok=True)

# ── API: Doc Types ────────────────────────────────────────────
@app.route("/api/doc_types/<pid>")
def api_doc_types(pid):
    return jsonify(db.get_doc_types(pid))

@app.route("/api/doc_types/<pid>", methods=["POST"])
def api_add_doc_type(pid):
    if not can_edit(pid): return jsonify(error="LOGIN_REQUIRED"), 403
    data = request.get_json(silent=True) or {}
    code = data.get("code","").strip().upper()
    name = data.get("name","").strip()
    if not code or not name:
        return jsonify(ok=False, error="Code and name required"), 400
    db.add_doc_type(pid, code, name)
    return jsonify(ok=True)

@app.route("/api/doc_types/<pid>/<dt_id>", methods=["DELETE"])
def api_delete_doc_type(pid, dt_id):
    if not can_edit(pid): return jsonify(error="LOGIN_REQUIRED"), 403
    db.delete_doc_type(pid, dt_id)
    return jsonify(ok=True)

# ── API: Columns ──────────────────────────────────────────────
@app.route("/api/columns/<pid>/<dt_id>")
def api_columns(pid, dt_id):
    return jsonify(db.get_columns(pid, dt_id))

@app.route("/api/columns/<pid>/<dt_id>", methods=["POST"])
def api_add_column(pid, dt_id):
    if not can_edit(pid): return jsonify(error="LOGIN_REQUIRED"), 403
    data = request.get_json(silent=True) or {}
    key  = "custom_" + data.get("label","").lower().replace(" ","_") + "_" + uuid.uuid4().hex[:6]
    db.add_column(pid, dt_id, key, data.get("label",""),
                  data.get("col_type","text"), data.get("list_name"))
    return jsonify(ok=True)

@app.route("/api/columns/visibility/<int:col_id>", methods=["POST"])
def api_col_visibility(col_id):
    data = request.get_json(silent=True) or {}
    db.set_col_visible(col_id, data.get("visible", True))
    return jsonify(ok=True)

@app.route("/api/columns/<int:col_id>", methods=["DELETE"])
def api_delete_col(col_id):
    db.delete_col(col_id)
    return jsonify(ok=True)

@app.route("/api/columns/rename/<int:col_id>", methods=["POST"])
def api_rename_col(col_id):
    if not current_user(): return jsonify(error="LOGIN_REQUIRED"), 403
    data = request.get_json(silent=True) or {}
    label = data.get("label","").strip()
    if not label: return jsonify(error="Label required"), 400
    db.rename_column(col_id, label)
    return jsonify(ok=True)

@app.route("/api/col_width/<pid>/<dt_id>", methods=["GET"])
def api_get_col_widths(pid, dt_id):
    return jsonify(db.get_col_widths(pid, dt_id))

@app.route("/api/col_width/<pid>/<dt_id>", methods=["POST"])
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
@app.route("/api/records/<pid>/<dt_id>")
def api_records(pid, dt_id):
    search  = request.args.get("search","")
    records = db.get_records(pid, dt_id, search=search)
    cols    = db.get_columns(pid, dt_id)
    date_col_keys = {c["col_key"] for c in cols if c.get("col_type") in ("date","auto_date")}
    for row in records:
        row["_expectedReplyDate"] = format_date(
            compute_expected_reply(row.get("issuedDate"), row.get("docNo")))
        issued   = row.get("issuedDate","")
        actual   = row.get("actualReplyDate","")
        dur = compute_duration(issued, actual)
        row["_duration"] = str(dur) if dur is not None else ""
        row["_overdue"]    = is_overdue(row.get("issuedDate"), row.get("docNo"), row.get("actualReplyDate"))
        row["_isRev"]      = extract_rev(row.get("docNo","")) > 0
        # Format ALL date columns (any col_type=date)
        for dk in date_col_keys:
            if dk in row and row[dk]:
                row["_fmt_" + dk] = format_date(row[dk])
        # Standard aliases
        row["_issuedFmt"]  = format_date(row.get("issuedDate",""))
        row["_replyFmt"]   = format_date(row.get("actualReplyDate",""))
    return jsonify(records=records, columns=cols, count=db.count_records(pid, dt_id))

@app.route("/api/records/<pid>/<dt_id>", methods=["POST"])
def api_save_record(pid, dt_id):
    if not can_edit(pid): return jsonify(error="LOGIN_REQUIRED"), 403
    data   = request.get_json(silent=True) or {}
    rec_id = data.pop("_id", None) or str(uuid.uuid4())
    db.save_record(pid, dt_id, rec_id, data)
    return jsonify(ok=True, id=rec_id)

@app.route("/api/records/<rec_id>", methods=["DELETE"])
def api_delete_record(rec_id):
    db.delete_record(rec_id)
    return jsonify(ok=True)

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
    prefix = dt["code"] if dt else dt_id
    return jsonify(next=get_next_doc_no(prefix, db.get_records(pid, dt_id)))

# ── API: Dropdown Lists ───────────────────────────────────────
@app.route("/api/lists/<pid>")
def api_lists(pid):
    return jsonify(db.get_lists(pid))

@app.route("/api/lists/<pid>", methods=["POST"])
def api_add_list_item(pid):
    if not can_edit(pid): return jsonify(error="LOGIN_REQUIRED"), 403
    data = request.get_json(silent=True) or {}
    db.add_list_item(pid, data.get("list_name",""), data.get("item",""))
    return jsonify(ok=True)

@app.route("/api/lists/<pid>", methods=["DELETE"])
def api_remove_list_item(pid):
    if not can_edit(pid): return jsonify(error="LOGIN_REQUIRED"), 403
    data = request.get_json(silent=True) or {}
    db.remove_list_item(pid, data.get("list_name",""), data.get("item",""))
    return jsonify(ok=True)

# ── API: Logos ────────────────────────────────────────────────
@app.route("/api/logo/<pid>/<key>")
def api_get_logo(pid, key):
    data = db.get_logo(pid, key)
    if not data: return "", 404
    raw = base64.b64decode(data)
    return send_file(io.BytesIO(raw), mimetype="image/png")

@app.route("/api/logo/<pid>/<key>", methods=["POST"])
def api_save_logo(pid, key):
    if not can_edit(pid): return jsonify(error="LOGIN_REQUIRED"), 403
    data = request.get_json(silent=True) or {}
    db.save_logo(pid, key, data.get("data",""))
    return jsonify(ok=True)

# ── API: Users ────────────────────────────────────────────────
@app.route("/api/users")
@require_superadmin
def api_users():
    return jsonify(db.get_all_users())

@app.route("/api/users", methods=["POST"])
@require_superadmin
def api_users_action():
    data   = request.get_json(silent=True) or {}
    action = data.get("action")
    if action == "add":
        uname = data.get("username","").strip().lower()
        pw    = data.get("password","")
        role  = data.get("role","viewer")
        if not uname or not pw:
            return jsonify(ok=False, error="Username and password required"), 400
        db.add_user(uname, pw, role)
        return jsonify(ok=True)
    if action == "delete":
        uname = data.get("username","")
        if uname == "admin":
            return jsonify(ok=False, error="Cannot delete super admin"), 400
        db.delete_user(uname)
        return jsonify(ok=True)
    if action == "change_password":
        db.change_pw(data.get("username",""), data.get("password",""))
        return jsonify(ok=True)
    if action == "assign":
        db.assign_project(data.get("username",""), data.get("project_id",""))
        return jsonify(ok=True)
    if action == "unassign":
        db.unassign_project(data.get("username",""), data.get("project_id",""))
        return jsonify(ok=True)
    return jsonify(ok=False, error="Unknown action"), 400

@app.route("/api/users/<username>/projects")
@require_superadmin
def api_user_projects(username):
    return jsonify(db.get_user_projects(username))

@app.route("/api/whoami")
def api_whoami():
    u = current_user()
    if not u: return jsonify(username="guest", role="guest", projects=[])
    projs = [] if u["role"] == "superadmin" else db.get_user_projects(u["username"])
    return jsonify(username=u["username"], role=u["role"], projects=projs)

@app.route("/api/change_password", methods=["POST"])
def api_change_own_pw():
    u = current_user()
    if not u: return jsonify(error="LOGIN_REQUIRED"), 403
    data = request.get_json(silent=True) or {}
    pw   = data.get("password","")
    if len(pw) < 4: return jsonify(ok=False, error="Min 4 characters"), 400
    db.change_pw(u["username"], pw)
    return jsonify(ok=True)

# ── Export Excel ──────────────────────────────────────────────
@app.route("/api/columns/reorder/<pid>/<dt_id>", methods=["POST"])
def api_reorder_cols(pid, dt_id):
    if not can_edit(pid): return jsonify(error="LOGIN_REQUIRED"), 403
    data = request.get_json(silent=True) or {}
    db.reorder_columns(pid, dt_id, data.get("order", []))
    return jsonify(ok=True)

@app.route("/api/export_all/<pid>")
def api_export_all(pid):
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    proj    = db.get_project(pid) or {}
    dts     = db.get_doc_types(pid)
    wb      = openpyxl.Workbook()
    wb.remove(wb.active)  # remove default sheet

    PRIMARY="1A3A5C"; PL="2563A8"; WHITE="FFFFFF"; ALT="F8FAFC"; OV="FFF5F5"; MUTED="9CA3AF"
    STATUS_XL = {
        "A - Approved":              ("BBF7D0","166534"),
        "B - Approved As Noted":     ("DCFCE7","14532D"),
        "B,C - Approved & Resubmit": ("FED7AA","7C2D12"),
        "C - Revise & Resubmit":     ("FCE7F3","831843"),
        "D - Review not Required":   ("FECACA","7F1D1D"),
        "Under Review":              ("FEF9C3","713F12"),
        "Cancelled":                 ("EF4444","FFFFFF"),
        "Open":                      ("FED7AA","7C2D12"),
        "Closed":                    ("BFDBFE","1E3A5F"),
        "Replied":                   ("D1FAE5","064E3B"),
        "Pending":                   ("E0E7FF","312E81"),
    }
    def fill(c): return PatternFill("solid", fgColor=c)
    def thin(): s=Side(style="thin",color="DDE3ED"); return Border(left=s,right=s,top=s,bottom=s)

    for dt in dts:
        cols    = [c for c in db.get_columns(pid, dt["id"]) if c["visible"]]
        records = db.get_records(pid, dt["id"])
        # always include tab (even if empty)

        ws = wb.create_sheet(title=dt["id"][:31])
        ws.sheet_view.showGridLines = False
        all_cols = [{"col_key":"_sr","label":"Sr."}] + [{"col_key":c["col_key"],"label":c["label"]} for c in cols]
        nc = len(all_cols)

        from openpyxl.utils import get_column_letter as gcl
        def mcell(row, val, bg, fg="FFFFFF", bold=False, sz=11):
            c = ws.cell(row=row, column=1, value=val)
            ws.merge_cells(f"A{row}:{gcl(nc)}{row}")
            c.font = Font(bold=bold, color=fg, size=sz, name="Arial")
            c.fill = fill(bg)
            c.alignment = Alignment(horizontal="left", vertical="center")
            return c

        mcell(1, f"{dt['name'].upper()} — {proj.get('name','')} ({proj.get('code','')})",
              PRIMARY, bold=True, sz=12)
        ws.row_dimensions[1].height = 30
        ws.row_dimensions[2].height = 4

        COL_W = {"_sr":5,"docNo":22,"discipline":14,"trade":14,"title":38,"floor":12,
                 "itemRef":16,"issuedDate":13,"expectedReplyDate":14,"actualReplyDate":13,
                 "status":24,"duration":10,"remarks":28,"fileLocation":18}
        CENTER = {"_sr","duration","issuedDate","expectedReplyDate","actualReplyDate"}

        for ci, col in enumerate(all_cols, 1):
            ws.column_dimensions[gcl(ci)].width = COL_W.get(col["col_key"],13)
            c = ws.cell(row=3, column=ci, value=col["label"])
            c.font = Font(bold=True,color=WHITE,size=10,name="Arial")
            c.fill = fill(PRIMARY)
            c.alignment = Alignment(horizontal="center",vertical="center",wrap_text=True)
            c.border = thin()
        ws.row_dimensions[3].height = 22

        sr = 1
        if not records:
            # Empty tab — write a "no records" row
            c = ws.cell(row=4, column=1, value="No records in this register")
            c.font = Font(italic=True, size=10, name="Arial", color="9CA3AF")
            ws.merge_cells(f"A4:{gcl(nc)}4")
            c.alignment = Alignment(horizontal="center", vertical="center")
            ws.row_dimensions[4].height = 24
        for ri, row in enumerate(records):
            rn = 4 + ri
            is_rev = extract_rev(row.get("docNo","")) > 0
            ov     = is_overdue(row.get("issuedDate"), row.get("docNo"), row.get("actualReplyDate"))
            bg     = OV if ov else (ALT if sr%2==0 else WHITE)
            ws.row_dimensions[rn].height = 18
            for ci, col in enumerate(all_cols, 1):
                key = col["col_key"]
                if key=="_sr":                   val = "" if is_rev else str(sr)
                elif key=="expectedReplyDate":   val = format_date(compute_expected_reply(row.get("issuedDate"),row.get("docNo")))
                elif key=="duration":            val = str(compute_duration(row.get("issuedDate"),row.get("actualReplyDate")) or "")
                elif key in ("issuedDate","actualReplyDate"): val = format_date(row.get(key,""))
                else:                            val = str(row.get(key,"") or "")
                c = ws.cell(row=rn, column=ci, value=val)
                c.border = thin()
                c.alignment = Alignment(vertical="center",
                                        horizontal="center" if key in CENTER else "left")
                if key=="status" and val:
                    bg2, fg2 = STATUS_XL.get(val, ("F3F4F6","374151"))
                    c.fill = fill(bg2); c.font = Font(bold=True,size=9,name="Arial",color=fg2)
                else:
                    c.fill = fill(bg)
                    c.font = Font(size=10,name="Arial",
                                  color=MUTED if is_rev else ("991B1B" if ov else "1E2A3A"))
            if not is_rev: sr += 1

    if not wb.sheetnames:
        ws = wb.create_sheet("Empty"); ws.cell(1,1,"No data")

    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    fname = f"{proj.get('code','DCR')}_All_Registers.xlsx"
    return send_file(buf, as_attachment=True, download_name=fname,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

@app.route("/api/export/<pid>/<dt_id>")
def api_export(pid, dt_id):
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    proj    = db.get_project(pid) or {}
    cols    = [c for c in db.get_columns(pid, dt_id) if c["visible"]]
    records = db.get_records(pid, dt_id)
    dts     = db.get_doc_types(pid)
    dt      = next((d for d in dts if d["id"] == dt_id), None)

    PRIMARY = "1A3A5C"; PL = "2563A8"; WHITE = "FFFFFF"
    ALT = "F8FAFC"; OV = "FFF5F5"; MUTED = "9CA3AF"

    STATUS_XL = {
        "A - Approved":              ("BBF7D0","166534"),
        "B - Approved As Noted":     ("DCFCE7","14532D"),
        "B,C - Approved & Resubmit": ("FED7AA","7C2D12"),
        "C - Revise & Resubmit":     ("FCE7F3","831843"),
        "D - Review not Required":   ("FECACA","7F1D1D"),
        "Under Review":              ("FEF9C3","713F12"),
        "Cancelled":                 ("EF4444","FFFFFF"),
        "Open":                      ("FED7AA","7C2D12"),
        "Closed":                    ("BFDBFE","1E3A5F"),
        "Replied":                   ("D1FAE5","064E3B"),
        "Pending":                   ("E0E7FF","312E81"),
    }

    def fill(c): return PatternFill("solid", fgColor=c)
    def thin(): s=Side(style="thin",color="DDE3ED"); return Border(left=s,right=s,top=s,bottom=s)

    wb = openpyxl.Workbook(); ws = wb.active
    ws.title = dt_id; ws.sheet_view.showGridLines = False

    all_cols = [{"col_key":"_sr","label":"Sr."}] + [{"col_key":c["col_key"],"label":c["label"]} for c in cols]
    nc = len(all_cols)

    def mcell(row, val, bg, fg="FFFFFF", bold=False, sz=11, halign="center"):
        c = ws.cell(row=row, column=1, value=val)
        ws.merge_cells(f"A{row}:{get_column_letter(nc)}{row}")
        c.font = Font(bold=bold, color=fg, size=sz, name="Arial")
        c.fill = fill(bg)
        c.alignment = Alignment(horizontal=halign, vertical="center")
        return c

    mcell(1, f"DOCUMENT CONTROL REGISTER  —  {(dt['name'] if dt else dt_id).upper()}",
          PRIMARY, bold=True, sz=13)
    ws.row_dimensions[1].height = 36

    info = "   |   ".join(f"{k}: {v}" for k,v in [
        ("Project",proj.get("name","")),("Code",proj.get("code","")),
        ("Client",proj.get("client","")),("Consultant",proj.get("mainConsultant","")),
        ("Exported",datetime.datetime.now().strftime("%d/%b/%Y %H:%M"))] if v)
    mcell(2, info, PL, sz=9); ws.row_dimensions[2].height=18
    ws.row_dimensions[3].height = 4

    COL_W = {"_sr":5,"docNo":22,"discipline":14,"trade":14,"title":38,"floor":12,
             "itemRef":16,"issuedDate":13,"expectedReplyDate":14,"actualReplyDate":13,
             "status":24,"duration":10,"remarks":28,"fileLocation":18}
    CENTER = {"_sr","duration","issuedDate","expectedReplyDate","actualReplyDate"}

    for ci, col in enumerate(all_cols, 1):
        ws.column_dimensions[get_column_letter(ci)].width = COL_W.get(col["col_key"],13)
        c = ws.cell(row=4, column=ci, value=col["label"])
        c.font = Font(bold=True,color=WHITE,size=10,name="Arial")
        c.fill = fill(PRIMARY)
        c.alignment = Alignment(horizontal="center",vertical="center",wrap_text=True)
        c.border = thin()
    ws.row_dimensions[4].height = 24

    sr = 1
    for ri, row in enumerate(records):
        rn   = 5 + ri
        is_rev = extract_rev(row.get("docNo","")) > 0
        ov     = is_overdue(row.get("issuedDate"), row.get("docNo"), row.get("actualReplyDate"))
        bg     = OV if ov else (ALT if sr%2==0 else WHITE)
        ws.row_dimensions[rn].height = 20

        for ci, col in enumerate(all_cols, 1):
            key = col["col_key"]
            if key=="_sr":                  val = "" if is_rev else str(sr)
            elif key=="expectedReplyDate":  val = format_date(compute_expected_reply(row.get("issuedDate"),row.get("docNo")))
            elif key=="duration":           val = str(compute_duration(row.get("issuedDate"),row.get("actualReplyDate")) or "")
            elif key=="issuedDate":         val = format_date(row.get(key,""))
            elif key=="actualReplyDate":    val = format_date(row.get(key,""))
            else:                           val = str(row.get(key,"") or "")

            c = ws.cell(row=rn, column=ci, value=val)
            c.border    = thin()
            c.alignment = Alignment(vertical="center",
                                    horizontal="center" if key in CENTER else "left")
            if key=="status" and val:
                bg2, fg2 = STATUS_XL.get(val, ("F3F4F6","374151"))
                c.fill = fill(bg2); c.font = Font(bold=True,size=9,name="Arial",color=fg2)
            elif key=="docNo":
                c.fill = fill(bg)
                c.font = Font(size=10,name="Consolas",bold=not is_rev,
                              color=MUTED if is_rev else PRIMARY)
            else:
                c.fill = fill(bg)
                c.font = Font(size=10,name="Arial",
                              color=MUTED if is_rev else ("991B1B" if ov else "1E2A3A"))
        if not is_rev: sr += 1

    tot = 5 + len(records)
    real = sum(1 for r in records if extract_rev(r.get("docNo",""))==0)
    mcell(tot, f"TOTAL: {real} documents  |  {len(records)} submissions", PRIMARY, sz=10, halign="left")
    ws.row_dimensions[tot].height = 22

    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    fname = f"{proj.get('code','DCR')}_{dt_id}_Register.xlsx"
    return send_file(buf, as_attachment=True, download_name=fname,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ── PDF Export ────────────────────────────────────────────────

def _build_pdf_for_dt(pid, dt_id, proj, buf=None):
    """Build a PDF for one document type. Returns BytesIO."""
    from reportlab.platypus import (SimpleDocTemplate, Table, TableStyle,
                                    Paragraph, Spacer, HRFlowable)
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.lib import colors as rl_colors

    buf = buf or io.BytesIO()
    cols    = [c for c in db.get_columns(pid, dt_id) if c["visible"]]
    records = db.get_records(pid, dt_id)
    dts     = db.get_doc_types(pid)
    dt      = next((d for d in dts if d["id"] == dt_id), {"name": dt_id, "code": dt_id})

    PAGE = landscape(A4)
    doc = SimpleDocTemplate(buf, pagesize=PAGE,
                            leftMargin=10*mm, rightMargin=10*mm,
                            topMargin=12*mm, bottomMargin=12*mm)
    styles = getSampleStyleSheet()
    DARK   = rl_colors.HexColor("#1A3A5C")
    LIGHT  = rl_colors.HexColor("#F8FAFC")
    OV_COL = rl_colors.HexColor("#FEF2F2")
    REV_COL= rl_colors.HexColor("#F0F9FF")
    ALT    = rl_colors.HexColor("#F0F4F8")

    STATUS_PDF = {
        "A - Approved":              rl_colors.HexColor("#BBF7D0"),
        "B - Approved As Noted":     rl_colors.HexColor("#DCFCE7"),
        "B,C - Approved & Resubmit": rl_colors.HexColor("#FED7AA"),
        "C - Revise & Resubmit":     rl_colors.HexColor("#FCE7F3"),
        "D - Review not Required":   rl_colors.HexColor("#FECACA"),
        "Under Review":              rl_colors.HexColor("#FEF9C3"),
        "Cancelled":                 rl_colors.HexColor("#EF4444"),
        "Open":                      rl_colors.HexColor("#FED7AA"),
        "Closed":                    rl_colors.HexColor("#BFDBFE"),
        "Replied":                   rl_colors.HexColor("#D1FAE5"),
        "Pending":                   rl_colors.HexColor("#E0E7FF"),
    }

    pstyle = ParagraphStyle("cell", fontName="Helvetica", fontSize=7, leading=9,
                             wordWrap="LTR", spaceAfter=0, spaceBefore=0)
    hstyle = ParagraphStyle("hdr", fontName="Helvetica-Bold", fontSize=7.5,
                             leading=10, textColor=rl_colors.white)

    hdr_cols = [{"col_key":"_sr","label":"Sr."}] + [{"col_key":c["col_key"],"label":c["label"]} for c in cols]

    # Column widths (points)
    W_MAP = {"_sr":18,"docNo":68,"discipline":55,"trade":55,"title":110,
             "floor":38,"itemRef":48,"issuedDate":46,"expectedReplyDate":48,
             "actualReplyDate":46,"status":80,"duration":34,"remarks":80,"fileLocation":60}
    page_w = PAGE[0] - 20*mm
    col_ws = [W_MAP.get(c["col_key"], 55) for c in hdr_cols]
    scale  = page_w / sum(col_ws)
    col_ws = [w * scale for w in col_ws]

    # Build header row
    hdr_row = [Paragraph(c["label"], hstyle) for c in hdr_cols]
    data    = [hdr_row]
    row_meta = []   # (bg, is_header)

    sr = 1
    for row in records:
        is_rev = extract_rev(row.get("docNo","")) > 0
        ov     = is_overdue(row.get("issuedDate"), row.get("docNo"), row.get("actualReplyDate"))
        cells  = []
        for c in hdr_cols:
            key = c["col_key"]
            if key == "_sr":
                val = "" if is_rev else str(sr)
            elif key == "expectedReplyDate":
                val = format_date(compute_expected_reply(row.get("issuedDate"), row.get("docNo")))
            elif key == "duration":
                val = str(compute_duration(row.get("issuedDate"), row.get("actualReplyDate")) or "")
            elif key in ("issuedDate","actualReplyDate"):
                val = format_date(row.get(key,""))
            else:
                val = str(row.get(key,"") or "")
            cells.append(Paragraph(val, pstyle))
        data.append(cells)
        if ov:     row_meta.append(OV_COL)
        elif is_rev: row_meta.append(REV_COL)
        elif sr % 2 == 0: row_meta.append(ALT)
        else:      row_meta.append(rl_colors.white)
        if not is_rev: sr += 1

    # Build table style
    ts = [
        ("BACKGROUND", (0,0), (-1,0), DARK),
        ("TEXTCOLOR",  (0,0), (-1,0), rl_colors.white),
        ("FONTNAME",   (0,0), (-1,0), "Helvetica-Bold"),
        ("FONTSIZE",   (0,0), (-1,-1), 7),
        ("ROWBACKGROUND", (0,0), (-1,0), DARK),
        ("VALIGN",     (0,0), (-1,-1), "MIDDLE"),
        ("TOPPADDING", (0,0), (-1,-1), 2),
        ("BOTTOMPADDING",(0,0),(-1,-1), 2),
        ("LEFTPADDING",(0,0),(-1,-1), 2),
        ("RIGHTPADDING",(0,0),(-1,-1), 2),
        ("GRID",       (0,0), (-1,-1), 0.3, rl_colors.HexColor("#CBD5E1")),
    ]
    # Row backgrounds
    for i, bg in enumerate(row_meta, start=1):
        ts.append(("BACKGROUND", (0,i), (-1,i), bg))
    # Status column color
    status_idx = next((i for i,c in enumerate(hdr_cols) if c["col_key"]=="status"), None)
    if status_idx is not None:
        for i, row in enumerate(records, start=1):
            sv = row.get("status","")
            bg = STATUS_PDF.get(sv)
            if bg: ts.append(("BACKGROUND", (status_idx,i), (status_idx,i), bg))

    tbl = Table(data, colWidths=col_ws, repeatRows=1)
    tbl.setStyle(TableStyle(ts))

    proj_name = proj.get("name","") if isinstance(proj, dict) else ""
    title_st  = ParagraphStyle("t", fontName="Helvetica-Bold", fontSize=12,
                                textColor=DARK, spaceAfter=4)
    sub_st    = ParagraphStyle("s", fontName="Helvetica", fontSize=9,
                                textColor=rl_colors.HexColor("#64748B"), spaceAfter=8)

    story = [
        Paragraph(f"{dt['name'].upper()} — {proj_name}", title_st),
        Paragraph(f"Project: {proj.get('code','')}  |  Exported: {datetime.datetime.now().strftime('%d-%m-%Y %H:%M')}", sub_st),
        HRFlowable(width="100%", thickness=1, color=DARK, spaceAfter=6),
        tbl,
    ]
    doc.build(story)
    buf.seek(0)
    return buf


@app.route("/api/export_pdf/<pid>/<dt_id>")
def api_export_pdf(pid, dt_id):
    proj = db.get_project(pid) or {}
    buf  = _build_pdf_for_dt(pid, dt_id, proj)
    fname = f"{proj.get('code','DCR')}_{dt_id}_Register.pdf"
    return send_file(buf, as_attachment=True, download_name=fname,
                     mimetype="application/pdf")


@app.route("/api/export_pdf_all/<pid>")
def api_export_pdf_all(pid):
    from reportlab.platypus import SimpleDocTemplate, PageBreak
    proj = db.get_project(pid) or {}
    dts  = db.get_doc_types(pid)

    # Build each DT as separate PDF bytes then merge
    try:
        from pypdf import PdfWriter
        writer = PdfWriter()
        for dt in dts:
            if not db.count_records(pid, dt["id"]): continue
            single_buf = _build_pdf_for_dt(pid, dt["id"], proj)
            from pypdf import PdfReader
            reader = PdfReader(single_buf)
            for page in reader.pages:
                writer.add_page(page)
        out = io.BytesIO()
        writer.write(out); out.seek(0)
    except ImportError:
        # Fallback: just export first DT or give error
        out = _build_pdf_for_dt(pid, dts[0]["id"], proj) if dts else io.BytesIO()

    fname = f"{proj.get('code','DCR')}_All_Registers.pdf"
    return send_file(out, as_attachment=True, download_name=fname,
                     mimetype="application/pdf")


# ── Import ────────────────────────────────────────────────────
@app.route("/api/columns/reorder", methods=["POST"])
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


@app.route("/api/import/<pid>/<dt_id>", methods=["POST"])
def api_import(pid, dt_id):
    if not can_edit(pid): return jsonify(error="LOGIN_REQUIRED"), 403
    data = request.get_json(silent=True) or {}
    b64  = data.get("file_b64","")
    ext  = data.get("ext","csv")
    if "," in b64: b64 = b64.split(",",1)[1]

    cols    = db.get_columns(pid, dt_id)
    col_map = {c["label"]: c["col_key"] for c in cols}
    imported = 0

    try:
        if ext in ("xlsx","xls"):
            import openpyxl, datetime as _dt
            wb  = openpyxl.load_workbook(io.BytesIO(base64.b64decode(b64)), data_only=True)
            ws  = wb.active; header = None
            for row in ws.iter_rows(values_only=True):
                vals = []
                for v in row:
                    if v is None: vals.append("")
                    elif isinstance(v, (_dt.datetime,_dt.date)):
                        vals.append((v.date() if isinstance(v,_dt.datetime) else v).strftime("%Y-%m-%d"))
                    else: vals.append(str(v).strip())
                if not any(vals): continue
                if header is None:
                    if any(v in col_map or v in ("Sr.","Document No.") for v in vals):
                        header = [col_map.get(v.strip(),v.strip()) for v in vals]
                    continue
                row_data = {header[i]:v for i,v in enumerate(vals)
                            if i<len(header) and header[i] and header[i] not in ("Sr.","sr","")}
                if any(row_data.values()):
                    db.save_record(pid, dt_id, str(uuid.uuid4()), row_data); imported+=1
        else:
            import csv, datetime as _dt
            text   = base64.b64decode(b64).decode("utf-8","ignore")
            header = None
            date_cols = {c["col_key"] for c in cols if c.get("col_type") in ("date","auto_date")}
            for line in csv.reader(io.StringIO(text)):
                if not header:
                    if any(c in line for c in ["Document No.","Sr.","docNo"]):
                        header = [col_map.get(h.strip(),h.strip()) for h in line]
                    continue
                if not any(line): continue
                row_data = {}
                for i,val in enumerate(line):
                    if i<len(header) and header[i] and header[i] not in ("sr","Sr.","Sr",""):
                        v = val.strip()
                        if header[i] in date_cols and v:
                            for fmt in ("%d/%b/%Y","%d-%b-%Y","%Y-%m-%d","%d/%m/%Y"):
                                try: v = _dt.datetime.strptime(v,fmt).strftime("%Y-%m-%d"); break
                                except: pass
                        row_data[header[i]] = v
                if any(row_data.values()):
                    db.save_record(pid, dt_id, str(uuid.uuid4()), row_data); imported+=1
    except Exception as e:
        return jsonify(ok=False, error=str(e)), 500

    return jsonify(ok=True, imported=imported)

# ═══════════════════════════════════════════════════════════════
# HTML RENDERING
# ═══════════════════════════════════════════════════════════════

def _user_info_html(u):
    if not u:
        return ('<a href="/login"><button class="tb-btn glow">🔐 Login</button></a>', "guest", "GUEST", "#fff3")
    role = u["role"]
    colors = {"superadmin":"rgba(240,165,0,.35)","admin":"rgba(255,255,255,.2)",
              "editor":"rgba(99,102,241,.3)","viewer":"rgba(255,255,255,.15)","guest":"rgba(255,255,255,.1)"}
    rbg = colors.get(role, "#fff3")
    labels = {"superadmin":"SUPER ADMIN","admin":"ADMIN","editor":"EDITOR","viewer":"VIEWER"}
    rlbl = labels.get(role, role.upper())
    btns = ""
    if role == "superadmin":
        btns += '<button class="tb-btn" onclick="openAdmin()">⚙ Admin</button>'
    btns += '<button class="tb-btn" onclick="changePw()">🔑</button>'
    btns += '<form action="/logout" method="post" style="display:inline"><button type="submit" class="tb-btn">⏻</button></form>'
    name = u["username"]
    return btns, name, rlbl, rbg

BASE_CSS = """
<style>
:root{--pr:#1a3a5c;--pl:#2563a8;--ac:#f0a500;--bg:#f0f4f8;--wh:#fff;--bd:#dde3ed;
  --tx:#1e2a3a;--mu:#6b7a94;--ok:#16a34a;--er:#ef4444;--wa:#f59e0b;--rd:6px}
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Segoe UI',Arial,sans-serif;background:var(--bg);color:var(--tx);font-size:13px}
#topbar{background:var(--pr);color:#fff;height:46px;display:flex;align-items:center;
  padding:0 14px;gap:8px;box-shadow:0 2px 8px rgba(0,0,0,.25);flex-shrink:0;position:relative;z-index:100}
#topbar .sp{flex:1}
.tb-btn{background:rgba(255,255,255,.15);border:1px solid rgba(255,255,255,.25);color:#fff;
  padding:5px 11px;border-radius:var(--rd);cursor:pointer;font-size:12px;font-family:inherit;
  text-decoration:none;display:inline-block;transition:background .15s}
.tb-btn:hover{background:rgba(255,255,255,.28)}
.tb-btn.glow{background:rgba(240,165,0,.3);border-color:rgba(240,165,0,.7);font-weight:700}
.overlay{position:fixed;inset:0;background:rgba(0,0,0,.55);z-index:1000;
  display:flex;align-items:center;justify-content:center;backdrop-filter:blur(3px)}
.overlay.hidden{display:none!important}
.modal{background:#fff;border-radius:10px;box-shadow:0 24px 64px rgba(0,0,0,.3);
  width:92%;max-width:600px;max-height:90vh;display:flex;flex-direction:column;
  animation:mIn .18s ease}
@keyframes mIn{from{transform:translateY(-14px);opacity:0}}
.mhdr{background:var(--pr);color:#fff;padding:12px 18px;font-weight:700;font-size:13px;
  display:flex;justify-content:space-between;align-items:center;flex-shrink:0;border-radius:10px 10px 0 0}
.mbody{padding:16px 18px;overflow-y:auto;flex:1}
.mfoot{padding:10px 18px;border-top:1px solid var(--bd);display:flex;justify-content:flex-end;
  gap:8px;background:var(--bg);flex-shrink:0;border-radius:0 0 10px 10px}
.xbtn{background:none;border:none;color:#fff;font-size:20px;cursor:pointer;opacity:.7;line-height:1}
.xbtn:hover{opacity:1}
.fgrid{display:grid;grid-template-columns:1fr 1fr;gap:12px}
.fg{display:flex;flex-direction:column;gap:4px}
.fg.full{grid-column:1/-1}
.fg label{font-size:10px;font-weight:700;color:var(--mu);text-transform:uppercase;letter-spacing:.4px}
.fg input,.fg select,.fg textarea{padding:7px 10px;border:1.5px solid var(--bd);
  border-radius:var(--rd);font-family:inherit;font-size:12px;outline:none;transition:border-color .2s}
.fg input:focus,.fg select:focus,.fg textarea:focus{border-color:var(--pl);box-shadow:0 0 0 2px rgba(37,99,168,.1)}
.btn{padding:7px 16px;border-radius:var(--rd);cursor:pointer;font-family:inherit;
  font-size:12px;font-weight:600;border:1px solid transparent;transition:all .15s}
.btn-pr{background:var(--pr);color:#fff}.btn-pr:hover{background:var(--pl)}
.btn-sc{background:var(--bg);color:var(--tx);border-color:var(--bd)}.btn-sc:hover{background:var(--bd)}
.btn-ok{background:var(--ok);color:#fff}
.btn-er{background:var(--er);color:#fff}
.btn-sm{padding:4px 10px;font-size:11px}
.stitle{font-size:11px;font-weight:700;color:var(--pr);text-transform:uppercase;
  letter-spacing:.5px;margin:14px 0 6px;padding-bottom:4px;border-bottom:2px solid var(--pr)}
.badge{display:inline-block;border-radius:10px;padding:2px 9px;font-size:10px;font-weight:700}
#toast{position:fixed;bottom:28px;right:18px;background:var(--pr);color:#fff;
  padding:10px 18px;border-radius:var(--rd);font-size:12px;z-index:9999;
  box-shadow:0 4px 16px rgba(0,0,0,.2);transform:translateY(80px);opacity:0;
  transition:all .3s;pointer-events:none;max-width:320px}
#toast.show{transform:none;opacity:1}
#toast.ok{background:#16a34a}#toast.er{background:#ef4444}#toast.wa{background:#f59e0b;color:#000}
</style>"""

SHARED_JS = """
<div id="toast"></div>
<script>
function toast(msg,type=''){
  const t=document.getElementById('toast');
  t.textContent=msg;t.className='show '+(type||'');
  clearTimeout(t._t);t._t=setTimeout(()=>t.className='',3200);
}
function openM(id){document.getElementById(id).classList.remove('hidden')}
function closeM(id){document.getElementById(id).classList.add('hidden')}
document.addEventListener('DOMContentLoaded',()=>{
  document.querySelectorAll('.overlay').forEach(o=>
    o.addEventListener('click',e=>{if(e.target===o)o.classList.add('hidden')}));
});
async function apiFetch(url,opts={}){
  const r=await fetch(url,{credentials:'include',headers:{'Content-Type':'application/json'},...opts});
  if(r.status===403){const d=await r.json().catch(()=>({}));
    if(d.error==='LOGIN_REQUIRED'){window.location='/login';return null;}
    throw new Error(d.error||'Forbidden');}
  if(!r.ok)throw new Error(await r.text());
  return r.json();
}
async function changePw(){
  const pw=prompt('New password (min 4 chars):');
  if(!pw||pw.length<4)return;
  const r=await apiFetch('/api/change_password',{method:'POST',body:JSON.stringify({password:pw})});
  if(r&&r.ok)toast('✔ Password changed','ok');else toast((r&&r.error)||'Error','er');
}
</script>"""


def render_login():
    return f"""<!DOCTYPE html><html lang="en"><head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>DCR — Login</title>
<style>
*{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:'Segoe UI',Arial,sans-serif;min-height:100vh;display:flex;align-items:center;
  justify-content:center;background:linear-gradient(135deg,#0f2640,#1a3a5c 60%,#2563a8)}}
.card{{background:#fff;border-radius:16px;box-shadow:0 24px 80px rgba(0,0,0,.4);width:100%;max-width:400px;overflow:hidden}}
.chdr{{background:linear-gradient(135deg,#1a3a5c,#2563a8);padding:32px;text-align:center}}
.chdr h1{{color:#fff;font-size:20px;margin-top:8px}}
.chdr p{{color:rgba(255,255,255,.6);font-size:12px;margin-top:4px}}
.cbody{{padding:28px 32px 32px}}
.fld{{margin-bottom:16px}}
.fld label{{display:block;font-size:11px;font-weight:700;color:#6b7a94;text-transform:uppercase;letter-spacing:.4px;margin-bottom:5px}}
.fld input{{width:100%;padding:11px 14px;border:1.5px solid #dde3ed;border-radius:8px;font-family:inherit;font-size:13px;outline:none;transition:border-color .2s}}
.fld input:focus{{border-color:#2563a8;box-shadow:0 0 0 3px rgba(37,99,168,.12)}}
.err{{background:#fef2f2;border:1px solid #fecaca;color:#dc2626;padding:9px 12px;border-radius:6px;font-size:12px;margin-bottom:14px;display:none}}
.btn-login{{width:100%;padding:13px;background:linear-gradient(135deg,#1a3a5c,#2563a8);color:#fff;border:none;
  border-radius:8px;font-family:inherit;font-size:14px;font-weight:700;cursor:pointer;transition:all .2s}}
.btn-login:hover{{transform:translateY(-1px);box-shadow:0 4px 16px rgba(26,58,92,.4)}}
.hint{{text-align:center;color:#9ca3af;font-size:11px;margin-top:16px}}
</style></head><body>
<div class="card">
  <div class="chdr"><div style="font-size:40px">📋</div>
    <h1>Document Control Register</h1><p>Sign in to continue</p></div>
  <div class="cbody">
    <div class="err" id="err"></div>
    <div class="fld"><label>Username</label><input id="un" type="text" autofocus autocomplete="username"></div>
    <div class="fld"><label>Password</label><input id="pw" type="password" autocomplete="current-password"></div>
    <button class="btn-login" onclick="login()">Sign In →</button>
    <p class="hint">Contact your administrator for credentials</p>
  </div>
</div>
<script>
document.getElementById('pw').onkeydown=e=>{{if(e.key==='Enter')login()}};
document.getElementById('un').onkeydown=e=>{{if(e.key==='Enter')document.getElementById('pw').focus()}};
async function login(){{
  const un=document.getElementById('un').value.trim();
  const pw=document.getElementById('pw').value;
  const err=document.getElementById('err');
  err.style.display='none';
  if(!un||!pw){{err.textContent='Please enter username and password';err.style.display='block';return;}}
  const r=await fetch('/login',{{method:'POST',credentials:'include',
    headers:{{'Content-Type':'application/json'}},body:JSON.stringify({{username:un,password:pw}})}});
  const d=await r.json();
  if(d.ok) window.location='/';
  else{{err.textContent=d.error||'Invalid credentials';err.style.display='block';}}
}}
</script></body></html>"""



def render_dashboard(u):
    btns, uname, rlbl, rbg = _user_info_html(u)
    role = u["role"] if u else "guest"

    return f"""<!DOCTYPE html><html lang="en"><head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>DCR \u2014 Dashboard</title>
{BASE_CSS}
<style>
body{{display:flex;flex-direction:column;min-height:100vh}}
.wrap{{max-width:1440px;margin:0 auto;padding:18px 14px;flex:1;width:100%}}
.kpi-grid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(130px,1fr));gap:10px;margin-bottom:16px}}
.kpi{{background:#fff;border-radius:8px;padding:12px 14px;box-shadow:0 1px 4px rgba(0,0,0,.07);border-left:4px solid var(--pr)}}
.kpi.ok{{border-left-color:var(--ok)}}.kpi.wa{{border-left-color:var(--wa)}}.kpi.er{{border-left-color:var(--er)}}
.kval{{font-size:26px;font-weight:800;color:var(--pr)}}
.kpi.ok .kval{{color:var(--ok)}}.kpi.wa .kval{{color:var(--wa)}}.kpi.er .kval{{color:var(--er)}}
.klbl{{font-size:10px;color:var(--mu);font-weight:700;text-transform:uppercase;letter-spacing:.4px;margin-top:2px}}
.pgrid{{display:grid;grid-template-columns:repeat(auto-fill,minmax(260px,1fr));gap:12px;margin-bottom:20px}}
.pcard{{background:#fff;border-radius:8px;box-shadow:0 2px 6px rgba(0,0,0,.08);overflow:hidden;
  text-decoration:none;color:inherit;display:block;transition:transform .15s,box-shadow .15s;position:relative}}
.pcard:hover{{transform:translateY(-2px);box-shadow:0 6px 18px rgba(0,0,0,.12)}}
.pchdr{{background:var(--pr);padding:10px 12px;display:flex;align-items:center;justify-content:space-between}}
.pcbody{{padding:10px 12px}}
.prow{{display:flex;justify-content:space-between;align-items:center;margin-bottom:4px}}
.prog{{height:4px;background:#eef1f7;border-radius:99px;overflow:hidden;margin-top:5px}}
.progf{{height:100%;border-radius:99px}}
.addcard{{background:#fff;border-radius:8px;border:2px dashed var(--bd);min-height:130px;
  display:flex;align-items:center;justify-content:center;flex-direction:column;gap:6px;
  cursor:pointer;transition:all .15s;color:var(--mu);font-size:12px}}
.addcard:hover{{border-color:var(--pr);color:var(--pr);background:#f7faff}}
.charts{{display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:16px}}
.ccard{{background:#fff;border-radius:8px;padding:12px;box-shadow:0 1px 4px rgba(0,0,0,.07)}}
.clbl{{font-size:10px;font-weight:700;color:var(--pr);text-transform:uppercase;letter-spacing:.4px;margin-bottom:8px}}
canvas{{max-height:190px}}
.urow{{display:flex;align-items:center;gap:8px;padding:6px 10px;background:var(--bg);border-radius:4px;font-size:12px;margin-bottom:4px}}
.pbtn{{padding:2px 8px;font-size:10px;border:1.5px solid var(--bd);background:#fff;border-radius:3px;cursor:pointer}}
.pbtn.on{{background:var(--pr);color:#fff;border-color:var(--pr)}}
.pbtn:hover:not(.on){{background:var(--bg)}}
#ld{{position:fixed;inset:0;background:rgba(15,38,64,.92);z-index:500;
  display:flex;align-items:center;justify-content:center;flex-direction:column;gap:14px}}
.spin{{width:40px;height:40px;border:4px solid rgba(255,255,255,.2);border-top-color:#f0a500;
  border-radius:50%;animation:spin .6s linear infinite}}
@keyframes spin{{to{{transform:rotate(360deg)}}}}
.psel-bar{{display:flex;align-items:center;gap:10px;background:#fff;padding:8px 12px;
  border-radius:8px;box-shadow:0 1px 4px rgba(0,0,0,.07);margin-bottom:14px;flex-wrap:wrap}}
.psel-bar label{{font-size:11px;font-weight:700;color:var(--mu);white-space:nowrap}}
.psel-bar select{{padding:5px 10px;border:1.5px solid var(--bd);border-radius:var(--rd);
  font-family:inherit;font-size:12px;outline:none}}
.dt-tbl{{width:100%;border-collapse:collapse;font-size:12px}}
.dt-tbl th{{background:var(--pr);color:#fff;padding:8px 10px;text-align:left;font-weight:600;white-space:nowrap}}
.dt-tbl td{{padding:6px 10px;border-bottom:1px solid #edf0f5}}
.dt-tbl tr:hover td{{background:#f0f4f8}}
.dt-tbl .alt td{{background:#fafbfd}}
.del-pbtn{{background:none;border:none;cursor:pointer;color:#ef4444;font-size:12px;padding:2px 5px;border-radius:3px;line-height:1}}
.del-pbtn:hover{{background:#fef2f2}}
@media(max-width:768px){{.charts{{grid-template-columns:1fr}}.pgrid{{grid-template-columns:1fr}}}}
</style>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
</head><body>

<div id="ld">
  <div class="spin"></div>
  <div style="color:rgba(255,255,255,.7);font-size:13px;font-weight:600">Loading dashboard...</div>
</div>

<div id="topbar">
  <span style="font-size:20px">📋</span>
  <span style="font-weight:700;font-size:14px">Document Control Register</span>
  <div class="sp"></div>
  {btns}
  <span style="color:rgba(255,255,255,.45);padding:0 4px">|</span>
  <span style="color:rgba(255,255,255,.8);font-size:11px">👤 {uname}
    <span style="background:{rbg};border-radius:3px;padding:1px 7px;font-size:9px;font-weight:700">{rlbl}</span>
  </span>
</div>

<div class="wrap">
  <div class="kpi-grid">
    <div class="kpi"><div class="kval" id="kpi-total">—</div><div class="klbl">Total Docs</div></div>
    <div class="kpi ok"><div class="kval" id="kpi-approved">—</div><div class="klbl">Approved</div></div>
    <div class="kpi wa"><div class="kval" id="kpi-pending">—</div><div class="klbl">Under Review</div></div>
    <div class="kpi er"><div class="kval" id="kpi-overdue">—</div><div class="klbl">Overdue</div></div>
    <div class="kpi" style="border-left-color:#7c3aed"><div class="kval" id="kpi-rejected" style="color:#7c3aed">—</div><div class="klbl">Rejected/Revise</div></div>
    <div class="kpi"><div class="kval" id="kpi-pct">—</div><div class="klbl">Completion %</div></div>
    <div class="kpi"><div class="kval" id="kpi-projs">—</div><div class="klbl">Projects</div></div>
  </div>

  <div class="psel-bar">
    <label>🔍 Project:</label>
    <select id="proj-sel" onchange="filterProject(this.value)">
      <option value="">All Projects</option>
    </select>
  </div>

  <div class="stitle">🗂 Projects</div>
  <div class="pgrid" id="pgrid"></div>

  <div class="charts">
    <div class="ccard"><div class="clbl">Documents by Project</div><canvas id="cProj"></canvas></div>
    <div class="ccard"><div class="clbl">Status Distribution</div><canvas id="cStatus"></canvas></div>
  </div>

  <div class="stitle">📋 Document Types Summary</div>
  <div style="background:#fff;border-radius:8px;box-shadow:0 1px 4px rgba(0,0,0,.07);overflow:hidden;margin-bottom:20px">
    <table class="dt-tbl">
      <thead><tr>
        <th>Project</th><th>Code</th><th>Type</th>
        <th style="text-align:center">Total</th>
        <th style="text-align:center">Approved</th>
        <th style="text-align:center">Pending</th>
        <th style="text-align:center">Rejected</th>
        <th style="text-align:center">Overdue</th>
        <th style="text-align:center">Open</th>
      </tr></thead>
      <tbody id="dt-tbody"></tbody>
    </table>
    <div id="dt-empty" style="text-align:center;padding:24px;color:var(--mu);display:none">No data</div>
  </div>

  <div class="stitle" style="margin-top:20px">🏗 Discipline Breakdown by Document Type</div>
  <div style="background:#fff;border-radius:8px;box-shadow:0 1px 4px rgba(0,0,0,.07);overflow:hidden;margin-bottom:24px">
    <table class="dt-tbl">
      <thead><tr>
        <th>Project</th><th>Doc Type</th><th>Discipline</th>
        <th style="text-align:center">Total</th>
        <th style="text-align:center">Approved</th>
        <th style="text-align:center">Pending</th>
        <th style="text-align:center;color:#c4b5fd">Rejected</th>
        <th style="text-align:center">Overdue</th>
      </tr></thead>
      <tbody id="disc-tbody"></tbody>
    </table>
    <div id="disc-empty" style="text-align:center;padding:24px;color:var(--mu);display:none">No data</div>
  </div>
</div>

<div class="overlay hidden" id="admin-modal">
  <div class="modal" style="max-width:780px">
    <div class="mhdr"><span>⚙ Admin Panel</span><button class="xbtn" onclick="closeM('admin-modal')">✕</button></div>
    <div class="mbody" id="admin-body"></div>
    <div class="mfoot"><button class="btn btn-sc" onclick="closeM('admin-modal')">Close</button></div>
  </div>
</div>

<div class="overlay hidden" id="newproj-modal">
  <div class="modal" style="max-width:420px">
    <div class="mhdr"><span>➕ New Project</span><button class="xbtn" onclick="closeM('newproj-modal')">✕</button></div>
    <div class="mbody">
      <div class="fgrid">
        <div class="fg"><label>Project ID</label><input id="np-id" placeholder="e.g. 24CP01"></div>
        <div class="fg"><label>Short Code</label><input id="np-code" placeholder="e.g. CP01"></div>
        <div class="fg full"><label>Project Name</label><input id="np-name" placeholder="Full project name"></div>
      </div>
    </div>
    <div class="mfoot">
      <button class="btn btn-sc" onclick="closeM('newproj-modal')">Cancel</button>
      <button class="btn btn-pr" id="cpbtn" onclick="createProject()">Create</button>
    </div>
  </div>
</div>

{SHARED_JS}
<script>
const ROLE='{role}';
let STATS=[]; let pChart,sChart;

async function init(){{
  try{{
    STATS=await apiFetch('/api/dashboard_stats');
    if(!STATS)return;
    document.getElementById('proj-sel').innerHTML='<option value="">All Projects</option>'+
      STATS.map(p=>`<option value="${{p.id}}">${{p.name}} (${{p.code}})</option>`).join('');
    renderAll('');
  }}finally{{document.getElementById('ld').style.display='none';}}
}}

function filterProject(pid){{renderAll(pid);}}

function getFiltered(pid){{return pid?STATS.filter(s=>s.id===pid):STATS;}}

function renderAll(pid){{
  const d=getFiltered(pid);
  updateKPIs(d);renderCards(d,pid);renderCharts(d);renderDTTable(d);renderDiscTable(d);
}}

function renderDiscTable(data){{
  const tbody=document.getElementById('disc-tbody');if(!tbody)return;
  const empty=document.getElementById('disc-empty');
  tbody.innerHTML='';let rows=0;
  data.forEach(p=>{{
    (p.dt_stats||[]).forEach(dt=>{{
      const disc=dt.disc_breakdown||[];if(!disc.length)return;
      disc.forEach((ds,di)=>{{
        rows++;
        const tr=document.createElement('tr');tr.className=rows%2===0?'alt':'';
        if(di===0){{
          const tdc=document.createElement('td');tdc.rowSpan=disc.length;tdc.style.cssText='font-size:11px;color:var(--mu);vertical-align:top;padding:6px 10px';tdc.textContent=p.code;tr.appendChild(tdc);
          const tdt=document.createElement('td');tdt.rowSpan=disc.length;tdt.style.cssText='font-weight:600;vertical-align:top;padding:6px 10px';tdt.textContent=dt.code+' '+dt.name;tr.appendChild(tdt);
        }}
        const mk=(v,c)=>{{const td=document.createElement('td');td.style.cssText='text-align:center;padding:6px 8px;font-weight:600;color:'+c;td.textContent=v||'0';return td;}};
        const tdisc=document.createElement('td');tdisc.style.cssText='padding:6px 10px';tdisc.textContent=ds.disc;tr.appendChild(tdisc);
        tr.appendChild(mk(ds.total,'#1a3a5c'));tr.appendChild(mk(ds.approved,'#16a34a'));
        tr.appendChild(mk(ds.pending,'#f59e0b'));tr.appendChild(mk(ds.rejected,'#7c3aed'));tr.appendChild(mk(ds.overdue,'#ef4444'));
        tbody.appendChild(tr);
      }});
    }});
  }});
  if(empty)empty.style.display=rows?'none':'block';
}}

function updateKPIs(d){{
  const t=d.reduce((s,p)=>s+p.total,0),ap=d.reduce((s,p)=>s+p.approved,0),
        pe=d.reduce((s,p)=>s+p.pending,0),ov=d.reduce((s,p)=>s+p.overdue,0);
  document.getElementById('kpi-total').textContent=t;
  document.getElementById('kpi-approved').textContent=ap;
  document.getElementById('kpi-pending').textContent=pe;
  document.getElementById('kpi-overdue').textContent=ov;
  document.getElementById('kpi-pct').textContent=(t?Math.round(ap/t*100):0)+'%';
  document.getElementById('kpi-projs').textContent=d.length;
}}

function renderCards(d,pid){{
  const g=document.getElementById('pgrid');g.innerHTML='';
  d.forEach(p=>{{
    const col=p.pct>=80?'#16a34a':p.pct>=50?'#f59e0b':'#ef4444';
    const a=document.createElement('a');a.href='/app?p='+p.id;a.className='pcard';
    a.innerHTML=`<div class="pchdr">
      <div><div style="color:rgba(255,255,255,.6);font-size:10px;font-weight:700">${{p.code}}</div>
        <div style="color:#fff;font-weight:700;font-size:13px">${{p.name}}</div></div>
      ${{p.overdue>0?`<span style="background:#ef4444;color:#fff;border-radius:10px;padding:1px 7px;font-size:10px;font-weight:700">⚠${{p.overdue}}</span>`:''}}
    </div>
    <div class="pcbody">
      ${{p.client?`<div style="font-size:10px;color:var(--mu);margin-bottom:5px">👤 ${{p.client}}</div>`:''}}
      <div class="prow"><span style="font-size:11px;color:var(--mu)">Total</span><b>${{p.total}}</b></div>
      <div class="prow"><span style="font-size:11px;color:var(--mu)">Approved</span><b style="color:#16a34a">${{p.approved}}</b></div>
      <div class="prow"><span style="font-size:11px;color:var(--mu)">Pending</span><b style="color:#f59e0b">${{p.pending}}</b></div>
      <div class="prog"><div class="progf" style="width:${{p.pct}}%;background:${{col}}"></div></div>
      <div style="font-size:10px;color:${{col}};font-weight:700;margin-top:3px">${{p.pct}}%</div>
      ${{p.can_edit?'<div style="margin-top:3px;font-size:10px;color:#2563a8">✏ Can edit</div>':''}}
      ${{ROLE==='superadmin'?`<button class="del-pbtn" onclick="event.preventDefault();delProject('${{p.id}}','${{p.name}}')">🗑 Delete</button>`:''}}
    </div>`;
    g.appendChild(a);
  }});
  if(ROLE==='superadmin'&&!pid){{
    const add=document.createElement('div');add.className='addcard';
    add.innerHTML='<span style="font-size:28px">➕</span><span>New Project</span>';
    add.onclick=()=>openM('newproj-modal');g.appendChild(add);
  }}
}}

function renderCharts(d){{
  if(pChart)pChart.destroy();
  pChart=new Chart(document.getElementById('cProj'),{{type:'bar',
    data:{{labels:d.map(s=>s.code),datasets:[
      {{label:'Approved',data:d.map(s=>s.approved),backgroundColor:'#16a34a'}},
      {{label:'Pending', data:d.map(s=>s.pending), backgroundColor:'#f59e0b'}},
      {{label:'Overdue', data:d.map(s=>s.overdue), backgroundColor:'#ef4444'}}]}},
    options:{{responsive:true,plugins:{{legend:{{position:'bottom'}}}},scales:{{y:{{beginAtZero:true}}}}}}}});
  const t=d.reduce((s,p)=>s+p.total,0),ap=d.reduce((s,p)=>s+p.approved,0),
        pe=d.reduce((s,p)=>s+p.pending,0),ov=d.reduce((s,p)=>s+p.overdue,0);
  if(sChart)sChart.destroy();
  sChart=new Chart(document.getElementById('cStatus'),{{type:'doughnut',
    data:{{labels:['Approved','Pending','Overdue','Other'],
      datasets:[{{data:[ap,pe,ov,Math.max(0,t-ap-pe-ov)],
        backgroundColor:['#16a34a','#f59e0b','#ef4444','#9ca3af'],borderWidth:2,borderColor:'#fff'}}]}},
    options:{{responsive:true,plugins:{{legend:{{position:'bottom'}}}}}}}});
}}

function renderDTTable(d){{
  const tbody=document.getElementById('dt-tbody'),empty=document.getElementById('dt-empty');
  tbody.innerHTML='';
  let rows=[],i=0;
  d.forEach(p=>{{(p.dt_stats||[]).forEach(dt=>{{rows.push({{...dt,pid:p.id,pcode:p.code}});}});}});
  if(!rows.length){{empty.style.display='block';return;}}
  empty.style.display='none';
  rows.forEach(r=>{{
    const pct=r.total?Math.round(r.approved/r.total*100):0;
    const col=pct>=80?'#16a34a':pct>=50?'#f59e0b':'#ef4444';
    const tr=document.createElement('tr');if(i%2)tr.className='alt';i++;
    tr.innerHTML=`<td style="font-size:10px;color:var(--mu)">${{r.pcode}}</td>
      <td style="font-weight:700;color:var(--pr)">${{r.dtCode||r.code||r.id}}</td>
      <td>${{r.dtName||r.name}}</td>
      <td style="text-align:center;font-weight:700">${{r.total}}</td>
      <td style="text-align:center;color:#16a34a;font-weight:600">${{r.approved}}</td>
      <td style="text-align:center;color:#f59e0b;font-weight:600">${{r.pending}}</td>
      <td style="text-align:center;color:#7c3aed;font-weight:600">${{r.rejected||0}}</td>
      <td style="text-align:center;color:#ef4444;font-weight:600">${{r.overdue}}</td>
      <td style="text-align:center"><a href="/app?p=${{r.pid}}&tab=${{r.id}}"
        style="padding:2px 9px;background:var(--pr);color:#fff;border-radius:3px;text-decoration:none;font-size:10px;font-weight:600">Open</a></td>`;
    tbody.appendChild(tr);
  }});
}}

async function openAdmin(){{
  const [users,projects]=await Promise.all([apiFetch('/api/users'),apiFetch('/api/projects')]);
  if(!users||!projects)return;
  const body=document.getElementById('admin-body');body.innerHTML='';
  const ut=document.createElement('div');ut.className='stitle';ut.textContent='👥 Users';body.appendChild(ut);
  for(const u of users){{
    const row=document.createElement('div');row.className='urow';
    row.innerHTML=`<span style="flex:1;font-weight:600">👤 ${{u.username}}</span>
      <span class="badge" style="background:#fef3c7;color:#92400e">${{u.role.toUpperCase()}}</span>
      ${{u.username!=='admin'?`<button class="btn btn-sc btn-sm" onclick="changeUserPw('${{u.username}}')">🔑</button>
        <button class="btn btn-er btn-sm" onclick="delUser('${{u.username}}')">✕</button>`:
        '<span style="font-size:10px;color:var(--mu)">(protected)</span>'}}`;
    body.appendChild(row);
    if(u.role!=='superadmin'){{
      const ad=document.createElement('div');
      ad.style.cssText='padding:4px 10px 10px 32px;border-bottom:1px solid var(--bd);margin-bottom:4px';
      ad.innerHTML='<div style="font-size:10px;color:var(--mu);margin-bottom:4px">Project access:</div>';
      const assigned=await apiFetch('/api/users/'+u.username+'/projects').catch(()=>[]);
      const pl=document.createElement('div');pl.style.cssText='display:flex;flex-wrap:wrap;gap:4px';
      projects.forEach(p=>{{
        const btn=document.createElement('button');btn.className='pbtn'+(assigned.includes(p.id)?' on':'');
        btn.textContent=p.code;btn.title=p.name;
        btn.onclick=async()=>{{const on=btn.classList.contains('on');
          await apiFetch('/api/users',{{method:'POST',body:JSON.stringify({{action:on?'unassign':'assign',username:u.username,project_id:p.id}})}});
          btn.classList.toggle('on');}};
        pl.appendChild(btn);}});
      ad.appendChild(pl);body.appendChild(ad);
    }}
  }}
  const at=document.createElement('div');at.className='stitle';at.textContent='➕ Add User';body.appendChild(at);
  body.innerHTML+=`<div style="display:grid;grid-template-columns:1fr 1fr 1fr auto;gap:8px;align-items:end">
    <div class="fg"><label>Username</label><input id="nu-name"></div>
    <div class="fg"><label>Role</label><select id="nu-role">
      <option value="editor">Editor</option><option value="viewer">Viewer</option>
      <option value="superadmin">Super Admin</option></select></div>
    <div class="fg"><label>Password</label><input id="nu-pw" type="password"></div>
    <button class="btn btn-pr btn-sm" style="margin-bottom:1px" onclick="addUser()">Add</button></div>`;
  openM('admin-modal');
}}
async function addUser(){{
  const n=document.getElementById('nu-name').value.trim().toLowerCase(),
        role=document.getElementById('nu-role').value,pw=document.getElementById('nu-pw').value;
  if(!n||!pw){{toast('Required','er');return;}}
  const r=await apiFetch('/api/users',{{method:'POST',body:JSON.stringify({{action:'add',username:n,role,password:pw}})}});
  if(r&&r.ok){{toast('✔ Added','ok');closeM('admin-modal');openAdmin();}}else toast((r&&r.error)||'Error','er');
}}
async function delUser(u){{if(!confirm('Delete '+u+'?'))return;
  const r=await apiFetch('/api/users',{{method:'POST',body:JSON.stringify({{action:'delete',username:u}})}});
  if(r&&r.ok){{toast('Deleted','wa');closeM('admin-modal');openAdmin();}}
}}
async function changeUserPw(u){{const pw=prompt('New pw for '+u+':');if(!pw)return;
  await apiFetch('/api/users',{{method:'POST',body:JSON.stringify({{action:'change_password',username:u,password:pw}})}});
  toast('✔ Changed','ok');
}}
async function delProject(id,name){{
  if(!confirm('Delete project "'+name+'" and ALL data? Cannot be undone!'))return;
  const r=await apiFetch('/api/projects/delete/'+id,{{method:'POST'}});
  if(r&&r.ok){{
    toast('Deleted','wa');
    STATS=STATS.filter(s=>s.id!==id);
    const sel=document.getElementById('proj-sel');
    [...sel.options].forEach(o=>{{if(o.value===id)o.remove();}});
    renderAll('');
  }}else toast('Error','er');
}}
async function createProject(){{
  const id=document.getElementById('np-id').value.trim().toUpperCase();
  const code=document.getElementById('np-code').value.trim().toUpperCase();
  const name=document.getElementById('np-name').value.trim();
  if(!id||!name||!code){{toast('All fields required','er');return;}}
  const btn=document.getElementById('cpbtn');
  btn.disabled=true;btn.textContent='⏳';
  let ok=false;
  try{{
    const resp=await fetch('/api/projects/create',{{
      method:'POST',credentials:'include',
      headers:{{'Content-Type':'application/json','X-Requested-With':'XMLHttpRequest'}},
      body:JSON.stringify({{id,name,code}})
    }});
    let r={{}};
    try{{r=await resp.json();}}catch(e){{}}
    if(resp.ok&&r.ok){{
      ok=true;
      const np={{id,name,code,client:'',total:0,approved:0,pending:0,overdue:0,pct:0,can_edit:true,dt_stats:[]}};
      STATS.push(np);
      const sel=document.getElementById('proj-sel');
      const o=document.createElement('option');o.value=id;o.textContent=name+' ('+code+')';sel.appendChild(o);
      ['np-id','np-code','np-name'].forEach(i=>{{const el=document.getElementById(i);if(el)el.value='';}});
      closeM('newproj-modal');
      renderAll('');
      toast('✔ Project created: '+name,'ok');
    }}else{{
      toast((r&&r.error)||('HTTP '+resp.status),'er');
    }}
  }}catch(e){{
    toast('Error: '+e.message,'er');
  }}
  btn.disabled=false;btn.textContent='Create';
}}

init();
</script></body></html>"""

def render_register(u, proj):
    from utils import STATUS_COLORS
    pid      = proj["id"]
    uname    = u["username"] if u else "guest"
    role     = u["role"]     if u else "guest"
    editable = can_edit(pid)
    btns, _, rlbl, rbg = _user_info_html(u)
    sc_json  = json.dumps(STATUS_COLORS)

    dts       = db.get_doc_types(pid)
    logo_l    = db.get_logo(pid,"logo_left")
    logo_r    = db.get_logo(pid,"logo_right")
    logo_html = ""
    if logo_l: logo_html += f'<img src="/api/logo/{pid}/logo_left" style="height:52px;object-fit:contain;margin-right:10px">'
    if logo_r: logo_html += f'<img src="/api/logo/{pid}/logo_right" style="height:52px;object-fit:contain;margin-left:auto;margin-right:10px">'

    PROJ_FIELDS = [("code","Code"),("name","Project Name"),("startDate","Start"),("endDate","End"),
                   ("client","Client"),("landlord","Landlord"),("pmo","PMO"),
                   ("mainConsultant","Consultant"),("mepConsultant","MEP"),("contractor","Contractor")]

    projbar = "".join(
        f'<div class="pf"><span class="pf-lbl">{lbl}</span>'
        f'<span class="pf-val" data-key="{key}">{proj.get(key,"—")}</span></div>'
        for key,lbl in PROJ_FIELDS)

    tabs_html = "".join(
        f'<button class="tab-btn" data-id="{dt["id"]}" onclick="switchTab(\'{dt["id"]}\')">'
        f'<span>{dt["code"]}</span><span class="tcnt" id="cnt-{dt["id"]}">0</span></button>'
        for dt in dts)

    _hol_btn = " <button class='tool-btn purple' onclick='openSettings()'>🗓 Holidays</button>" if role=='superadmin' else ''
    edit_btns = (f'<button class="tool-btn" onclick="addRecord()">➕ Add</button>'
                 f'<button class="tool-btn purple" onclick="manageColumns()">⚙ Columns</button>'
                 f'{_hol_btn}'
                 f'<button class="tool-btn" onclick="openLists()">📋 Lists</button>'
                 f'<button class="tool-btn" onclick="editProject()">🏗 Project</button>'
                 if editable else
                 '<span style="font-size:11px;color:rgba(255,255,255,.5);padding:4px 8px">'
                 '👁 Read-only — <a href="/login" style="color:#f0a500">login to edit</a></span>')

    return f"""<!DOCTYPE html><html lang="en"><head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>DCR — {proj.get("name","Register")}</title>
{BASE_CSS}
<style>
.hidden{{display:none!important}}
body{{height:100vh;display:flex;flex-direction:column;overflow:hidden}}
@media print{{
  #topbar,#tabbar,.toolrow,.bulkbar,#statusbar,.acts{{display:none!important}}
  body{{height:auto;overflow:visible}}
  #main{{overflow:visible;height:auto}}
  #tblwrap{{overflow:visible;height:auto}}
  #regtbl{{font-size:10px}}
  #regtbl th,#regtbl td{{padding:4px 6px!important;white-space:normal!important}}
  @page{{size:A4 landscape;margin:10mm}}
}}
#projbar{{background:#fff;border-bottom:2px solid var(--pr);padding:3px 12px;
  display:flex;align-items:center;overflow-x:auto;flex-shrink:0;gap:0}}
.pf{{display:flex;flex-direction:column;padding:0 10px;border-right:1px solid var(--bd)}}
.pf:last-of-type{{border-right:none}}
.pf-lbl{{font-size:9px;font-weight:700;color:var(--pr);text-transform:uppercase;letter-spacing:.4px}}
.pf-val{{font-size:11px;white-space:nowrap}}
#tabsbar{{background:#0f2640;display:flex;align-items:center;overflow-x:auto;flex-shrink:0;
  padding:0 8px;scrollbar-width:thin}}
#tabsbar::-webkit-scrollbar{{height:3px}}
#tabsbar::-webkit-scrollbar-thumb{{background:rgba(255,255,255,.3)}}
.tab-btn{{display:flex;align-items:center;gap:5px;padding:9px 12px;background:transparent;
  border:none;border-bottom:3px solid transparent;color:rgba(255,255,255,.55);cursor:pointer;
  font-family:inherit;font-size:11px;font-weight:600;white-space:nowrap;transition:all .15s}}
.tab-btn:hover{{color:#fff;background:rgba(255,255,255,.07)}}
.tab-btn.active{{color:#fff;border-bottom-color:var(--ac)}}
.tcnt{{background:rgba(255,255,255,.2);border-radius:10px;padding:1px 7px;font-size:10px;font-weight:700}}
.tab-btn.active .tcnt{{background:var(--ac);color:#000}}
.tab-add{{padding:5px 10px;background:rgba(255,255,255,.1);border:1px dashed rgba(255,255,255,.35);
  color:rgba(255,255,255,.7);border-radius:4px;cursor:pointer;font-size:16px;margin-left:6px}}
.tab-add:hover{{background:rgba(255,255,255,.2)}}
#toolbar{{background:#fff;border-bottom:1px solid var(--bd);padding:5px 10px;
  display:flex;align-items:center;gap:5px;flex-shrink:0;flex-wrap:wrap}}
.tool-btn{{display:flex;align-items:center;gap:4px;padding:5px 10px;background:var(--bg);
  border:1px solid var(--bd);border-radius:var(--rd);cursor:pointer;font-size:11px;
  font-family:inherit;color:var(--tx);transition:all .15s;white-space:nowrap}}
.tool-btn:hover{{background:var(--pr);color:#fff;border-color:var(--pr)}}
.tool-btn.purple:hover{{background:#7c3aed;border-color:#7c3aed}}
.tool-btn.teal:hover{{background:#0891b2;border-color:#0891b2}}
.tool-dd{{position:relative}}
.tool-dd-menu{{position:absolute;top:calc(100% + 4px);left:0;background:#fff;border:1.5px solid var(--bd);border-radius:6px;box-shadow:0 8px 24px rgba(0,0,0,.15);z-index:300;min-width:210px;overflow:hidden}}
.tool-dd-menu button{{display:block;width:100%;text-align:left;padding:9px 14px;border:none;background:none;cursor:pointer;font-size:12px;font-family:inherit;color:#1e2a3a;white-space:nowrap}}
.tool-dd-menu button:hover{{background:#f0f4f8;color:var(--pr)}}
#srchbox{{flex:1;min-width:150px;max-width:260px;padding:5px 10px 5px 28px;border:1px solid var(--bd);
  border-radius:var(--rd);font-family:inherit;font-size:12px;outline:none;
  background:#fff url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='13' height='13' viewBox='0 0 24 24' fill='none' stroke='%236b7a94' stroke-width='2'%3E%3Ccircle cx='11' cy='11' r='8'/%3E%3Cpath d='m21 21-4.35-4.35'/%3E%3C/svg%3E") no-repeat 7px center}}
#srchbox:focus{{border-color:var(--pl);box-shadow:0 0 0 2px rgba(37,99,168,.1)}}
#main{{flex:1;overflow:hidden;display:flex;flex-direction:column}}
#tblwrap{{flex:1;overflow:auto}}
table{{width:100%;border-collapse:collapse;min-width:900px;font-size:12px}}
thead{{position:sticky;top:0;z-index:10}}
th{{background:var(--pr);color:#fff;padding:8px;text-align:left;font-weight:600;
  white-space:nowrap;border-right:1px solid rgba(255,255,255,.1);cursor:pointer;user-select:none;position:relative}}
th:hover{{background:var(--pl)}}
.frow th{{background:#eef1f7;padding:2px 4px;cursor:default;position:sticky;top:33px;z-index:9}}
.frow th:hover{{background:#eef1f7}}
.frow input,.frow select{{width:100%;padding:3px 6px;border:1px solid var(--bd);
  border-radius:3px;font-size:10px;font-family:inherit;background:#fff;outline:none}}
td{{padding:5px 8px;border-bottom:1px solid #edf0f5;border-right:1px solid #f3f4f6;
  vertical-align:middle;max-width:220px;word-break:break-word}}
tr:hover td{{background:rgba(37,99,168,.10);transition:background .1s}}
tr.ov td{{background:#fff5f5}}
tr.rv td{{color:var(--mu)}}
tr.alt td{{background:#fafbfd}}
.sr{{text-align:center;color:var(--mu);font-size:10px;width:32px}}
.chkcell{{text-align:center;width:28px;padding:4px!important}}
.chkcell input{{width:14px;height:14px;cursor:pointer;accent-color:var(--pr)}}
.acts{{white-space:nowrap;width:64px}}
.act{{padding:2px 7px;border:1px solid var(--bd);background:#fff;border-radius:3px;cursor:pointer;font-size:11px}}
.act:hover{{background:var(--pr);color:#fff;border-color:var(--pr)}}
.act.del:hover{{background:var(--er);border-color:var(--er)}}
.sbadge{{display:inline-block;border-radius:10px;padding:2px 9px;font-size:10px;font-weight:700}}
.flink{{color:var(--pl);text-decoration:underline;cursor:pointer;font-size:11px}}
.ovdate{{color:#dc2626;font-weight:700}}
#bulkbar{{display:none;background:#1a3a5c;color:#fff;padding:5px 14px;align-items:center;gap:10px;font-size:12px;flex-shrink:0}}
#bulkbar.show{{display:flex}}
#sbar{{background:var(--pr);color:rgba(255,255,255,.75);padding:3px 14px;font-size:10px;display:flex;gap:16px;flex-shrink:0}}
.rz{{position:absolute;right:0;top:0;bottom:0;width:6px;cursor:col-resize;z-index:1}}
.rz:hover,.rz.rzg{{background:var(--ac)}}
.ms-con{{border:1px solid var(--bd);border-radius:var(--rd);min-height:34px;padding:3px;
  background:#fff;cursor:pointer;position:relative}}
.ms-tag{{display:inline-flex;align-items:center;gap:3px;background:var(--pr);color:#fff;
  border-radius:3px;padding:2px 7px;font-size:10px;margin:2px}}
.ms-rm{{cursor:pointer;opacity:.7}}
.ms-ph{{color:var(--mu);font-size:11px;padding:3px 5px}}
.ms-dd{{position:absolute;left:0;right:0;top:100%;background:#fff;border:1px solid var(--bd);
  border-radius:var(--rd);z-index:200;max-height:200px;overflow-y:auto;
  box-shadow:0 4px 16px rgba(0,0,0,.12);margin-top:2px}}
.ms-opt{{padding:6px 10px;cursor:pointer;font-size:11px;display:flex;align-items:center;gap:6px}}
.ms-opt:hover{{background:var(--bg)}}
.ms-opt.sel{{background:#eff6ff}}
.empty{{text-align:center;padding:60px 20px;color:var(--mu)}}
.slist{{list-style:none;display:flex;flex-direction:column;gap:3px;
  max-height:190px;overflow-y:auto;border:1px solid var(--bd);border-radius:var(--rd);padding:4px}}
.sitem{{display:flex;align-items:center;gap:8px;padding:4px 8px;background:var(--bg);
  border-radius:3px;font-size:11px}}
.sitem .nm{{flex:1}}
.sitem button{{padding:2px 8px;font-size:10px;border:1px solid var(--bd);background:#fff;
  border-radius:3px;cursor:pointer}}
.sitem button:hover{{background:var(--er);color:#fff;border-color:var(--er)}}
.addrow{{display:flex;gap:6px;margin-top:6px}}
.addrow input{{flex:1;padding:5px 8px;border:1px solid var(--bd);border-radius:3px;
  font-size:11px;font-family:inherit;outline:none}}
@media(max-width:768px){{
  .pf{{padding:0 6px;min-width:70px}}
  .tab-btn{{padding:7px 8px;font-size:10px}}
  .tool-btn{{padding:4px 7px;font-size:10px}}
  table{{font-size:11px}} td,th{{padding:4px 5px}}
  .modal{{width:97%;max-height:96vh}}
  .fgrid{{grid-template-columns:1fr}}
}}
</style></head><body>

<div id="topbar">
  <span style="font-size:20px">📋</span>
  <span style="font-weight:700;font-size:14px;letter-spacing:.3px">Document Control Register</span>
  <div class="sp"></div>
  <a href="/" class="tb-btn">📊 Dashboard</a>
  {btns}
  <span style="color:rgba(255,255,255,.45);padding:0 4px">|</span>
  <span style="color:rgba(255,255,255,.8);font-size:11px">👤 {uname}
    <span style="background:{rbg};border-radius:3px;padding:1px 7px;font-size:9px;font-weight:700">{rlbl}</span>
  </span>
</div>

<div id="projbar">
  {logo_html}
  {projbar}
  {'<button onclick="editProject()" style="margin-left:auto;background:var(--pr);color:#fff;border:none;padding:5px 12px;border-radius:var(--rd);cursor:pointer;font-size:11px;font-family:inherit;flex-shrink:0">✏ Edit</button>' if editable else ''}
</div>

<div id="tabsbar">
  {tabs_html}
  {'<button class="tab-add" onclick="addDocType()" title="Add Type">＋</button>' if editable else ''}
</div>

<div id="toolbar">
  {edit_btns}
  <div class="tool-dd" id="exp-dd">
    <button class="tool-btn teal" onclick="toggleExpDD(event)">📥 Export ▾</button>
    <div class="tool-dd-menu hidden" id="exp-menu">
      <button onclick="doExport();closeExpDD()">📊 Excel — This Tab</button>
      <button onclick="doExportAll();closeExpDD()">📊 Excel — All Tabs</button>
      <button onclick="doExportPDF();closeExpDD()">📄 PDF — This Tab</button>
      <button onclick="doExportAllPDF();closeExpDD()">📄 PDF — All Tabs</button>
    </div>
  </div>
  <button class="tool-btn teal" onclick="doPrint()">🖨 Print</button>
  {'<button class="tool-btn teal" onclick="openImport()">📤 Import</button>' if editable else ''}
  <input type="text" id="srchbox" placeholder="Search..." oninput="doSearch()">
</div>

<div id="bulkbar">
  <span id="bulkcnt">0 selected</span>
  <button class="btn btn-er btn-sm" onclick="bulkDel()">🗑 Delete</button>
  <button class="btn btn-sm" style="background:rgba(255,255,255,.15);color:#fff;border-color:rgba(255,255,255,.3)" onclick="clearSel()">✕ Cancel</button>
</div>

<div id="main">
  <div id="tblwrap">
    <table id="regtbl"><thead id="thead"></thead><tbody id="tbody"></tbody></table>
    <div class="empty hidden" id="empty" style="display:none"><div style="font-size:48px;margin-bottom:10px">📁</div><p style="color:var(--mu)">No records yet — click ➕ Add to create one</p></div>
  </div>
</div>

<div id="sbar">
  <span id="s-total">Total: 0</span>
  <span id="s-show">Showing: 0</span>
  <span id="s-ov">Overdue: 0</span>
  <span style="margin-left:auto" id="s-clock"></span>
</div>

<!-- ADD/EDIT RECORD -->
<div class="overlay hidden" id="rec-modal">
  <div class="modal" style="max-width:820px">
    <div class="mhdr"><span id="rec-title">Add Document</span>
      <button class="xbtn" onclick="closeM('rec-modal')">✕</button></div>
    <div class="mbody"><div class="fgrid" id="rec-form"></div></div>
    <div class="mfoot">
      <button class="btn btn-sc" onclick="closeM('rec-modal')">Cancel</button>
      <button class="btn btn-pr" onclick="saveRecord()">Save</button>
    </div>
  </div>
</div>

<!-- PROJECT MODAL -->
<div class="overlay hidden" id="proj-modal">
  <div class="modal" style="max-width:820px">
    <div class="mhdr"><span>🏗 Project Information</span>
      <button class="xbtn" onclick="closeM('proj-modal')">✕</button></div>
    <div class="mbody" id="proj-body"></div>
    <div class="mfoot">
      <button class="btn btn-sc" onclick="closeM('proj-modal')">Cancel</button>
      <button class="btn btn-pr" onclick="saveProject()">💾 Save Project</button>
    </div>
  </div>
</div>

<!-- LISTS -->
<div class="overlay hidden" id="lists-modal">
  <div class="modal">
    <div class="mhdr"><span>📋 Dropdown Lists</span>
      <button class="xbtn" onclick="closeM('lists-modal')">✕</button></div>
    <div class="mbody" id="lists-body"></div>
    <div class="mfoot"><button class="btn btn-sc" onclick="closeM('lists-modal')">Close</button></div>
  </div>
</div>

<!-- ADD DOC TYPE -->
<div class="overlay hidden" id="dt-modal">
  <div class="modal" style="max-width:420px">
    <div class="mhdr"><span>Add Document Type</span>
      <button class="xbtn" onclick="closeM('dt-modal')">✕</button></div>
    <div class="mbody">
      <div class="fgrid">
        <div class="fg"><label>Code</label><input id="dt-code" placeholder="e.g. MS"></div>
        <div class="fg"><label>Name</label><input id="dt-name" placeholder="e.g. Method Statement"></div>
      </div>
    </div>
    <div class="mfoot">
      <button class="btn btn-sc" onclick="closeM('dt-modal')">Cancel</button>
      <button class="btn btn-pr" onclick="saveDocType()">Add</button>
    </div>
  </div>
</div>

<!-- COLUMNS -->
<div class="overlay hidden" id="col-modal">
  <div class="modal">
    <div class="mhdr"><span>⚙ Manage Columns</span>
      <button class="xbtn" onclick="closeM('col-modal')">✕</button></div>
    <div class="mbody" id="col-body"></div>
    <div class="mfoot">
      <button class="btn btn-sc" onclick="openAddCol()">+ Add Column</button>
      <button class="btn btn-pr" onclick="closeM('col-modal');loadRecords()">Done</button>
    </div>
  </div>
</div>

<!-- ADD COLUMN -->
<div class="overlay hidden" id="addcol-modal">
  <div class="modal" style="max-width:460px">
    <div class="mhdr"><span>Add Column</span>
      <button class="xbtn" onclick="closeM('addcol-modal')">✕</button></div>
    <div class="mbody">
      <div class="fgrid">
        <div class="fg full"><label>Column Name</label><input id="col-name" placeholder="e.g. Review Duration"></div>
        <div class="fg full"><label>Type</label>
          <select id="col-type" onchange="onColType(this.value)">
            <option value="text">📝 Text</option>
            <option value="number">🔢 Number</option>
            <option value="date">📅 Date</option>
            <option value="dropdown">📋 Dropdown</option>
            <option value="link">🔗 Hyperlink</option>
            <option value="duration_calc">⏱ Duration (working days)</option>
          </select>
        </div>
        <div class="fg full" id="cg-list" style="display:none">
          <label>Dropdown Source</label><select id="col-list"></select></div>
        <div class="fg full" id="cg-ds" style="display:none">
          <label>Start Date Column</label><select id="col-ds"></select></div>
        <div class="fg full" id="cg-de" style="display:none">
          <label>End Date Column</label><select id="col-de"></select></div>
        <div class="fg full" id="cg-info" style="display:none">
          <div style="background:#f0f4fa;border-radius:6px;padding:9px 12px;font-size:11px;color:#1a3a5c">
            ⏱ Working days between two date columns — excludes Fridays + Egyptian holidays
          </div>
        </div>
      </div>
    </div>
    <div class="mfoot">
      <button class="btn btn-sc" onclick="closeM('addcol-modal')">Cancel</button>
      <button class="btn btn-pr" onclick="saveAddCol()">Add</button>
    </div>
  </div>
</div>

<!-- IMPORT -->
<div class="overlay hidden" id="import-modal">
  <div class="modal" style="max-width:480px">
    <div class="mhdr"><span>📤 Import Excel / CSV</span>
      <button class="xbtn" onclick="closeM('import-modal')">✕</button></div>
    <div class="mbody">
      <p style="font-size:12px;color:var(--mu);margin-bottom:12px">Select .xlsx or .csv file. Headers must match register columns.</p>
      <input type="file" id="imp-file" accept=".csv,.xlsx,.xls" style="font-size:12px">
    </div>
    <div class="mfoot">
      <button class="btn btn-sc" onclick="closeM('import-modal')">Cancel</button>
      <button class="btn btn-pr" onclick="doImport()">Import</button>
    </div>
  </div>
</div>

<!-- ADMIN MODAL (register page) -->
<div class="overlay hidden" id="admin-modal">
  <div class="modal" style="max-width:780px">
    <div class="mhdr"><span>⚙ Admin Panel</span><button class="xbtn" onclick="closeM('admin-modal')">✕</button></div>
    <div class="mbody" id="admin-body"></div>
    <div class="mfoot"><button class="btn btn-sc" onclick="closeM('admin-modal')">Close</button></div>
  </div>
</div>

{SHARED_JS}
<script>
const PID='{pid}', ROLE='{role}', CAN_EDIT={'true' if editable else 'false'};
const SC={sc_json};
const PROJ_FIELDS=[['code','Code'],['name','Project Name'],['startDate','Start Date'],['endDate','End Date'],
  ['client','Client'],['landlord','Landlord'],['pmo','PMO'],['mainConsultant','Consultant'],
  ['mepConsultant','MEP'],['contractor','Contractor']];
const state={{tab:null,cols:[],recs:null,sortCol:null,sortDir:'asc',filters:{{}},editId:null,lists:{{}}}};

// Init
(async()=>{{
  await Promise.all([loadDTs(), loadLists()]);
  updateClock(); setInterval(updateClock,60000);
}})();

function updateClock(){{document.getElementById('s-clock').textContent=new Date().toLocaleString('en-GB');}}

async function loadDTs(keepTab=false){{
  const dts=await apiFetch('/api/doc_types/'+PID); if(!dts)return;
  renderTabs(dts);
  await refreshCounts();
  if(!keepTab){{
    const tab=new URLSearchParams(location.search).get('tab');
    if(dts.length) switchTab(tab&&dts.find(d=>d.id===tab)?tab:dts[0].id);
  }}
}}

async function refreshCounts(){{
  const cnts=await apiFetch('/api/counts/'+PID); if(!cnts)return;
  Object.entries(cnts).forEach(([id,n])=>{{const el=document.getElementById('cnt-'+id);if(el)el.textContent=n;}});
}}

async function loadLists(force=false){{
  if(!force&&Object.keys(state.lists).length) return;
  const d=await apiFetch('/api/lists/'+PID); if(d) state.lists=d;
}}

function renderTabs(dts){{
  const bar=document.getElementById('tabsbar');
  bar.querySelectorAll('.tab-btn').forEach(b=>b.remove());
  const addBtn=bar.querySelector('.tab-add');
  dts.forEach(dt=>{{
    const btn=document.createElement('button');
    btn.className='tab-btn'+(dt.id===state.tab?' active':'');
    btn.dataset.id=dt.id; btn.title=dt.name;
    btn.innerHTML=`<span>${{dt.code}}</span><span class="tcnt" id="cnt-${{dt.id}}">0</span>`;
    btn.onclick=()=>switchTab(dt.id);
    if(CAN_EDIT) btn.oncontextmenu=e=>{{e.preventDefault();tabMenu(dt.id,e);}};
    bar.insertBefore(btn,addBtn||null);
  }});
}}

function switchTab(id){{
  state.tab=id;state.recs=null;state.filters={{}};state.sortCol=null;
  document.getElementById('srchbox').value='';
  document.querySelectorAll('.tab-btn').forEach(b=>b.classList.toggle('active',b.dataset.id===id));
  loadRecords();
}}

function tabMenu(id,e){{
  const old=document.getElementById('tabctx');if(old)old.remove();
  const m=document.createElement('div');m.id='tabctx';
  m.style.cssText=`position:fixed;left:${{e.clientX}}px;top:${{e.clientY}}px;background:#fff;
    border:1px solid var(--bd);border-radius:6px;box-shadow:0 4px 12px rgba(0,0,0,.15);z-index:9999;min-width:140px;overflow:hidden`;
  m.innerHTML=`<div onclick="delDT('${{id}}')" style="padding:8px 14px;cursor:pointer;font-size:12px;color:#ef4444"
    onmouseover="this.style.background='#fef2f2'" onmouseout="this.style.background=''">🗑 Delete Type</div>`;
  document.body.appendChild(m);
  setTimeout(()=>document.addEventListener('click',()=>m.remove(),{{once:true}}),10);
}}

async function delDT(id){{
  if(!confirm('Delete type and ALL its records?'))return;
  await apiFetch('/api/doc_types/'+PID+'/'+id,{{method:'DELETE'}});
  if(state.tab===id)state.tab=null;
  await loadDTs(); toast('Deleted','wa');
}}

async function loadRecords(){{
  if(!state.tab)return;
  const search=document.getElementById('srchbox').value.trim();
  const [data, widths]=await Promise.all([
    apiFetch('/api/records/'+PID+'/'+state.tab+(search?'?search='+encodeURIComponent(search):'')),
    apiFetch('/api/col_width/'+PID+'/'+state.tab)
  ]);
  if(!data)return;
  state.recs=data.records; state.cols=data.columns.filter(c=>c.visible);
  state.colWidths=widths||{{}};
  const cnt=document.getElementById('cnt-'+state.tab); if(cnt)cnt.textContent=data.count;
  buildHead(); requestAnimationFrame(()=>initRz()); renderRows();
}}

function buildHead(){{
  const head=document.getElementById('thead'); head.innerHTML='';
  const hr=document.createElement('tr');
  const chk=document.createElement('th'); chk.className='chkcell';
  if(CAN_EDIT)chk.innerHTML='<input type="checkbox" id="chkall" onchange="selAll(this.checked)">';
  hr.appendChild(chk);
  const sr=document.createElement('th');sr.textContent='Sr.';sr.style.cssText='width:34px;cursor:default;white-space:normal;word-break:break-word';hr.appendChild(sr);
  state.cols.forEach(col=>{{
    const th=document.createElement('th');th.dataset.key=col.col_key;
    const w=state.colWidths&&state.colWidths[col.col_key];
    if(w)th.style.cssText='width:'+w+'px;min-width:'+w+'px;max-width:'+w+'px;white-space:normal;word-break:break-word';
    else th.style.cssText='white-space:normal;word-break:break-word';
    const sortInd=state.sortCol===col.col_key?(state.sortDir==='asc'?' ↑':' ↓'):'';
    th.innerHTML='<span class="th-lbl">'+col.label+sortInd+'</span>';
    if(!['auto_date','auto_num'].includes(col.col_type))th.onclick=()=>sortBy(col.col_key);
    else th.style.cursor='default';
    hr.appendChild(th);
  }});
  const at=document.createElement('th');at.textContent='Actions';at.style.cssText='width:64px;cursor:default';hr.appendChild(at);
  head.appendChild(hr);
  const fr=document.createElement('tr');fr.className='frow';
  fr.appendChild(document.createElement('th'));fr.appendChild(document.createElement('th'));
  state.cols.forEach(col=>{{
    const th=document.createElement('th');
    if(['auto_date','auto_num'].includes(col.col_type)){{fr.appendChild(th);return;}}
    if(col.col_type==='dropdown'&&col.list_name){{
      const sel=document.createElement('select');
      sel.innerHTML='<option value="">All</option>'+(state.lists[col.list_name]||[]).map(o=>`<option ${{state.filters[col.col_key]===o?'selected':''}}>${{o}}</option>`).join('');
      sel.onchange=()=>{{state.filters[col.col_key]=sel.value;renderRows();}};
      th.appendChild(sel);
    }}else{{
      const inp=document.createElement('input');inp.value=state.filters[col.col_key]||'';
      inp.oninput=()=>{{state.filters[col.col_key]=inp.value;renderRows();}};
      th.appendChild(inp);
    }}
    fr.appendChild(th);
  }});
  fr.appendChild(document.createElement('th'));
  head.appendChild(fr);
}}

function parseDocNo(docNo){{
  // Extract base number and rev from "DS-001 REV00" format
  const m=(docNo||'').match(/^([A-Za-z]+)-([0-9]+) REV([0-9]+)$/i);
  if(m)return{{prefix:m[1],num:parseInt(m[2]),rev:parseInt(m[3])}};
  return{{prefix:docNo||'',num:0,rev:0}};
}}

function sortByDocNo(rows){{
  // Group by base (prefix+num), then sort revs within group
  const groups={{}};
  rows.forEach(r=>{{
    const p=parseDocNo(r.docNo||'');
    const key=p.prefix+'-'+String(p.num).padStart(6,'0');
    if(!groups[key])groups[key]={{baseNum:p.num,prefix:p.prefix,rows:[]}};
    groups[key].rows.push({{...r,_parsedNum:p.num,_parsedRev:p.rev}});
  }});
  // Sort groups by base number, then rows within group by rev
  const sorted=[];
  Object.values(groups)
    .sort((a,b)=>a.prefix.localeCompare(b.prefix)||a.baseNum-b.baseNum)
    .forEach(g=>{{
      g.rows.sort((a,b)=>a._parsedRev-b._parsedRev);
      g.rows.forEach(r=>sorted.push(r));
    }});
  return sorted;
}}

function validateDocNo(docNo,existingRecs,editId){{
  if(!docNo)return'Document No. is required';
  const p=parseDocNo(docNo);
  if(!p.num&&!docNo.includes('-'))return'Invalid format. Use: CODE-001 REV00';
  // Duplicate check
  const dup=existingRecs.find(r=>r._id!==editId&&(r.docNo||'').toLowerCase()===docNo.toLowerCase());
  if(dup)return'Document No. already exists: '+docNo;
  // REV check: base must exist
  if(p.rev>0){{
    const hasBase=existingRecs.some(r=>{{
      const bp=parseDocNo(r.docNo||'');
      return bp.prefix===p.prefix&&bp.num===p.num&&bp.rev===0&&r._id!==editId;
    }});
    if(!hasBase)return'Cannot add REV'+String(p.rev).padStart(2,'0')+' — REV00 not found for this document';
  }}
  // Sequence gap check (warning only)
  if(p.rev===0&&p.num>1){{
    const nums=existingRecs
      .filter(r=>{{const bp=parseDocNo(r.docNo||'');return bp.prefix===p.prefix&&bp.rev===0&&r._id!==editId;}})
      .map(r=>parseDocNo(r.docNo||'').num);
    const maxExist=nums.length?Math.max(...nums):0;
    if(p.num>maxExist+1)return'GAP: '+p.prefix+'-'+String(maxExist+1).padStart(3,'0')+' to '+p.prefix+'-'+String(p.num-1).padStart(3,'0')+' are missing';
  }}
  return null;
}}

function checkGap(docNo,recs,editId){{
  const p=parseDocNo(docNo);if(!p||p.rev>0)return null;
  // Find max existing base num for same prefix
  let maxN=0;
  recs.forEach(r=>{{
    if(r._id===editId)return;
    const rp=parseDocNo(r.docNo||'');
    if(rp&&rp.prefix===p.prefix&&rp.rev===0)maxN=Math.max(maxN,rp.num);
  }});
  if(p.num>maxN+1)return`Sequence gap detected: Next expected is ${{p.prefix}}-${{String(maxN+1).padStart(3,'0')}} REV00, but you entered ${{p.prefix}}-${{String(p.num).padStart(3,'0')}} REV00`;
  return null;
}}

function renderRows(){{
  const body=document.getElementById('tbody');body.innerHTML='';
  let rows=state.recs.filter(r=>{{
    for(const[k,v]of Object.entries(state.filters)){{if(v&&!String(r[k]||'').toLowerCase().includes(v.toLowerCase()))return false;}}
    return true;
  }});
  if(state.sortCol){{
    rows.sort((a,b)=>{{
      const va=String(a[state.sortCol]||'').toLowerCase(),vb=String(b[state.sortCol]||'').toLowerCase();
      return state.sortDir==='asc'?(va>vb?1:-1):(va<vb?1:-1);
    }});
  }}else{{
    rows=sortByDocNo(rows);
  }}
  const emptyEl=document.getElementById('empty');
  const tblEl=document.getElementById('regtbl');
  if(rows.length===0){{
    emptyEl.style.display='block';
    tblEl.style.display='none';
  }}else{{
    emptyEl.style.display='none';
    tblEl.style.display='';
  }}
  let sr=1;
  rows.forEach((row,idx)=>{{
    const tr=document.createElement('tr');
    if(row._overdue)tr.classList.add('ov');
    else if(row._isRev)tr.classList.add('rv');
    else if(idx%2===1)tr.classList.add('alt');
    const tc=document.createElement('td');tc.className='chkcell';
    if(CAN_EDIT){{const cb=document.createElement('input');cb.type='checkbox';cb.dataset.id=row._id;cb.onchange=updBulk;tc.appendChild(cb);}}
    tr.appendChild(tc);
    const tsr=document.createElement('td');tsr.className='sr';tsr.textContent=row._isRev?'':sr;tr.appendChild(tsr);
    state.cols.forEach(col=>{{
      const td=document.createElement('td');const key=col.col_key;let val='';
      if(key==='expectedReplyDate'){{val=row._expectedReplyDate||'';if(row._overdue&&val)td.classList.add('ovdate');}}
      else if(key==='duration')val=row._duration||'';
      else if(col.col_type==='duration_calc'){{
        const[ds,de]=(col.list_name||'issuedDate,actualReplyDate').split(',');
        val=calcWD(row[ds.trim()]||'',row[de.trim()]||'');
      }}
      else if(key==='issuedDate')val=row._issuedFmt||'';
      else if(key==='actualReplyDate')val=row._replyFmt||'';
      else if(col.col_type==='date'||col.col_type==='auto_date'){{
        val=row['_fmt_'+key]||row[key]||'';  // use pre-formatted version
      }}
      else if(key==='status'){{
        val=row[key]||'';
        if(val){{
          td.innerHTML=val.split(',').map(s=>{{s=s.trim();const[bg,fg]=SC[s]||['e5e7eb','374151'];
            return `<span class="sbadge" style="background:#${{bg}};color:#${{fg}}">${{s}}</span>`;}}).join('');
          tr.appendChild(td);return;
        }}
      }}
      else if(key==='fileLocation'){{
        const url=row[key]||'';
        if(url){{td.innerHTML=`<a class="flink" href="${{url}}" target="_blank">View</a>`;tr.appendChild(td);return;}}
      }}
      else val=String(row[key]||'');
      td.textContent=val;tr.appendChild(td);
    }});
    const ta=document.createElement('td');ta.className='acts';
    if(CAN_EDIT)ta.innerHTML=`<button class="act" onclick="editRec('${{row._id}}')">✏</button> <button class="act del" onclick="delRec('${{row._id}}')">🗑</button>`;
    else ta.innerHTML='<span style="color:var(--mu);font-size:10px">—</span>';
    tr.appendChild(ta);body.appendChild(tr);
    if(!row._isRev)sr++;
  }});
  const ov=state.recs.filter(r=>r._overdue).length;
  document.getElementById('s-total').textContent='Total: '+state.recs.length;
  document.getElementById('s-show').textContent='Showing: '+rows.length;
  document.getElementById('s-ov').textContent='Overdue: '+ov;
}}

function calcWD(s,e){{
  if(!s||!e)return'';
  try{{let a=new Date(s),b=new Date(e);if(isNaN(a)||isNaN(b)||b<=a)return'0';
    let n=0,c=new Date(a);c.setDate(c.getDate()+1);
    while(c<=b){{if(c.getDay()!==5)n++;c.setDate(c.getDate()+1);}}return String(n);}}
  catch{{return'';}}
}}

function sortBy(key){{
  state.sortDir=state.sortCol===key?(state.sortDir==='asc'?'desc':'asc'):'asc';
  state.sortCol=key;buildHead();renderRows();
}}

let _rzDrag=null;
document.addEventListener('mousemove',e=>{{
  if(!_rzDrag)return;
  const w=Math.max(40,_rzDrag.sw+e.clientX-_rzDrag.sx);
  _rzDrag.th.style.width=w+'px';_rzDrag.th.style.minWidth=w+'px';
}});
document.addEventListener('mouseup',()=>{{
  if(!_rzDrag)return;
  _rzDrag.rz.classList.remove('rzg');
  document.body.style.cursor='';document.body.style.userSelect='none';
  _rzDrag=null;
}});
let _rzActive=null;
function initRz(){{
  document.querySelectorAll('#regtbl thead tr:first-child th[data-key]').forEach(th=>{{
    if(th.querySelector('.rz'))return;
    const rz=document.createElement('div');rz.className='rz';th.appendChild(rz);
    rz.addEventListener('mousedown',e=>{{
      e.stopPropagation();e.preventDefault();
      _rzActive={{th,sx:e.clientX,sw:th.offsetWidth,key:th.dataset.key}};
      rz.classList.add('rzg');
      document.body.style.cursor='col-resize';
      document.body.style.userSelect='none';
    }});
  }});
}}
document.addEventListener('mousemove',e=>{{
  if(!_rzActive)return;
  const w=Math.max(40,_rzActive.sw+(e.clientX-_rzActive.sx));
  _rzActive.th.style.width=w+'px';
  _rzActive.th.style.minWidth=w+'px';
  _rzActive.th.style.maxWidth=w+'px';
}});
let _rzSaveTimer=null;
document.addEventListener('mouseup',()=>{{
  if(!_rzActive)return;
  const {{th,key}}=_rzActive;
  const w=th.offsetWidth;
  th.querySelector('.rz')?.classList.remove('rzg');
  _rzActive=null;
  document.body.style.cursor='';
  document.body.style.userSelect='';
  // Save width if superadmin
  if(ROLE==='superadmin'&&key){{
    clearTimeout(_rzSaveTimer);
    _rzSaveTimer=setTimeout(()=>{{
      apiFetch('/api/col_width/'+PID+'/'+state.tab,{{
        method:'POST',body:JSON.stringify({{col_key:key,width_px:w}})
      }}).then(r=>{{if(r&&r.ok){{if(!state.colWidths)state.colWidths={{}};state.colWidths[key]=w;}}}});
    }},400);
  }}
}});

let _st;
function doSearch(){{clearTimeout(_st);_st=setTimeout(()=>loadRecords(),250);}}

// Add/Edit Record
function addRecord(){{state.editId=null;document.getElementById('rec-title').textContent='Add Document';buildForm(null);openM('rec-modal');}}
function editRec(id){{state.editId=id;const row=state.recs.find(r=>r._id===id);if(!row)return;document.getElementById('rec-title').textContent='Edit Document';buildForm(row);openM('rec-modal');}}

async function buildForm(row){{
  const allCols=await apiFetch('/api/columns/'+PID+'/'+state.tab);if(!allCols)return;
  const AUTO=new Set(['expectedReplyDate','duration','_duration','_duration_today']);
  const grid=document.getElementById('rec-form');grid.innerHTML='';
  let nextNo='';
  if(!row){{const r=await apiFetch('/api/next_doc_no/'+PID+'/'+state.tab);nextNo=r?.next||'';}}
  for(const col of allCols){{
    if(AUTO.has(col.col_key))continue;
    const key=col.col_key;
    const full=['title','remarks','fileLocation','itemRef'].includes(key);
    const grp=document.createElement('div');grp.className='fg'+(full?' full':'');
    const lbl=document.createElement('label');lbl.textContent=col.label;grp.appendChild(lbl);
    const val=row?.[key]||'';
    if(col.col_type==='date'){{const inp=document.createElement('input');inp.type='date';inp.id='f-'+key;inp.value=val;grp.appendChild(inp);}}
    else if(col.col_type==='dropdown'&&col.list_name){{grp.appendChild(buildMS(key,state.lists[col.list_name]||[],val));}}
    else if(col.col_type==='docno'){{
      const inp=document.createElement('input');inp.id='f-'+key;
      if(row){{inp.value=val;}}
      else{{inp.value=nextNo;inp.placeholder=nextNo?'':state.tab+'-001 REV00';}}
      inp.style.cssText='font-family:Consolas,monospace;font-weight:600';
      grp.appendChild(inp);
    }}
    else if(key==='remarks'){{const ta=document.createElement('textarea');ta.id='f-'+key;ta.value=val;ta.rows=3;grp.appendChild(ta);}}
    else{{const inp=document.createElement('input');inp.id='f-'+key;inp.value=val;if(col.col_type==='link')inp.placeholder='https://...';grp.appendChild(inp);}}
    grid.appendChild(grp);
  }}
}}

function buildMS(key,options,init){{
  const sel=init?init.split(',').map(s=>s.trim()).filter(Boolean):[];
  const con=document.createElement('div');con.className='ms-con';con.id='f-'+key;con.dataset.value=init||'';
  function render(){{con.innerHTML='';sel.forEach(v=>{{const t=document.createElement('span');t.className='ms-tag';t.innerHTML=`${{v}} <span class="ms-rm" data-v="${{v}}">✕</span>`;t.querySelector('.ms-rm').onclick=e=>{{e.stopPropagation();sel.splice(sel.indexOf(v),1);con.dataset.value=sel.join(', ');render();}};con.appendChild(t);}});if(!sel.length)con.innerHTML='<span class="ms-ph">Select...</span>';con.dataset.value=sel.join(', ');}}
  con.onclick=e=>{{if(e.target.classList.contains('ms-rm'))return;const ex=document.querySelector('.ms-dd');if(ex){{ex.remove();return;}}
    const dd=document.createElement('div');dd.className='ms-dd';
    options.forEach(opt=>{{const it=document.createElement('div');it.className='ms-opt'+(sel.includes(opt)?' sel':'');it.innerHTML=`<input type="checkbox" ${{sel.includes(opt)?'checked':''}} style="pointer-events:none"> ${{opt}}`;it.onclick=ev=>{{ev.stopPropagation();if(sel.includes(opt))sel.splice(sel.indexOf(opt),1);else sel.push(opt);con.dataset.value=sel.join(', ');render();it.classList.toggle('sel',sel.includes(opt));it.querySelector('input').checked=sel.includes(opt);}};dd.appendChild(it);}});
    con.style.position='relative';con.appendChild(dd);}};
  document.addEventListener('click',e=>{{if(!con.contains(e.target))con.querySelector('.ms-dd')?.remove();}},true);
  render();return con;
}}

function durChoice(docNo){{
  return new Promise(resolve=>{{
    const ov=document.createElement('div');ov.className='overlay';ov.style.cssText='position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:9999;display:flex;align-items:center;justify-content:center';
    ov.innerHTML=`<div style="background:#fff;border-radius:12px;padding:28px 24px;max-width:380px;width:90%;box-shadow:0 20px 60px rgba(0,0,0,.3)">
      <div style="font-size:16px;font-weight:700;margin-bottom:8px">⏱ Duration</div>
      <p style="font-size:13px;color:#64748b;margin-bottom:20px">Issue date is set but no reply date yet.<br>How should Duration be shown?</p>
      <div style="display:flex;flex-direction:column;gap:8px">
        <button id="dur-today" style="padding:10px 16px;background:#1a3a5c;color:#fff;border:none;border-radius:6px;cursor:pointer;font-size:13px">📅 Calculate from issue date → Today</button>
        <button id="dur-empty" style="padding:10px 16px;background:#f1f5f9;color:#374151;border:none;border-radius:6px;cursor:pointer;font-size:13px">— Leave Duration empty</button>
        <button id="dur-cancel" style="padding:8px;background:none;border:none;cursor:pointer;font-size:12px;color:#94a3b8">Cancel</button>
      </div></div>`;
    document.body.appendChild(ov);
    ov.querySelector('#dur-today').onclick=()=>{{document.body.removeChild(ov);resolve('today');}};
    ov.querySelector('#dur-empty').onclick=()=>{{document.body.removeChild(ov);resolve('empty');}};
    ov.querySelector('#dur-cancel').onclick=()=>{{document.body.removeChild(ov);resolve(null);}};
  }});
}}

async function saveRecord(){{
  const allCols=await apiFetch('/api/columns/'+PID+'/'+state.tab);if(!allCols)return;
  const AUTO=new Set(['expectedReplyDate','duration','_duration','_duration_today']);
  const data={{}};
  for(const col of allCols){{
    if(AUTO.has(col.col_key))continue;
    const el=document.getElementById('f-'+col.col_key);if(!el)continue;
    data[col.col_key]=el.classList.contains('ms-con')?el.dataset.value||'':el.tagName==='TEXTAREA'?el.value.trim():el.value.trim();
  }}
  if(!data.docNo){{toast('Document No. required','er');return;}}
    // Duration is computed server-side automatically
  const valErr=validateDocNo(data.docNo,state.recs||[],state.editId);
  if(valErr&&valErr.startsWith('GAP')){{
    if(!confirm('⚠ Sequence gap detected: '+valErr.replace('GAP:','')+'. Continue anyway?'))return;
  }}else if(valErr){{toast('⚠ '+valErr,'er');return;}}
  if(state.editId)data._id=state.editId;
  const r=await apiFetch('/api/records/'+PID+'/'+state.tab,{{method:'POST',body:JSON.stringify(data)}});
  if(r&&r.ok){{closeM('rec-modal');const savedTab=state.tab;await loadRecords();await refreshCounts();toast(state.editId?'Updated':'Added','ok');}}
  else toast('Error saving','er');
}}

async function delRec(id){{
  if(!confirm('Delete this record?'))return;
  await apiFetch('/api/records/'+id,{{method:'DELETE'}});
  await loadRecords();await refreshCounts();toast('Deleted','wa');
}}

// Bulk
function updBulk(){{
  const checked=document.querySelectorAll('.chkcell input[data-id]:checked');
  document.getElementById('bulkcnt').textContent=checked.length+' selected';
  document.getElementById('bulkbar').classList.toggle('show',checked.length>0);
  const all=document.querySelectorAll('.chkcell input[data-id]');
  const ca=document.getElementById('chkall');if(ca)ca.checked=all.length>0&&checked.length===all.length;
}}
function selAll(v){{document.querySelectorAll('.chkcell input[data-id]').forEach(cb=>cb.checked=v);updBulk();}}
function clearSel(){{document.querySelectorAll('.chkcell input').forEach(cb=>cb.checked=false);updBulk();}}
async function bulkDel(){{
  const ids=[...document.querySelectorAll('.chkcell input[data-id]:checked')].map(cb=>cb.dataset.id);
  if(!ids.length||!confirm('Delete '+ids.length+' records?'))return;
  let ok=0;for(const id of ids){{const r=await apiFetch('/api/records/'+id,{{method:'DELETE'}});if(r&&r.ok)ok++;}}
  clearSel();await loadRecords();await refreshCounts();toast('✔ Deleted '+ok,'ok');
}}

// Project Modal
async function editProject(){{
  const proj=await apiFetch('/api/project/'+PID);if(!proj)return;
  const body=document.getElementById('proj-body');body.innerHTML='';
  const grid=document.createElement('div');grid.className='fgrid';
  PROJ_FIELDS.forEach(([key,lbl])=>{{
    const grp=document.createElement('div');grp.className='fg';
    const label=document.createElement('label');label.textContent=lbl;
    const inp=document.createElement('input');inp.id='pf-'+key;inp.value=proj[key]||'';
    grp.appendChild(label);grp.appendChild(inp);grid.appendChild(grp);
  }});
  body.appendChild(grid);
  // Logo section
  const lt=document.createElement('div');lt.className='stitle';lt.textContent='Company Logos';body.appendChild(lt);
  const lg=document.createElement('div');lg.style.cssText='display:grid;grid-template-columns:1fr 1fr;gap:12px';
  for(const[k,lbl]of[['logo_left','Left Logo'],['logo_right','Right Logo']]){{
    const d=document.createElement('div');d.style.cssText='border:1px solid var(--bd);border-radius:6px;padding:10px;text-align:center;background:var(--bg)';
    const lbel=document.createElement('div');lbel.style.cssText='font-size:10px;font-weight:700;color:var(--mu);text-transform:uppercase;margin-bottom:6px';lbel.textContent=lbl;d.appendChild(lbel);
    const img=document.createElement('img');img.id='lp-'+k;img.style.cssText='max-height:52px;max-width:100%;object-fit:contain;display:block;margin:0 auto 6px';d.appendChild(img);
    fetch('/api/logo/'+PID+'/'+k).then(r=>r.ok?r.blob():null).then(b=>{{if(b)img.src=URL.createObjectURL(b);}});
    const fi=document.createElement('input');fi.type='file';fi.accept='image/*';fi.style.cssText='width:100%;font-size:10px;margin-bottom:4px';
    fi.onchange=async e=>{{const f=e.target.files[0];if(!f)return;const b64=await new Promise(res=>{{const fr=new FileReader();fr.onload=e2=>res(e2.target.result);fr.readAsDataURL(f);}});img.src=b64;await apiFetch('/api/logo/'+PID+'/'+k,{{method:'POST',body:JSON.stringify({{data:b64.split(',')[1]}})}});toast('Logo saved','ok');}};
    d.appendChild(fi);
    const clr=document.createElement('button');clr.className='btn btn-sc btn-sm';clr.textContent='Remove';clr.style.fontSize='9px';
    clr.onclick=async()=>{{await apiFetch('/api/logo/'+PID+'/'+k,{{method:'POST',body:JSON.stringify({{data:''}})}});img.src='';toast('Removed','wa');}};
    d.appendChild(clr);lg.appendChild(d);
  }}
  body.appendChild(lg);
  openM('proj-modal');
}}

async function saveProject(){{
  const data={{}};
  PROJ_FIELDS.forEach(([key])=>{{const el=document.getElementById('pf-'+key);if(el)data[key]=el.value.trim();}});
  const r=await apiFetch('/api/project/'+PID,{{method:'POST',body:JSON.stringify(data)}});
  if(r===null)return;
  if(r.ok){{
    closeM('proj-modal');toast('✔ Project saved!','ok');
    PROJ_FIELDS.forEach(([key])=>{{const el=document.querySelector('#projbar .pf-val[data-key="'+key+'"]');if(el)el.textContent=data[key]||'—';}});
  }}else toast('Save failed','er');
}}

// Doc Types
function addDocType(){{document.getElementById('dt-code').value='';document.getElementById('dt-name').value='';openM('dt-modal');}}
async function saveDocType(){{
  const code=document.getElementById('dt-code').value.trim().toUpperCase();
  const name=document.getElementById('dt-name').value.trim();
  if(!code||!name){{toast('Code and name required','er');return;}}
  await apiFetch('/api/doc_types/'+PID,{{method:'POST',body:JSON.stringify({{code,name}})}});
  closeM('dt-modal');await loadDTs();switchTab(code);toast('✔ Type added','ok');
}}

// Lists
async function openLists(){{
  await loadLists(true);
  const body=document.getElementById('lists-body');body.innerHTML='';
  for(const[ln,items]of Object.entries(state.lists)){{
    const t=document.createElement('div');t.className='stitle';t.textContent=ln.charAt(0).toUpperCase()+ln.slice(1);body.appendChild(t);
    const ul=document.createElement('ul');ul.className='slist';
    items.forEach(item=>{{const li=document.createElement('li');li.className='sitem';li.innerHTML=`<span class="nm">${{item}}</span><button onclick="rmItem('${{ln}}','${{item.replace(/'/g,"\\'")}}'  ,this)">Remove</button>`;ul.appendChild(li);}});
    body.appendChild(ul);
    const ar=document.createElement('div');ar.className='addrow';
    ar.innerHTML=`<input id="new-${{ln}}" placeholder="New item..."><button class="btn btn-ok btn-sm" onclick="addItem('${{ln}}')">Add</button>`;
    body.appendChild(ar);
  }}
  body.innerHTML+=`<div class="stitle">New List</div><div class="addrow"><input id="new-list" placeholder="List name"><button class="btn btn-pr btn-sm" onclick="mkList()">Create</button></div>`;
  openM('lists-modal');
}}
async function addItem(ln){{
  const inp=document.getElementById('new-'+ln);const val=inp?.value.trim();if(!val)return;
  await apiFetch('/api/lists/'+PID,{{method:'POST',body:JSON.stringify({{list_name:ln,item:val}})}});
  await loadLists(true);openLists();
}}
async function rmItem(ln,item,btn){{
  await apiFetch('/api/lists/'+PID,{{method:'DELETE',body:JSON.stringify({{list_name:ln,item}})}});
  btn.closest('li').remove();await loadLists(true);
}}
async function mkList(){{
  const name=document.getElementById('new-list')?.value.trim().toLowerCase().replace(/[\\s]+/g,'_');if(!name)return;
  await apiFetch('/api/lists/'+PID,{{method:'POST',body:JSON.stringify({{list_name:name,item:'Item 1'}})}});
  await loadLists(true);openLists();
}}

// Column drag-to-reorder
function initColDrag(ul){{
  let drag=null,over=null;
  ul.querySelectorAll('li').forEach(li=>{{
    li.draggable=true;
    li.ondragstart=e=>{{drag=li;li.style.opacity='.4';e.dataTransfer.effectAllowed='move';}};
    li.ondragend=()=>{{drag.style.opacity='';drag=null;ul.querySelectorAll('li').forEach(l=>l.style.background='');}};
    li.ondragover=e=>{{e.preventDefault();if(li===drag)return;over=li;
      ul.querySelectorAll('li').forEach(l=>l.style.background='');
      li.style.background='#e0f2fe';}};
    li.ondrop=e=>{{e.preventDefault();if(!drag||drag===li)return;
      const items=[...ul.querySelectorAll('li')];
      const di=items.indexOf(drag),oi=items.indexOf(li);
      if(di<oi)ul.insertBefore(drag,li.nextSibling);else ul.insertBefore(drag,li);
      saveColOrder();}};
  }});
}}
async function saveColOrder(){{
  const ids=[...document.querySelectorAll('#col-sortable li')].map(li=>parseInt(li.dataset.id));
  await apiFetch('/api/columns/reorder/'+PID+'/'+state.tab,{{method:'POST',body:JSON.stringify({{order:ids}})}});
  toast('✔ Order saved','ok');
}}

// Columns
async function manageColumns(){{
  const cols=await apiFetch('/api/columns/'+PID+'/'+state.tab);if(!cols)return;
  const body=document.getElementById('col-body');body.innerHTML='';
  const ul=document.createElement('ul');ul.id='col-sortable';ul.className='slist';ul.style.maxHeight='400px';
  cols.forEach(col=>{{
    const li=document.createElement('li');li.className='sitem';li.dataset.id=col.id;li.draggable=true;
    li.style.cssText='cursor:default;align-items:center;gap:7px';
    const grip=document.createElement('span');grip.textContent='⠿';
    grip.style.cssText='cursor:grab;color:var(--mu);font-size:16px;flex-shrink:0';
    grip.title='Drag to reorder';
    const chk=document.createElement('input');chk.type='checkbox';chk.checked=col.visible;
    chk.onchange=()=>toggleCol(col.id,chk.checked);
    const nm=document.createElement('span');nm.textContent=col.label;nm.style.cssText='flex:1;font-size:12px';
    const tp=document.createElement('span');tp.textContent=col.col_type;
    tp.style.cssText='font-size:9px;background:#e0e7ff;color:#3730a3;padding:1px 6px;border-radius:3px;flex-shrink:0';
    const ren_btn=document.createElement('button');ren_btn.textContent='✏';
    ren_btn.title='Rename';ren_btn.className='btn btn-sc btn-sm';
    ren_btn.style.cssText='flex-shrink:0;padding:2px 6px;font-size:11px';
    ren_btn.onclick=async()=>{{
      const newLbl=prompt('New label for column:',col.label);
      if(!newLbl||!newLbl.trim())return;
      const r=await apiFetch('/api/columns/rename/'+col.id,{{method:'POST',body:JSON.stringify({{label:newLbl.trim()}})}});
      if(r&&r.ok){{nm.textContent=newLbl.trim();toast('✔ Renamed','ok');}}
      else toast('Error','er');
    }};
    const db_btn=document.createElement('button');db_btn.textContent='🗑';
    db_btn.title='Delete';db_btn.className='btn btn-er btn-sm';
    db_btn.style.cssText='flex-shrink:0;padding:2px 6px;font-size:11px';
    db_btn.onclick=()=>deleteCol(col.id,col.col_key,db_btn);
    li.appendChild(grip);li.appendChild(chk);li.appendChild(nm);li.appendChild(tp);li.appendChild(ren_btn);li.appendChild(db_btn);
    ul.appendChild(li);
  }});
  initColDrag(ul);
  body.appendChild(ul);openM('col-modal');
}}
async function toggleCol(id,v){{await apiFetch('/api/columns/visibility/'+id,{{method:'POST',body:JSON.stringify({{visible:v}})}});}}
async function deleteCol(id,key,btn){{const warn=key==='docNo'?'⚠ WARNING: Deleting Document No. column will break the register! Are you sure?':'Delete this column?';if(!confirm(warn))return;await apiFetch('/api/columns/'+id,{{method:'DELETE'}});btn.closest('li').remove();}}

function onColType(v){{
  document.getElementById('cg-list').style.display=v==='dropdown'?'flex':'none';
  document.getElementById('cg-ds').style.display=v==='duration_calc'?'flex':'none';
  document.getElementById('cg-de').style.display=v==='duration_calc'?'flex':'none';
  document.getElementById('cg-info').style.display=v==='duration_calc'?'block':'none';
}}
async function openAddCol(){{
  await loadLists();
  const ls=document.getElementById('col-list');
  ls.innerHTML=Object.keys(state.lists).map(k=>`<option value="${{k}}">${{k}}</option>`).join('');
  const all=await apiFetch('/api/columns/'+PID+'/'+state.tab);
  const dates=(all||[]).filter(c=>['date','auto_date'].includes(c.col_type));
  const dopts=dates.map(c=>`<option value="${{c.col_key}}">${{c.label}}</option>`).join('')||'<option value="issuedDate">Issued Date</option>';
  document.getElementById('col-ds').innerHTML=dopts;
  document.getElementById('col-de').innerHTML=dopts;
  document.getElementById('col-name').value='';document.getElementById('col-type').value='text';onColType('text');
  openM('addcol-modal');
}}
async function saveAddCol(){{
  const name=document.getElementById('col-name').value.trim();
  const type=document.getElementById('col-type').value;
  const list=type==='dropdown'?document.getElementById('col-list').value:null;
  const ds=type==='duration_calc'?document.getElementById('col-ds').value:null;
  const de=type==='duration_calc'?document.getElementById('col-de').value:null;
  if(!name){{toast('Name required','er');return;}}
  if(type==='duration_calc'&&ds===de){{toast('Start and end must differ','er');return;}}
  await apiFetch('/api/columns/'+PID+'/'+state.tab,{{method:'POST',body:JSON.stringify({{label:name,col_type:type,list_name:type==='duration_calc'?(ds+','+de):list}})}});
  closeM('addcol-modal');closeM('col-modal');await loadRecords();toast('✔ Column added','ok');
}}

// Admin Panel (same as dashboard)
async function openAdmin(){{
  const [users,projects]=await Promise.all([apiFetch('/api/users'),apiFetch('/api/projects')]);
  if(!users||!projects) return;
  const body=document.getElementById('admin-body'); body.innerHTML='';
  const utitle=document.createElement('div');utitle.className='stitle';utitle.textContent='👥 Users';body.appendChild(utitle);
  for(const u of users){{
    const row=document.createElement('div');row.className='urow';
    row.innerHTML=`<span style="flex:1;font-weight:600">👤 ${{u.username}}</span>
      <span class="badge" style="background:#fef3c7;color:#92400e">${{u.role.toUpperCase()}}</span>
      ${{u.username!=='admin'?`<button class="btn btn-sc btn-sm" onclick="chgPw('${{u.username}}')">🔑 PW</button>
        <button class="btn btn-er btn-sm" onclick="delUsr('${{u.username}}')">✕</button>`:
        '<span style="font-size:10px;color:var(--mu)">(protected)</span>'}}`;
    body.appendChild(row);
    if(u.role!=='superadmin'){{
      const ad=document.createElement('div');ad.style.cssText='padding:4px 10px 10px 32px;border-bottom:1px solid var(--bd);margin-bottom:4px';
      ad.innerHTML='<div style="font-size:10px;color:var(--mu);margin-bottom:4px">Project access:</div>';
      const assigned=await apiFetch('/api/users/'+u.username+'/projects').catch(()=>[]);
      const pl=document.createElement('div');pl.style.cssText='display:flex;flex-wrap:wrap;gap:4px';
      projects.forEach(p=>{{
        const btn=document.createElement('button');btn.className='pbtn'+(assigned.includes(p.id)?' on':'');
        btn.textContent=p.code;btn.title=p.name;
        btn.onclick=async()=>{{const on=btn.classList.contains('on');
          await apiFetch('/api/users',{{method:'POST',body:JSON.stringify({{action:on?'unassign':'assign',username:u.username,project_id:p.id}})}});
          btn.classList.toggle('on');toast((btn.classList.contains('on')?'✔ Assigned: ':'Removed: ')+p.name,'ok');}};
        pl.appendChild(btn);}});
      ad.appendChild(pl);body.appendChild(ad);
    }}
  }}
  const at=document.createElement('div');at.className='stitle';at.textContent='➕ Add User';body.appendChild(at);
  const ar=document.createElement('div');
  ar.innerHTML=`<div style="display:grid;grid-template-columns:1fr 1fr 1fr auto;gap:8px;align-items:end">
    <div class="fg"><label>Username</label><input id="nu-name" placeholder="username"></div>
    <div class="fg"><label>Role</label><select id="nu-role">
      <option value="editor">Editor</option><option value="viewer">Viewer</option>
      <option value="superadmin">Super Admin</option></select></div>
    <div class="fg"><label>Password</label><input id="nu-pw" type="password"></div>
    <button class="btn btn-pr btn-sm" style="margin-bottom:1px" onclick="addUsr()">Add</button></div>`;
  body.appendChild(ar);openM('admin-modal');
}}
async function addUsr(){{
  const name=document.getElementById('nu-name').value.trim().toLowerCase();
  const role=document.getElementById('nu-role').value;
  const pw=document.getElementById('nu-pw').value;
  if(!name||!pw){{toast('Username and password required','er');return;}}
  const r=await apiFetch('/api/users',{{method:'POST',body:JSON.stringify({{action:'add',username:name,role,password:pw}})}});
  if(r&&r.ok){{toast('✔ User added','ok');closeM('admin-modal');openAdmin();}}else toast((r&&r.error)||'Error','er');
}}
async function delUsr(u){{if(!confirm('Delete user: '+u+'?'))return;
  const r=await apiFetch('/api/users',{{method:'POST',body:JSON.stringify({{action:'delete',username:u}})}});
  if(r&&r.ok){{toast('Deleted','wa');closeM('admin-modal');openAdmin();}}
}}
async function chgPw(u){{const pw=prompt('New password for '+u+':');if(!pw)return;
  await apiFetch('/api/users',{{method:'POST',body:JSON.stringify({{action:'change_password',username:u,password:pw}})}});
  toast('✔ Password changed','ok');
}}

// Export/Import
function doExport(){{if(state.tab)window.location='/api/export/'+PID+'/'+state.tab;}}
function doExportPDF(){{if(state.tab)window.location='/api/export_pdf/'+PID+'/'+state.tab;}}
function doExportAllPDF(){{window.location='/api/export_pdf_all/'+PID;}}
function toggleExpDD(e){{e.stopPropagation();document.getElementById('exp-menu').classList.toggle('hidden');}}
function closeExpDD(){{document.getElementById('exp-menu').classList.add('hidden');}}
document.addEventListener('click',e=>{{if(!e.target.closest('#exp-dd'))closeExpDD();}});
function doPrint(){{
  const orig=document.title;
  document.title='DCR Print';
  window.print();
  document.title=orig;
}}
// Holidays Settings
async function openSettings(){{
  const r=await apiFetch('/api/settings/holidays');if(!r)return;
  let hols=[...(r.holidays||[])];
  const ov=document.createElement('div');ov.className='overlay';
  ov.style.cssText='position:fixed;inset:0;background:rgba(0,0,0,.55);z-index:9999;display:flex;align-items:center;justify-content:center';
  function buildUI(){{
    ov.innerHTML=`<div style="background:#fff;border-radius:12px;padding:24px;max-width:520px;width:95%;max-height:85vh;overflow:hidden;display:flex;flex-direction:column;box-shadow:0 20px 60px rgba(0,0,0,.3)">
      <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:14px">
        <div style="font-size:16px;font-weight:800;color:#1a3a5c">🗓 Public Holidays</div>
        <button id="hol-close" style="background:none;border:none;font-size:20px;cursor:pointer;color:#64748b">✕</button>
      </div>
      <p style="font-size:12px;color:#64748b;margin-bottom:12px">These dates are excluded from Duration and Expected Reply calculations (along with Fridays).</p>
      <div style="display:flex;gap:8px;margin-bottom:12px">
        <input type="date" id="hol-inp" style="flex:1;padding:6px 10px;border:1.5px solid #e2e8f0;border-radius:6px;font-family:inherit;font-size:12px">
        <button id="hol-add-btn" style="padding:6px 14px;background:#1a3a5c;color:#fff;border:none;border-radius:6px;cursor:pointer;font-size:12px;font-weight:600">+ Add</button>
      </div>
      <div id="hol-list" style="flex:1;overflow-y:auto;border:1px solid #e2e8f0;border-radius:6px;padding:8px;min-height:180px;max-height:340px"></div>
      <div style="margin-top:12px;font-size:11px;color:#94a3b8">${{hols.length}} holidays total</div>
      <div style="display:flex;gap:8px;margin-top:12px;justify-content:flex-end">
        <button id="hol-cancel" style="padding:7px 16px;border:1.5px solid #e2e8f0;background:#fff;border-radius:6px;cursor:pointer;font-size:12px">Cancel</button>
        <button id="hol-save" style="padding:7px 16px;background:#16a34a;color:#fff;border:none;border-radius:6px;cursor:pointer;font-size:12px;font-weight:600">💾 Save Holidays</button>
      </div></div>`;
    renderHL();
    ov.querySelector('#hol-close').onclick=()=>ov.remove();
    ov.querySelector('#hol-cancel').onclick=()=>ov.remove();
    ov.querySelector('#hol-add-btn').onclick=()=>{{
      const v=ov.querySelector('#hol-inp').value;
      if(!v)return;if(hols.includes(v)){{toast('Already added','wa');return;}}
      hols.push(v);renderHL();ov.querySelector('#hol-inp').value='';
    }};
    ov.querySelector('#hol-save').onclick=async()=>{{
      const btn=ov.querySelector('#hol-save');btn.disabled=true;btn.textContent='Saving...';
      const res=await apiFetch('/api/settings/holidays',{{method:'POST',body:JSON.stringify({{holidays:hols}})}});
      btn.disabled=false;btn.textContent='💾 Save Holidays';
      if(res&&res.ok){{toast('✔ '+res.count+' holidays saved','ok');ov.remove();}}
      else toast('Error','er');
    }};
  }}
  function renderHL(){{
    const ul=ov.querySelector('#hol-list');if(!ul)return;ul.innerHTML='';
    if(!hols.length){{ul.innerHTML='<div style="color:#94a3b8;text-align:center;padding:30px;font-size:12px">No holidays — click + Add to add dates</div>';return;}}
    [...hols].sort().forEach(d=>{{
      const row=document.createElement('div');
      row.style.cssText='display:flex;align-items:center;justify-content:space-between;padding:5px 8px;border-radius:4px;margin-bottom:3px;background:#f8fafc;font-size:12px';
      const dt=new Date(d+'T00:00:00');
      const fmt=dt.getDate().toString().padStart(2,'0')+'-'+(dt.getMonth()+1).toString().padStart(2,'0')+'-'+dt.getFullYear();
      const days=['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];
      const day=days[dt.getDay()];
      row.innerHTML=`<span style="color:#1e293b">${{fmt}} <span style="color:#94a3b8;font-size:10px">(${{day}})</span></span>
        <button data-d="${{d}}" style="background:none;border:none;color:#ef4444;cursor:pointer;font-size:15px;line-height:1;padding:2px 6px">✕</button>`;
      row.querySelector('button').onclick=e=>{{const dv=e.currentTarget.dataset.d;hols=hols.filter(h=>h!==dv);renderHL();}};
      ul.appendChild(row);
    }});
  }}
  buildUI();
  document.body.appendChild(ov);
}}

function doExportAll(){{window.location='/api/export_all/'+PID;}}
function openImport(){{openM('import-modal');}}
async function doImport(){{
  const file=document.getElementById('imp-file').files[0];if(!file)return;
  const ext=file.name.split('.').pop().toLowerCase();
  const btn=document.querySelector('#import-modal .btn-pr');
  btn.disabled=true;btn.textContent='⏳ Importing...';
  try{{
    const b64=await new Promise((res,rej)=>{{const fr=new FileReader();fr.onload=e=>res(e.target.result);fr.onerror=rej;fr.readAsDataURL(file);}});
    const r=await apiFetch('/api/import/'+PID+'/'+state.tab,{{method:'POST',body:JSON.stringify({{file_b64:b64,ext}})}});
    if(r===null)return;
    closeM('import-modal');switchTab(state.tab);await loadDTs();toast('✔ Imported '+r.imported+' records','ok');
  }}catch(e){{toast('Error: '+e.message,'er');}}
  finally{{btn.disabled=false;btn.textContent='Import';}}
}}
</script></body></html>"""


if __name__ == "__main__":
    db.init()
    db.cleanup_sessions()
    port = int(os.environ.get("PORT", 5000))
    print(f"[DCR Flask] Running on http://localhost:{port}")
    app.run(host="0.0.0.0", port=port, debug=not IS_RENDER)
