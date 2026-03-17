"""db.py - DCR Flask - Clean PostgreSQL database layer"""
import os, json, uuid, datetime, hashlib, secrets
import psycopg2
from psycopg2.extras import RealDictCursor
from psycopg2.pool import ThreadedConnectionPool

DATABASE_URL = os.environ.get("DATABASE_URL", "")
_pool = None

def get_pool():
    global _pool
    if _pool is None:
        _pool = ThreadedConnectionPool(1, 10, DATABASE_URL,
            sslmode="require", connect_timeout=10,
            keepalives=1, keepalives_idle=30,
            keepalives_interval=10, keepalives_count=5)
    return _pool

class DB:
    def __enter__(self):
        self.conn = get_pool().getconn()
        self.conn.autocommit = False
        return self.conn.cursor(cursor_factory=RealDictCursor)
    def __exit__(self, exc, *_):
        if exc: self.conn.rollback()
        else:   self.conn.commit()
        get_pool().putconn(self.conn)

def q(sql, params=(), one=False):
    with DB() as cur:
        cur.execute(sql, params)
        rows = cur.fetchone() if one else cur.fetchall()
    if one:  return dict(rows) if rows else None
    return [dict(r) for r in rows] if rows else []

def exe(sql, params=()):
    with DB() as cur:
        cur.execute(sql, params)

def exem(sql, param_list):
    with DB() as cur:
        cur.executemany(sql, param_list)

# ── Schema ────────────────────────────────────────────────────
SCHEMA = """
CREATE TABLE IF NOT EXISTS users (
    username   TEXT PRIMARY KEY,
    pw_hash    TEXT NOT NULL,
    role       TEXT NOT NULL DEFAULT 'viewer'
);
CREATE TABLE IF NOT EXISTS sessions (
    token      TEXT PRIMARY KEY,
    username   TEXT NOT NULL,
    role       TEXT NOT NULL,
    expires_at TIMESTAMPTZ NOT NULL
);
CREATE TABLE IF NOT EXISTS projects (
    id         TEXT PRIMARY KEY,
    name       TEXT NOT NULL,
    code       TEXT NOT NULL,
    data       JSONB NOT NULL DEFAULT '{}'
);
CREATE TABLE IF NOT EXISTS user_projects (
    username   TEXT NOT NULL,
    project_id TEXT NOT NULL,
    PRIMARY KEY (username, project_id)
);
CREATE TABLE IF NOT EXISTS logos (
    project_id TEXT NOT NULL,
    logo_key   TEXT NOT NULL,
    image_data TEXT,
    PRIMARY KEY (project_id, logo_key)
);
CREATE TABLE IF NOT EXISTS doc_types (
    id         TEXT NOT NULL,
    project_id TEXT NOT NULL,
    name       TEXT NOT NULL,
    code       TEXT NOT NULL,
    sort_order INTEGER DEFAULT 0,
    PRIMARY KEY (id, project_id)
);
CREATE TABLE IF NOT EXISTS columns_config (
    id          SERIAL PRIMARY KEY,
    project_id  TEXT NOT NULL,
    dt_id       TEXT NOT NULL,
    col_key     TEXT NOT NULL,
    label       TEXT NOT NULL,
    col_type    TEXT NOT NULL DEFAULT 'text',
    list_name   TEXT,
    visible     BOOLEAN DEFAULT TRUE,
    sort_order  INTEGER DEFAULT 0,
    UNIQUE (project_id, dt_id, col_key)
);
CREATE TABLE IF NOT EXISTS records (
    id          TEXT PRIMARY KEY,
    project_id  TEXT NOT NULL,
    dt_id       TEXT NOT NULL,
    data        JSONB NOT NULL DEFAULT '{}',
    created_at  TIMESTAMPTZ DEFAULT NOW(),
    updated_at  TIMESTAMPTZ DEFAULT NOW()
);
CREATE INDEX IF NOT EXISTS idx_records_proj_dt ON records(project_id, dt_id);
CREATE TABLE IF NOT EXISTS app_settings (
    key        TEXT PRIMARY KEY,
    value      JSONB NOT NULL DEFAULT '{}'
);
CREATE TABLE IF NOT EXISTS col_widths (
    project_id TEXT NOT NULL,
    dt_id      TEXT NOT NULL,
    col_key    TEXT NOT NULL,
    width_px   INTEGER NOT NULL DEFAULT 120,
    PRIMARY KEY (project_id, dt_id, col_key)
);
CREATE TABLE IF NOT EXISTS dropdown_lists (
    id          SERIAL PRIMARY KEY,
    project_id  TEXT NOT NULL,
    list_name   TEXT NOT NULL,
    item_value  TEXT NOT NULL,
    sort_order  INTEGER DEFAULT 0,
    meta        TEXT DEFAULT NULL,
    UNIQUE (project_id, list_name, item_value)
);
ALTER TABLE dropdown_lists ADD COLUMN IF NOT EXISTS meta TEXT DEFAULT NULL;
"""

DEFAULT_DOC_TYPES = [
    ("DS","Document Submittal",0),("SD","Shop Drawing Submittal",1),
    ("MS","Material Submittal",2),("IR","Inspection Request",3),
    ("MIR","Material Inspection Request",4),("RFI","Request for Information",5),
    ("WN","Work Notification",6),("QS","Quantity Surveyor",7),
    ("ABD","As Built Drawing",8),("CVI","Confirmation of Verbal Instruction",9),
    ("PCQ","Potential Change Questionnaire",10),("SI","Site Instructions",11),
    ("EI","Engineer's Instruction",12),("SVP","Safety Violation with Penalty",13),
    ("NOC","Notice of Change",14),("NCR","Non-Conformance Report",15),
    ("MOM","Minutes of Meetings",16),("IOM","Internal Office Memo",17),
    ("PR","Requisition",18),("LTR","Letters",19),
]

DEFAULT_COLS = [
    ("docNo","Document No.","docno",None,0),
    ("discipline","Discipline","dropdown","discipline",1),
    ("trade","Sub-Trade","dropdown","trade",2),
    ("title","Title","text",None,3),
    ("floor","Floor","dropdown","floor",4),
    ("itemRef","Item Ref./DWG No.","text",None,5),
    ("issuedDate","Issued Date","date",None,6),
    ("expectedReplyDate","Expected Reply","auto_date",None,7),
    ("actualReplyDate","Actual Reply","date",None,8),
    ("status","Status","dropdown","status",9),
    ("duration","Duration (days)","auto_num",None,10),
    ("remarks","Remarks","text",None,11),
    ("fileLocation","File Location","link",None,12),
]

# Default meta categories for known status values
DEFAULT_STATUS_META = {
    "A - Approved":              "approved",
    "B - Approved As Noted":     "approved",
    "B,C - Approved & Resubmit": "approved",
    "C - Revise & Resubmit":     "rejected",
    "D - Review not Required":   "rejected",
    "Under Review":              "pending",
    "Pending":                   "pending",
    "Open":                      "pending",
    "Not Closed":                "pending",
    "Replied":                   "approved",
    "Closed":                    "approved",
    "Cancelled":                 "cancelled",
}

DEFAULT_LISTS = {
    "discipline": ["Electrical","Mechanical","Civil","Structural","Architecture","General","Others"],
    "trade":["Lighting","Power","Light Current","Fire Alarm","Low Voltage","Medium Voltage",
             "Control","HVAC","Plumbing","Fire Fighting","General","Others"],
    "floor":["Basement Floor 1","Basement Floor 2","Ground Floor","First Floor","Second Floor",
             "Third Floor","Fourth Floor","Roof Floor","Upper Roof"],
    "status":["A - Approved","B - Approved As Noted","B,C - Approved & Resubmit",
              "C - Revise & Resubmit","D - Review not Required","Under Review",
              "Cancelled","Open","Closed","Replied","Pending"],
}

def init():
    stmts = [s.strip() for s in SCHEMA.strip().split(";") if s.strip()]
    for stmt in stmts:
        try:
            with DB() as cur: cur.execute(stmt)
        except Exception as e:
            if "already exists" not in str(e).lower():
                print(f"[DB] Schema note: {e}")
    _ensure_admin()

def _ensure_admin():
    r = q("SELECT COUNT(*) as c FROM users", one=True)
    if not r or r["c"] == 0:
        exe("INSERT INTO users(username,pw_hash,role) VALUES(%s,%s,%s) ON CONFLICT DO NOTHING",
            ("admin", hash_pw("admin123"), "superadmin"))

# ── Auth ──────────────────────────────────────────────────────
def hash_pw(pw): return hashlib.sha256(pw.encode()).hexdigest()

def verify_pw(username, pw):
    u = q("SELECT pw_hash FROM users WHERE username=%s", (username,), one=True)
    return u and u["pw_hash"] == hash_pw(pw)

def get_user(username):
    return q("SELECT * FROM users WHERE username=%s", (username,), one=True)

def get_all_users():
    return q("SELECT username, role FROM users ORDER BY username")

def add_user(username, pw, role="viewer"):
    exe("INSERT INTO users(username,pw_hash,role) VALUES(%s,%s,%s) ON CONFLICT DO NOTHING",
        (username, hash_pw(pw), role))

def delete_user(username):
    exe("DELETE FROM sessions WHERE username=%s", (username,))
    exe("DELETE FROM user_projects WHERE username=%s", (username,))
    exe("DELETE FROM users WHERE username=%s", (username,))

def change_pw(username, new_pw):
    exe("UPDATE users SET pw_hash=%s WHERE username=%s", (hash_pw(new_pw), username))

# ── Sessions ──────────────────────────────────────────────────
SESSION_TTL = 8 * 3600

def create_session(username, role):
    token = secrets.token_hex(32)
    exp   = datetime.datetime.utcnow() + datetime.timedelta(seconds=SESSION_TTL)
    exe("INSERT INTO sessions(token,username,role,expires_at) VALUES(%s,%s,%s,%s)",
        (token, username, role, exp))
    return token

def get_session(token):
    if not token: return None
    r = q("SELECT username,role FROM sessions WHERE token=%s AND expires_at > NOW()",
          (token,), one=True)
    return r

def delete_session(token):
    exe("DELETE FROM sessions WHERE token=%s", (token,))

def cleanup_sessions():
    exe("DELETE FROM sessions WHERE expires_at <= NOW()")

# ── User ↔ Project ────────────────────────────────────────────
def get_user_projects(username):
    return [r["project_id"] for r in
            q("SELECT project_id FROM user_projects WHERE username=%s", (username,))]

def assign_project(username, project_id):
    exe("INSERT INTO user_projects(username,project_id) VALUES(%s,%s) ON CONFLICT DO NOTHING",
        (username, project_id))

def unassign_project(username, project_id):
    exe("DELETE FROM user_projects WHERE username=%s AND project_id=%s", (username, project_id))

def get_project_users(project_id):
    return [r["username"] for r in
            q("SELECT username FROM user_projects WHERE project_id=%s", (project_id,))]

# ── Projects ──────────────────────────────────────────────────
def get_projects():
    return q("SELECT * FROM projects ORDER BY id")

def get_project(pid):
    r = q("SELECT * FROM projects WHERE id=%s", (pid,), one=True)
    if not r: return None
    data = r.get("data") or {}
    if isinstance(data, str):
        try: data = json.loads(data)
        except: data = {}
    return {"id": r["id"], "name": r["name"], "code": r["code"], **data}

def save_project(pid, name, code, extra: dict):
    jdata = json.dumps(extra)
    exe("""INSERT INTO projects(id,name,code,data) VALUES(%s,%s,%s,%s::jsonb)
           ON CONFLICT(id) DO UPDATE SET name=EXCLUDED.name, code=EXCLUDED.code, data=EXCLUDED.data""",
        (pid, name, code, jdata))

def create_project(pid, name, code, creator=None):
    exe("INSERT INTO projects(id,name,code,data) VALUES(%s,%s,%s,'{}') ON CONFLICT DO NOTHING",
        (pid, name, code))
    _seed_project(pid)
    if creator: assign_project(creator, pid)

def delete_project(pid):
    for tbl in ("records","columns_config","doc_types","dropdown_lists","logos","user_projects","projects"):
        exe(f"DELETE FROM {tbl} WHERE {'id' if tbl=='projects' else 'project_id'}=%s", (pid,))

def _seed_project(pid):
    for code, name, order in DEFAULT_DOC_TYPES:
        exe("INSERT INTO doc_types(id,project_id,name,code,sort_order) VALUES(%s,%s,%s,%s,%s) ON CONFLICT DO NOTHING",
            (code, pid, name, code, order))
        for ck, lbl, ct, ln, so in DEFAULT_COLS:
            exe("""INSERT INTO columns_config(project_id,dt_id,col_key,label,col_type,list_name,visible,sort_order)
                   VALUES(%s,%s,%s,%s,%s,%s,true,%s) ON CONFLICT DO NOTHING""",
                (pid, code, ck, lbl, ct, ln, so))
    for ln, items in DEFAULT_LISTS.items():
        for i, item in enumerate(items):
            meta = DEFAULT_STATUS_META.get(item) if ln == "status" else None
            exe("INSERT INTO dropdown_lists(project_id,list_name,item_value,sort_order,meta)"
                " VALUES(%s,%s,%s,%s,%s) ON CONFLICT(project_id,list_name,item_value)"
                " DO UPDATE SET meta=COALESCE(dropdown_lists.meta, EXCLUDED.meta)",
                (pid, ln, item, i, meta))

# ── Logos ─────────────────────────────────────────────────────
def get_logo(pid, key):
    r = q("SELECT image_data FROM logos WHERE project_id=%s AND logo_key=%s", (pid,key), one=True)
    return r["image_data"] if r else None

def save_logo(pid, key, data):
    exe("""INSERT INTO logos(project_id,logo_key,image_data) VALUES(%s,%s,%s)
           ON CONFLICT(project_id,logo_key) DO UPDATE SET image_data=EXCLUDED.image_data""",
        (pid, key, data))

# ── Doc Types ─────────────────────────────────────────────────
def get_doc_types(pid):
    return q("SELECT * FROM doc_types WHERE project_id=%s ORDER BY sort_order, id", (pid,))

def add_doc_type(pid, code, name):
    r = q("SELECT COALESCE(MAX(sort_order),0)+1 as n FROM doc_types WHERE project_id=%s", (pid,), one=True)
    order = r["n"] if r else 0
    exe("INSERT INTO doc_types(id,project_id,name,code,sort_order) VALUES(%s,%s,%s,%s,%s) ON CONFLICT DO NOTHING",
        (code, pid, name, code, order))
    for ck, lbl, ct, ln, so in DEFAULT_COLS:
        exe("""INSERT INTO columns_config(project_id,dt_id,col_key,label,col_type,list_name,visible,sort_order)
               VALUES(%s,%s,%s,%s,%s,%s,true,%s) ON CONFLICT DO NOTHING""",
            (pid, code, ck, lbl, ct, ln, so))

def delete_doc_type(pid, dt_id):
    exe("DELETE FROM doc_types WHERE id=%s AND project_id=%s", (dt_id, pid))
    exe("DELETE FROM records WHERE dt_id=%s AND project_id=%s", (dt_id, pid))
    exe("DELETE FROM columns_config WHERE dt_id=%s AND project_id=%s", (dt_id, pid))

# ── Columns ───────────────────────────────────────────────────
def get_columns(pid, dt_id, visible_only=False):
    sql = "SELECT * FROM columns_config WHERE project_id=%s AND dt_id=%s"
    if visible_only: sql += " AND visible=true"
    sql += " ORDER BY sort_order"
    return q(sql, (pid, dt_id))

def add_column(pid, dt_id, col_key, label, col_type, list_name=None):
    r = q("SELECT COALESCE(MAX(sort_order),0)+1 as n FROM columns_config WHERE project_id=%s AND dt_id=%s",
          (pid, dt_id), one=True)
    so = r["n"] if r else 0
    exe("""INSERT INTO columns_config(project_id,dt_id,col_key,label,col_type,list_name,visible,sort_order)
           VALUES(%s,%s,%s,%s,%s,%s,true,%s) ON CONFLICT DO NOTHING""",
        (pid, dt_id, col_key, label, col_type, list_name, so))

def set_col_visible(col_id, visible):
    exe("UPDATE columns_config SET visible=%s WHERE id=%s", (visible, col_id))

def delete_col(col_id):
    exe("DELETE FROM columns_config WHERE id=%s", (col_id,))

# ── Records ───────────────────────────────────────────────────
def get_records(pid, dt_id, search=""):
    rows = q("SELECT id, data, created_at FROM records WHERE project_id=%s AND dt_id=%s ORDER BY created_at",
             (pid, dt_id))
    result = []
    for row in rows:
        d = row["data"] if isinstance(row["data"], dict) else {}
        d["_id"]      = row["id"]
        d["_created"] = str(row.get("created_at",""))
        result.append(d)
    if search:
        sq = search.lower()
        result = [r for r in result if any(sq in str(v).lower() for v in r.values())]
    return result

def count_records(pid, dt_id):
    r = q("SELECT COUNT(*) as c FROM records WHERE project_id=%s AND dt_id=%s", (pid, dt_id), one=True)
    return r["c"] if r else 0

def save_record(pid, dt_id, rec_id, data: dict):
    clean = {k: v for k, v in data.items() if not k.startswith("_")}
    jdata = json.dumps(clean)
    exe("""INSERT INTO records(id,project_id,dt_id,data,updated_at) VALUES(%s,%s,%s,%s::jsonb,NOW())
           ON CONFLICT(id) DO UPDATE SET data=EXCLUDED.data, updated_at=NOW()""",
        (rec_id, pid, dt_id, jdata))

def delete_record(rec_id):
    exe("DELETE FROM records WHERE id=%s", (rec_id,))

# ── Dropdown Lists ────────────────────────────────────────────
def get_lists(pid):
    names = [r["list_name"] for r in
             q("SELECT DISTINCT list_name FROM dropdown_lists WHERE project_id=%s ORDER BY list_name", (pid,))]
    return {n: [r["item_value"] for r in
                q("SELECT item_value FROM dropdown_lists WHERE project_id=%s AND list_name=%s ORDER BY sort_order",
                  (pid, n))] for n in names}

def get_lists_with_meta(pid):
    """Return {list_name: [{value, meta}]} for status-aware lists."""
    rows = q("SELECT list_name, item_value, meta FROM dropdown_lists"
             " WHERE project_id=%s ORDER BY list_name, sort_order", (pid,))
    result = {}
    for r in rows:
        result.setdefault(r["list_name"], []).append(
            {"value": r["item_value"], "meta": r["meta"]})
    return result

def get_status_meta_map(pid):
    """Return {status_value: meta} for all status lists in this project."""
    rows = q("SELECT item_value, meta FROM dropdown_lists"
             " WHERE project_id=%s AND list_name LIKE %s AND meta IS NOT NULL",
             (pid, "status%"))
    return {r["item_value"]: r["meta"] for r in rows}

def set_list_item_meta(pid, list_name, item_value, meta):
    exe("UPDATE dropdown_lists SET meta=%s"
        " WHERE project_id=%s AND list_name=%s AND item_value=%s",
        (meta, pid, list_name, item_value))

def add_list_item(pid, list_name, item):
    r = q("SELECT COALESCE(MAX(sort_order),0)+1 as n FROM dropdown_lists WHERE project_id=%s AND list_name=%s",
          (pid, list_name), one=True)
    exe("INSERT INTO dropdown_lists(project_id,list_name,item_value,sort_order) VALUES(%s,%s,%s,%s) ON CONFLICT DO NOTHING",
        (pid, list_name, item, r["n"] if r else 0))

def remove_list_item(pid, list_name, item):
    exe("DELETE FROM dropdown_lists WHERE project_id=%s AND list_name=%s AND item_value=%s",
        (pid, list_name, item))

# ── Fast Dashboard Stats (single query) ──────────────────────
def get_dashboard_stats():
    """One big SQL call — replaces N×M individual queries."""
    import re, json as _json
    # Pre-load status meta for all projects
    meta_rows = q("SELECT project_id, item_value, meta FROM dropdown_lists"
                  " WHERE list_name LIKE %s AND meta IS NOT NULL", ("status%",))
    # meta_map[pid][status_value] = meta
    meta_map = {}
    for r in meta_rows:
        meta_map.setdefault(r["project_id"], {})[r["item_value"]] = r["meta"]
    from utils import is_overdue

    projects = get_projects()
    if not projects:
        return []

    pids = [p["id"] for p in projects]
    ph   = ",".join(["%s"] * len(pids))

    # All records in one query
    all_recs = q(f"""
        SELECT project_id, dt_id, data
        FROM records
        WHERE project_id IN ({ph})
    """, pids)

    # All doc_types in one query
    all_dts = q(f"""
        SELECT project_id, id, code, name, sort_order
        FROM doc_types
        WHERE project_id IN ({ph})
        ORDER BY project_id, sort_order, id
    """, pids)

    dt_map   = {}  # pid -> list of dts
    for dt in all_dts:
        dt_map.setdefault(dt["project_id"], []).append(dt)

    rec_map  = {}  # (pid, dt_id) -> list of data dicts
    for row in all_recs:
        key = (row["project_id"], row["dt_id"])
        d   = row["data"] if isinstance(row["data"], dict) else {}
        rec_map.setdefault(key, []).append(d)

    result = []
    for p in projects:
        pid   = p["id"]
        pdata = p.get("data") or {}
        if isinstance(pdata, str):
            try: pdata = _json.loads(pdata)
            except: pdata = {}

        dts      = dt_map.get(pid, [])
        dt_stats = []
        total = approved = pending = overdue_cnt = total_rj = 0

        # Which doc types have expectedReplyDate and status columns
        exp_col_rows = q(
            "SELECT DISTINCT dt_id, col_key FROM columns_config"
            " WHERE project_id=%s AND col_key IN ('expectedReplyDate','status')",
            (pid,))
        dt_has_exp_reply = {}
        dt_has_status    = {}
        for r in exp_col_rows:
            if r["col_key"] == "expectedReplyDate":
                dt_has_exp_reply[r["dt_id"]] = True
            elif r["col_key"] == "status":
                dt_has_status[r["dt_id"]] = True

        for dt in dts:
            rows  = rec_map.get((pid, dt["id"]), [])
            t=ap=pe=ov=rj = 0
            disc_map = {}

            # ── Approach 1: group by base doc, use LATEST rev status ──
            # Collect ALL revisions per base document first
            doc_groups = {}  # base_key → list of (rev, row)
            for d in rows:
                doc_no = d.get("docNo", "")
                m = re.search(r"REV(\d+)", doc_no, re.IGNORECASE)
                rev = int(m.group(1)) if m else 0
                base = re.sub(r"\s*REV\d+$", "", doc_no, flags=re.IGNORECASE).strip()
                doc_groups.setdefault(base, []).append((rev, d))

            pid_meta = meta_map.get(pid, {})

            for base, revisions in doc_groups.items():
                revisions.sort(key=lambda x: x[0])   # sort ascending by rev
                latest_rev, d = revisions[-1]
                doc_no = d.get("docNo", "")
                status = d.get("status", "")
                disc   = d.get("discipline","") or "—"
                # Docs from types WITHOUT status column → count in total only (no status bucket)
                if not dt_has_status.get(dt["id"], False):
                    t += 1
                    is_ov = is_overdue(d.get("issuedDate"), doc_no, d.get("actualReplyDate"), dt_has_exp_reply.get(dt["id"], False))
                    if is_ov: ov += 1
                    ds = disc_map.setdefault(disc, {"total":0,"approved":0,"pending":0,"rejected":0,"overdue":0})
                    ds["total"] += 1
                    if is_ov: ds["overdue"] += 1
                    continue
                # Resolve meta: DB first, then fallback to DEFAULT_STATUS_META
                meta = pid_meta.get(status) or DEFAULT_STATUS_META.get(status) or "pending"
                if meta == "cancelled": continue   # skip cancelled entirely
                t += 1
                is_ap = (meta == "approved")
                is_rj = (meta == "rejected")
                is_pe = (meta == "pending")
                is_ov = is_overdue(d.get("issuedDate"), doc_no, d.get("actualReplyDate"), dt_has_exp_reply.get(dt["id"], False))
                if is_ap: ap += 1
                if is_pe: pe += 1
                if is_rj: rj += 1
                if is_ov: ov += 1
                # Discipline breakdown
                ds = disc_map.setdefault(disc, {"total":0,"approved":0,"pending":0,"rejected":0,"overdue":0})
                ds["total"] += 1
                if is_ap: ds["approved"] += 1
                if is_pe: ds["pending"]  += 1
                if is_rj: ds["rejected"] += 1
                if is_ov: ds["overdue"]  += 1

            disc_list = [{"disc":k,"total":v["total"],"approved":v["approved"],
                          "pending":v["pending"],"rejected":v["rejected"],"overdue":v["overdue"]}
                         for k,v in sorted(disc_map.items())]
            dt_stats.append({"id":dt["id"],"code":dt["code"],"name":dt["name"],
                             "total":t,"approved":ap,"pending":pe,"overdue":ov,"rejected":rj,
                             "disc_breakdown":disc_list})
            total+=t; approved+=ap; pending+=pe; overdue_cnt+=ov; total_rj+=rj

        pct = round(approved / total * 100) if total else 0
        result.append({
            "id": pid, "name": p["name"], "code": p["code"],
            "client": pdata.get("client","") if isinstance(pdata, dict) else "",
            "total": total, "approved": approved, "pending": pending,
            "overdue": overdue_cnt, "rejected": total_rj, "pct": pct, "dt_stats": dt_stats
        })
    return result



def get_monthly_trend(pid=None):
    """Returns last 6 months of submission counts per month."""
    import re as _re
    from utils import DEFAULT_STATUS_META
    where = "WHERE project_id=%s" if pid else "WHERE 1=1"
    params = (pid,) if pid else ()
    rows = q(f"SELECT data FROM records {where}", params)
    months = {}
    import datetime
    today = datetime.date.today()
    for i in range(5, -1, -1):
        m = (today.month - i - 1) % 12 + 1
        y = today.year - ((today.month - i - 1) // 12)
        key = f"{y:04d}-{m:02d}"
        months[key] = {"submitted": 0, "approved": 0}
    for row in rows:
        d = row["data"] if isinstance(row["data"], dict) else {}
        issued = d.get("issuedDate", "") or ""
        if not issued: continue
        try:
            key = issued[:7]
        except: continue
        if key not in months: continue
        months[key]["submitted"] += 1
        meta = DEFAULT_STATUS_META.get(d.get("status", ""), "pending")
        if meta == "approved":
            months[key]["approved"] += 1
    return [{"month": k, "submitted": v["submitted"], "approved": v["approved"]}
            for k, v in sorted(months.items())]


def get_aging_report(pid=None):
    """Returns pending docs grouped by days lapsed: 1-7, 8-14, 15-21, >21."""
    from utils import is_overdue
    import datetime, re as _re
    where = "WHERE project_id=%s" if pid else "WHERE 1=1"
    params = (pid,) if pid else ()
    rows = q(f"SELECT data FROM records {where}", params)
    buckets = {"1-7": 0, "8-14": 0, "15-21": 0, ">21": 0}
    today = datetime.date.today()
    for row in rows:
        d = row["data"] if isinstance(row["data"], dict) else {}
        if d.get("actualReplyDate"): continue
        issued = d.get("issuedDate", "")
        if not issued: continue
        try:
            dt = datetime.date.fromisoformat(str(issued)[:10])
        except: continue
        days = (today - dt).days
        if days < 1: continue
        if days <= 7:   buckets["1-7"]   += 1
        elif days <= 14: buckets["8-14"]  += 1
        elif days <= 21: buckets["15-21"] += 1
        else:           buckets[">21"]   += 1
    return [{"range": k, "count": v} for k, v in buckets.items()]


def get_quality_report(pid=None):
    """Returns doc quality: how many docs needed 0,1,2,3+ revisions."""
    import re as _re
    where = "WHERE project_id=%s" if pid else "WHERE 1=1"
    params = (pid,) if pid else ()
    rows = q(f"SELECT data FROM records {where}", params)
    doc_revs = {}
    for row in rows:
        d = row["data"] if isinstance(row["data"], dict) else {}
        doc_no = d.get("docNo", "") or ""
        m = _re.search(r"REV(\d+)", doc_no, _re.IGNORECASE)
        rev = int(m.group(1)) if m else 0
        base = _re.sub(r"\s*REV\d+$", "", doc_no, flags=_re.IGNORECASE).strip()
        if base:
            doc_revs[base] = max(doc_revs.get(base, 0), rev)
    buckets = {"0": 0, "1": 0, "2": 0, "3+": 0}
    for rev in doc_revs.values():
        if rev == 0:   buckets["0"]  += 1
        elif rev == 1: buckets["1"]  += 1
        elif rev == 2: buckets["2"]  += 1
        else:          buckets["3+"] += 1
    return [{"revisions": k, "count": v} for k, v in buckets.items()]


def get_overdue_records(pid=None):
    """Returns overdue records — only for doc types WITH expectedReplyDate column."""
    from utils import is_overdue
    import datetime as _dt
    where = "WHERE project_id=%s" if pid else "WHERE 1=1"
    params = (pid,) if pid else ()
    rows = q(f"""SELECT r.project_id, r.dt_id, r.data, d.code as dt_code, d.name as dt_name
                 FROM records r
                 JOIN doc_types d ON d.id=r.dt_id AND d.project_id=r.project_id
                 {where.replace("project_id", "r.project_id")}""", params)
    # Only dt_ids with expectedReplyDate column
    if pid:
        exp_rows = q("SELECT DISTINCT dt_id FROM columns_config WHERE project_id=%s AND col_key='expectedReplyDate'", (pid,))
    else:
        exp_rows = q("SELECT DISTINCT dt_id FROM columns_config WHERE col_key='expectedReplyDate'")
    dt_with_exp = {r["dt_id"] for r in exp_rows}
    result = []
    for row in rows:
        if row["dt_id"] not in dt_with_exp: continue
        d = row["data"] if isinstance(row["data"], dict) else {}
        if d.get("actualReplyDate"): continue
        doc_no = d.get("docNo", "") or ""
        issued = d.get("issuedDate", "") or ""
        if not issued: continue
        if is_overdue(issued, doc_no, None, True):
            try:
                dt = _dt.date.fromisoformat(str(issued)[:10])
                days = (_dt.date.today() - dt).days
            except: days = 0
            result.append({
                "project_id": row["project_id"],
                "dt_code": row["dt_code"],
                "dt_name": row["dt_name"],
                "docNo": doc_no,
                "title": d.get("title", ""),
                "discipline": d.get("discipline", ""),
                "issuedDate": issued,
                "days_overdue": days,
                "status": d.get("status", ""),
            })
    result.sort(key=lambda x: x["days_overdue"], reverse=True)
    return result

def rename_column(col_id, new_label):
    exe("UPDATE columns_config SET label=%s WHERE id=%s", (new_label.strip(), col_id))

def get_col_widths(pid, dt_id):
    rows = q("SELECT col_key, width_px FROM col_widths WHERE project_id=%s AND dt_id=%s", (pid, dt_id))
    return {r["col_key"]: r["width_px"] for r in rows}

def save_col_width(pid, dt_id, col_key, width_px):
    exe(
        "INSERT INTO col_widths(project_id,dt_id,col_key,width_px) VALUES(%s,%s,%s,%s)"
        " ON CONFLICT(project_id,dt_id,col_key) DO UPDATE SET width_px=EXCLUDED.width_px",
        (pid, dt_id, col_key, width_px))

def get_setting(key, default=None):
    r = q("SELECT value FROM app_settings WHERE key=%s", (key,), one=True)
    if r is None: return default
    import json
    v = r["value"]
    return v if not isinstance(v, str) else json.loads(v)

def save_setting(key, value):
    import json
    exe(
        "INSERT INTO app_settings(key,value) VALUES(%s,%s::jsonb)"
        " ON CONFLICT(key) DO UPDATE SET value=EXCLUDED.value",
        (key, json.dumps(value)))

def reorder_columns(pid, dt_id, ordered_ids):
    """Update sort_order for columns given a list of ids in desired order."""
    for i, col_id in enumerate(ordered_ids):
        exe("UPDATE columns_config SET sort_order=%s WHERE id=%s AND project_id=%s",
            (i, col_id, pid))
