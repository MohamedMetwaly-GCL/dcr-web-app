"""db.py - DCR Flask - Clean PostgreSQL database layer"""
import logging
import os, json, uuid, datetime, hashlib, secrets
import bcrypt as _bcrypt
import psycopg2
from psycopg2.extras import RealDictCursor
from psycopg2.pool import ThreadedConnectionPool

DATABASE_URL = os.environ.get("DATABASE_URL", "")
_pool = None
logger = logging.getLogger(__name__)

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
CREATE TABLE IF NOT EXISTS pr_items (
    id          TEXT PRIMARY KEY,
    record_id   TEXT NOT NULL REFERENCES records(id) ON DELETE CASCADE,
    item_name   TEXT NOT NULL,
    unit        TEXT,
    quantity    NUMERIC,
    remarks     TEXT,
    sort_order  INTEGER DEFAULT 0,
    created_at  TIMESTAMPTZ DEFAULT NOW(),
    updated_at  TIMESTAMPTZ DEFAULT NOW()
);
CREATE INDEX IF NOT EXISTS idx_pr_items_record ON pr_items(record_id);
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
CREATE TABLE IF NOT EXISTS audit_log (
    id          SERIAL PRIMARY KEY,
    ts          TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    username    TEXT NOT NULL,
    action      TEXT NOT NULL,
    project_id  TEXT,
    dt_id       TEXT,
    record_id   TEXT,
    doc_no      TEXT,
    field_name  TEXT,
    old_value   TEXT,
    new_value   TEXT,
    detail      TEXT
);
CREATE INDEX IF NOT EXISTS idx_audit_ts ON audit_log(ts DESC);
CREATE INDEX IF NOT EXISTS idx_audit_proj ON audit_log(project_id, ts DESC);
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

def _is_non_workflow_dt(dt_code="", dt_name=""):
    code = str(dt_code or "").strip().lower()
    name = str(dt_name or "").strip().lower()
    return code in ("pr", "ltr") or "requisition" in name or "letter" in name

def _is_pr_dt(dt_code="", dt_name=""):
    code = str(dt_code or "").strip().lower()
    name = str(dt_name or "").strip().lower()
    return code == "pr" or "requisition" in name or "purchase request" in name

def init():
    stmts = [s.strip() for s in SCHEMA.strip().split(";") if s.strip()]
    for stmt in stmts:
        try:
            with DB() as cur: cur.execute(stmt)
        except Exception as e:
            if "already exists" not in str(e).lower():
                logger.warning("db_schema_note error=%s", e)
    _ensure_admin()

def _ensure_admin():
    r = q("SELECT COUNT(*) as c FROM users", one=True)
    if not r or r["c"] == 0:
        exe("INSERT INTO users(username,pw_hash,role) VALUES(%s,%s,%s) ON CONFLICT DO NOTHING",
            ("admin", hash_pw("admin123"), "superadmin"))

# ── Auth ──────────────────────────────────────────────────────
def hash_pw(pw: str) -> str:
    """Always produces a bcrypt hash. Used for new passwords and upgrades."""
    return _bcrypt.hashpw(pw.encode(), _bcrypt.gensalt()).decode()

def _is_bcrypt(h: str) -> bool:
    return h.startswith("$2b$") or h.startswith("$2a$")

def verify_pw(username: str, pw: str) -> bool:
    """
    Verify password with automatic lazy migration from SHA-256 to bcrypt.
    - If stored hash is already bcrypt → verify directly with bcrypt.
    - If stored hash is legacy SHA-256 → verify with SHA-256, then silently
      re-hash with bcrypt and update the DB row in the same call.
    This means every user is upgraded on their next successful login with
    zero disruption: no forced logout, no data loss, no migration script.
    """
    u = q("SELECT pw_hash FROM users WHERE username=%s", (username,), one=True)
    if not u:
        return False
    stored = u["pw_hash"]
    if _is_bcrypt(stored):
        return _bcrypt.checkpw(pw.encode(), stored.encode())
    # Legacy SHA-256 path
    legacy_hash = hashlib.sha256(pw.encode()).hexdigest()
    if stored == legacy_hash:
        # Password correct — silently upgrade to bcrypt
        exe("UPDATE users SET pw_hash=%s WHERE username=%s",
            (hash_pw(pw), username))
        return True
    return False

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
    if not isinstance(data, dict): data = {}
    # Never let data override the core fields
    data.pop("id", None); data.pop("name", None); data.pop("code", None)
    return {"id": r["id"], "name": r["name"], "code": r["code"], **data}

def save_project(pid, name, code, extra: dict):
    if isinstance(extra, dict):
        extra.pop("id", None); extra.pop("name", None); extra.pop("code", None)
    jdata = json.dumps(extra or {})
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
def get_record_by_id(rec_id):
    """
    Return full row context: _id, _project_id, _dt_id, _updated_at,
    plus all document data fields unpacked at the top level.
    The _ prefix on injected keys matches the existing convention (_id, _created)
    so no downstream code needs to change for read operations.
    Returns None if record not found.
    """
    r = q("""SELECT id, project_id, dt_id, data, updated_at
             FROM records WHERE id=%s""", (rec_id,), one=True)
    if not r:
        return None
    data = r["data"] if isinstance(r["data"], dict) else {}
    return {
        "_id":         r["id"],
        "_project_id": r["project_id"],
        "_dt_id":      r["dt_id"],
        "_updated_at": str(r["updated_at"] or ""),
        **data,
    }

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

# ── PR Items ─────────────────────────────────────────────────
def get_pr_items(record_id):
    return q("""SELECT id, record_id, item_name, unit, quantity, remarks, sort_order
                FROM pr_items
                WHERE record_id=%s
                ORDER BY sort_order, created_at, id""", (record_id,))

def get_pr_items_for_records(record_ids):
    if not record_ids: return {}
    record_ids = [r for r in record_ids if r]
    if not record_ids: return {}
    rows = q("""SELECT id, record_id, item_name, unit, quantity, remarks, sort_order
                FROM pr_items
                WHERE record_id = ANY(%s)
                ORDER BY record_id, sort_order, created_at, id""", (record_ids,))
    out = {}
    for r in rows:
        out.setdefault(r["record_id"], []).append(r)
    return out

def save_pr_items(record_id, items):
    exe("DELETE FROM pr_items WHERE record_id=%s", (record_id,))
    if not items: return 0
    for i, it in enumerate(items):
        exe("""INSERT INTO pr_items(id,record_id,item_name,unit,quantity,remarks,sort_order,updated_at)
               VALUES(%s,%s,%s,%s,%s,%s,%s,NOW())""",
            (str(uuid.uuid4()), record_id, it.get("item_name",""),
             it.get("unit"), it.get("quantity"), it.get("remarks"), i))
    return len(items)

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

        # Which doc types have expectedReplyDate, status, and discipline columns
        exp_col_rows = q(
            "SELECT DISTINCT dt_id, col_key FROM columns_config"
            " WHERE project_id=%s AND col_key IN ('expectedReplyDate','status','discipline')"
            " AND visible=TRUE",
            (pid,))
        dt_has_exp_reply = {}
        dt_has_status    = {}
        dt_has_discipline = {}
        for r in exp_col_rows:
            if r["col_key"] == "expectedReplyDate":
                dt_has_exp_reply[r["dt_id"]] = True
            elif r["col_key"] == "status":
                dt_has_status[r["dt_id"]] = True
            elif r["col_key"] == "discipline":
                dt_has_discipline[r["dt_id"]] = True

        for dt in dts:
            rows  = rec_map.get((pid, dt["id"]), [])
            t=ap=pe=ov=rj = 0
            disc_map = {}
            dt_is_non_workflow = _is_non_workflow_dt(dt.get("code",""), dt.get("name",""))
            dt_is_pr = _is_pr_dt(dt.get("code",""), dt.get("name",""))

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
                    _has_both = dt_has_exp_reply.get(dt["id"], False) and dt_has_status.get(dt["id"], False) and not dt_is_non_workflow
                    is_ov = is_overdue(d.get("issuedDate"), doc_no, d.get("actualReplyDate"), _has_both)
                    if is_ov: ov += 1
                    if dt_has_discipline.get(dt["id"], False):
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
                _has_both = dt_has_exp_reply.get(dt["id"], False) and dt_has_status.get(dt["id"], False) and not dt_is_non_workflow
                is_ov = is_overdue(d.get("issuedDate"), doc_no, d.get("actualReplyDate"), _has_both)
                if is_ap: ap += 1
                if is_pe: pe += 1
                if is_rj: rj += 1
                if is_ov: ov += 1
                # Discipline breakdown
                # Only add to discipline breakdown if this dt has a discipline column
                if dt_has_discipline.get(dt["id"], False):
                    ds = disc_map.setdefault(disc, {"total":0,"approved":0,"pending":0,"rejected":0,"overdue":0})
                    ds["total"] += 1
                    if is_ap: ds["approved"] += 1
                    if is_pe: ds["pending"]  += 1
                    if is_rj: ds["rejected"] += 1
                    if is_ov: ds["overdue"]  += 1

            disc_list = [{"disc":k,"total":v["total"],"approved":v["approved"],
                          "pending":v["pending"],"rejected":v["rejected"],"overdue":v["overdue"]}
                         for k,v in sorted(disc_map.items())]
            if not dt_is_pr:
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


def get_data_quality_summary(pid=None):
    dt_rows = q("SELECT project_id, id, code, name FROM doc_types WHERE project_id=%s", (pid,)) if pid else q("SELECT project_id, id, code, name FROM doc_types")
    if not dt_rows:
        return {
            "total_records": 0,
            "issued_date_count": 0,
            "issued_date_pct": 0,
            "workflow_records": 0,
            "status_count": 0,
            "status_pct": 0,
            "actual_reply_count": 0,
            "actual_reply_pct": 0,
            "pr_records": 0,
            "pr_structured_count": 0,
            "pr_structured_pct": 0,
        }

    dt_meta = {(d["project_id"], d["id"]): d for d in dt_rows}
    col_rows = q("""
        SELECT project_id, dt_id, col_key
        FROM columns_config
        WHERE visible=TRUE AND col_key IN ('expectedReplyDate','status')
    """ + (" AND project_id=%s" if pid else ""), (pid,)) if pid else q("""
        SELECT project_id, dt_id, col_key
        FROM columns_config
        WHERE visible=TRUE AND col_key IN ('expectedReplyDate','status')
    """)
    dt_has_exp = set()
    dt_has_status = set()
    for r in col_rows:
        key = (r["project_id"], r["dt_id"])
        if r["col_key"] == "expectedReplyDate":
            dt_has_exp.add(key)
        elif r["col_key"] == "status":
            dt_has_status.add(key)

    workflow_dt_keys = set()
    pr_dt_keys = set()
    for key, dt in dt_meta.items():
        if key in dt_has_exp and key in dt_has_status and not _is_non_workflow_dt(dt.get("code",""), dt.get("name","")):
            workflow_dt_keys.add(key)
        code = str(dt.get("code","")).strip().upper()
        name = str(dt.get("name","")).strip().lower()
        if code == "PR" or "requisition" in name or "purchase request" in name:
            pr_dt_keys.add(key)

    recs = q("SELECT id, project_id, dt_id, data FROM records WHERE project_id=%s", (pid,)) if pid else q("SELECT id, project_id, dt_id, data FROM records")
    total_records = len(recs)
    issued_date_count = 0
    workflow_records = 0
    status_count = 0
    actual_reply_count = 0
    pr_record_ids = []

    for r in recs:
        data = r["data"] if isinstance(r["data"], dict) else {}
        if data.get("issuedDate"):
            issued_date_count += 1
        dt_key = (r["project_id"], r["dt_id"])
        if dt_key in workflow_dt_keys:
            workflow_records += 1
            if data.get("status"):
                status_count += 1
            if data.get("actualReplyDate"):
                actual_reply_count += 1
        if dt_key in pr_dt_keys:
            pr_record_ids.append(r["id"])

    pr_structured_count = 0
    if pr_record_ids:
        row = q("SELECT COUNT(DISTINCT record_id) as c FROM pr_items WHERE record_id = ANY(%s)", (pr_record_ids,), one=True)
        pr_structured_count = row["c"] if row and row.get("c") is not None else 0

    def pct(part, whole):
        return round(part / whole * 100) if whole else 0

    return {
        "total_records": total_records,
        "issued_date_count": issued_date_count,
        "issued_date_pct": pct(issued_date_count, total_records),
        "workflow_records": workflow_records,
        "status_count": status_count,
        "status_pct": pct(status_count, workflow_records),
        "actual_reply_count": actual_reply_count,
        "actual_reply_pct": pct(actual_reply_count, workflow_records),
        "pr_records": len(pr_record_ids),
        "pr_structured_count": pr_structured_count,
        "pr_structured_pct": pct(pr_structured_count, len(pr_record_ids)),
    }


def get_action_required_summary(pid=None, limit=10, pending_threshold=14):
    from utils import compute_duration

    projects = get_projects()
    project_codes = {p["id"]: p["code"] for p in projects}

    overdue_rows = get_overdue_records(pid)[:limit]
    top_overdue = [
        {**r, "project_code": project_codes.get(r["project_id"], r["project_id"])}
        for r in overdue_rows
    ]

    dt_rows = q("SELECT project_id, id, code, name FROM doc_types WHERE project_id=%s", (pid,)) if pid else q("SELECT project_id, id, code, name FROM doc_types")
    dt_meta = {(d["project_id"], d["id"]): d for d in dt_rows}

    meta_rows = q("SELECT project_id, item_value, meta FROM dropdown_lists WHERE project_id=%s AND list_name LIKE %s AND meta IS NOT NULL", (pid, "status%")) if pid else q("SELECT project_id, item_value, meta FROM dropdown_lists WHERE list_name LIKE %s AND meta IS NOT NULL", ("status%",))
    meta_map = {}
    for r in meta_rows:
        meta_map.setdefault(r["project_id"], {})[r["item_value"]] = r["meta"]

    col_rows = q("""
        SELECT project_id, dt_id, col_key
        FROM columns_config
        WHERE project_id=%s AND visible=TRUE AND col_key IN ('expectedReplyDate','status')
    """, (pid,)) if pid else q("""
        SELECT project_id, dt_id, col_key
        FROM columns_config
        WHERE visible=TRUE AND col_key IN ('expectedReplyDate','status')
    """)
    dt_has_exp = set()
    dt_has_status = set()
    for r in col_rows:
        key = (r["project_id"], r["dt_id"])
        if r["col_key"] == "expectedReplyDate":
            dt_has_exp.add(key)
        elif r["col_key"] == "status":
            dt_has_status.add(key)

    workflow_dt_keys = {
        key for key, dt in dt_meta.items()
        if key in dt_has_exp and key in dt_has_status and not _is_non_workflow_dt(dt.get("code",""), dt.get("name",""))
    }

    rows = q("SELECT id, project_id, dt_id, data, created_at FROM records WHERE project_id=%s ORDER BY created_at DESC", (pid,)) if pid else q("SELECT id, project_id, dt_id, data, created_at FROM records ORDER BY created_at DESC")
    pending_longest = []
    recent_rejected = []

    for row in rows:
        dt_key = (row["project_id"], row["dt_id"])
        if dt_key not in workflow_dt_keys:
            continue
        d = row["data"] if isinstance(row["data"], dict) else {}
        status = d.get("status", "")
        meta = meta_map.get(row["project_id"], {}).get(status) or DEFAULT_STATUS_META.get(status) or "pending"
        if meta == "cancelled":
            continue
        dt = dt_meta.get(dt_key, {})
        base = {
            "project_id": row["project_id"],
            "project_code": project_codes.get(row["project_id"], row["project_id"]),
            "dt_id": row["dt_id"],
            "dt_code": dt.get("code", row["dt_id"]),
            "docNo": d.get("docNo", ""),
            "title": d.get("title", ""),
            "discipline": d.get("discipline", ""),
            "status": status,
            "created_at": str(row.get("created_at") or ""),
        }
        if meta == "pending" and not d.get("actualReplyDate") and d.get("issuedDate"):
            days_delay = compute_duration(d.get("issuedDate"), None) or 0
            if days_delay > pending_threshold:
                pending_longest.append({**base, "issuedDate": d.get("issuedDate", ""), "days_delay": days_delay})
        elif meta == "rejected":
            recent_rejected.append(base)

    pending_longest.sort(key=lambda x: x["days_delay"], reverse=True)
    recent_rejected.sort(key=lambda x: x["created_at"], reverse=True)

    return {
        "top_overdue": top_overdue,
        "pending_longest": pending_longest[:limit],
        "recent_rejected": recent_rejected[:limit],
        "pending_threshold": pending_threshold,
    }


def get_pr_analytics_summary(pid=None, limit=5):
    projects = get_projects()
    project_map = {p["id"]: p for p in projects}

    pr_rows = q("""
        SELECT r.id, r.project_id, r.data, d.code, d.name
        FROM records r
        JOIN doc_types d ON d.id=r.dt_id AND d.project_id=r.project_id
        WHERE r.project_id=%s
          AND (
               UPPER(COALESCE(d.code, ''))='PR'
            OR LOWER(COALESCE(d.name, '')) LIKE %s
            OR LOWER(COALESCE(d.name, '')) LIKE %s
          )
    """, (pid, "%requisition%", "%purchase request%")) if pid else q("""
        SELECT r.id, r.project_id, r.data, d.code, d.name
        FROM records r
        JOIN doc_types d ON d.id=r.dt_id AND d.project_id=r.project_id
        WHERE UPPER(COALESCE(d.code, ''))='PR'
           OR LOWER(COALESCE(d.name, '')) LIKE %s
           OR LOWER(COALESCE(d.name, '')) LIKE %s
    """, ("%requisition%", "%purchase request%"))

    total_pr_records = len(pr_rows)
    if not total_pr_records:
        return {
            "total_pr_records": 0,
            "top_projects": [],
            "top_trades": [],
        }

    project_counts = {}
    trade_counts = {}
    for row in pr_rows:
        row_pid = row["project_id"]
        project_counts[row_pid] = project_counts.get(row_pid, 0) + 1
        data = row["data"] if isinstance(row["data"], dict) else {}
        trade = str(data.get("discipline") or "").strip() or "Unspecified"
        trade_counts[trade] = trade_counts.get(trade, 0) + 1

    top_projects = []
    for row_pid, pr_count in sorted(project_counts.items(), key=lambda x: (-x[1], x[0]))[:limit]:
        proj = project_map.get(row_pid, {})
        top_projects.append({
            "project_id": row_pid,
            "project_code": proj.get("code", row_pid),
            "project_name": proj.get("name", row_pid),
            "pr_count": pr_count,
        })

    top_trades = [
        {"trade": trade, "pr_count": pr_count}
        for trade, pr_count in sorted(trade_counts.items(), key=lambda x: (-x[1], x[0]))[:limit]
    ]

    return {
        "total_pr_records": total_pr_records,
        "top_projects": top_projects,
        "top_trades": top_trades,
    }


def get_monthly_trend(pid=None):
    """Returns last 6 months of submission counts per month."""
    import re as _re
    where = "WHERE project_id=%s" if pid else "WHERE 1=1"
    params = (pid,) if pid else ()
    rows = q(f"SELECT data FROM records {where}", params)
    months = {}
    import datetime
    today = datetime.date.today()
    # Show last 12 months to capture more data
    for i in range(11, -1, -1):
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
    """Returns pending docs grouped by days lapsed — only for types with visible expectedReplyDate."""
    from utils import compute_duration
    where = "WHERE r.project_id=%s" if pid else "WHERE 1=1"
    params = (pid,) if pid else ()
    if pid:
        exp_rows = q(
            "SELECT dt_id, col_key FROM columns_config"
            " WHERE project_id=%s AND col_key IN ('expectedReplyDate','status') AND visible=TRUE",
            (pid,))
    else:
        exp_rows = q(
            "SELECT dt_id, col_key FROM columns_config"
            " WHERE col_key IN ('expectedReplyDate','status') AND visible=TRUE")
    dt_has_exp_a    = set()
    dt_has_status_a = set()
    for r in exp_rows:
        if r["col_key"] == "expectedReplyDate": dt_has_exp_a.add(r["dt_id"])
        elif r["col_key"] == "status":          dt_has_status_a.add(r["dt_id"])
    dt_with_exp = dt_has_exp_a & dt_has_status_a
    dt_rows = q("SELECT id, code, name FROM doc_types WHERE project_id=%s", (pid,)) if pid else q("SELECT id, code, name FROM doc_types")
    dt_with_exp = {dt_id for dt_id in dt_with_exp
                   if not _is_non_workflow_dt(
                       next((d["code"] for d in dt_rows if d["id"] == dt_id), ""),
                       next((d["name"] for d in dt_rows if d["id"] == dt_id), "")
                   )}
    rows = q(f"SELECT r.dt_id, r.data FROM records r {where}", params)
    buckets = {"1-7": 0, "8-14": 0, "15-21": 0, ">21": 0}
    for row in rows:
        if row["dt_id"] not in dt_with_exp: continue
        d = row["data"] if isinstance(row["data"], dict) else {}
        if d.get("actualReplyDate"): continue
        issued = d.get("issuedDate", "")
        if not issued: continue
        days = compute_duration(issued, None) or 0
        if days < 1: continue
        if days <= 7:    buckets["1-7"]   += 1
        elif days <= 14: buckets["8-14"]  += 1
        elif days <= 21: buckets["15-21"] += 1
        else:            buckets[">21"]   += 1
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
    """Returns overdue records — only for doc types with BOTH visible status AND visible expectedReplyDate."""
    from utils import is_overdue
    import datetime as _dt
    where = "WHERE project_id=%s" if pid else "WHERE 1=1"
    params = (pid,) if pid else ()
    rows = q(f"""SELECT r.project_id, r.dt_id, r.data, d.code as dt_code, d.name as dt_name
                 FROM records r
                 JOIN doc_types d ON d.id=r.dt_id AND d.project_id=r.project_id
                 {where.replace("project_id", "r.project_id")}""", params)
    # Only dt_ids that have BOTH visible status AND visible expectedReplyDate
    # (types like Letters/PR that have no status column are excluded)
    if pid:
        exp_rows = q(
            "SELECT dt_id, col_key FROM columns_config"
            " WHERE project_id=%s AND col_key IN ('expectedReplyDate','status') AND visible=TRUE",
            (pid,))
    else:
        exp_rows = q(
            "SELECT dt_id, col_key FROM columns_config"
            " WHERE col_key IN ('expectedReplyDate','status') AND visible=TRUE")
    dt_has_exp    = set()
    dt_has_status = set()
    for r in exp_rows:
        if r["col_key"] == "expectedReplyDate": dt_has_exp.add(r["dt_id"])
        elif r["col_key"] == "status":          dt_has_status.add(r["dt_id"])
    # Must have BOTH to be considered overdue-eligible
    dt_with_exp = dt_has_exp & dt_has_status
    dt_rows = q("SELECT id, code, name FROM doc_types WHERE project_id=%s", (pid,)) if pid else q("SELECT id, code, name FROM doc_types")
    dt_with_exp = {dt_id for dt_id in dt_with_exp
                   if not _is_non_workflow_dt(
                       next((d["code"] for d in dt_rows if d["id"] == dt_id), ""),
                       next((d["name"] for d in dt_rows if d["id"] == dt_id), "")
                   )}
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
                from utils import compute_duration
                days = compute_duration(issued, None) or 0
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


# ── Audit Log ─────────────────────────────────────────────────
def log_action(username, action, project_id=None, dt_id=None, record_id=None,
               doc_no=None, field_name=None, old_value=None, new_value=None, detail=None):
    """Write one audit entry. Never raises — silently ignores errors."""
    try:
        exe("""INSERT INTO audit_log(username,action,project_id,dt_id,record_id,
               doc_no,field_name,old_value,new_value,detail)
               VALUES(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)""",
            (username, action, project_id, dt_id, record_id,
             doc_no, field_name,
             str(old_value)[:500] if old_value is not None else None,
             str(new_value)[:500] if new_value is not None else None,
             str(detail)[:500] if detail else None))
    except Exception as e:
        logger.error("audit_log_failed action=%s project_id=%s dt_id=%s record_id=%s error=%s",
                     action, project_id, dt_id, record_id, e)

def get_audit_log(project_id=None, username=None, action=None,
                  limit=200, offset=0):
    """Fetch audit entries with optional filters."""
    conditions = []
    params = []
    if project_id:
        conditions.append("project_id=%s"); params.append(project_id)
    if username:
        conditions.append("username=%s"); params.append(username)
    if action:
        conditions.append("action=%s"); params.append(action)
    where = ("WHERE " + " AND ".join(conditions)) if conditions else ""
    params += [limit, offset]
    return q(f"""SELECT id,ts,username,action,project_id,dt_id,
                        doc_no,field_name,old_value,new_value,detail
                 FROM audit_log {where}
                 ORDER BY ts DESC
                 LIMIT %s OFFSET %s""", params)

def get_audit_actions():
    """Return distinct action types for filter dropdown."""
    rows = q("SELECT DISTINCT action FROM audit_log ORDER BY action")
    return [r["action"] for r in rows]


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
