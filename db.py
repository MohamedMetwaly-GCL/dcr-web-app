"""db.py - DCR Flask - Clean PostgreSQL database layer"""
import logging
import os, json, uuid, datetime, hashlib, secrets, re
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
    role       TEXT NOT NULL DEFAULT 'viewer',
    email      TEXT
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
    is_dc      BOOLEAN DEFAULT FALSE,
    PRIMARY KEY (username, project_id)
);
CREATE TABLE IF NOT EXISTS project_distribution (
    id SERIAL PRIMARY KEY,
    project_id TEXT NOT NULL,
    doc_type_id TEXT NOT NULL,
    event_type TEXT NOT NULL,
    emails JSONB NOT NULL DEFAULT '[]',
    UNIQUE(project_id, doc_type_id, event_type)
);
CREATE TABLE IF NOT EXISTS notification_queue (
    id SERIAL PRIMARY KEY,
    project_id TEXT NOT NULL,
    recipient_email TEXT NOT NULL,
    subject TEXT NOT NULL,
    body_html TEXT NOT NULL,
    status TEXT DEFAULT 'pending',
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
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
    data       JSONB NOT NULL DEFAULT '{}',
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
    row_type    TEXT NOT NULL DEFAULT 'item',
    item_name   TEXT NOT NULL,
    unit        TEXT,
    quantity    NUMERIC,
    remarks     TEXT,
    sort_order  INTEGER DEFAULT 0,
    created_at  TIMESTAMPTZ DEFAULT NOW(),
    updated_at  TIMESTAMPTZ DEFAULT NOW()
);
ALTER TABLE pr_items ADD COLUMN IF NOT EXISTS row_type TEXT NOT NULL DEFAULT 'item';
ALTER TABLE users ADD COLUMN IF NOT EXISTS email TEXT;
ALTER TABLE user_projects ADD COLUMN IF NOT EXISTS is_dc BOOLEAN DEFAULT FALSE;
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
ALTER TABLE doc_types ADD COLUMN IF NOT EXISTS data JSONB NOT NULL DEFAULT '{}';
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
    ("fileLocation","Drive Link","url",None,12),
]

NOC_STATUS_VALUES = [
    "Accepted & to proceed with Part C",
    "Rejected",
    "Cancelled",
    "Under Review",
    "Information Required",
    "Pending",
]

NOC_COLS = [
    ("docNo", "NOC No", "docno", None, 0, True),
    ("title", "NOC Subject", "text", None, 1, True),
    ("nocDescription", "NOC Description", "text", None, 2, False),
    ("originatingDocument", "Originating Document", "text", None, 3, True),
    ("remarks", "Remarks", "text", None, 4, False),
    ("partAIssueDate", "Part A Issue Date", "date", None, 5, False),
    ("partBReturnDate", "Part B Return Date", "date", None, 6, False),
    ("partBStatus", "Part B Status", "dropdown", "noc_part_b_status", 7, True),
    ("partCIssueDate", "Part C Issue Date", "date", None, 8, True),
    ("submittedCost", "Submitted Cost (EGP)", "number", None, 9, True),
    ("partDReturnDate", "Part D Return Date", "date", None, 10, True),
    ("partDStatus", "Part D Status", "dropdown", "noc_part_d_status", 11, True),
    ("finalApprovedCost", "Final Approved Cost (EGP)", "number", None, 12, True),
    ("voNo", "VO No", "text", None, 13, True),
    ("voIssueDate", "VO Issue Date", "date", None, 14, False),
    ("voBaseValue", "VO Base Value (EGP)", "number", None, 15, False),
    ("voValueWithSIAndVAT", "VO Value (EGP) Including SI & VAT", "number", None, 16, True),
    ("fileLocation", "Drive Link", "url", None, 17, True),
]

LTR_DIRECTION_VALUES = ["Sent", "Received"]

LTR_COLS = [
    ("docNo", "Letter Ref", "docno", None, 0, True),
    ("title", "Subject", "text", None, 1, True),
    ("direction", "Direction", "dropdown", "letter_direction", 2, True),
    ("fromParty", "From Party", "dropdown", "correspondence_parties", 3, True),
    ("toParty", "To Party", "dropdown", "correspondence_parties", 4, True),
    ("issuedDate", "Issue Date", "date", None, 5, True),
    ("receivedDate", "Received Date", "date", None, 6, True),
    ("remarks", "Remarks", "text", None, 7, False),
    ("parentLetterId", "Parent Letter ID", "text", None, 8, False),
    ("parentLetterRef", "Response Ref", "text", None, 9, True),
    ("fileLocation", "Drive Link", "url", None, 10, True),
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

def _normalize_status_text(value):
    return " ".join(str(value or "").strip().lower().split())

def _normalize_meta_value(value):
    text = _normalize_status_text(value)
    if text in ("approved", "rejected", "pending", "cancelled"):
        return text
    if text in ("open", "under review", "not closed"):
        return "pending"
    if text in ("closed", "replied"):
        return "approved"
    return None

def _status_alias_meta_map(pid_meta=None):
    alias_map = {}
    sources = []
    if pid_meta:
        sources.extend(pid_meta.items())
    sources.extend(DEFAULT_STATUS_META.items())
    for raw_key, meta in sources:
        key = str(raw_key or "").strip()
        meta_value = _normalize_meta_value(meta)
        if not key:
            continue
        if not meta_value:
            continue
        alias_map.setdefault(_normalize_status_text(key), meta_value)
        m = re.match(r"^\s*([A-Za-z](?:\s*[,/&+]\s*[A-Za-z])*)\b", key)
        if m:
            lead = re.sub(r"[^A-Za-z]+", "", m.group(1)).lower()
            if lead:
                alias_map.setdefault(lead, meta_value)
                chars = re.findall(r"[A-Za-z]", m.group(1))
                for ch in chars:
                    alias_map.setdefault(ch.lower(), meta_value)
    return alias_map

def _status_meta_from_tokens(status_value, pid_meta=None):
    text = _normalize_status_text(status_value)
    if not text:
        return None
    alias_map = _status_alias_meta_map(pid_meta)
    token_texts = [t.strip() for t in re.split(r"\s*(?:&|,|/|\+|\band\b)\s*", text) if t.strip()]
    metas = []
    for token in token_texts:
        compact = re.sub(r"[^a-z0-9]+", "", token)
        meta = alias_map.get(token) or alias_map.get(compact)
        if meta:
            metas.append(meta)
    if not metas:
        return None
    metas = set(metas)
    for meta in ("cancelled", "rejected", "approved", "pending"):
        if meta in metas:
            return meta
    return None

import functools

@functools.lru_cache(maxsize=4096)
def _resolve_status_meta_cached(status_value, pid_meta_tuple):
    pid_meta = dict(pid_meta_tuple) if pid_meta_tuple else None
    status = str(status_value or "").strip()
    if not status:
        return "pending"
    direct_meta = _normalize_meta_value(pid_meta[status]) if pid_meta and status in pid_meta else None
    if direct_meta:
        return direct_meta
    default_meta = _normalize_meta_value(DEFAULT_STATUS_META.get(status))
    if default_meta:
        return default_meta
    normalized = _normalize_status_text(status)
    if pid_meta:
        for key, meta in pid_meta.items():
            if _normalize_status_text(key) == normalized:
                normalized_meta = _normalize_meta_value(meta)
                if normalized_meta:
                    return normalized_meta
    for key, meta in DEFAULT_STATUS_META.items():
        if _normalize_status_text(key) == normalized:
            normalized_meta = _normalize_meta_value(meta)
            if normalized_meta:
                return normalized_meta
    try:
        token_meta = _status_meta_from_tokens(status, pid_meta)
    except Exception as e:
        logger.warning("status_token_resolution_failed status=%r error=%s", status, e)
        token_meta = None
    if token_meta:
        return token_meta
    return "pending"

def resolve_status_meta(status_value, pid_meta=None):
    pid_meta_tuple = tuple(sorted(pid_meta.items())) if pid_meta else None
    return _resolve_status_meta_cached(status_value, pid_meta_tuple)

def _is_non_workflow_dt(dt_code="", dt_name=""):
    code = str(dt_code or "").strip().lower()
    name = str(dt_name or "").strip().lower()
    return code in ("pr", "ltr", "si", "ei") or "requisition" in name or "letter" in name or "site instruction" in name or "engineer instruction" in name

def _is_pr_dt(dt_code="", dt_name=""):
    code = str(dt_code or "").strip().lower()
    name = str(dt_name or "").strip().lower()
    return code == "pr" or "requisition" in name or "purchase request" in name

def _is_noc_dt(dt_code="", dt_name=""):
    code = str(dt_code or "").strip().lower()
    name = str(dt_name or "").strip().lower()
    return code == "noc" or "notice of change" in name

def _is_ltr_dt(dt_code="", dt_name=""):
    code = str(dt_code or "").strip().lower()
    name = str(dt_name or "").strip().lower()
    return code == "ltr" or "letter" in name or "correspondence" in name

def _doc_type_col_specs(dt_code="", dt_name=""):
    if _is_noc_dt(dt_code, dt_name):
        return NOC_COLS
    if _is_ltr_dt(dt_code, dt_name):
        return LTR_COLS
    return [(ck, lbl, ct, ln, so, True) for ck, lbl, ct, ln, so in DEFAULT_COLS]

def _ensure_noc_lists(pid):
    for ln in ("noc_part_b_status", "noc_part_d_status"):
        for i, item in enumerate(NOC_STATUS_VALUES):
            meta = DEFAULT_STATUS_META.get(item, "pending" if item in ("Under Review", "Information Required", "Pending") else ("cancelled" if item == "Cancelled" else ("approved" if item == "Accepted & to proceed with Part C" else "rejected")))
            exe("INSERT INTO dropdown_lists(project_id,list_name,item_value,sort_order,meta)"
                " VALUES(%s,%s,%s,%s,%s) ON CONFLICT(project_id,list_name,item_value)"
                " DO UPDATE SET sort_order=EXCLUDED.sort_order, meta=EXCLUDED.meta",
                (pid, ln, item, i, meta))

def _ensure_ltr_lists(pid):
    for i, item in enumerate(LTR_DIRECTION_VALUES):
        exe("INSERT INTO dropdown_lists(project_id,list_name,item_value,sort_order,meta)"
            " VALUES(%s,%s,%s,%s,%s) ON CONFLICT(project_id,list_name,item_value)"
            " DO UPDATE SET sort_order=EXCLUDED.sort_order",
            (pid, "letter_direction", item, i, None))

def _sync_noc_doc_type(pid, dt_id="NOC", dt_name="Notice of Change"):
    exe("INSERT INTO doc_types(id,project_id,name,code,sort_order)"
        " VALUES(%s,%s,%s,%s,COALESCE((SELECT sort_order FROM doc_types WHERE id=%s AND project_id=%s),14))"
        " ON CONFLICT(id,project_id) DO UPDATE SET name=EXCLUDED.name, code=EXCLUDED.code",
        (dt_id, pid, dt_name, "NOC", dt_id, pid))
    specs = _doc_type_col_specs("NOC", dt_name)
    keep_keys = {ck for ck, *_ in specs}
    for ck, lbl, ct, ln, so, visible in specs:
        exe("""
            INSERT INTO columns_config(project_id,dt_id,col_key,label,col_type,list_name,visible,sort_order)
            VALUES(%s,%s,%s,%s,%s,%s,%s,%s)
            ON CONFLICT(project_id,dt_id,col_key) DO UPDATE
            SET label=EXCLUDED.label,
                col_type=EXCLUDED.col_type,
                list_name=EXCLUDED.list_name,
                sort_order=COALESCE(columns_config.sort_order, EXCLUDED.sort_order)
        """, (pid, dt_id, ck, lbl, ct, ln, visible, so))
    _ensure_noc_lists(pid)

def _sync_ltr_doc_type(pid, dt_id="LTR", dt_name="Letters"):
    specs = _doc_type_col_specs("LTR", dt_name)
    blocked_keys = ["description", "status"]
    exe(
        "DELETE FROM columns_config WHERE project_id=%s AND dt_id=%s AND col_key = ANY(%s)",
        (pid, dt_id, blocked_keys),
    )
    for ck, lbl, ct, ln, so, visible in specs:
        exe("""
            INSERT INTO columns_config(project_id,dt_id,col_key,label,col_type,list_name,visible,sort_order)
            VALUES(%s,%s,%s,%s,%s,%s,%s,%s)
            ON CONFLICT(project_id,dt_id,col_key) DO UPDATE
            SET col_type=EXCLUDED.col_type,
                list_name=EXCLUDED.list_name,
                sort_order=COALESCE(columns_config.sort_order, EXCLUDED.sort_order),
                visible=CASE
                    WHEN EXCLUDED.col_key='fileLocation' THEN TRUE
                    WHEN columns_config.col_key='parentLetterId' THEN FALSE
                    ELSE columns_config.visible
                END
        """, (pid, dt_id, ck, lbl, ct, ln, visible, so))
    _ensure_ltr_lists(pid)

def _sync_all_noc_doc_types():
    projects = q("SELECT id FROM projects")
    for p in projects:
        pid = p["id"]
        for dt in q("SELECT id, name, code FROM doc_types WHERE project_id=%s", (pid,)):
            if _is_noc_dt(dt["code"], dt["name"]):
                _sync_noc_doc_type(pid, dt["id"], dt.get("name", "Notice of Change"))

def _sync_all_ltr_doc_types():
    projects = q("SELECT id FROM projects")
    for p in projects:
        pid = p["id"]
        for dt in q("SELECT id, name, code FROM doc_types WHERE project_id=%s", (pid,)):
            if _is_ltr_dt(dt["code"], dt["name"]):
                _sync_ltr_doc_type(pid, dt["id"], dt.get("name", "Letters"))

def _cleanup_custom_columns():
    exe("DELETE FROM columns_config WHERE col_key LIKE %s AND (label ILIKE %s OR label ILIKE %s) AND dt_id IN (SELECT id FROM doc_types WHERE UPPER(code)='LTR' OR name ILIKE %s OR name ILIKE %s)",
        ('custom_%', '%file location%', '%drive link%', '%letter%', '%correspondence%'))

def _global_rename_filelocation():
    exe("UPDATE columns_config SET label='Drive Link', col_type='url' WHERE col_key='fileLocation'")

def init():
    stmts = [s.strip() for s in SCHEMA.strip().split(";") if s.strip()]
    for stmt in stmts:
        try:
            with DB() as cur: cur.execute(stmt)
        except Exception as e:
            if "already exists" not in str(e).lower():
                logger.warning("db_schema_note error=%s", e)
    _ensure_admin()
    _cleanup_custom_columns()
    _global_rename_filelocation()
    _sync_all_noc_doc_types()
    _sync_all_ltr_doc_types()

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
    return q("SELECT username, role, email FROM users ORDER BY username")

def add_user(username, pw, role="viewer", email=None):
    exe("INSERT INTO users(username,pw_hash,role,email) VALUES(%s,%s,%s,%s) ON CONFLICT DO NOTHING",
        (username, hash_pw(pw), role, email))

def delete_user(username):
    exe("DELETE FROM sessions WHERE username=%s", (username,))
    exe("DELETE FROM user_projects WHERE username=%s", (username,))
    exe("DELETE FROM users WHERE username=%s", (username,))

def change_pw(username, new_pw):
    exe("UPDATE users SET pw_hash=%s WHERE username=%s", (hash_pw(new_pw), username))

def set_user_role(username, role):
    role = str(role or "").strip().lower()
    if role not in ("viewer", "editor", "admin", "superadmin"):
        return False
    exe("UPDATE users SET role=%s WHERE username=%s", (role, username))
    exe("UPDATE sessions SET role=%s WHERE username=%s", (role, username))
    return True

def set_user_email(username, email):
    exe("UPDATE users SET email=%s WHERE username=%s", (email, username))

def set_user_project_dc(username, project_id, is_dc):
    exe("UPDATE user_projects SET is_dc=%s WHERE username=%s AND project_id=%s", (is_dc, username, project_id))

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
    # Returns a list of dicts: {"project_id": "...", "is_dc": True/False}
    return [{"project_id": r["project_id"], "is_dc": r.get("is_dc", False)} for r in
            q("SELECT project_id, is_dc FROM user_projects WHERE username=%s", (username,))]

def assign_project(username, project_id, is_dc=False):
    exe("INSERT INTO user_projects(username,project_id,is_dc) VALUES(%s,%s,%s) ON CONFLICT (username,project_id) DO UPDATE SET is_dc=EXCLUDED.is_dc",
        (username, project_id, is_dc))

def unassign_project(username, project_id):
    exe("DELETE FROM user_projects WHERE username=%s AND project_id=%s", (username, project_id))

def get_project_users(project_id):
    return [{"username": r["username"], "is_dc": r.get("is_dc", False)} for r in
            q("SELECT username, is_dc FROM user_projects WHERE project_id=%s", (project_id,))]

def get_project_dc(project_id):
    """Return the DC user record (with email) for a given project, or None."""
    rows = q(
        """SELECT u.username, u.email
           FROM users u
           JOIN user_projects up ON u.username = up.username
           WHERE up.project_id = %s AND up.is_dc = TRUE
           LIMIT 1""",
        (project_id,), one=True
    )
    return rows

def user_is_dc(username, project_id):
    """Return True if the given user is the DC for the project."""
    r = q("SELECT is_dc FROM user_projects WHERE username=%s AND project_id=%s",
          (username, project_id), one=True)
    return bool(r and r.get("is_dc"))

def get_project_users_full(project_id):
    """Return all users assigned to the project PLUS all admins/superadmins."""
    rows = q("""
        SELECT username, role 
        FROM users 
        WHERE role IN ('admin', 'superadmin')
        UNION
        SELECT u.username, u.role
        FROM user_projects up
        JOIN users u ON u.username = up.username
        WHERE up.project_id = %s
        ORDER BY username
    """, (project_id,))
    return [dict(r) for r in rows]

def get_daily_digest(pid, doc_type_ids):
    from datetime import date
    today_str = date.today().isoformat()
    if not doc_type_ids: return {"received": [], "issued": [], "replied": []}
    rows = q("SELECT id, dt_id, data FROM records WHERE project_id=%s AND dt_id = ANY(%s)", (pid, list(doc_type_ids)))
    received, issued, replied = [], [], []
    for r in rows:
        d = r.get("data") or {}
        if d.get("receivedDate") == today_str: received.append(r)
        if d.get("issuedDate") == today_str or d.get("partAIssueDate") == today_str or d.get("partCIssueDate") == today_str: issued.append(r)
        if d.get("actualReplyDate") == today_str or d.get("partBReturnDate") == today_str or d.get("partDReturnDate") == today_str: replied.append(r)
    def format_rec(rec):
        d = rec.get("data") or {}
        return {
            "id": rec["id"], "dt_id": rec["dt_id"],
            "docNo": d.get("docNo") or d.get("record_id") or "Untitled",
            "title": d.get("title") or d.get("subject") or d.get("nocDescription") or "No Subject",
            "status": d.get("status") or d.get("partDStatus") or d.get("partBStatus") or ""
        }
    return {
        "received": [format_rec(x) for x in received],
        "issued": [format_rec(x) for x in issued],
        "replied": [format_rec(x) for x in replied],
    }

# ── Distribution Matrix ────────────────────────────────────────
def get_distribution(project_id):
    """Return all distribution rows for a project as a nested dict:
       {doc_type_id: {event_type: [emails]}}"""
    rows = q("SELECT doc_type_id, event_type, emails FROM project_distribution WHERE project_id=%s",
             (project_id,))
    result = {}
    for r in rows:
        result.setdefault(r["doc_type_id"], {})[r["event_type"]] = r["emails"]
    return result

def upsert_distribution(project_id, doc_type_id, event_type, emails):
    """Insert or update a distribution row."""
    exe("""INSERT INTO project_distribution(project_id, doc_type_id, event_type, emails)
           VALUES(%s,%s,%s,%s)
           ON CONFLICT(project_id, doc_type_id, event_type)
           DO UPDATE SET emails=EXCLUDED.emails""",
        (project_id, doc_type_id, event_type, json.dumps(emails)))

# ── Notification Queue ─────────────────────────────────────────
def enqueue_notification(project_id, recipient_email, subject, body_html):
    """Add an email to the notification queue."""
    exe("""INSERT INTO notification_queue(project_id, recipient_email, subject, body_html)
           VALUES(%s,%s,%s,%s)""",
        (project_id, recipient_email, subject, body_html))

def get_pending_notifications(limit=50):
    """Return pending notifications for the background worker to send."""
    return q("SELECT * FROM notification_queue WHERE status='pending' ORDER BY created_at LIMIT %s",
             (limit,))

def mark_notification_sent(nid):
    exe("UPDATE notification_queue SET status='sent' WHERE id=%s", (nid,))

def mark_notification_failed(nid):
    exe("UPDATE notification_queue SET status='failed' WHERE id=%s", (nid,))

def _clean_project_ids(project_ids):
    if project_ids is None:
        return None
    seen = set()
    cleaned = []
    for pid in project_ids:
        val = str(pid or "").strip()
        if not val or val in seen:
            continue
        seen.add(val)
        cleaned.append(val)
    return cleaned

def _project_scope_clause(column, pid=None, project_ids=None):
    if pid:
        return f"WHERE {column}=%s", [pid]
    ids = _clean_project_ids(project_ids)
    if ids is None:
        return "WHERE 1=1", []
    if not ids:
        return "WHERE 1=0", []
    ph = ",".join(["%s"] * len(ids))
    return f"WHERE {column} IN ({ph})", ids

# ── Projects ──────────────────────────────────────────────────
def get_projects(project_ids=None):
    order_sql = """ORDER BY
        CASE WHEN (data->>'dashboard_order') ~ '^-?[0-9]+$' THEN (data->>'dashboard_order')::int ELSE 999999 END,
        id"""
    ids = _clean_project_ids(project_ids)
    if ids is None:
        return q(f"SELECT * FROM projects {order_sql}")
    if not ids:
        return []
    ph = ",".join(["%s"] * len(ids))
    return q(f"SELECT * FROM projects WHERE id IN ({ph}) {order_sql}", ids)


def reorder_projects(order):
    for i, pid in enumerate([str(v).strip() for v in (order or []) if str(v).strip()]):
        r = q("SELECT data FROM projects WHERE id=%s", (pid,), one=True)
        if not r:
            continue
        data = r.get("data") or {}
        if isinstance(data, str):
            try: data = json.loads(data)
            except Exception: data = {}
        if not isinstance(data, dict):
            data = {}
        data["dashboard_order"] = i
        exe("UPDATE projects SET data=%s::jsonb WHERE id=%s", (json.dumps(data), pid))

def get_projects_for_user(username, role=None):
    user_role = str(role or "").strip().lower()
    if user_role in ("superadmin", "admin"):
        return get_projects()
    return get_projects(get_user_projects(username))

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

def get_expected_reply_rule(pid, dt_id=None):
    """Read the project rule and optionally apply doc-type day overrides."""
    try:
        from utils import normalize_expected_reply_rule, apply_doc_type_expected_reply_override
        project = get_project(pid) or {}
        rule = normalize_expected_reply_rule(project.get("expected_reply_rule"))
        if dt_id:
            override = get_doc_type_expected_reply_override(pid, dt_id)
            return apply_doc_type_expected_reply_override(rule, override)
        return rule
    except Exception:
        from utils import DEFAULT_EXPECTED_REPLY_RULE
        return dict(DEFAULT_EXPECTED_REPLY_RULE)

def save_project(pid, name, code, extra: dict):
    if isinstance(extra, dict):
        extra.pop("id", None); extra.pop("name", None); extra.pop("code", None)
    extra = extra or {}
    current = q("SELECT data FROM projects WHERE id=%s", (pid,), one=True)
    current_data = current.get("data") if current else {}
    if isinstance(current_data, str):
        try: current_data = json.loads(current_data)
        except Exception: current_data = {}
    if isinstance(current_data, dict) and "dashboard_order" in current_data and "dashboard_order" not in extra:
        extra["dashboard_order"] = current_data.get("dashboard_order")
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
        for ck, lbl, ct, ln, so, visible in _doc_type_col_specs(code, name):
            exe("""INSERT INTO columns_config(project_id,dt_id,col_key,label,col_type,list_name,visible,sort_order)
                   VALUES(%s,%s,%s,%s,%s,%s,%s,%s) ON CONFLICT DO NOTHING""",
                (pid, code, ck, lbl, ct, ln, visible, so))
    for ln, items in DEFAULT_LISTS.items():
        for i, item in enumerate(items):
            meta = DEFAULT_STATUS_META.get(item) if ln == "status" else None
            exe("INSERT INTO dropdown_lists(project_id,list_name,item_value,sort_order,meta)"
                " VALUES(%s,%s,%s,%s,%s) ON CONFLICT(project_id,list_name,item_value)"
                " DO UPDATE SET meta=COALESCE(dropdown_lists.meta, EXCLUDED.meta)",
                (pid, ln, item, i, meta))
    _ensure_noc_lists(pid)
    _ensure_ltr_lists(pid)
    _sync_all_noc_doc_types()
    _sync_all_ltr_doc_types()

# ── Logos ─────────────────────────────────────────────────────
def get_logo(pid, key):
    r = q("SELECT image_data FROM logos WHERE project_id=%s AND logo_key=%s", (pid,key), one=True)
    return r["image_data"] if r else None

def save_logo(pid, key, data):
    exe("""INSERT INTO logos(project_id,logo_key,image_data) VALUES(%s,%s,%s)
           ON CONFLICT(project_id,logo_key) DO UPDATE SET image_data=EXCLUDED.image_data""",
        (pid, key, data))

# ── Doc Types ─────────────────────────────────────────────────
def _normalize_doc_type_row(row):
    if not row:
        return None
    data = row.get("data") or {}
    if isinstance(data, str):
        try: data = json.loads(data)
        except Exception: data = {}
    if not isinstance(data, dict): data = {}
    data.pop("id", None); data.pop("project_id", None); data.pop("name", None); data.pop("code", None); data.pop("sort_order", None)
    return {"id": row["id"], "project_id": row["project_id"], "name": row["name"], "code": row["code"], "sort_order": row.get("sort_order", 0), **data}

def get_doc_types(pid):
    rows = q("SELECT * FROM doc_types WHERE project_id=%s ORDER BY sort_order, id", (pid,))
    return [_normalize_doc_type_row(r) for r in rows]

def get_doc_type(pid, dt_id):
    return _normalize_doc_type_row(q("SELECT * FROM doc_types WHERE id=%s AND project_id=%s", (dt_id, pid), one=True))

def get_doc_type_expected_reply_override(pid, dt_id):
    try:
        from utils import normalize_doc_type_expected_reply_override
        dt = get_doc_type(pid, dt_id) or {}
        return normalize_doc_type_expected_reply_override(dt.get("expected_reply_override"))
    except Exception:
        from utils import DEFAULT_DOC_TYPE_EXPECTED_REPLY_OVERRIDE
        return dict(DEFAULT_DOC_TYPE_EXPECTED_REPLY_OVERRIDE)

def save_doc_type_expected_reply_override(pid, dt_id, override):
    dt = get_doc_type(pid, dt_id) or {}
    data = {k: v for k, v in dt.items() if k not in ("id", "project_id", "name", "code", "sort_order")}
    data["expected_reply_override"] = override or {}
    exe("UPDATE doc_types SET data=%s::jsonb WHERE id=%s AND project_id=%s", (json.dumps(data), dt_id, pid))

def add_doc_type(pid, code, name, expected_reply_override=None):
    r = q("SELECT COALESCE(MAX(sort_order),0)+1 as n FROM doc_types WHERE project_id=%s", (pid,), one=True)
    order = r["n"] if r else 0
    data = {"expected_reply_override": expected_reply_override or {}}
    exe("INSERT INTO doc_types(id,project_id,name,code,sort_order,data) VALUES(%s,%s,%s,%s,%s,%s::jsonb) ON CONFLICT DO NOTHING",
        (code, pid, name, code, order, json.dumps(data)))
    for ck, lbl, ct, ln, so, visible in _doc_type_col_specs(code, name):
        exe("""INSERT INTO columns_config(project_id,dt_id,col_key,label,col_type,list_name,visible,sort_order)
               VALUES(%s,%s,%s,%s,%s,%s,%s,%s) ON CONFLICT DO NOTHING""",
            (pid, code, ck, lbl, ct, ln, visible, so))
    if _is_noc_dt(code, name):
        _sync_noc_doc_type(pid, code, name)
    if _is_ltr_dt(code, name):
        _sync_ltr_doc_type(pid, code, name)

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

def get_record_by_doc_no(pid, dt_id, doc_no):
    norm = str(doc_no or "").strip()
    if not norm:
        return None
    r = q("""
        SELECT id, project_id, dt_id, data, updated_at
        FROM records
        WHERE project_id=%s
          AND dt_id=%s
          AND LOWER(BTRIM(COALESCE(data->>'docNo',''))) = LOWER(BTRIM(%s))
        ORDER BY created_at
        LIMIT 1
    """, (pid, dt_id, norm), one=True)
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

def _norm_ltr_text(v):
    return "".join(ch for ch in str(v or "").strip().lower() if ch.isalnum())

def _get_ltr_dt_id(pid):
    dt = q("""
        SELECT id, name, code
        FROM doc_types
        WHERE project_id=%s
        ORDER BY sort_order, id
    """, (pid,))
    match = next((d for d in dt if _is_ltr_dt(d.get("code",""), d.get("name",""))), None)
    return match["id"] if match else None

def _ltr_field_key_from_cols(cols, role):
    wanted = {
        "docNo": {"docno", "letterref"},
        "title": {"title", "subject"},
        "direction": {"direction"},
        "fromParty": {"fromparty", "from"},
        "toParty": {"toparty", "to"},
        "issuedDate": {"issueddate"},
        "receivedDate": {"receiveddate"},
        "parentLetterId": {"parentletterid"},
    }.get(role, set())
    for c in cols or []:
        key = _norm_ltr_text(c.get("col_key"))
        label = _norm_ltr_text(c.get("label"))
        if key in wanted or label in wanted:
            return c.get("col_key")
    return role

def get_letter_parent_options(pid, exclude_id=None):
    dt_id = _get_ltr_dt_id(pid)
    if not dt_id:
        return []
    cols = get_columns(pid, dt_id)
    doc_key = _ltr_field_key_from_cols(cols, "docNo")
    subject_key = _ltr_field_key_from_cols(cols, "title")
    direction_key = _ltr_field_key_from_cols(cols, "direction")
    from_key = _ltr_field_key_from_cols(cols, "fromParty")
    to_key = _ltr_field_key_from_cols(cols, "toParty")
    issued_key = _ltr_field_key_from_cols(cols, "issuedDate")
    received_key = _ltr_field_key_from_cols(cols, "receivedDate")
    params = [pid, dt_id]
    sql = """
        SELECT id, data, created_at
        FROM records
        WHERE project_id=%s AND dt_id=%s
    """
    if exclude_id:
        sql += " AND id<>%s"
        params.append(exclude_id)
    sql += " ORDER BY created_at DESC, id DESC LIMIT 200"
    rows = q(sql, params)
    out = []
    for row in rows:
        data = row["data"] if isinstance(row.get("data"), dict) else {}
        issued = str(data.get(issued_key) or "").strip()
        received = str(data.get(received_key) or "").strip()
        out.append({
            "id": row["id"],
            "doc_no": str(data.get(doc_key) or "").strip(),
            "subject": str(data.get(subject_key) or "").strip(),
            "direction": str(data.get(direction_key) or "").strip(),
            "from_party": str(data.get(from_key) or "").strip(),
            "to_party": str(data.get(to_key) or "").strip(),
            "date": issued or received,
        })
    return out

def _letter_sort_key(node):
    data = node.get("data") if isinstance(node, dict) else {}
    issued = str(data.get("issuedDate") or data.get("partAIssueDate") or "").strip()
    received = str(data.get("receivedDate") or "").strip()
    created = node.get("created_at")
    created_txt = ""
    if created:
        created_txt = created.isoformat() if hasattr(created, "isoformat") else str(created)
    primary = issued or received or created_txt or ""
    has_explicit_date = 0 if (issued or received) else 1
    doc_no = str(data.get("docNo") or "").strip()
    return (has_explicit_date, primary, created_txt, doc_no, str(node.get("id") or ""))

def _get_letter_family_context(pid):
    dt_id = _get_ltr_dt_id(pid)
    if not dt_id:
        return None
    cols = get_columns(pid, dt_id)
    keys = {
        "docNo": _ltr_field_key_from_cols(cols, "docNo"),
        "title": _ltr_field_key_from_cols(cols, "title"),
        "direction": _ltr_field_key_from_cols(cols, "direction"),
        "fromParty": _ltr_field_key_from_cols(cols, "fromParty"),
        "toParty": _ltr_field_key_from_cols(cols, "toParty"),
        "issuedDate": _ltr_field_key_from_cols(cols, "issuedDate"),
        "receivedDate": _ltr_field_key_from_cols(cols, "receivedDate"),
        "parentLetterId": _ltr_field_key_from_cols(cols, "parentLetterId"),
    }
    rows = q("""
        SELECT id, data, created_at
        FROM records
        WHERE project_id=%s AND dt_id=%s
        ORDER BY created_at, id
    """, (pid, dt_id))
    nodes = {}
    parent_map = {}
    children = {}
    for row in rows:
        data = row["data"] if isinstance(row.get("data"), dict) else {}
        rec_id = row["id"]
        node = {
            "id": rec_id,
            "data": data,
            "created_at": row.get("created_at"),
        }
        nodes[rec_id] = node
        parent_id = str(data.get(keys["parentLetterId"]) or "").strip()
        if parent_id and parent_id != rec_id:
            parent_map[rec_id] = parent_id
    for child_id, parent_id in parent_map.items():
        if parent_id in nodes:
            children.setdefault(parent_id, []).append(child_id)
    for parent_id, ids in children.items():
        ids.sort(key=lambda rec_id: _letter_sort_key(nodes.get(rec_id, {})))
    return {
        "dt_id": dt_id,
        "keys": keys,
        "nodes": nodes,
        "parent_map": parent_map,
        "children": children,
    }

def _get_letter_thread_root(record_id, nodes, parent_map):
    root_id = record_id
    seen = set()
    while root_id in parent_map and parent_map[root_id] in nodes and root_id not in seen:
        seen.add(root_id)
        root_id = parent_map[root_id]
    return root_id

def _build_letter_payload(context, rec_id, level=0, current_id=None):
    node = context["nodes"][rec_id]
    data = node["data"]
    keys = context["keys"]
    issued = str(data.get(keys["issuedDate"]) or "").strip()
    received = str(data.get(keys["receivedDate"]) or "").strip()
    created = node.get("created_at")
    created_txt = created.isoformat() if hasattr(created, "isoformat") else str(created or "")
    return {
        "id": rec_id,
        "parent_id": context["parent_map"].get(rec_id, ""),
        "doc_no": str(data.get(keys["docNo"]) or "").strip(),
        "subject": str(data.get(keys["title"]) or "").strip(),
        "direction": str(data.get(keys["direction"]) or "").strip(),
        "from_party": str(data.get(keys["fromParty"]) or "").strip(),
        "to_party": str(data.get(keys["toParty"]) or "").strip(),
        "date": issued or received,
        "created_at": created_txt,
        "level": level,
        "is_current": rec_id == current_id,
    }

def get_letter_thread(pid, record_id):
    context = _get_letter_family_context(pid)
    if not context:
        return None
    if record_id not in context["nodes"]:
        return None
    root_id = _get_letter_thread_root(record_id, context["nodes"], context["parent_map"])
    items = []
    def _walk(rec_id, level, seen_ids):
        if rec_id in seen_ids:
            return
        seen_ids.add(rec_id)
        items.append(_build_letter_payload(context, rec_id, level, record_id))
        for child_id in context["children"].get(rec_id, []):
            _walk(child_id, level + 1, seen_ids)
    _walk(root_id, 0, set())
    return {
        "root_id": root_id,
        "current_id": record_id,
        "items": items,
    }

def get_letter_timeline(pid, record_id):
    context = _get_letter_family_context(pid)
    if not context:
        return None
    if record_id not in context["nodes"]:
        return None
    root_id = _get_letter_thread_root(record_id, context["nodes"], context["parent_map"])
    family_ids = []
    def _collect(rec_id, seen_ids):
        if rec_id in seen_ids:
            return
        seen_ids.add(rec_id)
        family_ids.append(rec_id)
        for child_id in context["children"].get(rec_id, []):
            _collect(child_id, seen_ids)
    _collect(root_id, set())
    items = [_build_letter_payload(context, rec_id, 0, record_id) for rec_id in family_ids]
    items.sort(key=lambda item: (
        0 if item.get("date") else 1,
        str(item.get("date") or item.get("created_at") or ""),
        str(item.get("created_at") or ""),
        str(item.get("doc_no") or ""),
        str(item.get("id") or ""),
    ))
    return {
        "root_id": root_id,
        "current_id": record_id,
        "items": items,
    }

def _get_ltr_dashboard_stats_map(project_ids=None):
    projects = get_projects(project_ids)
    if not projects:
        return {}
    pids = [p["id"] for p in projects]
    stats_map = {
        pid: {"total": 0, "sent": 0, "received": 0, "party_stats": [], "top_parties": []}
        for pid in pids
    }
    ph = ",".join(["%s"] * len(pids))
    dt_rows = q(f"""
        SELECT project_id, id, code, name
        FROM doc_types
        WHERE project_id IN ({ph})
    """, pids)
    ltr_dt_map = {}
    key_map = {}
    for dt in dt_rows:
        if not _is_ltr_dt(dt.get("code", ""), dt.get("name", "")):
            continue
        pid = dt["project_id"]
        ltr_dt_map[pid] = dt["id"]
        cols = get_columns(pid, dt["id"])
        key_map[pid] = {
            "direction": _ltr_field_key_from_cols(cols, "direction"),
            "fromParty": _ltr_field_key_from_cols(cols, "fromParty"),
            "toParty": _ltr_field_key_from_cols(cols, "toParty"),
        }
    if not ltr_dt_map:
        return stats_map
    dt_ids = list(ltr_dt_map.values())
    dt_ph = ",".join(["%s"] * len(dt_ids))
    rows = q(f"""
        SELECT id, project_id, dt_id, data
        FROM records
        WHERE project_id IN ({ph})
          AND dt_id IN ({dt_ph})
    """, pids + dt_ids)
    party_maps = {pid: {} for pid in pids}
    for row in rows:
        pid = row["project_id"]
        if ltr_dt_map.get(pid) != row["dt_id"]:
            continue
        data = row["data"] if isinstance(row.get("data"), dict) else {}
        keys = key_map.get(pid, {})
        direction = str(data.get(keys.get("direction")) or "").strip().lower()
        stats = stats_map.setdefault(pid, {"total": 0, "sent": 0, "received": 0, "party_stats": [], "top_parties": []})
        stats["total"] += 1
        if direction == "sent":
            stats["sent"] += 1
        elif direction == "received":
            stats["received"] += 1
        parties = {
            str(data.get(keys.get("fromParty")) or "").strip(),
            str(data.get(keys.get("toParty")) or "").strip(),
        }
        parties = {party for party in parties if party}
        for party in parties:
            pstats = party_maps.setdefault(pid, {}).setdefault(party, {"party": party, "total": 0, "sent": 0, "received": 0})
            pstats["total"] += 1
            if direction == "sent":
                pstats["sent"] += 1
            elif direction == "received":
                pstats["received"] += 1
    for pid, party_map in party_maps.items():
        ranked = sorted(
            party_map.values(),
            key=lambda item: (-item["total"], -item["sent"], -item["received"], item["party"].lower())
        )
        stats_map[pid]["party_stats"] = ranked
        stats_map[pid]["top_parties"] = ranked[:5]
    return stats_map

def get_ltr_dashboard_stats(pid=None, project_ids=None):
    if pid:
        return _get_ltr_dashboard_stats_map([pid]).get(pid, {
            "total": 0,
            "sent": 0,
            "received": 0,
            "party_stats": [],
            "top_parties": [],
        })
    stats_map = _get_ltr_dashboard_stats_map(project_ids)
    party_totals = {}
    out = {"total": 0, "sent": 0, "received": 0, "party_stats": [], "top_parties": []}
    for stats in stats_map.values():
        out["total"] += int(stats.get("total", 0) or 0)
        out["sent"] += int(stats.get("sent", 0) or 0)
        out["received"] += int(stats.get("received", 0) or 0)
        for party_row in stats.get("party_stats", []):
            party = str(party_row.get("party") or "").strip()
            if not party:
                continue
            bucket = party_totals.setdefault(party, {"party": party, "total": 0, "sent": 0, "received": 0})
            bucket["total"] += int(party_row.get("total", 0) or 0)
            bucket["sent"] += int(party_row.get("sent", 0) or 0)
            bucket["received"] += int(party_row.get("received", 0) or 0)
    ranked = sorted(
        party_totals.values(),
        key=lambda item: (-item["total"], -item["sent"], -item["received"], item["party"].lower())
    )
    out["party_stats"] = ranked
    out["top_parties"] = ranked[:5]
    return out

def merge_record_data(existing_data: dict, incoming_data: dict):
    merged = {k: v for k, v in (existing_data or {}).items() if not str(k).startswith("_")}
    for key, value in (incoming_data or {}).items():
        if str(key).startswith("_"):
            continue
        if value is None:
            continue
        if isinstance(value, str):
            if not value.strip():
                continue
            merged[key] = value
            continue
        if str(value).strip():
            merged[key] = value
    return merged

def delete_record(rec_id):
    exe("DELETE FROM records WHERE id=%s", (rec_id,))


def get_records_meta(record_ids):
    record_ids = [r for r in (record_ids or []) if r]
    if not record_ids:
        return []
    return q("""
        SELECT id, project_id, dt_id, COALESCE(data->>'docNo','') AS doc_no
        FROM records
        WHERE id = ANY(%s)
    """, (record_ids,))


def delete_records_bulk(record_ids):
    record_ids = [r for r in (record_ids or []) if r]
    if not record_ids:
        return 0
    with DB() as cur:
        cur.execute("DELETE FROM records WHERE id = ANY(%s)", (record_ids,))
        return cur.rowcount or 0

# ── PR Items ─────────────────────────────────────────────────
def get_pr_items(record_id):
    return q("""SELECT id, record_id, row_type, item_name, unit, quantity, remarks, sort_order
                FROM pr_items
                WHERE record_id=%s
                ORDER BY sort_order, created_at, id""", (record_id,))

def get_pr_items_for_records(record_ids):
    if not record_ids: return {}
    record_ids = [r for r in record_ids if r]
    if not record_ids: return {}
    rows = q("""SELECT id, record_id, row_type, item_name, unit, quantity, remarks, sort_order
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
        exe("""INSERT INTO pr_items(id,record_id,row_type,item_name,unit,quantity,remarks,sort_order,updated_at)
               VALUES(%s,%s,%s,%s,%s,%s,%s,%s,NOW())""",
            (str(uuid.uuid4()), record_id, it.get("row_type","item"),
             it.get("item_name",""), it.get("unit"), it.get("quantity"), it.get("remarks"), i))
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

def rename_list_item(pid, list_name, old_item, new_item):
    old_item = str(old_item or "").strip()
    new_item = str(new_item or "").strip()
    if not old_item or not new_item or old_item == new_item:
        return 0
    with DB() as cur:
        cur.execute("""
            UPDATE dropdown_lists
            SET item_value=%s
            WHERE project_id=%s AND list_name=%s AND item_value=%s
        """, (new_item, pid, list_name, old_item))
        renamed = cur.rowcount or 0
        if not renamed:
            return 0
        cur.execute("""
            SELECT DISTINCT dt_id, col_key
            FROM columns_config
            WHERE project_id=%s AND list_name=%s AND col_type='dropdown'
        """, (pid, list_name))
        targets = cur.fetchall() or []
        for t in targets:
            cur.execute("""
                SELECT id, data
                FROM records
                WHERE project_id=%s AND dt_id=%s
            """, (pid, t["dt_id"]))
            rows = cur.fetchall() or []
            for row in rows:
                data = row["data"] if isinstance(row["data"], dict) else {}
                raw = data.get(t["col_key"])
                if raw is None:
                    continue
                tokens = [p.strip() for p in str(raw).split(",")]
                changed = False
                out = []
                for tok in tokens:
                    if tok == old_item:
                        out.append(new_item)
                        changed = True
                    elif tok:
                        out.append(tok)
                if not changed:
                    continue
                data[t["col_key"]] = ", ".join(out)
                cur.execute("""
                    UPDATE records
                    SET data=%s::jsonb, updated_at=NOW()
                    WHERE id=%s
                """, (json.dumps(data), row["id"]))
        return renamed

def reorder_list_items(pid, list_name, ordered_items):
    for i, item in enumerate([str(v).strip() for v in (ordered_items or []) if str(v).strip()]):
        exe("""
            UPDATE dropdown_lists
            SET sort_order=%s
            WHERE project_id=%s AND list_name=%s AND item_value=%s
        """, (i, pid, list_name, item))

def remove_list_item(pid, list_name, item):
    exe("DELETE FROM dropdown_lists WHERE project_id=%s AND list_name=%s AND item_value=%s",
        (pid, list_name, item))

# ── Fast Dashboard Stats (single query) ──────────────────────
def get_dashboard_stats(project_ids=None):
    """One big SQL call — replaces N×M individual queries."""
    import re, json as _json
    projects = get_projects(project_ids)
    if not projects:
        return []

    pids = [p["id"] for p in projects]
    ltr_stats_map = _get_ltr_dashboard_stats_map(pids)
    ph   = ",".join(["%s"] * len(pids))

    meta_rows = q(f"""
        SELECT project_id, item_value, meta
        FROM dropdown_lists
        WHERE project_id IN ({ph})
          AND list_name LIKE %s
          AND meta IS NOT NULL
    """, pids + ["status%"])
    # meta_map[pid][status_value] = meta
    meta_map = {}
    for r in meta_rows:
        meta_map.setdefault(r["project_id"], {})[r["item_value"]] = r["meta"]
    from utils import is_overdue

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
            dt_is_ltr = _is_ltr_dt(dt.get("code",""), dt.get("name",""))

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

            dt_rules = {dt["id"]: get_expected_reply_rule(pid, dt["id"]) for dt in dts}

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
                    status_val = d.get("status")
                    action_val = d.get("action")
                    dt_rule = dt_rules.get(dt["id"])
                    is_ov = is_overdue(d.get("issuedDate"), doc_no, d.get("actualReplyDate"), _has_both, rule=dt_rule, status=status_val, action=action_val)
                    if is_ov: ov += 1
                    if dt_has_discipline.get(dt["id"], False):
                        ds = disc_map.setdefault(disc, {"total":0,"approved":0,"pending":0,"rejected":0,"overdue":0})
                        ds["total"] += 1
                        if is_ov: ds["overdue"] += 1
                    continue
                # Resolve meta: DB first, then fallback to DEFAULT_STATUS_META
                meta = resolve_status_meta(status, pid_meta)
                if meta == "cancelled": continue   # skip cancelled entirely
                t += 1
                is_ap = (meta == "approved")
                is_rj = (meta == "rejected")
                is_pe = (meta == "pending")
                _has_both = dt_has_exp_reply.get(dt["id"], False) and dt_has_status.get(dt["id"], False) and not dt_is_non_workflow
                status_val = d.get("status")
                action_val = d.get("action")
                dt_rule = dt_rules.get(dt["id"])
                is_ov = is_pe and is_overdue(d.get("issuedDate"), doc_no, d.get("actualReplyDate"), _has_both, rule=dt_rule, status=status_val, action=action_val)
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
            if not dt_is_pr and not dt_is_ltr:
                dt_stats.append({"id":dt["id"],"code":dt["code"],"name":dt["name"],
                                 "total":t,"approved":ap,"pending":pe,"overdue":ov,"rejected":rj,
                                 "disc_breakdown":disc_list})
            total+=t; approved+=ap; pending+=pe; overdue_cnt+=ov; total_rj+=rj

        pct = round(approved / total * 100) if total else 0
        result.append({
            "id": pid, "name": p["name"], "code": p["code"],
            "client": pdata.get("client","") if isinstance(pdata, dict) else "",
            "total": total, "approved": approved, "pending": pending,
            "overdue": overdue_cnt, "rejected": total_rj, "pct": pct, "dt_stats": dt_stats,
            "ltr": ltr_stats_map.get(pid, {"total": 0, "sent": 0, "received": 0, "party_stats": [], "top_parties": []})
        })
    return result


def get_data_quality_summary(pid=None, project_ids=None):
    where, params = _project_scope_clause("project_id", pid, project_ids)
    dt_rows = q(f"SELECT project_id, id, code, name FROM doc_types {where}", params)
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
    col_where, col_params = _project_scope_clause("project_id", pid, project_ids)
    col_rows = q(f"""
        SELECT project_id, dt_id, col_key
        FROM columns_config
        {col_where} AND visible=TRUE AND col_key IN ('expectedReplyDate','status')
    """, col_params)
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

    rec_where, rec_params = _project_scope_clause("project_id", pid, project_ids)
    recs = q(f"SELECT id, project_id, dt_id, data FROM records {rec_where}", rec_params)
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


def get_action_required_summary(pid=None, limit=10, pending_threshold=14, project_ids=None):
    from utils import compute_duration

    projects = get_projects(project_ids)
    project_codes = {p["id"]: p["code"] for p in projects}

    overdue_rows = get_overdue_records(pid, project_ids=project_ids)[:limit]
    top_overdue = [
        {**r, "project_code": project_codes.get(r["project_id"], r["project_id"])}
        for r in overdue_rows
    ]

    dt_where, dt_params = _project_scope_clause("project_id", pid, project_ids)
    dt_rows = q(f"SELECT project_id, id, code, name FROM doc_types {dt_where}", dt_params)
    dt_meta = {(d["project_id"], d["id"]): d for d in dt_rows}

    list_where, list_params = _project_scope_clause("project_id", pid, project_ids)
    meta_rows = q(f"""
        SELECT project_id, item_value, meta
        FROM dropdown_lists
        {list_where} AND list_name LIKE %s AND meta IS NOT NULL
    """, list_params + ["status%"])
    meta_map = {}
    for r in meta_rows:
        meta_map.setdefault(r["project_id"], {})[r["item_value"]] = r["meta"]

    col_where, col_params = _project_scope_clause("project_id", pid, project_ids)
    col_rows = q(f"""
        SELECT project_id, dt_id, col_key
        FROM columns_config
        {col_where} AND visible=TRUE AND col_key IN ('expectedReplyDate','status')
    """, col_params)
    dt_has_exp = set()
    dt_has_status = set()
    for r in col_rows:
        key = (r["project_id"], r["dt_id"])
        if r["col_key"] == "expectedReplyDate":
            dt_has_exp.add(key)
        elif r["col_key"] == "status":
            dt_has_status.add(key)
            
    dt_rules = {r["dt_id"]: get_expected_reply_rule(pid, r["dt_id"]) for r in col_rows}

    workflow_dt_keys = {
        key for key, dt in dt_meta.items()
        if key in dt_has_exp and key in dt_has_status and not _is_non_workflow_dt(dt.get("code",""), dt.get("name",""))
    }

    rec_where, rec_params = _project_scope_clause("project_id", pid, project_ids)
    rows = q(f"SELECT id, project_id, dt_id, data, created_at FROM records {rec_where} ORDER BY created_at DESC", rec_params)
    pending_longest = []
    recent_rejected = []

    for row in rows:
        dt_key = (row["project_id"], row["dt_id"])
        if dt_key not in workflow_dt_keys:
            continue
        d = row["data"] if isinstance(row["data"], dict) else {}
        status = d.get("status", "")
        meta = resolve_status_meta(status, meta_map.get(row["project_id"], {}))
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
            status_val = d.get("status")
            action_val = d.get("action")
            dt_rule = dt_rules.get(row["dt_id"])
            days_delay = compute_duration(d.get("issuedDate"), None, rule=dt_rule, status=status_val, action=action_val) or 0
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


def get_pr_analytics_summary(pid=None, limit=5, project_ids=None):
    projects = get_projects(project_ids)
    project_map = {p["id"]: p for p in projects}
    allowed_ids = _clean_project_ids(project_ids)
    if pid:
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
        """, (pid, "%requisition%", "%purchase request%"))
    elif allowed_ids is not None:
        if not allowed_ids:
            pr_rows = []
        else:
            ph = ",".join(["%s"] * len(allowed_ids))
            pr_rows = q(f"""
                SELECT r.id, r.project_id, r.data, d.code, d.name
                FROM records r
                JOIN doc_types d ON d.id=r.dt_id AND d.project_id=r.project_id
                WHERE r.project_id IN ({ph})
                  AND (
                       UPPER(COALESCE(d.code, ''))='PR'
                    OR LOWER(COALESCE(d.name, '')) LIKE %s
                    OR LOWER(COALESCE(d.name, '')) LIKE %s
                  )
            """, allowed_ids + ["%requisition%", "%purchase request%"])
    else:
        pr_rows = q("""
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
            "top_project_name": "",
            "top_project_count": 0,
            "top_trade_name": "",
            "top_trade_count": 0,
            "trade_count_total": 0,
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

    top_project = top_projects[0] if top_projects else {}
    top_trade = top_trades[0] if top_trades else {}

    return {
        "total_pr_records": total_pr_records,
        "top_projects": top_projects,
        "top_trades": top_trades,
        "top_project_name": top_project.get("project_name", ""),
        "top_project_count": top_project.get("pr_count", 0),
        "top_trade_name": top_trade.get("trade", ""),
        "top_trade_count": top_trade.get("pr_count", 0),
        "trade_count_total": len(trade_counts),
    }


def get_monthly_trend(pid=None, project_ids=None):
    """Returns last 6 months of submission counts per month."""
    import re as _re
    where, params = _project_scope_clause("project_id", pid, project_ids)
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
        meta = resolve_status_meta(d.get("status", ""))
        if meta == "approved":
            months[key]["approved"] += 1
    return [{"month": k, "submitted": v["submitted"], "approved": v["approved"]}
            for k, v in sorted(months.items())]


def get_aging_report(pid=None, project_ids=None):
    """Returns pending docs grouped by days lapsed — only for types with visible expectedReplyDate."""
    from utils import compute_duration
    where, params = _project_scope_clause("r.project_id", pid, project_ids)
    col_where, col_params = _project_scope_clause("project_id", pid, project_ids)
    exp_rows = q(
        "SELECT dt_id, col_key FROM columns_config"
        f" {col_where} AND col_key IN ('expectedReplyDate','status') AND visible=TRUE",
        col_params)
    dt_has_exp_a    = set()
    dt_has_status_a = set()
    for r in exp_rows:
        if r["col_key"] == "expectedReplyDate": dt_has_exp_a.add(r["dt_id"])
        elif r["col_key"] == "status":          dt_has_status_a.add(r["dt_id"])
        
    dt_rules = {r["dt_id"]: get_expected_reply_rule(pid, r["dt_id"]) for r in exp_rows}
    dt_with_exp = dt_has_exp_a & dt_has_status_a
    dt_where, dt_params = _project_scope_clause("project_id", pid, project_ids)
    dt_rows = q(f"SELECT id, code, name FROM doc_types {dt_where}", dt_params)
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
        status_val = d.get("status")
        action_val = d.get("action")
        dt_rule = dt_rules.get(row["dt_id"])
        days = compute_duration(issued, None, rule=dt_rule, status=status_val, action=action_val) or 0
        if days < 1: continue
        if days <= 7:    buckets["1-7"]   += 1
        elif days <= 14: buckets["8-14"]  += 1
        elif days <= 21: buckets["15-21"] += 1
        else:            buckets[">21"]   += 1
    return [{"range": k, "count": v} for k, v in buckets.items()]

def get_quality_report(pid=None, project_ids=None):
    """Returns doc quality: how many docs needed 0,1,2,3+ revisions."""
    import re as _re
    where, params = _project_scope_clause("project_id", pid, project_ids)
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


def get_overdue_records(pid=None, project_ids=None):
    """Returns overdue records — only for doc types with BOTH visible status AND visible expectedReplyDate."""
    from utils import is_overdue
    import datetime as _dt
    where, params = _project_scope_clause("r.project_id", pid, project_ids)
    rows = q(f"""SELECT r.project_id, r.dt_id, r.data, d.code as dt_code, d.name as dt_name
                 FROM records r
                 JOIN doc_types d ON d.id=r.dt_id AND d.project_id=r.project_id
                 {where}""", params)
    # Only dt_ids that have BOTH visible status AND visible expectedReplyDate
    # (types like Letters/PR that have no status column are excluded)
    col_where, col_params = _project_scope_clause("project_id", pid, project_ids)
    cols = q(
        "SELECT dt_id, col_key FROM columns_config"
        f" {col_where} AND col_key IN ('expectedReplyDate','status') AND visible=TRUE",
        col_params)
    dt_has_exp    = set()
    dt_has_status = set()
    for r in cols:
        if r["col_key"] == "expectedReplyDate": dt_has_exp.add(r["dt_id"])
        elif r["col_key"] == "status":          dt_has_status.add(r["dt_id"])
    dt_rules = {r["dt_id"]: get_expected_reply_rule(pid, r["dt_id"]) for r in cols}
    dt_with_exp = dt_has_exp & dt_has_status
    dt_where, dt_params = _project_scope_clause("project_id", pid, project_ids)
    dt_rows = q(f"SELECT id, code, name FROM doc_types {dt_where}", dt_params)
    dt_with_exp = {dt_id for dt_id in dt_with_exp
                   if not _is_non_workflow_dt(
                       next((d["code"] for d in dt_rows if d["id"] == dt_id), ""),
                       next((d["name"] for d in dt_rows if d["id"] == dt_id), "")
                   )}
    
    list_where, list_params = _project_scope_clause("project_id", pid, project_ids)
    meta_rows = q(f"SELECT project_id, item_value, meta FROM dropdown_lists {list_where} AND list_name LIKE %s AND meta IS NOT NULL", list_params + ["status%"])
    meta_map = {}
    for r in meta_rows:
        meta_map.setdefault(r["project_id"], {})[r["item_value"]] = r["meta"]

    result = []
    for row in rows:
        if row["dt_id"] not in dt_with_exp: continue
        d = row["data"] if isinstance(row["data"], dict) else {}
        if d.get("actualReplyDate"): continue
        doc_no = d.get("docNo", "") or ""
        issued = d.get("issuedDate", "") or ""
        if not issued: continue
        status_val = d.get("status")
        action_val = d.get("action")
        dt_rule = dt_rules.get(row["dt_id"])
        
        meta = resolve_status_meta(status_val, meta_map.get(row["project_id"], {}))
        if meta != "pending":
            continue

        if is_overdue(issued, doc_no, None, True, rule=dt_rule, status=status_val, action=action_val):
            try:
                from utils import compute_duration
                days = compute_duration(issued, None, rule=dt_rule, status=status_val, action=action_val) or 0
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

def update_record_link(doc_no, link):
    """Updates the fileLocation (Drive Link) of a record by its docNo."""
    import json
    rows = q("SELECT id, data FROM records WHERE data->>'docNo' = %s", (doc_no,))
    for r in rows:
        d = r["data"] if isinstance(r["data"], dict) else {}
        d["fileLocation"] = link
        exe("UPDATE records SET data=%s WHERE id=%s", (json.dumps(d), r["id"]))
