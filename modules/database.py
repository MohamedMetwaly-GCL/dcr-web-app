"""
database.py - DCR v6
Supabase PostgreSQL (Production) + SQLite (Local Dev)
New: sessions in DB, multi-project, user-project assignments
"""
import json, os, uuid, datetime, hashlib, secrets

DATABASE_URL = os.environ.get("DATABASE_URL", "")
USE_POSTGRES  = bool(DATABASE_URL)

if not USE_POSTGRES:
    import sqlite3
    from pathlib import Path
    _DB = Path(__file__).parent.parent / "DCR_Database.db"

if USE_POSTGRES:
    import psycopg2
    from psycopg2.extras import RealDictCursor
    from psycopg2.pool import ThreadedConnectionPool
    _pool = None

    def _get_pool():
        global _pool
        if _pool is None:
            _pool = ThreadedConnectionPool(1, 10, DATABASE_URL, sslmode="require",
                                           connect_timeout=10, keepalives=1,
                                           keepalives_idle=30, keepalives_interval=10,
                                           keepalives_count=5)
        return _pool

    def _reset_pool():
        global _pool
        try:
            if _pool: _pool.closeall()
        except Exception: pass
        _pool = None

    class _PGConn:
        def __init__(self):
            for attempt in range(2):
                try:
                    self.conn = _get_pool().getconn()
                    self.conn.autocommit = True
                    self.conn.cursor().execute("SELECT 1")
                    self.conn.autocommit = False
                    break
                except Exception:
                    try: _get_pool().putconn(self.conn)
                    except Exception: pass
                    if attempt == 0: _reset_pool()
                    else: raise

        def __enter__(self):
            return self.conn.cursor(cursor_factory=RealDictCursor)

        def __exit__(self, exc, *_):
            if exc:
                try: self.conn.rollback()
                except Exception: pass
            else:
                self.conn.commit()
            try: _get_pool().putconn(self.conn)
            except Exception: pass


# ── Query helpers ─────────────────────────────────────────────
def _pg_sql(sql):
    return sql.replace("?", "%s")

def _q(sql, params=(), one=False):
    if USE_POSTGRES:
        with _PGConn() as cur:
            cur.execute(_pg_sql(sql), params)
            rows = cur.fetchone() if one else cur.fetchall()
        if one: return dict(rows) if rows else None
        return [dict(r) for r in rows] if rows else []
    else:
        conn = sqlite3.connect(str(_DB), check_same_thread=False)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA journal_mode=WAL")
        try:
            cur = conn.execute(sql, params)
            rows = cur.fetchone() if one else cur.fetchall()
            if one: return dict(rows) if rows else None
            return [dict(r) for r in rows] if rows else []
        finally: conn.close()

def _exe(sql, params=()):
    if USE_POSTGRES:
        with _PGConn() as cur:
            cur.execute(_pg_sql(sql), params)
    else:
        conn = sqlite3.connect(str(_DB), check_same_thread=False)
        conn.execute("PRAGMA journal_mode=WAL")
        try:
            conn.execute(sql, params)
            conn.commit()
        finally: conn.close()


# ── Schema ────────────────────────────────────────────────────
_SCHEMA_PG = """
CREATE TABLE IF NOT EXISTS users (
    username TEXT PRIMARY KEY,
    pw_hash  TEXT NOT NULL,
    role     TEXT NOT NULL DEFAULT 'viewer',
    created_at TIMESTAMPTZ DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS sessions (
    token      TEXT PRIMARY KEY,
    username   TEXT NOT NULL,
    role       TEXT NOT NULL,
    expires_at TIMESTAMPTZ NOT NULL,
    created_at TIMESTAMPTZ DEFAULT NOW()
);
CREATE INDEX IF NOT EXISTS idx_sessions_token ON sessions(token);

CREATE TABLE IF NOT EXISTS projects (
    id         TEXT PRIMARY KEY,
    name       TEXT NOT NULL,
    code       TEXT NOT NULL,
    data       TEXT NOT NULL DEFAULT '{}',
    created_at TIMESTAMPTZ DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS user_projects (
    username   TEXT NOT NULL,
    project_id TEXT NOT NULL,
    can_edit   INTEGER DEFAULT 1,
    PRIMARY KEY (username, project_id)
);

CREATE TABLE IF NOT EXISTS logos (
    project_id  TEXT NOT NULL,
    logo_key    TEXT NOT NULL,
    image_data  TEXT,
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
    doc_type_id TEXT NOT NULL,
    col_key     TEXT NOT NULL,
    label       TEXT NOT NULL,
    col_type    TEXT NOT NULL,
    list_name   TEXT,
    visible     INTEGER DEFAULT 1,
    sort_order  INTEGER DEFAULT 0,
    width       INTEGER DEFAULT 120,
    UNIQUE(project_id, doc_type_id, col_key)
);

CREATE TABLE IF NOT EXISTS records (
    id          TEXT PRIMARY KEY,
    project_id  TEXT NOT NULL,
    doc_type_id TEXT NOT NULL,
    data        TEXT NOT NULL DEFAULT '{}',
    created_at  TIMESTAMPTZ DEFAULT NOW(),
    updated_at  TIMESTAMPTZ DEFAULT NOW()
);
CREATE INDEX IF NOT EXISTS idx_records_proj_dt ON records(project_id, doc_type_id);

CREATE TABLE IF NOT EXISTS dropdown_lists (
    id         SERIAL PRIMARY KEY,
    project_id TEXT NOT NULL,
    list_name  TEXT NOT NULL,
    item_value TEXT NOT NULL,
    sort_order INTEGER DEFAULT 0,
    UNIQUE(project_id, list_name, item_value)
);

CREATE TABLE IF NOT EXISTS settings (
    project_id TEXT NOT NULL,
    key        TEXT NOT NULL,
    value      TEXT,
    PRIMARY KEY (project_id, key)
);
"""

_SCHEMA_SQLITE = """
CREATE TABLE IF NOT EXISTS users (
    username TEXT PRIMARY KEY, pw_hash TEXT NOT NULL,
    role TEXT NOT NULL DEFAULT 'viewer', created_at TEXT DEFAULT (datetime('now')));
CREATE TABLE IF NOT EXISTS sessions (
    token TEXT PRIMARY KEY, username TEXT NOT NULL, role TEXT NOT NULL,
    expires_at TEXT NOT NULL, created_at TEXT DEFAULT (datetime('now')));
CREATE INDEX IF NOT EXISTS idx_sessions_token ON sessions(token);
CREATE TABLE IF NOT EXISTS projects (
    id TEXT PRIMARY KEY, name TEXT NOT NULL, code TEXT NOT NULL,
    data TEXT NOT NULL DEFAULT '{}', created_at TEXT DEFAULT (datetime('now')));
CREATE TABLE IF NOT EXISTS user_projects (
    username TEXT NOT NULL, project_id TEXT NOT NULL, can_edit INTEGER DEFAULT 1,
    PRIMARY KEY (username, project_id));
CREATE TABLE IF NOT EXISTS logos (
    project_id TEXT NOT NULL, logo_key TEXT NOT NULL, image_data TEXT,
    PRIMARY KEY (project_id, logo_key));
CREATE TABLE IF NOT EXISTS doc_types (
    id TEXT NOT NULL, project_id TEXT NOT NULL, name TEXT NOT NULL,
    code TEXT NOT NULL, sort_order INTEGER DEFAULT 0,
    PRIMARY KEY (id, project_id));
CREATE TABLE IF NOT EXISTS columns_config (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    project_id TEXT NOT NULL, doc_type_id TEXT NOT NULL,
    col_key TEXT NOT NULL, label TEXT NOT NULL, col_type TEXT NOT NULL,
    list_name TEXT, visible INTEGER DEFAULT 1, sort_order INTEGER DEFAULT 0,
    width INTEGER DEFAULT 120,
    UNIQUE(project_id, doc_type_id, col_key));
CREATE TABLE IF NOT EXISTS records (
    id TEXT PRIMARY KEY, project_id TEXT NOT NULL, doc_type_id TEXT NOT NULL,
    data TEXT NOT NULL DEFAULT '{}',
    created_at TEXT DEFAULT (datetime('now')),
    updated_at TEXT DEFAULT (datetime('now')));
CREATE INDEX IF NOT EXISTS idx_records_proj_dt ON records(project_id, doc_type_id);
CREATE TABLE IF NOT EXISTS dropdown_lists (
    id INTEGER PRIMARY KEY AUTOINCREMENT, project_id TEXT NOT NULL,
    list_name TEXT NOT NULL, item_value TEXT NOT NULL, sort_order INTEGER DEFAULT 0,
    UNIQUE(project_id, list_name, item_value));
CREATE TABLE IF NOT EXISTS settings (
    project_id TEXT NOT NULL, key TEXT NOT NULL, value TEXT,
    PRIMARY KEY (project_id, key));
"""

DEFAULT_DOC_TYPES = [
    ("DS","Document Submittal",0), ("SD","Shop Drawing Submittal",1),
    ("MS","Material Submittal",2), ("IR","Inspection Request",3),
    ("MIR","Material Inspection Request",4), ("RFI","Request for Information",5),
    ("WN","Work Notification",6), ("QS","Quantity Surveyor",7),
    ("ABD","As Built Drawing",8), ("CVI","Confirmation of Verbal Instruction",9),
    ("PCQ","Potential Change Questionnaire",10), ("SI","Site Instructions",11),
    ("EI","Engineer's Instruction",12), ("SVP","Safety Violation with Penalty",13),
    ("NOC","Notice of Change",14), ("NCR","Non-Conformance Report",15),
    ("MOM","Minutes of Meetings",16), ("IOM","Internal Office Memo",17),
    ("PR","Requisition",18), ("LTR","Letters",19),
]

DEFAULT_COLS = [
    ("docNo","Document No.","docno",None,0,160),
    ("discipline","Discipline","dropdown","discipline",1,110),
    ("trade","Sub-Trade","dropdown","trade",2,120),
    ("title","Title","text",None,3,200),
    ("floor","Floor","dropdown","floor",4,120),
    ("itemRef","Item Ref./DWG No.","text",None,5,150),
    ("issuedDate","Issued Date","date",None,6,100),
    ("expectedReplyDate","Expected Reply","auto_date",None,7,100),
    ("actualReplyDate","Actual Reply","date",None,8,100),
    ("status","Status","dropdown","status",9,160),
    ("duration","Duration (days)","auto_num",None,10,90),
    ("remarks","Remarks","text",None,11,180),
    ("fileLocation","File Location","link",None,12,100),
]

DEFAULT_LISTS = {
    "discipline": ["Electrical","Mechanical","Civil","Structural","Architecture","General","Others"],
    "trade": ["Lighting","Power","Light Current","Fire Alarm","Low Voltage","Medium Voltage",
              "Control","HVAC","Plumbing","Fire Fighting","General","Others"],
    "floor": ["Basement Floor 1","Basement Floor 2","Ground Floor","First Floor","Second Floor",
              "Third Floor","Fourth Floor","Roof Floor","Upper Roof"],
    "status": ["A - Approved","B - Approved As Noted","B,C - Approved & Resubmit",
               "C - Revise & Resubmit","D - Review not Required","Under Review",
               "Cancelled","Open","Closed","Replied","Pending"],
}


def init_db():
    if USE_POSTGRES:
        # Each statement in its own transaction to avoid cascade failures
        stmts = [s.strip() for s in _SCHEMA_PG.strip().split(";") if s.strip()]
        for stmt in stmts:
            try:
                with _PGConn() as cur:
                    cur.execute(stmt)
            except Exception as e:
                msg = str(e).lower()
                if "already exists" in msg or "duplicate" in msg:
                    pass  # Ignore — table/index already there
                else:
                    print(f"[DCR] Schema warning ({stmt[:40]}): {e}")
    else:
        from pathlib import Path
        _DB.parent.mkdir(parents=True, exist_ok=True)
        conn = sqlite3.connect(str(_DB), check_same_thread=False)
        conn.executescript(_SCHEMA_SQLITE)
        conn.commit(); conn.close()

    _migrate_old_data()
    _ensure_super_admin()


def _migrate_old_data():
    """Migrate v5 single-project data → v6 multi-project schema.
    Uses raw SQL with explicit column names to avoid schema conflicts."""
    if not USE_POSTGRES:
        return  # SQLite: fresh install only

    try:
        # Check old 'project' table exists (v5 schema marker)
        r = _q("SELECT to_regclass(%s) as t", ("project",), one=True)
        if not r or not r.get("t"):
            return  # No v5 data

        # Check if already migrated
        existing = _q("SELECT COUNT(*) as c FROM projects", one=True)
        if existing and existing.get("c", 0) > 0:
            return  # Already done

        print("[DCR] Starting v5 → v6 migration...")

        # Read old project data
        old_proj_rows = _q("SELECT key, value FROM project")
        if not old_proj_rows:
            return
        old_proj = {r["key"]: r["value"] for r in old_proj_rows}
        proj_id   = (old_proj.get("code") or "P001").strip().upper()
        proj_name = old_proj.get("name") or "Migrated Project"
        proj_code = old_proj.get("code") or "P001"
        proj_data = json.dumps({k: v for k, v in old_proj.items()})

        _exe("INSERT INTO projects(id,name,code,data) VALUES(%s,%s,%s,%s) ON CONFLICT DO NOTHING",
             (proj_id, proj_name, proj_code, proj_data))

        # Migrate doc_types (old schema: id, name, code, sort_order — no project_id)
        try:
            rows = _q("SELECT id, name, code, sort_order FROM doc_types WHERE project_id IS NULL OR project_id = ''")
        except Exception:
            try:
                rows = _q("SELECT id, name, code, sort_order FROM doc_types")
            except Exception:
                rows = []
        migrated_dt_ids = set()
        for dt in rows:
            if dt.get("id") and dt.get("name"):
                _exe("INSERT INTO doc_types(id,project_id,name,code,sort_order) VALUES(%s,%s,%s,%s,%s) ON CONFLICT DO NOTHING",
                     (dt["id"], proj_id, dt["name"], dt.get("code", dt["id"]), dt.get("sort_order", 0)))
                migrated_dt_ids.add(dt["id"])

        # Migrate columns_config (old schema: id, doc_type_id, col_key, label, col_type, list_name, visible, sort_order, width)
        try:
            cols = _q("SELECT id, doc_type_id, col_key, label, col_type, list_name, visible, sort_order, width FROM columns_config WHERE project_id IS NULL OR project_id = ''")
        except Exception:
            try:
                cols = _q("SELECT id, doc_type_id, col_key, label, col_type, list_name, visible, sort_order, width FROM columns_config")
            except Exception:
                cols = []
        for c in cols:
            _exe("""INSERT INTO columns_config(project_id,doc_type_id,col_key,label,col_type,list_name,visible,sort_order,width)
                VALUES(%s,%s,%s,%s,%s,%s,%s,%s,%s) ON CONFLICT DO NOTHING""",
                (proj_id, c["doc_type_id"], c["col_key"], c["label"], c["col_type"],
                 c.get("list_name"), c.get("visible", 1), c.get("sort_order", 0), c.get("width", 120)))

        # Migrate records (old schema: id, doc_type_id, data, created_at, updated_at)
        try:
            recs = _q("SELECT id, doc_type_id, data, created_at, updated_at FROM records WHERE project_id IS NULL OR project_id = ''")
        except Exception:
            try:
                recs = _q("SELECT id, doc_type_id, data, created_at, updated_at FROM records")
            except Exception:
                recs = []
        for r in recs:
            _exe("""INSERT INTO records(id,project_id,doc_type_id,data,created_at,updated_at)
                VALUES(%s,%s,%s,%s,%s,%s) ON CONFLICT DO NOTHING""",
                (r["id"], proj_id, r["doc_type_id"], r.get("data","{}"),
                 r.get("created_at",""), r.get("updated_at","")))

        # Migrate dropdown_lists (old schema: id, list_name, item_value, sort_order)
        try:
            lists = _q("SELECT list_name, item_value, sort_order FROM dropdown_lists WHERE project_id IS NULL OR project_id = ''")
        except Exception:
            try:
                lists = _q("SELECT list_name, item_value, sort_order FROM dropdown_lists")
            except Exception:
                lists = []
        for item in lists:
            _exe("INSERT INTO dropdown_lists(project_id,list_name,item_value,sort_order) VALUES(%s,%s,%s,%s) ON CONFLICT DO NOTHING",
                 (proj_id, item["list_name"], item["item_value"], item.get("sort_order", 0)))

        # Migrate logos (old schema had company_key or logo_key)
        try:
            logos = _q("SELECT company_key, image_data FROM logos")
            for logo in logos:
                _exe("INSERT INTO logos(project_id,logo_key,image_data) VALUES(%s,%s,%s) ON CONFLICT DO NOTHING",
                     (proj_id, logo.get("company_key",""), logo.get("image_data","")))
        except Exception:
            pass  # logos table might already be new schema

        # Migrate users from settings table
        try:
            s = _q("SELECT value FROM settings WHERE key=%s", ("users_json",), one=True)
            if s and s.get("value"):
                users = json.loads(s["value"])
                for uname, udata in users.items():
                    _exe("INSERT INTO users(username,pw_hash,role) VALUES(%s,%s,%s) ON CONFLICT DO NOTHING",
                         (uname, udata.get("pw",""), udata.get("role","viewer")))
                    _exe("INSERT INTO user_projects(username,project_id,can_edit) VALUES(%s,%s,1) ON CONFLICT DO NOTHING",
                         (uname, proj_id))
        except Exception as e:
            print(f"[DCR] Users migration note: {e}")

        print(f"[DCR] ✅ Migration complete → project '{proj_id}' with {len(recs)} records")
    except Exception as e:
        print(f"[DCR] Migration skipped (likely fresh install): {e}")


def _table_exists(name):
    try:
        if USE_POSTGRES:
            r = _q("SELECT to_regclass(%s) as t", (name,), one=True)
            return r and r.get("t") is not None
        else:
            r = _q("SELECT name FROM sqlite_master WHERE type='table' AND name=?", (name,), one=True)
            return bool(r)
    except: return False


def _ensure_super_admin():
    """Create default super admin if no users exist."""
    r = _q("SELECT COUNT(*) as c FROM users", one=True)
    if r and r.get("c", 0) == 0:
        pw = _hash_pw("admin123")
        if USE_POSTGRES:
            _exe("INSERT INTO users(username,pw_hash,role) VALUES(?,?,?) ON CONFLICT DO NOTHING",
                 ("admin", pw, "superadmin"))
        else:
            _exe("INSERT OR IGNORE INTO users(username,pw_hash,role) VALUES(?,?,?)",
                 ("admin", pw, "superadmin"))


def _hash_pw(pw):
    return hashlib.sha256(pw.encode()).hexdigest()


# ── Sessions (stored in DB — survive restarts) ─────────────────
SESSION_TTL = 8 * 3600  # 8 hours

def create_session(username, role):
    token = secrets.token_hex(32)
    expires = (datetime.datetime.utcnow() + datetime.timedelta(seconds=SESSION_TTL)).isoformat()
    if USE_POSTGRES:
        _exe("INSERT INTO sessions(token,username,role,expires_at) VALUES(?,?,?,?) ON CONFLICT DO NOTHING",
             (token, username, role, expires))
    else:
        _exe("INSERT OR IGNORE INTO sessions(token,username,role,expires_at) VALUES(?,?,?,?)",
             (token, username, role, expires))
    return token

def get_session(token):
    if not token: return None
    now = datetime.datetime.utcnow().isoformat()
    r = _q("SELECT username, role FROM sessions WHERE token=? AND expires_at > ?",
           (token, now), one=True)
    return r if r else None

def delete_session(token):
    _exe("DELETE FROM sessions WHERE token=?", (token,))

def cleanup_sessions():
    now = datetime.datetime.utcnow().isoformat()
    _exe("DELETE FROM sessions WHERE expires_at <= ?", (now,))


# ── Users ─────────────────────────────────────────────────────
def get_user(username):
    return _q("SELECT * FROM users WHERE username=?", (username,), one=True)

def get_all_users():
    return _q("SELECT username, role FROM users ORDER BY username")

def add_user(username, pw, role="viewer"):
    if USE_POSTGRES:
        _exe("INSERT INTO users(username,pw_hash,role) VALUES(?,?,?) ON CONFLICT DO NOTHING",
             (username, _hash_pw(pw), role))
    else:
        _exe("INSERT OR IGNORE INTO users(username,pw_hash,role) VALUES(?,?,?)",
             (username, _hash_pw(pw), role))

def delete_user(username):
    _exe("DELETE FROM users WHERE username=?", (username,))
    _exe("DELETE FROM user_projects WHERE username=?", (username,))
    _exe("DELETE FROM sessions WHERE username=?", (username,))

def change_password(username, new_pw):
    _exe("UPDATE users SET pw_hash=? WHERE username=?", (_hash_pw(new_pw), username))

def verify_password(username, pw):
    u = get_user(username)
    return u and u["pw_hash"] == _hash_pw(pw)


# ── User ↔ Project assignments ────────────────────────────────
def get_user_projects(username):
    """Projects a user can edit."""
    return [r["project_id"] for r in
            _q("SELECT project_id FROM user_projects WHERE username=?", (username,))]

def assign_user_project(username, project_id, can_edit=1):
    if USE_POSTGRES:
        _exe("INSERT INTO user_projects(username,project_id,can_edit) VALUES(?,?,?) ON CONFLICT DO NOTHING",
             (username, project_id, can_edit))
    else:
        _exe("INSERT OR IGNORE INTO user_projects(username,project_id,can_edit) VALUES(?,?,?)",
             (username, project_id, can_edit))

def remove_user_project(username, project_id):
    _exe("DELETE FROM user_projects WHERE username=? AND project_id=?", (username, project_id))

def get_project_users(project_id):
    return _q("SELECT username FROM user_projects WHERE project_id=?", (project_id,))


# ── Projects ──────────────────────────────────────────────────
def get_all_projects():
    return _q("SELECT * FROM projects ORDER BY created_at")

def get_project(project_id):
    r = _q("SELECT * FROM projects WHERE id=?", (project_id,), one=True)
    if not r: return {}
    try:
        data = json.loads(r.get("data") or "{}")
    except: data = {}
    data["id"]   = r["id"]
    data["name"] = r["name"]
    data["code"] = r["code"]
    return data

def save_project(project_id, data: dict):
    name = data.get("name", project_id)
    code = data.get("code", project_id)
    jdata = json.dumps({k: v for k, v in data.items() if k not in ("id","name","code")})
    if USE_POSTGRES:
        _exe("""INSERT INTO projects(id,name,code,data) VALUES(?,?,?,?)
                ON CONFLICT(id) DO UPDATE SET name=EXCLUDED.name, code=EXCLUDED.code, data=EXCLUDED.data""",
             (project_id, name, code, jdata))
    else:
        _exe("INSERT OR REPLACE INTO projects(id,name,code,data) VALUES(?,?,?,?)",
             (project_id, name, code, jdata))

def create_project(project_id, name, code, creator_username=None):
    if USE_POSTGRES:
        _exe("INSERT INTO projects(id,name,code,data) VALUES(?,?,?,?) ON CONFLICT DO NOTHING",
             (project_id, name, code, "{}"))
    else:
        _exe("INSERT OR IGNORE INTO projects(id,name,code,data) VALUES(?,?,?,?)",
             (project_id, name, code, "{}"))
    # Seed default doc types, columns, dropdown lists
    _seed_project(project_id)
    # Assign creator
    if creator_username:
        assign_user_project(creator_username, project_id)

def delete_project(project_id):
    for tbl in ("projects","doc_types","columns_config","records",
                "dropdown_lists","logos","settings","user_projects"):
        _exe(f"DELETE FROM {tbl} WHERE project_id=?", (project_id,))

def _seed_project(project_id):
    for (code, name, order) in DEFAULT_DOC_TYPES:
        if USE_POSTGRES:
            _exe("INSERT INTO doc_types(id,project_id,name,code,sort_order) VALUES(?,?,?,?,?) ON CONFLICT DO NOTHING",
                 (code, project_id, name, code, order))
        else:
            _exe("INSERT OR IGNORE INTO doc_types(id,project_id,name,code,sort_order) VALUES(?,?,?,?,?)",
                 (code, project_id, name, code, order))
        for col in DEFAULT_COLS:
            if USE_POSTGRES:
                _exe("""INSERT INTO columns_config(project_id,doc_type_id,col_key,label,col_type,list_name,visible,sort_order,width)
                    VALUES(?,?,?,?,?,?,1,?,?) ON CONFLICT DO NOTHING""",
                    (project_id, code, col[0], col[1], col[2], col[3], col[4], col[5]))
            else:
                _exe("""INSERT OR IGNORE INTO columns_config(project_id,doc_type_id,col_key,label,col_type,list_name,visible,sort_order,width)
                    VALUES(?,?,?,?,?,?,1,?,?)""",
                    (project_id, code, col[0], col[1], col[2], col[3], col[4], col[5]))
    for ln, items in DEFAULT_LISTS.items():
        for i, item in enumerate(items):
            if USE_POSTGRES:
                _exe("INSERT INTO dropdown_lists(project_id,list_name,item_value,sort_order) VALUES(?,?,?,?) ON CONFLICT DO NOTHING",
                     (project_id, ln, item, i))
            else:
                _exe("INSERT OR IGNORE INTO dropdown_lists(project_id,list_name,item_value,sort_order) VALUES(?,?,?,?)",
                     (project_id, ln, item, i))


# ── Logos ─────────────────────────────────────────────────────
def get_logo(project_id, key):
    r = _q("SELECT image_data FROM logos WHERE project_id=? AND logo_key=?", (project_id, key), one=True)
    return r["image_data"] if r else None

def save_logo(project_id, key, data):
    if USE_POSTGRES:
        _exe("""INSERT INTO logos(project_id,logo_key,image_data) VALUES(?,?,?)
                ON CONFLICT(project_id,logo_key) DO UPDATE SET image_data=EXCLUDED.image_data""",
             (project_id, key, data))
    else:
        _exe("INSERT OR REPLACE INTO logos(project_id,logo_key,image_data) VALUES(?,?,?)",
             (project_id, key, data))


# ── Doc Types ─────────────────────────────────────────────────
def get_doc_types(project_id):
    return _q("SELECT * FROM doc_types WHERE project_id=? ORDER BY sort_order, id", (project_id,))

def add_doc_type(project_id, id_, name, code):
    r = _q("SELECT MAX(sort_order) as m FROM doc_types WHERE project_id=?", (project_id,), one=True)
    order = (r["m"] or 0) + 1 if r else 0
    if USE_POSTGRES:
        _exe("INSERT INTO doc_types(id,project_id,name,code,sort_order) VALUES(?,?,?,?,?) ON CONFLICT DO NOTHING",
             (id_, project_id, name, code, order))
    else:
        _exe("INSERT OR IGNORE INTO doc_types(id,project_id,name,code,sort_order) VALUES(?,?,?,?,?)",
             (id_, project_id, name, code, order))
    for col in DEFAULT_COLS:
        if USE_POSTGRES:
            _exe("""INSERT INTO columns_config(project_id,doc_type_id,col_key,label,col_type,list_name,visible,sort_order,width)
                VALUES(?,?,?,?,?,?,1,?,?) ON CONFLICT DO NOTHING""",
                (project_id, id_, col[0], col[1], col[2], col[3], col[4], col[5]))
        else:
            _exe("""INSERT OR IGNORE INTO columns_config(project_id,doc_type_id,col_key,label,col_type,list_name,visible,sort_order,width)
                VALUES(?,?,?,?,?,?,1,?,?)""",
                (project_id, id_, col[0], col[1], col[2], col[3], col[4], col[5]))

def delete_doc_type(project_id, id_):
    _exe("DELETE FROM doc_types WHERE id=? AND project_id=?", (id_, project_id))
    _exe("DELETE FROM records WHERE doc_type_id=? AND project_id=?", (id_, project_id))
    _exe("DELETE FROM columns_config WHERE doc_type_id=? AND project_id=?", (id_, project_id))


# ── Columns ───────────────────────────────────────────────────
def get_columns(project_id, dt_id, visible_only=False):
    sql = "SELECT * FROM columns_config WHERE project_id=? AND doc_type_id=?"
    params = [project_id, dt_id]
    if visible_only:
        sql += " AND visible=1"
    sql += " ORDER BY sort_order"
    return _q(sql, params)

def add_column(project_id, dt_id, col_key, label, col_type, list_name=None):
    r = _q("SELECT MAX(sort_order) as m FROM columns_config WHERE project_id=? AND doc_type_id=?",
           (project_id, dt_id), one=True)
    order = (r["m"] or 0) + 1 if r else 0
    if USE_POSTGRES:
        _exe("""INSERT INTO columns_config(project_id,doc_type_id,col_key,label,col_type,list_name,visible,sort_order,width)
            VALUES(?,?,?,?,?,?,1,?,120) ON CONFLICT DO NOTHING""",
            (project_id, dt_id, col_key, label, col_type, list_name, order))
    else:
        _exe("""INSERT OR IGNORE INTO columns_config(project_id,doc_type_id,col_key,label,col_type,list_name,visible,sort_order,width)
            VALUES(?,?,?,?,?,?,1,?,120)""",
            (project_id, dt_id, col_key, label, col_type, list_name, order))

def update_col_visibility(col_id, visible):
    _exe("UPDATE columns_config SET visible=? WHERE id=?", (1 if visible else 0, col_id))

def delete_column(col_id):
    _exe("DELETE FROM columns_config WHERE id=?", (col_id,))


# ── Records ───────────────────────────────────────────────────
def get_records(project_id, dt_id, search=""):
    rows = _q("SELECT * FROM records WHERE project_id=? AND doc_type_id=? ORDER BY created_at",
              (project_id, dt_id))
    result = []
    for row in rows:
        try: data = json.loads(row["data"])
        except: data = {}
        data["_id"]      = row["id"]
        data["_created"] = str(row.get("created_at",""))
        result.append(data)
    if search:
        sq = search.lower()
        result = [r for r in result if any(sq in str(v).lower() for v in r.values())]
    return result

def get_record_count(project_id, dt_id):
    r = _q("SELECT COUNT(*) as c FROM records WHERE project_id=? AND doc_type_id=?",
           (project_id, dt_id), one=True)
    return r["c"] if r else 0

def save_record(project_id, dt_id, rec_id, data: dict):
    now   = datetime.datetime.utcnow().isoformat()
    clean = {k: v for k, v in data.items() if not k.startswith("_")}
    jdata = json.dumps(clean, ensure_ascii=False)
    if USE_POSTGRES:
        _exe("""INSERT INTO records(id,project_id,doc_type_id,data,updated_at) VALUES(?,?,?,?,?)
                ON CONFLICT(id) DO UPDATE SET data=EXCLUDED.data, updated_at=EXCLUDED.updated_at""",
             (rec_id, project_id, dt_id, jdata, now))
    else:
        _exe("INSERT OR REPLACE INTO records(id,project_id,doc_type_id,data,updated_at) VALUES(?,?,?,?,?)",
             (rec_id, project_id, dt_id, jdata, now))

def delete_record(rec_id):
    _exe("DELETE FROM records WHERE id=?", (rec_id,))


# ── Dropdown Lists ────────────────────────────────────────────
def get_all_dropdown_lists(project_id):
    names = [r["list_name"] for r in
             _q("SELECT DISTINCT list_name FROM dropdown_lists WHERE project_id=? ORDER BY list_name",
                (project_id,))]
    return {n: [r["item_value"] for r in
                _q("SELECT item_value FROM dropdown_lists WHERE project_id=? AND list_name=? ORDER BY sort_order",
                   (project_id, n))]
            for n in names}

def add_dropdown_item(project_id, list_name, item):
    r = _q("SELECT MAX(sort_order) as m FROM dropdown_lists WHERE project_id=? AND list_name=?",
           (project_id, list_name), one=True)
    order = (r["m"] or 0) + 1 if r else 0
    if USE_POSTGRES:
        _exe("INSERT INTO dropdown_lists(project_id,list_name,item_value,sort_order) VALUES(?,?,?,?) ON CONFLICT DO NOTHING",
             (project_id, list_name, item, order))
    else:
        _exe("INSERT OR IGNORE INTO dropdown_lists(project_id,list_name,item_value,sort_order) VALUES(?,?,?,?)",
             (project_id, list_name, item, order))

def remove_dropdown_item(project_id, list_name, item):
    _exe("DELETE FROM dropdown_lists WHERE project_id=? AND list_name=? AND item_value=?",
         (project_id, list_name, item))


# ── Settings ──────────────────────────────────────────────────
def get_setting(project_id, key, default=None):
    r = _q("SELECT value FROM settings WHERE project_id=? AND key=?", (project_id, key), one=True)
    return r["value"] if r else default

def save_setting(project_id, key, value):
    if USE_POSTGRES:
        _exe("""INSERT INTO settings(project_id,key,value) VALUES(?,?,?)
                ON CONFLICT(project_id,key) DO UPDATE SET value=EXCLUDED.value""",
             (project_id, key, str(value)))
    else:
        _exe("INSERT OR REPLACE INTO settings(project_id,key,value) VALUES(?,?,?)",
             (project_id, key, str(value)))
