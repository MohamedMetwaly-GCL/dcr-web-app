"""
database.py
Supabase (PostgreSQL) في Production على Render
SQLite fallback للتطوير المحلي
"""
import json, os, uuid, datetime

# ── بيئة التشغيل ──────────────────────────────────────────────────────────
DATABASE_URL = os.environ.get("DATABASE_URL", "")   # Render بيحقنها تلقائياً
USE_POSTGRES  = bool(DATABASE_URL)

# ── SQLite (local dev) ────────────────────────────────────────────────────
if not USE_POSTGRES:
    import sqlite3
    from pathlib import Path
    _DB = Path(__file__).parent.parent / "DCR_Database.db"

# ── PostgreSQL (Supabase) ─────────────────────────────────────────────────
if USE_POSTGRES:
    import psycopg2
    from psycopg2.extras import RealDictCursor
    from psycopg2.pool import ThreadedConnectionPool
    _pool = None

    def _get_pool():
        global _pool
        if _pool is None:
            _pool = ThreadedConnectionPool(1, 10, DATABASE_URL, sslmode="require")
        return _pool

    class _PGConn:
        def __init__(self):
            self.conn = _get_pool().getconn()
            self.conn.autocommit = False

        def __enter__(self):
            return self.conn.cursor(cursor_factory=RealDictCursor)

        def __exit__(self, exc, *_):
            if exc:
                self.conn.rollback()
            else:
                self.conn.commit()
            _get_pool().putconn(self.conn)


# ── Unified query helpers ─────────────────────────────────────────────────
def _q(sql, params=(), one=False):
    if USE_POSTGRES:
        sql = _pg_sql(sql)
        with _PGConn() as cur:
            cur.execute(sql, params)
            rows = cur.fetchone() if one else cur.fetchall()
        if one:
            return dict(rows) if rows else None
        return [dict(r) for r in rows] if rows else []
    else:
        conn = sqlite3.connect(str(_DB), check_same_thread=False)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA journal_mode=WAL")
        try:
            cur = conn.execute(sql, params)
            rows = cur.fetchone() if one else cur.fetchall()
            if one:
                return dict(rows) if rows else None
            return [dict(r) for r in rows] if rows else []
        finally:
            conn.close()


def _exe(sql, params=()):
    if USE_POSTGRES:
        sql = _pg_sql(sql)
        with _PGConn() as cur:
            cur.execute(sql, params)
    else:
        conn = sqlite3.connect(str(_DB), check_same_thread=False)
        conn.execute("PRAGMA journal_mode=WAL")
        try:
            conn.execute(sql, params)
            conn.commit()
        finally:
            conn.close()


def _exe_many(sql, param_list):
    if USE_POSTGRES:
        sql = _pg_sql(sql)
        with _PGConn() as cur:
            cur.executemany(sql, param_list)
    else:
        conn = sqlite3.connect(str(_DB), check_same_thread=False)
        try:
            conn.executemany(sql, param_list)
            conn.commit()
        finally:
            conn.close()


def _pg_sql(sql):
    """Convert SQLite ? placeholders to PostgreSQL %s"""
    return sql.replace("?", "%s")


# ── Schema ────────────────────────────────────────────────────────────────
_SCHEMA_PG = """
CREATE TABLE IF NOT EXISTS project (
    key TEXT PRIMARY KEY, value TEXT);

CREATE TABLE IF NOT EXISTS logos (
    company_key TEXT PRIMARY KEY, image_data TEXT);

CREATE TABLE IF NOT EXISTS doc_types (
    id TEXT PRIMARY KEY, name TEXT NOT NULL,
    code TEXT NOT NULL, sort_order INTEGER DEFAULT 0);

CREATE TABLE IF NOT EXISTS columns_config (
    id SERIAL PRIMARY KEY,
    doc_type_id TEXT NOT NULL,
    col_key TEXT NOT NULL,
    label TEXT NOT NULL,
    col_type TEXT NOT NULL,
    list_name TEXT,
    visible INTEGER DEFAULT 1,
    sort_order INTEGER DEFAULT 0,
    width INTEGER DEFAULT 120,
    UNIQUE(doc_type_id, col_key));

CREATE TABLE IF NOT EXISTS records (
    id TEXT PRIMARY KEY,
    doc_type_id TEXT NOT NULL,
    data TEXT NOT NULL,
    created_at TEXT DEFAULT (to_char(NOW(),'YYYY-MM-DD HH24:MI:SS')),
    updated_at TEXT DEFAULT (to_char(NOW(),'YYYY-MM-DD HH24:MI:SS')));

CREATE INDEX IF NOT EXISTS idx_records_dt ON records(doc_type_id);

CREATE TABLE IF NOT EXISTS dropdown_lists (
    id SERIAL PRIMARY KEY,
    list_name TEXT NOT NULL,
    item_value TEXT NOT NULL,
    sort_order INTEGER DEFAULT 0,
    UNIQUE(list_name, item_value));

CREATE TABLE IF NOT EXISTS settings (
    key TEXT PRIMARY KEY, value TEXT);
"""

_SCHEMA_SQLITE = """
CREATE TABLE IF NOT EXISTS project (key TEXT PRIMARY KEY, value TEXT);
CREATE TABLE IF NOT EXISTS logos (company_key TEXT PRIMARY KEY, image_data TEXT);
CREATE TABLE IF NOT EXISTS doc_types (
    id TEXT PRIMARY KEY, name TEXT NOT NULL, code TEXT NOT NULL, sort_order INTEGER DEFAULT 0);
CREATE TABLE IF NOT EXISTS columns_config (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    doc_type_id TEXT NOT NULL, col_key TEXT NOT NULL, label TEXT NOT NULL,
    col_type TEXT NOT NULL, list_name TEXT, visible INTEGER DEFAULT 1,
    sort_order INTEGER DEFAULT 0, width INTEGER DEFAULT 120,
    UNIQUE(doc_type_id, col_key));
CREATE TABLE IF NOT EXISTS records (
    id TEXT PRIMARY KEY, doc_type_id TEXT NOT NULL, data TEXT NOT NULL,
    created_at TEXT DEFAULT (datetime('now')), updated_at TEXT DEFAULT (datetime('now')));
CREATE INDEX IF NOT EXISTS idx_records_dt ON records(doc_type_id);
CREATE TABLE IF NOT EXISTS dropdown_lists (
    id INTEGER PRIMARY KEY AUTOINCREMENT, list_name TEXT NOT NULL,
    item_value TEXT NOT NULL, sort_order INTEGER DEFAULT 0,
    UNIQUE(list_name, item_value));
CREATE TABLE IF NOT EXISTS settings (key TEXT PRIMARY KEY, value TEXT);
"""


def init_db():
    if USE_POSTGRES:
        with _PGConn() as cur:
            for stmt in _SCHEMA_PG.strip().split(";"):
                stmt = stmt.strip()
                if stmt:
                    cur.execute(stmt)
    else:
        from pathlib import Path
        _DB.parent.mkdir(parents=True, exist_ok=True)
        conn = sqlite3.connect(str(_DB), check_same_thread=False)
        conn.executescript(_SCHEMA_SQLITE)
        conn.commit()
        conn.close()
    _seed_if_empty()


def _seed_if_empty():
    # Only seed if doc_types is empty
    if _q("SELECT COUNT(*) as c FROM doc_types", one=True).get("c", 0) > 0:
        return

    # Project
    proj = {
        "code": "24CP01", "name": "EKH New Office Building (H24A)",
        "startDate": "2025-03-16", "endDate": "2026-12-31",
        "client": "EKH", "landlord": "Misr Italia", "pmo": "PMO",
        "mainConsultant": "SCAS", "mepConsultant": "Style Design",
        "contractor": "GCL", "extraFields": "[]",
    }
    for k, v in proj.items():
        _exe("INSERT INTO project(key,value) VALUES(?,?) ON CONFLICT(key) DO NOTHING", (k, v))

    # Doc types
    dts = [
        ("DS","Document Submittal","DS",0), ("SD","Shop Drawing Submittal","SD",1),
        ("MS","Material Submittal","MS",2), ("IR","Inspection Request","IR",3),
        ("MIR","Material Inspection Request","MIR",4), ("RFI","Request for Information","RFI",5),
        ("WN","Work Notification","WN",6), ("QS","Quantity Surveyor","QS",7),
        ("ABD","As Built Drawing","ABD",8), ("CVI","Confirmation of Verbal Instruction","CVI",9),
        ("PCQ","Potential Change Questionnaire","PCQ",10), ("SI","Site Instructions","SI",11),
        ("EI","Engineer's Instruction","EI",12), ("SVP","Safety Violation with Penalty","SVP",13),
        ("NOC","Notice of Change","NOC",14), ("NCR","Non-Conformance Report","NCR",15),
        ("MOM","Minutes of Meetings","MOM",16), ("IOM","Internal Office Memo","IOM",17),
        ("PR","Requisition","PR",18), ("LTR","Letters","LTR",19),
    ]
    for dt in dts:
        _exe("INSERT INTO doc_types(id,name,code,sort_order) VALUES(?,?,?,?) ON CONFLICT DO NOTHING", dt)

    # Dropdown lists
    lists = {
        "discipline": ["Electrical","Mechanical","Civil","Structural","Architecture","General","Others"],
        "trade": ["Lighting","Power","Light Current","Fire Alarm","Low Voltage","Medium Voltage",
                  "Control","HVAC","Plumbing","Fire Fighting","General","Others"],
        "floor": ["Basement Floor 1","Basement Floor 2","Ground Floor","First Floor","Second Floor",
                  "Third Floor","Fourth Floor","Roof Floor","Upper Roof"],
        "status": ["A - Approved","B - Approved As Noted","B,C - Approved & Resubmit",
                   "C - Revise & Resubmit","D - Review not Required","Under Review",
                   "Cancelled","Open","Closed","Replied","Pending"],
    }
    for ln, items in lists.items():
        for i, item in enumerate(items):
            _exe("INSERT INTO dropdown_lists(list_name,item_value,sort_order) VALUES(?,?,?) ON CONFLICT DO NOTHING",
                 (ln, item, i))

    # Default columns for all doc types
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
    for (dt_id, *_) in dts:
        for col in DEFAULT_COLS:
            _exe("""INSERT INTO columns_config
                (doc_type_id,col_key,label,col_type,list_name,visible,sort_order,width)
                VALUES(?,?,?,?,?,1,?,?) ON CONFLICT DO NOTHING""",
                (dt_id, col[0], col[1], col[2], col[3], col[4], col[5]))


# ── ON CONFLICT syntax fix for Postgres ──────────────────────────────────
# psycopg2 doesn't like "ON CONFLICT DO NOTHING" without column spec in some cases
# We handle this by wrapping inserts in try/except for seeds


# ── Project ───────────────────────────────────────────────────────────────
def get_project():
    rows = _q("SELECT key, value FROM project")
    return {r["key"]: r["value"] for r in rows}


def save_project(data: dict):
    for k, v in data.items():
        if USE_POSTGRES:
            _exe("""INSERT INTO project(key,value) VALUES(?,?)
                    ON CONFLICT(key) DO UPDATE SET value=EXCLUDED.value""", (k, str(v)))
        else:
            _exe("INSERT OR REPLACE INTO project(key,value) VALUES(?,?)", (k, str(v)))


# ── Logos ─────────────────────────────────────────────────────────────────
def get_logo(key):
    r = _q("SELECT image_data FROM logos WHERE company_key=?", (key,), one=True)
    return r["image_data"] if r else None


def save_logo(key, data):
    if USE_POSTGRES:
        _exe("""INSERT INTO logos(company_key,image_data) VALUES(?,?)
                ON CONFLICT(company_key) DO UPDATE SET image_data=EXCLUDED.image_data""", (key, data))
    else:
        _exe("INSERT OR REPLACE INTO logos(company_key,image_data) VALUES(?,?)", (key, data))


# ── Doc Types ─────────────────────────────────────────────────────────────
def get_doc_types():
    return _q("SELECT * FROM doc_types ORDER BY sort_order, id")


def add_doc_type(id_, name, code):
    r = _q("SELECT MAX(sort_order) as m FROM doc_types", one=True)
    order = (r["m"] or 0) + 1 if r else 0
    try:
        if USE_POSTGRES:
            _exe("INSERT INTO doc_types(id,name,code,sort_order) VALUES(?,?,?,?) ON CONFLICT DO NOTHING",
                 (id_, name, code, order))
        else:
            _exe("INSERT OR IGNORE INTO doc_types(id,name,code,sort_order) VALUES(?,?,?,?)",
                 (id_, name, code, order))
        # Seed default columns
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
        for col in DEFAULT_COLS:
            if USE_POSTGRES:
                _exe("""INSERT INTO columns_config(doc_type_id,col_key,label,col_type,list_name,visible,sort_order,width)
                    VALUES(?,?,?,?,?,1,?,?) ON CONFLICT DO NOTHING""",
                    (id_, col[0], col[1], col[2], col[3], col[4], col[5]))
            else:
                _exe("""INSERT OR IGNORE INTO columns_config(doc_type_id,col_key,label,col_type,list_name,visible,sort_order,width)
                    VALUES(?,?,?,?,?,1,?,?)""",
                    (id_, col[0], col[1], col[2], col[3], col[4], col[5]))
    except Exception as e:
        print(f"add_doc_type error: {e}")


def rename_doc_type(id_, new_name):
    _exe("UPDATE doc_types SET name=? WHERE id=?", (new_name, id_))


def delete_doc_type(id_):
    _exe("DELETE FROM doc_types WHERE id=?", (id_,))
    _exe("DELETE FROM records WHERE doc_type_id=?", (id_,))
    _exe("DELETE FROM columns_config WHERE doc_type_id=?", (id_,))


# ── Columns ───────────────────────────────────────────────────────────────
def get_columns(dt_id, visible_only=False):
    sql = "SELECT * FROM columns_config WHERE doc_type_id=?"
    if visible_only:
        sql += " AND visible=1"
    sql += " ORDER BY sort_order"
    return _q(sql, (dt_id,))


def add_column(dt_id, col_key, label, col_type, list_name=None):
    r = _q("SELECT MAX(sort_order) as m FROM columns_config WHERE doc_type_id=?", (dt_id,), one=True)
    order = (r["m"] or 0) + 1 if r else 0
    if USE_POSTGRES:
        _exe("""INSERT INTO columns_config(doc_type_id,col_key,label,col_type,list_name,visible,sort_order,width)
            VALUES(?,?,?,?,?,1,?,120) ON CONFLICT DO NOTHING""",
            (dt_id, col_key, label, col_type, list_name, order))
    else:
        _exe("""INSERT OR IGNORE INTO columns_config(doc_type_id,col_key,label,col_type,list_name,visible,sort_order,width)
            VALUES(?,?,?,?,?,1,?,120)""",
            (dt_id, col_key, label, col_type, list_name, order))


def update_col_visibility(col_id, visible):
    _exe("UPDATE columns_config SET visible=? WHERE id=?", (1 if visible else 0, col_id))


def delete_column(col_id):
    _exe("DELETE FROM columns_config WHERE id=?", (col_id,))


# ── Records ───────────────────────────────────────────────────────────────
def get_records(dt_id, search="", filters=None):
    rows = _q("SELECT * FROM records WHERE doc_type_id=? ORDER BY created_at", (dt_id,))
    result = []
    for row in rows:
        data = json.loads(row["data"])
        data["_id"]      = row["id"]
        data["_created"] = row["created_at"]
        result.append(data)
    if search:
        sq = search.lower()
        result = [r for r in result if any(sq in str(v).lower() for v in r.values())]
    if filters:
        for k, v in filters.items():
            if v:
                result = [r for r in result if v.lower() in str(r.get(k, "")).lower()]
    return result


def get_record_count(dt_id):
    r = _q("SELECT COUNT(*) as c FROM records WHERE doc_type_id=?", (dt_id,), one=True)
    return r["c"] if r else 0


def save_record(dt_id, rec_id, data: dict):
    now  = datetime.datetime.now().isoformat()
    clean = {k: v for k, v in data.items() if not k.startswith("_")}
    jdata = json.dumps(clean, ensure_ascii=False)
    if USE_POSTGRES:
        _exe("""INSERT INTO records(id,doc_type_id,data,updated_at) VALUES(?,?,?,?)
                ON CONFLICT(id) DO UPDATE SET data=EXCLUDED.data, updated_at=EXCLUDED.updated_at""",
             (rec_id, dt_id, jdata, now))
    else:
        _exe("INSERT OR REPLACE INTO records(id,doc_type_id,data,updated_at) VALUES(?,?,?,?)",
             (rec_id, dt_id, jdata, now))


def delete_record(rec_id):
    _exe("DELETE FROM records WHERE id=?", (rec_id,))


# ── Dropdown Lists ────────────────────────────────────────────────────────
def get_dropdown_list(name):
    return [r["item_value"] for r in
            _q("SELECT item_value FROM dropdown_lists WHERE list_name=? ORDER BY sort_order", (name,))]


def get_all_dropdown_lists():
    names = [r["list_name"] for r in
             _q("SELECT DISTINCT list_name FROM dropdown_lists ORDER BY list_name")]
    return {n: get_dropdown_list(n) for n in names}


def add_dropdown_item(list_name, item):
    r = _q("SELECT MAX(sort_order) as m FROM dropdown_lists WHERE list_name=?", (list_name,), one=True)
    order = (r["m"] or 0) + 1 if r else 0
    if USE_POSTGRES:
        _exe("INSERT INTO dropdown_lists(list_name,item_value,sort_order) VALUES(?,?,?) ON CONFLICT DO NOTHING",
             (list_name, item, order))
    else:
        _exe("INSERT OR IGNORE INTO dropdown_lists(list_name,item_value,sort_order) VALUES(?,?,?)",
             (list_name, item, order))


def remove_dropdown_item(list_name, item):
    _exe("DELETE FROM dropdown_lists WHERE list_name=? AND item_value=?", (list_name, item))


# ── Settings ──────────────────────────────────────────────────────────────
def get_setting(key, default=None):
    r = _q("SELECT value FROM settings WHERE key=?", (key,), one=True)
    return r["value"] if r else default


def save_setting(key, value):
    if USE_POSTGRES:
        _exe("""INSERT INTO settings(key,value) VALUES(?,?)
                ON CONFLICT(key) DO UPDATE SET value=EXCLUDED.value""", (key, str(value)))
    else:
        _exe("INSERT OR REPLACE INTO settings(key,value) VALUES(?,?)", (key, str(value)))
