"""db.py - DCR Flask - Clean PostgreSQL database layer"""
import logging
import os, json, uuid, datetime, hashlib, secrets, re, threading
from contextlib import contextmanager
import bcrypt as _bcrypt
import psycopg2
from psycopg2.extras import RealDictCursor
from psycopg2.pool import ThreadedConnectionPool

DATABASE_URL = os.environ.get("DATABASE_URL", "")
_pool = None
_tx_local = threading.local()
logger = logging.getLogger(__name__)

def get_pool():
    global _pool
    if _pool is None:
        _pool = ThreadedConnectionPool(1, 10, DATABASE_URL,
            sslmode="require", connect_timeout=10,
            keepalives=1, keepalives_idle=30,
            keepalives_interval=10, keepalives_count=5)
    return _pool

class TransactionAbortedError(RuntimeError):
    """Raised when an operation runs in or tries to commit an aborted nested transaction."""


def _sanitize_db_error(err):
    msg = str(err or "")
    msg = re.sub(r"://[^/\s@]+@", "://***:***@", msg)
    msg = re.sub(r"(?i)(password\s*=\s*)(\S+)", r"\1***", msg)
    return f"{type(err).__name__}: {msg}" if msg else type(err).__name__


def _safe_close_conn(conn):
    if conn is None:
        return
    try:
        conn.close()
    except Exception as e:
        logger.error("Failed closing broken DB connection: %s", _sanitize_db_error(e))


def _return_conn_to_pool(conn, broken=False):
    """Safely return a DB connection to the pool (closing broken connections) and log errors without credentials."""
    if conn is None:
        return
    try:
        pool = get_pool()
    except Exception as e:
        logger.error("Failed acquiring DB pool while returning connection: %s", _sanitize_db_error(e))
        _safe_close_conn(conn)
        return
    is_broken = bool(broken or getattr(conn, "closed", 0))
    try:
        if is_broken:
            try:
                pool.putconn(conn, close=True)
            except TypeError:
                _safe_close_conn(conn)
                pool.putconn(conn)
        else:
            pool.putconn(conn)
    except Exception as e:
        logger.error("Failed returning DB connection to pool (broken=%s): %s", is_broken, _sanitize_db_error(e))
        _safe_close_conn(conn)


@contextmanager
def transaction():
    """Run multiple db.q / db.exe calls in one shared, lazy-acquired database transaction.

    The pool connection is acquired only when an unmocked DB() executes a query, so offline
    unit tests that mock db.exe / db.q without a database pool continue to work seamlessly.

    Nested-transaction contract:
    - Nested blocks share the outer transaction lazily.
    - Any exception escaping a nested block marks the transaction rollback-only
      (``_tx_local.aborted = True``) even if caught by outer Python code.
    - Subsequent DB() queries inside an aborted transaction raise TransactionAbortedError,
      and exiting the outer block rolls back cleanly instead of committing partial writes.
    """
    depth = getattr(_tx_local, "depth", 0)
    is_outer = (depth == 0)
    if is_outer:
        _tx_local.conn = None
        _tx_local.aborted = False
    _tx_local.depth = depth + 1
    broken = False
    try:
        yield
        if is_outer:
            conn = getattr(_tx_local, "conn", None)
            if getattr(_tx_local, "aborted", False):
                if conn is not None:
                    try:
                        conn.rollback()
                    except Exception as rb_err:
                        broken = True
                        logger.error("Outer transaction rollback after abort failed: %s", _sanitize_db_error(rb_err))
                raise TransactionAbortedError("Outer transaction rolled back after nested transaction failure")
            if conn is not None:
                try:
                    conn.commit()
                except Exception as cm_err:
                    broken = True
                    logger.error("Transaction commit failed: %s", _sanitize_db_error(cm_err))
                    raise
    except BaseException:
        if not is_outer:
            _tx_local.aborted = True
        else:
            conn = getattr(_tx_local, "conn", None)
            if conn is not None:
                try:
                    conn.rollback()
                except Exception as rb_err:
                    broken = True
                    logger.error("Transaction rollback failed: %s", _sanitize_db_error(rb_err))
        raise
    finally:
        _tx_local.depth = depth
        if is_outer:
            conn = getattr(_tx_local, "conn", None)
            _tx_local.conn = None
            _tx_local.aborted = False
            if conn is not None:
                _return_conn_to_pool(conn, broken=broken)

class DB:
    def __enter__(self):
        self.cur = None
        self.conn = None
        if getattr(_tx_local, "depth", 0) > 0:
            self.in_tx = True
            if getattr(_tx_local, "aborted", False):
                raise TransactionAbortedError("Cannot execute DB operation in an aborted transaction")
            if getattr(_tx_local, "conn", None) is None:
                conn = get_pool().getconn()
                try:
                    conn.autocommit = False
                    cur = conn.cursor(cursor_factory=RealDictCursor)
                except BaseException as err:
                    _tx_local.aborted = True
                    logger.error("Failed preparing tx DB connection/cursor: %s", _sanitize_db_error(err))
                    try:
                        conn.rollback()
                    except Exception as rb_err:
                        logger.error("Rollback after tx setup failure failed: %s", _sanitize_db_error(rb_err))
                    _return_conn_to_pool(conn, broken=True)
                    raise
                _tx_local.conn = conn
                self.conn = conn
                self.cur = cur
                return self.cur
            self.conn = _tx_local.conn
            try:
                self.cur = self.conn.cursor(cursor_factory=RealDictCursor)
            except BaseException as err:
                _tx_local.aborted = True
                logger.error("Failed creating cursor on tx connection: %s", _sanitize_db_error(err))
                raise
            return self.cur
        else:
            self.in_tx = False
            conn = get_pool().getconn()
            try:
                conn.autocommit = False
                cur = conn.cursor(cursor_factory=RealDictCursor)
            except BaseException as err:
                logger.error("Failed preparing DB connection/cursor: %s", _sanitize_db_error(err))
                try:
                    conn.rollback()
                except Exception as rb_err:
                    logger.error("Rollback after setup failure failed: %s", _sanitize_db_error(rb_err))
                _return_conn_to_pool(conn, broken=True)
                raise
            self.conn = conn
            self.cur = cur
            return self.cur

    def __exit__(self, exc, *_):
        cursor_err = False
        try:
            if getattr(self, "cur", None) is not None:
                self.cur.close()
        except Exception as err:
            cursor_err = True
            logger.error("Failed closing DB cursor: %s", _sanitize_db_error(err))
        if self.in_tx:
            if exc is not None or cursor_err:
                _tx_local.aborted = True
            return False
        broken = cursor_err
        try:
            if exc:
                try:
                    self.conn.rollback()
                except Exception as rb_err:
                    broken = True
                    logger.error("DB rollback failed: %s", _sanitize_db_error(rb_err))
            else:
                try:
                    self.conn.commit()
                except Exception as cm_err:
                    broken = True
                    logger.error("DB commit failed: %s", _sanitize_db_error(cm_err))
                    try:
                        self.conn.rollback()
                    except Exception as rb_err:
                        logger.error("DB rollback after commit failure failed: %s", _sanitize_db_error(rb_err))
                    raise
        finally:
            _return_conn_to_pool(self.conn, broken=broken)

def q(sql, params=(), one=False):
    with DB() as cur:
        cur.execute(sql, params)
        rows = cur.fetchone() if one else cur.fetchall()
    if one:  return dict(rows) if rows else None
    return [dict(r) for r in rows] if rows else []

def exe(sql, params=()):
    with DB() as cur:
        cur.execute(sql, params)
        return cur.rowcount

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
CREATE INDEX IF NOT EXISTS idx_records_docno ON records ((data->>'docNo'));
CREATE INDEX IF NOT EXISTS idx_records_status ON records ((data->>'status'));
CREATE INDEX IF NOT EXISTS idx_records_discipline ON records ((data->>'discipline'));
CREATE INDEX IF NOT EXISTS idx_records_issued_date ON records ((data->>'issuedDate'));
CREATE INDEX IF NOT EXISTS idx_records_data_gin ON records USING GIN (data);
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
ALTER TABLE pr_items ADD COLUMN IF NOT EXISTS po_ref TEXT;
ALTER TABLE pr_items ADD COLUMN IF NOT EXISTS po_qty NUMERIC DEFAULT 0;
ALTER TABLE pr_items ADD COLUMN IF NOT EXISTS delivered_qty NUMERIC DEFAULT 0;
ALTER TABLE users ADD COLUMN IF NOT EXISTS email TEXT;
ALTER TABLE user_projects ADD COLUMN IF NOT EXISTS is_dc BOOLEAN DEFAULT FALSE;
CREATE INDEX IF NOT EXISTS idx_pr_items_record ON pr_items(record_id);
CREATE TABLE IF NOT EXISTS app_settings (
    key        TEXT PRIMARY KEY,
    value      JSONB NOT NULL DEFAULT '{}'
);
CREATE TABLE IF NOT EXISTS dashboard_cache (
    project_id TEXT PRIMARY KEY,
    cache_data JSONB NOT NULL,
    updated_at TIMESTAMPTZ DEFAULT NOW()
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
    detail      TEXT,
    ip_address  TEXT DEFAULT NULL,
    user_agent  TEXT DEFAULT NULL,
    device      TEXT DEFAULT NULL
);
ALTER TABLE audit_log ADD COLUMN IF NOT EXISTS ip_address TEXT DEFAULT NULL;
ALTER TABLE audit_log ADD COLUMN IF NOT EXISTS user_agent TEXT DEFAULT NULL;
ALTER TABLE audit_log ADD COLUMN IF NOT EXISTS device TEXT DEFAULT NULL;
CREATE INDEX IF NOT EXISTS idx_audit_ts ON audit_log(ts DESC);
CREATE INDEX IF NOT EXISTS idx_audit_proj ON audit_log(project_id, ts DESC);
CREATE INDEX IF NOT EXISTS idx_audit_user ON audit_log(username, ts DESC);
CREATE INDEX IF NOT EXISTS idx_audit_action ON audit_log(action, ts DESC);

CREATE TABLE IF NOT EXISTS workspaces (
    id          TEXT PRIMARY KEY,
    code        TEXT NOT NULL UNIQUE,
    name        TEXT NOT NULL,
    name_ar     TEXT NOT NULL,
    icon        TEXT DEFAULT '📁',
    type        TEXT NOT NULL DEFAULT 'engineering',
    description TEXT DEFAULT '',
    sort_order  INTEGER DEFAULT 0,
    is_active   BOOLEAN DEFAULT TRUE,
    created_at  TIMESTAMPTZ DEFAULT NOW()
);
ALTER TABLE workspaces ADD COLUMN IF NOT EXISTS description TEXT DEFAULT '';
CREATE TABLE IF NOT EXISTS user_workspaces (
    username     TEXT NOT NULL,
    workspace_id TEXT NOT NULL REFERENCES workspaces(id) ON DELETE CASCADE,
    role         TEXT NOT NULL DEFAULT 'viewer',
    created_at   TIMESTAMPTZ DEFAULT NOW(),
    PRIMARY KEY (username, workspace_id)
);
ALTER TABLE projects ADD COLUMN IF NOT EXISTS workspace_id TEXT DEFAULT 'ENG';
CREATE INDEX IF NOT EXISTS idx_projects_workspace ON projects(workspace_id);
ALTER TABLE users ADD COLUMN IF NOT EXISTS primary_workspace TEXT DEFAULT 'ENG';
ALTER TABLE users ADD COLUMN IF NOT EXISTS can_download BOOLEAN DEFAULT FALSE;
"""

DEFAULT_OM_DOC_TYPES = [
    ("WO", "Work Orders", 0),
    ("PPM", "Preventive Maintenance", 1),
    ("CM", "Corrective Maintenance", 2),
    ("INC", "Incident Reports", 3),
    ("MR", "Maintenance Requests", 4),
    ("SLA", "Service Level Reports", 5),
    ("COR", "Correspondences", 6),
]

DEFAULT_OM_COLS = [
    ("docNo", "Document No.", "docno", None, 0, True),
    ("facility", "Facility / Location", "text", None, 1, True),
    ("equipment", "Equipment / Asset", "text", None, 2, True),
    ("title", "Description / Subject", "text", None, 3, True),
    ("priority", "Priority", "dropdown", "priority", 4, True),
    ("issuedDate", "Issue Date", "date", None, 5, True),
    ("targetDate", "Target Completion", "date", None, 6, True),
    ("actualReplyDate", "Completion Date", "date", None, 7, True),
    ("status", "Status", "dropdown", "om_status", 8, True),
    ("remarks", "Remarks / Notes", "text", None, 9, True),
    ("fileLocation", "Drive Link", "url", None, 10, True),
]

DEFAULT_OM_LISTS = {
    "priority": ["Low", "Medium", "High", "Critical"],
    "om_status": ["Open", "In Progress", "Pending", "Completed", "Closed", "Cancelled"]
}

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
    if text in ("approved", "rejected", "pending", "cancelled", "info_closed"):
        return text
    if text in ("open", "under review", "not closed"):
        return "pending"
    if text in ("closed", "replied", "info", "for information", "info closed"):
        return "info_closed"
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
    for meta in ("cancelled", "rejected", "approved", "info_closed", "pending"):
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
    exe("DELETE FROM columns_config WHERE col_key LIKE %s AND label ILIKE %s", ('custom_%', '%drive link%'))
    exe("""INSERT INTO columns_config(project_id, dt_id, col_key, label, col_type, visible, sort_order)
           SELECT project_id, id, 'fileLocation', 'Drive Link', 'url', TRUE, 100
           FROM doc_types dt
           WHERE NOT EXISTS (
               SELECT 1 FROM columns_config c 
               WHERE c.project_id = dt.project_id AND c.dt_id = dt.id AND c.col_key = 'fileLocation'
           )
           ON CONFLICT DO NOTHING""")

def _sync_all_hse_ptw_subtypes():
    """Ensure active hse_ptw_subtype lists contain the approved permit types and default codes."""
    from hse_defaults import HSE_LISTS, HSE_SUBTYPES, PTW_SUBTYPE_CODES

    exe("DELETE FROM dropdown_lists WHERE list_name=%s AND item_value=%s",
        ("hse_ptw_subtype", "Maintenance Work"))

    hse_projects = q("SELECT id FROM projects WHERE UPPER(COALESCE(workspace_id, 'ENG')) = 'HSE'")
    ptw_subtypes = HSE_SUBTYPES.get("PTW", {})
    default_values = HSE_LISTS.get("hse_ptw_subtype", ())
    for p in hse_projects:
        pid = p["id"]
        ptw_dt = q("SELECT id, data FROM doc_types WHERE project_id=%s AND id='PTW'", (pid,), one=True)
        if ptw_dt:
            dt_data = ptw_dt.get("data") or {}
            if isinstance(dt_data, str):
                try:
                    dt_data = json.loads(dt_data)
                except Exception:
                    dt_data = {}
            if isinstance(dt_data, dict):
                cur_subs = dt_data.get("legacy_subtypes")
                if cur_subs != ptw_subtypes:
                    dt_data["legacy_subtypes"] = ptw_subtypes
                    exe("UPDATE doc_types SET data=%s::jsonb WHERE id='PTW' AND project_id=%s",
                        (json.dumps(dt_data, ensure_ascii=False), pid))
        for order, value in enumerate(default_values):
            code = PTW_SUBTYPE_CODES.get(value)
            exe("""
                INSERT INTO dropdown_lists(project_id, list_name, item_value, sort_order, meta)
                VALUES(%s, 'hse_ptw_subtype', %s, %s, %s)
                ON CONFLICT(project_id, list_name, item_value)
                DO UPDATE SET meta = CASE
                    WHEN dropdown_lists.meta IS NULL
                      OR TRIM(dropdown_lists.meta) = ''
                      OR UPPER(TRIM(dropdown_lists.meta)) IN ('PENDING','APPROVED','REJECTED','CANCELLED','INFO_CLOSED')
                    THEN EXCLUDED.meta
                    ELSE dropdown_lists.meta
                END
            """, (pid, value, order, code))


def init():
    stmts = [s.strip() for s in SCHEMA.strip().split(";") if s.strip()]
    for stmt in stmts:
        try:
            with DB() as cur: cur.execute(stmt)
        except Exception as e:
            if "already exists" not in str(e).lower():
                logger.warning("db_schema_note error=%s", e)
    ensure_audit_schema()
    _ensure_admin()
    _ensure_workspaces()
    _cleanup_custom_columns()
    _global_rename_filelocation()
    _sync_all_noc_doc_types()
    _sync_all_ltr_doc_types()
    _sync_all_hse_ptw_subtypes()

def _ensure_admin():
    r = q("SELECT COUNT(*) as c FROM users", one=True)
    if not r or r["c"] == 0:
        exe("INSERT INTO users(username,pw_hash,role) VALUES(%s,%s,%s) ON CONFLICT DO NOTHING",
            ("admin", hash_pw("admin123"), "superadmin"))

def _ensure_workspaces():
    try:
        exe("""INSERT INTO workspaces(id, code, name, name_ar, icon, type, sort_order, is_active, description)
               VALUES('ENG', 'ENG', 'Engineering & Projects', 'المشروعات الهندسية', '🏗️', 'engineering', 1, TRUE, 'EPC & Site Submittals')
               ON CONFLICT (id) DO UPDATE SET code = 'ENG', name = 'Engineering & Projects'""")
        exe("""INSERT INTO workspaces(id, code, name, name_ar, icon, type, sort_order, is_active, description)
               VALUES('OM', 'O&M', 'Operation & Maintenance (O&M)', 'الصيانة والتشغيل', '⚙️', 'om', 2, TRUE, 'Facilities & Maintenance Contracts')
               ON CONFLICT (id) DO UPDATE SET code = 'O&M', name = 'Operation & Maintenance (O&M)'""")
        exe("UPDATE workspaces SET description = 'EPC & Site Submittals' WHERE id = 'ENG' AND description IS NULL")
        exe("UPDATE workspaces SET description = 'Facilities & Maintenance Contracts' WHERE id = 'OM' AND description IS NULL")
        exe("UPDATE workspaces SET sort_order = 1 WHERE id = 'ENG'")
        exe("UPDATE workspaces SET sort_order = 2 WHERE id = 'OM'")
        default_depts = [
            ('HR', 'HR', 'Human Resources', 'الموارد البشرية', '👥', 'hr', 3, 'Personnel archives, employee onboarding, credential files, attendance and corporate policies.'),
            ('FIN', 'FIN', 'Finance & Accounting', 'المالية والحسابات', '💰', 'finance', 4, 'Interim Payment Certificates (IPC), subcontractor invoices, project budgets and cost audits.'),
            ('PROC', 'PROC', 'Procurement', 'المشتريات والتعاقدات', '🛒', 'procurement', 5, 'Purchase Orders (PO), vendor quotations, material requisition packages and approved suppliers.'),
            ('WH', 'WH', 'Warehousing', 'المخازن والمستودعات', '🏬', 'warehousing', 6, 'Central store inventory, Material Receipt Notes (MRN), site dispatch and stock balances.'),
            ('HSE', 'HSE', 'Health, Safety & Environment (HSE)', 'السلامة والصحة المهنية', '🛡️', 'hse', 7, 'Occupational safety audits, site inspection checklists, incident registers and ISO certifications.'),
            ('TND', 'TND', 'Tendering', 'العطاءات والمناقصات', '📑', 'tendering', 8, 'Tender submissions, BOQ bill of quantities, tender addendums, commercial bids and award tracking.'),
        ]
        for ws_id, code, name, name_ar, icon, ws_type, s_ord, desc in default_depts:
            existing = q(
                "SELECT id, description, icon FROM workspaces WHERE UPPER(id) = %s OR UPPER(code) = %s OR LOWER(name) = LOWER(%s) LIMIT 1",
                (ws_id, code, name),
                one=True
            )
            if not existing:
                exe("""INSERT INTO workspaces(id, code, name, name_ar, icon, type, sort_order, is_active, description)
                       VALUES(%s, %s, %s, %s, %s, %s, %s, TRUE, %s)
                       ON CONFLICT (id) DO NOTHING""",
                    (ws_id, code, name, name_ar, icon, ws_type, s_ord, desc))
            else:
                if not (existing.get("description") or "").strip():
                    exe("UPDATE workspaces SET description = %s WHERE id = %s", (desc, existing["id"]))
                if not (existing.get("icon") or "").strip() or existing.get("icon") == "📁":
                    exe("UPDATE workspaces SET icon = %s WHERE id = %s", (icon, existing["id"]))
        exe("""
            WITH ranked AS (
                SELECT id, ROW_NUMBER() OVER (ORDER BY created_at, id) + 2 as new_order
                FROM workspaces
                WHERE id NOT IN ('ENG', 'OM')
            )
            UPDATE workspaces w
            SET sort_order = r.new_order
            FROM ranked r
            WHERE w.id = r.id AND (w.sort_order IS NULL OR w.sort_order <= 2)
        """)
        exe("UPDATE projects SET workspace_id = 'ENG' WHERE workspace_id IS NULL OR workspace_id = ''")
        exe("""
            INSERT INTO user_workspaces (username, workspace_id, role)
            SELECT username, 'ENG', 'admin'
            FROM users
            WHERE role = 'superadmin'
            ON CONFLICT DO NOTHING
        """)
        exe("""
            INSERT INTO user_workspaces (username, workspace_id, role)
            SELECT username, 'OM', 'admin'
            FROM users
            WHERE role = 'superadmin'
            ON CONFLICT DO NOTHING
        """)
        exe("""
            INSERT INTO user_workspaces (username, workspace_id, role)
            SELECT username, COALESCE(primary_workspace, 'ENG'), 
                   CASE WHEN role = 'superadmin' THEN 'admin' ELSE role END
            FROM users
            WHERE username NOT IN (SELECT username FROM user_workspaces)
            ON CONFLICT DO NOTHING
        """)

        exe("ALTER TABLE user_projects ADD COLUMN IF NOT EXISTS allowed_doc_types JSONB DEFAULT NULL")
    except Exception as e:
        logger.warning("_ensure_workspaces note: %s", e)

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
    return q("SELECT username, role, email, COALESCE(primary_workspace, 'ENG') as primary_workspace, COALESCE(can_download, FALSE) as can_download FROM users ORDER BY username")

def add_user(username, pw, role="viewer", email=None, primary_workspace="ENG", can_download=False):
    ws = primary_workspace or "ENG"
    exe("INSERT INTO users(username,pw_hash,role,email,primary_workspace,can_download) VALUES(%s,%s,%s,%s,%s,%s) ON CONFLICT DO NOTHING",
        (username, hash_pw(pw), role, email, ws, bool(can_download)))
    assign_user_workspace(username, ws, role)

def delete_user(username):
    exe("DELETE FROM sessions WHERE username=%s", (username,))
    exe("DELETE FROM user_projects WHERE username=%s", (username,))
    exe("DELETE FROM user_workspaces WHERE username=%s", (username,))
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

def set_user_can_download(username, can_download=False):
    exe("UPDATE users SET can_download=%s WHERE username=%s", (bool(can_download), username))
    return True

def can_user_download(username, project_id=None):
    """Check if a user is permitted to download files (workspace-aware when project_id is provided)."""
    if not username:
        return False
    try:
        u = get_user(username)
        if not u or not isinstance(u, dict):
            return False
        global_role = str(u.get("role", "")).strip().lower()
        if global_role == "superadmin":
            return True
        if project_id:
            proj = get_project(project_id) or resolve_project_alias(project_id)
            if not proj:
                return False
            ws_id = proj.get("workspace_id") or "ENG"
            role = get_user_workspace_role(username, ws_id)
        else:
            role = global_role
        if role in ("superadmin", "admin", "editor"):
            return True
        return bool(u.get("can_download", False))
    except Exception:
        return False

def set_user_project_dc(username, project_id, is_dc):
    exe("UPDATE user_projects SET is_dc=%s WHERE username=%s AND project_id=%s", (is_dc, username, project_id))

# ── Workspaces & Modular Scoping ──────────────────────────────
def get_workspaces(active_only=True):
    where = "WHERE is_active = TRUE" if active_only else ""
    return q(f"SELECT id, code, name, name_ar, icon, type, sort_order, is_active, COALESCE(description, '') as description FROM workspaces {where} ORDER BY sort_order, id")

def get_workspace(ws_id):
    if not ws_id: return None
    return q("SELECT id, code, name, name_ar, icon, type, sort_order, is_active, COALESCE(description, '') as description FROM workspaces WHERE id=%s", (ws_id,), one=True)

def create_workspace(ws_id, code, name, name_ar, icon="📁", ws_type="general", sort_order=0, description=""):
    exe("""INSERT INTO workspaces(id, code, name, name_ar, icon, type, sort_order, is_active, description)
           VALUES(%s, %s, %s, %s, %s, %s, %s, TRUE, %s)
           ON CONFLICT(id) DO UPDATE SET code=EXCLUDED.code, name=EXCLUDED.name,
           name_ar=EXCLUDED.name_ar, icon=EXCLUDED.icon, type=EXCLUDED.type, sort_order=EXCLUDED.sort_order,
           description=EXCLUDED.description""",
        (ws_id, code, name, name_ar, icon, ws_type, sort_order, description or ''))

def update_workspace(ws_id, code=None, name=None, name_ar=None, icon=None, sort_order=None, is_active=None, description=None):
    fields = []
    params = []
    if code is not None:
        fields.append("code = %s")
        params.append(code)
    if name is not None:
        fields.append("name = %s")
        params.append(name)
    if name_ar is not None:
        fields.append("name_ar = %s")
        params.append(name_ar)
    if icon is not None:
        fields.append("icon = %s")
        params.append(icon)
    if sort_order is not None:
        fields.append("sort_order = %s")
        params.append(sort_order)
    if is_active is not None:
        fields.append("is_active = %s")
        params.append(bool(is_active))
    if description is not None:
        fields.append("description = %s")
        params.append(description)
    if not fields:
        return
    params.append(ws_id)
    sql = f"UPDATE workspaces SET {', '.join(fields)} WHERE id = %s"
    exe(sql, tuple(params))

def get_user_workspaces(username):
    """Return all workspaces accessible to a user, with their role in each."""
    u = get_user(username)
    if not u:
        return []
    role = str(u.get("role", "")).strip().lower()
    if role == "superadmin":
        all_ws = get_workspaces(active_only=True)
        return [{**ws, "user_role": "admin", "is_superadmin": True} for ws in all_ws]

    rows = q("""
        SELECT w.id, w.code, w.name, w.name_ar, w.icon, w.type, w.sort_order, w.is_active,
               COALESCE(uw.role, u.role, 'viewer') as user_role,
               FALSE as is_superadmin
        FROM workspaces w
        LEFT JOIN user_workspaces uw ON w.id = uw.workspace_id AND uw.username = %s
        LEFT JOIN users u ON u.username = %s
        WHERE w.is_active = TRUE
          AND (uw.username IS NOT NULL OR u.primary_workspace = w.id)
        ORDER BY w.sort_order, w.id
    """, (username, username))
    return rows

def is_user_workspace_admin(username, workspace_id=None):
    """Check if user is admin of a specific workspace, or admin of ANY workspace."""
    u = get_user(username)
    if not u:
        return False
    if str(u.get("role", "")).strip().lower() == "superadmin":
        return True

    if workspace_id:
        ws_upper = str(workspace_id).strip().upper()
        row = q("SELECT role FROM user_workspaces WHERE username=%s AND UPPER(workspace_id)=%s", (username, ws_upper), one=True)
        if row is not None:
            return str(row.get("role") or "").strip().lower() in ("admin", "superadmin")
        primary = str(u.get("primary_workspace") or "ENG").strip().upper()
        if str(u.get("role", "")).strip().lower() in ("admin", "superadmin") and primary == ws_upper:
            return True
        return False
    else:
        return bool(get_admin_workspaces_for_user(username))

def get_admin_workspaces_for_user(username):
    """Return workspace IDs where this user has admin role."""
    u = get_user(username)
    if not u:
        return []
    if str(u.get("role", "")).strip().lower() == "superadmin":
        return [str(w["id"]).upper() for w in get_workspaces(active_only=True)]

    ws_ids = set()
    explicit_roles = {}
    rows = q("SELECT workspace_id, role FROM user_workspaces WHERE username=%s", (username,))
    for r in (rows or []):
        if r.get("workspace_id"):
            wid = str(r["workspace_id"]).strip().upper()
            role = str(r.get("role", "")).strip().lower()
            explicit_roles[wid] = role
            if role in ("admin", "superadmin"):
                ws_ids.add(wid)

    if str(u.get("role", "")).strip().lower() in ("admin", "superadmin"):
        pw = (u.get("primary_workspace") or "ENG").upper()
        if pw not in explicit_roles or explicit_roles[pw] in ("admin", "superadmin"):
            ws_ids.add(pw)
    return list(ws_ids)

def assign_user_workspace(username, workspace_id, role="viewer"):
    role = str(role or "viewer").strip().lower()
    if role not in ("admin", "editor", "viewer"):
        role = "viewer"
    exe("""INSERT INTO user_workspaces(username, workspace_id, role)
           VALUES(%s, %s, %s)
           ON CONFLICT (username, workspace_id) DO UPDATE SET role = EXCLUDED.role""",
        (username, workspace_id, role))

def get_user_workspace_role(username, workspace_id):
    """Return user's specific role in workspace_id without leaking users.role across workspaces."""
    u = get_user(username)
    if not u:
        return "viewer"
    if str(u.get("role", "")).strip().lower() == "superadmin":
        return "admin"
    ws_upper = str(workspace_id or "").strip().upper()
    if not ws_upper:
        return "viewer"
    row = q("SELECT role FROM user_workspaces WHERE username=%s AND UPPER(workspace_id)=%s", (username, ws_upper), one=True)
    if row is not None:
        return str(row.get("role") or "viewer").strip().lower()
    primary = str(u.get("primary_workspace") or "ENG").strip().upper()
    if ws_upper == primary:
        return str(u.get("role") or "viewer").strip().lower()
    return "viewer"

def set_user_workspace_role(username, workspace_id, role):
    role = str(role or "viewer").strip().lower()
    if role not in ("admin", "editor", "viewer"):
        role = "viewer"
    exe("""INSERT INTO user_workspaces(username, workspace_id, role)
           VALUES(%s, %s, %s)
           ON CONFLICT (username, workspace_id) DO UPDATE SET role = EXCLUDED.role""",
        (username, workspace_id, role))

def remove_user_workspace(username, workspace_id):
    exe("DELETE FROM user_workspaces WHERE username=%s AND workspace_id=%s", (username, workspace_id))

def set_user_primary_workspace(username, ws_id):
    exe("UPDATE users SET primary_workspace=%s WHERE username=%s", (ws_id, username))

def set_user_workspaces(username, workspace_ids, default_role="viewer"):
    """
    Synchronize user's assigned workspaces.
    - Adds newly assigned workspaces.
    - Deletes revoked workspaces.
    - Cascades deletion of user_projects for revoked workspaces.
    - Updates primary_workspace if current primary was revoked.
    """
    ws_ids = [str(w).strip().upper() for w in workspace_ids if str(w).strip()]
    if not ws_ids:
        return False

    u = get_user(username)
    if not u:
        return False

    curr_primary = u.get("primary_workspace") or "ENG"
    if curr_primary not in ws_ids:
        new_primary = ws_ids[0]
        exe("UPDATE users SET primary_workspace=%s WHERE username=%s", (new_primary, username))

    for wid in ws_ids:
        exe("""
            INSERT INTO user_workspaces (username, workspace_id, role)
            VALUES (%s, %s, %s)
            ON CONFLICT (username, workspace_id) DO NOTHING
        """, (username, wid, default_role))

    exe("""
        DELETE FROM user_workspaces
        WHERE username=%s AND NOT (workspace_id = ANY(%s))
    """, (username, ws_ids))

    exe("""
        DELETE FROM user_projects
        WHERE username=%s
          AND project_id IN (
              SELECT id FROM projects WHERE NOT (workspace_id = ANY(%s))
          )
    """, (username, ws_ids))
    return True

def get_users_for_admin(admin_username):
    """
    If admin_username is superadmin: returns all users.
    If admin_username is a Department Admin: returns only users belonging to the admin's workspace(s),
    excluding superadmins.
    """
    u = get_user(admin_username)
    if not u:
        return []
    if str(u.get("role", "")).strip().lower() == "superadmin":
        return q("""
            SELECT u.username, u.role, u.email, COALESCE(u.primary_workspace, 'ENG') as primary_workspace,
                   COALESCE(u.can_download, FALSE) as can_download,
                   COALESCE(
                       (SELECT json_agg(json_build_object('workspace_id', uw.workspace_id, 'role', uw.role))
                        FROM user_workspaces uw WHERE uw.username = u.username), '[]'::json
                   ) as workspaces
            FROM users u
            ORDER BY u.username
        """)

    admin_ws = get_admin_workspaces_for_user(admin_username)
    if not admin_ws:
        return []

    rows = q("""
        SELECT u.username, u.role, u.email, COALESCE(u.primary_workspace, 'ENG') as primary_workspace,
               COALESCE(u.can_download, FALSE) as can_download,
               COALESCE(
                   (SELECT json_agg(json_build_object('workspace_id', uw.workspace_id, 'role', uw.role))
                    FROM user_workspaces uw WHERE uw.username = u.username), '[]'::json
               ) as workspaces
        FROM users u
        WHERE u.role != 'superadmin'
          AND (
              u.primary_workspace = ANY(%s)
              OR EXISTS (
                  SELECT 1 FROM user_workspaces uw2
                  WHERE uw2.username = u.username AND uw2.workspace_id = ANY(%s)
              )
          )
        ORDER BY u.username
    """, (admin_ws, admin_ws))
    return rows

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
    # Returns a list of dicts: {"project_id": "...", "is_dc": True/False, "allowed_doc_types": [...]}
    out = []
    for r in (q("SELECT project_id, is_dc, allowed_doc_types FROM user_projects WHERE username=%s", (username,)) or []):
        entry = {"project_id": r["project_id"], "is_dc": r.get("is_dc", False)}
        if "allowed_doc_types" in r:
            entry["allowed_doc_types"] = r["allowed_doc_types"]
        out.append(entry)
    return out

def get_all_user_projects():
    rows = q("SELECT username, project_id, is_dc, allowed_doc_types FROM user_projects")
    res = {}
    for r in rows:
        res.setdefault(r["username"], []).append({
            "project_id": r["project_id"],
            "is_dc": r.get("is_dc", False),
            "allowed_doc_types": r.get("allowed_doc_types")
        })
    return res

def set_user_project_doc_types(username, project_id, allowed_doc_types):
    if allowed_doc_types is not None and not isinstance(allowed_doc_types, list):
        allowed_doc_types = None
    if isinstance(allowed_doc_types, list):
        allowed_val = json.dumps([str(x) for x in allowed_doc_types])
    else:
        allowed_val = None

    if allowed_val is not None:
        exe("""
            UPDATE user_projects 
            SET allowed_doc_types = %s::jsonb 
            WHERE username=%s AND project_id=%s
        """, (allowed_val, username, project_id))
    else:
        exe("""
            UPDATE user_projects 
            SET allowed_doc_types = NULL 
            WHERE username=%s AND project_id=%s
        """, (username, project_id))
    return True

def _build_project_resolution_index(all_projs=None):
    """Build exact-id, case-insensitive-id, and code/name alias lookup tables across all projects."""
    projs = q("SELECT * FROM projects ORDER BY id") or [] if all_projs is None else list(all_projs)
    by_exact_id = {}
    by_ci_id = {}
    by_alias = {}
    for p in projs:
        canon_id = str(p.get("id") or "").strip()
        if not canon_id:
            continue
        by_exact_id[canon_id] = p
        by_ci_id.setdefault(canon_id.upper(), []).append(p)
        for field in ("code", "name"):
            val = str(p.get(field) or "").strip().upper()
            if val:
                by_alias.setdefault(val, set()).add(canon_id)
    return by_exact_id, by_ci_id, by_alias


def _resolve_single_project_from_index(raw_identifier, by_exact_id, by_ci_id, by_alias):
    """Resolve one project ID or unambiguous legacy alias using canonical DB rules."""
    raw = str(raw_identifier or "").strip()
    if not raw:
        return None
    if raw in by_exact_id:
        return by_exact_id[raw]
    raw_up = raw.upper()
    ci_matches = by_ci_id.get(raw_up) or []
    if len(ci_matches) == 1:
        return ci_matches[0]
    if len(ci_matches) > 1:
        return None
    alias_ids = by_alias.get(raw_up) or set()
    if len(alias_ids) == 1:
        return by_exact_id.get(next(iter(alias_ids)))
    return None


def _intersect_doc_type_lists(restricted_lists, doc_types):
    """Conservatively intersect multiple restricted allowed_doc_types lists for one project."""
    if not restricted_lists:
        return []
    for lst in restricted_lists:
        if not lst:
            return []

    dt_list = list(doc_types or [])
    dt_by_token_up = {}
    ordered_dt_ids = []
    for dt in dt_list:
        dt_id = str(dt.get("id") or "").strip()
        dt_code = str(dt.get("code") or "").strip()
        canon_dt = dt_id or dt_code
        if not canon_dt:
            continue
        if canon_dt not in ordered_dt_ids:
            ordered_dt_ids.append(canon_dt)
        if dt_id:
            dt_by_token_up[dt_id.upper()] = canon_dt
        if dt_code:
            dt_by_token_up[dt_code.upper()] = canon_dt

    per_row_sets = []
    token_display = {}
    for lst in restricted_lists:
        matched_set = set()
        for tok in lst:
            t_str = str(tok or "").strip()
            if not t_str:
                continue
            t_up = t_str.upper()
            canon_tok = dt_by_token_up.get(t_up)
            if canon_tok:
                norm_key = canon_tok.upper()
                matched_set.add(norm_key)
                token_display.setdefault(norm_key, canon_tok)
            else:
                matched_set.add(t_up)
                token_display.setdefault(t_up, t_str)
        if not matched_set:
            return []
        per_row_sets.append(matched_set)

    common = set.intersection(*per_row_sets)
    if not common:
        return []
    result = []
    seen = set()
    for canon_dt in ordered_dt_ids:
        key = canon_dt.upper()
        if key in common and key not in seen:
            seen.add(key)
            result.append(canon_dt)
    for tok in restricted_lists[0]:
        t_str = str(tok or "").strip()
        if not t_str:
            continue
        canon_tok = dt_by_token_up.get(t_str.upper()) or t_str
        key = canon_tok.upper()
        if key in common and key not in seen:
            seen.add(key)
            result.append(token_display.get(key, canon_tok))
    return result


def resolve_user_project_grant(username, project_id, doc_types=None, all_projs=None, user_rows=None):
    """Resolve canonical project membership and conservative effective doc-type/DC grants.

    - Canonical `projects.id` and unambiguous legacy `code`/`name` assignment rows in
      `user_projects` are resolved together against the project catalog.
    - Ambiguous project identifiers or missing assignment rows fail closed (`assigned=False`,
      `allowed_doc_types=[]`, never unrestricted `None`).
    - When canonical and legacy rows coexist on the same `(username, project)`, a specific
      restricted list overrides unrestricted `None`, and multiple restricted lists intersect
      conservatively (returning `[]` when disjoint).
    """
    denied = {
        "assigned": False,
        "canonical_project_id": None,
        "matched_rows": [],
        "is_dc": False,
        "allowed_doc_types": [],
        "unrestricted": False,
    }
    if not username or not project_id:
        return denied

    by_exact_id, by_ci_id, by_alias = _build_project_resolution_index(all_projs=all_projs)
    if not by_exact_id:
        return denied
    canon_proj = _resolve_single_project_from_index(project_id, by_exact_id, by_ci_id, by_alias)
    if not canon_proj:
        return denied
    canon_pid = canon_proj["id"]

    rows = get_user_projects(username) if user_rows is None else list(user_rows)
    matched_rows = []
    for r in (rows or []):
        raw_pid = r.get("project_id")
        if not raw_pid:
            continue
        r_proj = _resolve_single_project_from_index(raw_pid, by_exact_id, by_ci_id, by_alias)
        if r_proj and r_proj["id"] == canon_pid:
            matched_rows.append(r)

    if not matched_rows:
        return {
            "assigned": False,
            "canonical_project_id": canon_pid,
            "matched_rows": [],
            "is_dc": False,
            "allowed_doc_types": [],
            "unrestricted": False,
        }

    is_dc = any(bool(r.get("is_dc")) for r in matched_rows)
    parsed_grants = []
    has_restriction = False
    for r in matched_rows:
        if "allowed_doc_types" in r:
            val = r.get("allowed_doc_types")
        else:
            # Compatibility fallback for minimal unit-test mocks that return project_id only in get_user_projects()
            sub = q("SELECT allowed_doc_types FROM user_projects WHERE username=%s AND project_id=%s",
                    (username, r["project_id"]), one=True)
            if not sub:
                has_restriction = True
                parsed_grants.append([])
                continue
            val = sub.get("allowed_doc_types")

        if val is None:
            parsed_grants.append(None)
        elif isinstance(val, str):
            try:
                decoded = json.loads(val)
            except Exception:
                decoded = []
            has_restriction = True
            if isinstance(decoded, list):
                parsed_grants.append([str(x).strip() for x in decoded if str(x).strip()])
            else:
                parsed_grants.append([])
        elif isinstance(val, list):
            has_restriction = True
            parsed_grants.append([str(x).strip() for x in val if str(x).strip()])
        else:
            has_restriction = True
            parsed_grants.append([])

    if not has_restriction:
        return {
            "assigned": True,
            "canonical_project_id": canon_pid,
            "matched_rows": matched_rows,
            "is_dc": is_dc,
            "allowed_doc_types": None,
            "unrestricted": True,
        }

    restricted_lists = [g for g in parsed_grants if g is not None]
    proj_dts = doc_types if doc_types is not None else (get_doc_types(canon_pid) or [])
    effective_dts = _intersect_doc_type_lists(restricted_lists, proj_dts)
    return {
        "assigned": True,
        "canonical_project_id": canon_pid,
        "matched_rows": matched_rows,
        "is_dc": is_dc,
        "allowed_doc_types": effective_dts,
        "unrestricted": False,
    }


def get_user_project_doc_types(username, project_id):
    return resolve_user_project_grant(username, project_id)["allowed_doc_types"]

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
    """Return True if the given user is the DC for the project (canonical ID or unambiguous legacy alias)."""
    return bool(resolve_user_project_grant(username, project_id)["is_dc"])

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

def get_daily_digest(pid, doc_type_ids, target_date=None):
    from datetime import date
    today_str = target_date if target_date else date.today().isoformat()
    if not doc_type_ids: return {"received": [], "issued": [], "replied": []}
    
    # Query only records that match today's date in the relevant fields using PostgreSQL JSONB operators
    sql = """
        SELECT id, project_id, dt_id, data FROM records 
        WHERE project_id=%s AND dt_id = ANY(%s)
        AND data IS NOT NULL
        AND jsonb_typeof(data) = 'object'
        AND (
            SUBSTRING(data->>'receivedDate', 1, 10) = %s OR
            SUBSTRING(data->>'issuedDate', 1, 10) = %s OR
            SUBSTRING(data->>'partAIssueDate', 1, 10) = %s OR
            SUBSTRING(data->>'partCIssueDate', 1, 10) = %s OR
            SUBSTRING(data->>'actualReplyDate', 1, 10) = %s OR
            SUBSTRING(data->>'partBReturnDate', 1, 10) = %s OR
            SUBSTRING(data->>'partDReturnDate', 1, 10) = %s
        )
    """
    params = (pid, list(doc_type_ids), today_str, today_str, today_str, today_str, today_str, today_str, today_str)
    rows = q(sql, params)
    
    received, issued, replied = [], [], []
    for r in rows:
        d = r.get("data") or {}
        if str(d.get("receivedDate"))[:10] == today_str: received.append(r)
        if str(d.get("issuedDate"))[:10] == today_str or str(d.get("partAIssueDate"))[:10] == today_str or str(d.get("partCIssueDate"))[:10] == today_str: issued.append(r)
        if str(d.get("actualReplyDate"))[:10] == today_str or str(d.get("partBReturnDate"))[:10] == today_str or str(d.get("partDReturnDate"))[:10] == today_str: replied.append(r)
        
    project_names = {}
    for p in get_projects():
        project_names[p["id"]] = p["name"]
        if p.get("code"): project_names[p["code"]] = p["name"]
        
    dt_rows = q("SELECT id, code, name FROM doc_types WHERE project_id=%s", (pid,))
    dt_map = {r["id"]: r for r in dt_rows}

    def format_rec(rec):
        d = rec.get("data") or {}
        proj_name = project_names.get(rec["project_id"], rec["project_id"])
        dt = dt_map.get(rec.get("dt_id"), {})
        return {
            "id": str(rec.get("id", "")),
            "project_id": str(rec.get("project_id", "")),
            "project_name": str(proj_name),
            "dt_id": str(rec.get("dt_id", "")),
            "docNo": str(d.get("docNo") or d.get("record_id") or "Untitled"),
            "title": str(d.get("title") or d.get("subject") or d.get("nocDescription") or "No Subject"),
            "status": str(d.get("status") or d.get("partDStatus") or d.get("partBStatus") or ""),
            "discipline": str(d.get("discipline") or d.get("trade") or ""),
            "is_letter": _is_ltr_dt(dt.get("code", ""), dt.get("name", "")),
            "direction": str(d.get("direction", "")).strip().lower()
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
    """Insert or update a distribution row, delete if empty."""
    if not emails:
        exe("DELETE FROM project_distribution WHERE project_id=%s AND doc_type_id=%s AND event_type=%s", 
            (project_id, doc_type_id, event_type))
        return
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
        if isinstance(pid, dict):
            pid = pid.get("project_id") or pid.get("id") or pid.get("code") or pid.get("name")
        val = str(pid or "").strip()
        if not val or val in seen:
            continue
        seen.add(val)
        cleaned.append(val)
    return cleaned

def _project_scope_clause(column, pid=None, project_ids=None):
    ids = _clean_project_ids(project_ids)
    if pid:
        if ids is not None and pid not in ids:
            return "WHERE 1=0", []
        return f"WHERE {column}=%s", [pid]
    if ids is None:
        return "WHERE 1=1", []
    if not ids:
        return "WHERE 1=0", []
    ph = ",".join(["%s"] * len(ids))
    return f"WHERE {column} IN ({ph})", ids


def _dt_allowed_in_scope(pid, dt_id, doc_type_scope, dt_code=""):
    """Check whether (pid, dt_id) is permitted under doc_type_scope (None means unrestricted)."""
    if doc_type_scope is None:
        return True
    if pid not in doc_type_scope:
        return False
    allowed = doc_type_scope[pid]
    if allowed is None:
        return True
    allowed_set = {str(x).strip().upper() for x in allowed if str(x).strip()}
    if not allowed_set:
        return False
    if str(dt_id or "").strip().upper() in allowed_set:
        return True
    if dt_code and str(dt_code).strip().upper() in allowed_set:
        return True
    return False


# ── Projects ──────────────────────────────────────────────────
def resolve_project_identifiers(raw_identifiers, workspace_id=None):
    """Resolve stored project IDs or legacy aliases into canonical project IDs.

    Rules (Finding 1):
    1. Exact ``projects.id`` matches strictly on canonical ID and never expands to aliases.
    2. Unique case-insensitive ``projects.id`` matches strictly on that canonical ID.
    3. Multiple case-insensitive ID matches are rejected before alias lookup.
    4. Legacy ``code`` / ``name`` aliases resolve ONLY when unambiguous across ALL projects
       (exactly one distinct canonical ``projects.id`` matches). Ambiguous aliases (>= 2 matches)
       are rejected and resolve to nothing.
    """
    raw_list = _clean_project_ids(raw_identifiers)
    if not raw_list:
        return []
    project_index = _build_project_resolution_index()

    ws_up = str(workspace_id).strip().upper() if workspace_id else None
    resolved = []
    seen = set()
    for raw in raw_list:
        canon_proj = _resolve_single_project_from_index(raw, *project_index)
        if not canon_proj:
            continue
        if ws_up:
            p_ws = str(canon_proj.get("workspace_id") or "ENG").strip().upper()
            if p_ws != ws_up:
                continue
        cid = canon_proj["id"]
        if cid not in seen:
            seen.add(cid)
            resolved.append(cid)
    return resolved


def resolve_project_alias(project_id):
    """Resolve one project identifier (canonical id or unambiguous code/name) to its normalized dict."""
    pid = str(project_id or "").strip()
    if not pid:
        return None
    exact = get_project(pid)
    if exact:
        return exact
    resolved = resolve_project_identifiers([pid])
    if len(resolved) == 1:
        return get_project(resolved[0])
    return None


def get_projects(project_ids=None, workspace_id=None):
    order_sql = """ORDER BY
        CASE WHEN (data->>'dashboard_order') ~ '^-?[0-9]+$' THEN (data->>'dashboard_order')::int ELSE 999999 END,
        id"""
    ids = _clean_project_ids(project_ids)
    where_clauses = []
    params = []

    if workspace_id:
        where_clauses.append("COALESCE(workspace_id, 'ENG') = %s")
        params.append(workspace_id)

    if ids is not None:
        if not ids:
            return []
        ph = ",".join(["%s"] * len(ids))
        where_clauses.append(f"id IN ({ph})")
        params.extend(ids)

    where_sql = ("WHERE " + " AND ".join(where_clauses)) if where_clauses else ""
    return q(f"SELECT * FROM projects {where_sql} {order_sql}", tuple(params))


def _extract_dashboard_order(data):
    if isinstance(data, str):
        try:
            data = json.loads(data)
        except Exception:
            return None
    if not isinstance(data, dict):
        return None
    val = data.get("dashboard_order")
    if isinstance(val, bool) or val is None:
        return None
    s = str(val).strip()
    if re.match(r"^-?[0-9]+$", s):
        return int(s)
    return None


def _compute_admin_slot_orders(merged_rows, admin_ids):
    """Compute integer `dashboard_order` values for admin projects while pinning fixed projects.

    Fixed (read-only) rows keep their existing `dashboard_order` (effective SQL sort key `999999`
    when unset). Each contiguous run of admin rows between two fixed rows receives strictly
    increasing integers strictly between the adjacent fixed sort keys (reusing existing admin slot
    integers when they already fit that exact interval).
    """
    admin_set = set(admin_ids)
    n = len(merged_rows)
    if not any(r["id"] not in admin_set for r in merged_rows):
        return {r["id"]: idx for idx, r in enumerate(merged_rows)}

    def _fixed_key(row):
        v = _extract_dashboard_order(row.get("data"))
        return v if v is not None else 999999

    assigned = {}
    i = 0
    while i < n:
        if merged_rows[i]["id"] not in admin_set:
            i += 1
            continue
        run_start = i
        while i < n and merged_rows[i]["id"] in admin_set:
            i += 1
        run = merged_rows[run_start:i]
        left_key = _fixed_key(merged_rows[run_start - 1]) if run_start > 0 else None
        right_key = _fixed_key(merged_rows[i]) if i < n else None

        existing_vals = sorted(
            v for v in (_extract_dashboard_order(r.get("data")) for r in run) if v is not None
        )
        can_reuse = (
            len(existing_vals) == len(run)
            and len(set(existing_vals)) == len(run)
            and all(
                (left_key is None or v > left_key) and (right_key is None or v < right_key)
                for v in existing_vals
            )
        )
        if can_reuse:
            vals = existing_vals
        else:
            m = len(run)
            if left_key is None and right_key is None:
                vals = list(range(m))
            elif left_key is None:
                vals = list(range(m)) if right_key >= m else [right_key - m + k for k in range(m)]
            elif right_key is None:
                vals = [left_key + 1 + k for k in range(m)]
            else:
                gap = right_key - left_key - 1
                if gap >= m:
                    vals = [left_key + 1 + k for k in range(m)]
                else:
                    vals = [left_key if gap <= 0 else left_key + min(k + 1, gap) for k in range(m)]
        for r, val in zip(run, vals):
            assigned[r["id"]] = int(val)
    return assigned


def reorder_projects(order, visible_projects=None):
    """Persist relative order of admin projects atomically via `jsonb_set`, preserving fixed positions."""
    clean_order = []
    seen = set()
    for v in (order or []):
        pid = str(v or "").strip()
        if pid and pid not in seen:
            seen.add(pid)
            clean_order.append(pid)
    if not clean_order:
        return []

    scope_ids = list(clean_order)
    if visible_projects:
        for vp in visible_projects:
            vid = str(vp.get("id") if isinstance(vp, dict) else vp or "").strip()
            if vid and vid not in seen:
                seen.add(vid)
                scope_ids.append(vid)

    with transaction():
        ph = ",".join(["%s"] * len(scope_ids))
        locked_rows = q(
            f"SELECT id, data FROM projects WHERE id IN ({ph}) ORDER BY id FOR UPDATE",
            tuple(sorted(scope_ids)),
        ) or []
        if not locked_rows:
            return []

        by_id = {r["id"]: r for r in locked_rows}
        admin_in_db = [pid for pid in clean_order if pid in by_id]
        if not admin_in_db:
            return []
        admin_set = set(admin_in_db)

        current_sorted = sorted(
            locked_rows,
            key=lambda r: (
                v if (v := _extract_dashboard_order(r.get("data"))) is not None else 999999,
                r["id"],
            ),
        )
        merged_rows = []
        admin_iter = iter(by_id[pid] for pid in admin_in_db)
        for row in current_sorted:
            if row["id"] in admin_set:
                merged_rows.append(next(admin_iter))
            else:
                merged_rows.append(row)

        slot_orders = _compute_admin_slot_orders(merged_rows, admin_set)
        for pid in admin_in_db:
            order_val = slot_orders[pid]
            exe(
                """UPDATE projects
                   SET data = jsonb_set(
                       CASE WHEN data IS NOT NULL AND jsonb_typeof(data) = 'object' THEN data ELSE '{}'::jsonb END,
                       '{dashboard_order}',
                       to_jsonb(%s::int),
                       true
                   )
                   WHERE id = %s""",
                (int(order_val), pid),
            )
        return [r["id"] for r in merged_rows]

def get_projects_for_user(username, role=None, workspace_id=None):
    user_role = str(role or "").strip().lower()
    if not user_role and username:
        u = get_user(username)
        user_role = str((u or {}).get("role") or "").strip().lower()
    if user_role == "superadmin":
        return get_projects(workspace_id=workspace_id)
    admin_ws = [str(w).strip().upper() for w in (get_admin_workspaces_for_user(username) or []) if str(w).strip()]
    ws_up = str(workspace_id).strip().upper() if workspace_id else None
    if admin_ws:
        if ws_up:
            if ws_up in admin_ws:
                return get_projects(workspace_id=workspace_id)
            user_projs = get_user_projects(username)
            raw_pids = [p["project_id"] for p in user_projs if p.get("project_id")]
            pids = resolve_project_identifiers(raw_pids, workspace_id=workspace_id)
            if not pids:
                return []
            return get_projects(pids, workspace_id=workspace_id)
        all_projs = []
        seen_ids = set()
        for aws in admin_ws:
            for p in get_projects(workspace_id=aws):
                if p["id"] not in seen_ids:
                    seen_ids.add(p["id"])
                    all_projs.append(p)
        user_projs = get_user_projects(username)
        raw_pids = [p["project_id"] for p in user_projs if p.get("project_id")]
        pids = [cid for cid in resolve_project_identifiers(raw_pids) if cid not in seen_ids]
        if pids:
            for p in get_projects(pids):
                if p["id"] not in seen_ids:
                    seen_ids.add(p["id"])
                    all_projs.append(p)
        all_projs.sort(
            key=lambda p: (
                v if (v := _extract_dashboard_order(p.get("data"))) is not None else 999999,
                str(p.get("id") or ""),
            )
        )
        return all_projs
    user_projs = get_user_projects(username)
    raw_pids = [p["project_id"] for p in user_projs if p.get("project_id")]
    pids = resolve_project_identifiers(raw_pids, workspace_id=workspace_id)
    if not pids:
        return []
    return get_projects(pids, workspace_id=workspace_id)

def get_project(pid):
    r = q("SELECT * FROM projects WHERE id=%s", (pid,), one=True)
    if not r: return None
    data = r.get("data") or {}
    if isinstance(data, str):
        try: data = json.loads(data)
        except: data = {}
    if not isinstance(data, dict): data = {}
    ws_id = r.get("workspace_id") or data.get("workspace_id") or "ENG"
    # Never let data override the core fields
    data.pop("id", None); data.pop("name", None); data.pop("code", None); data.pop("workspace_id", None)
    return {"id": r["id"], "name": r["name"], "code": r["code"], "workspace_id": ws_id, **data}

def update_project_settings(pid, settings_dict):
    r = q("SELECT data FROM projects WHERE id=%s", (pid,), one=True)
    if not r: return False
    data = r.get("data") or {}
    if isinstance(data, str):
        try: data = json.loads(data)
        except: data = {}
    if not isinstance(data, dict): data = {}
    data.update(settings_dict)
    exe("UPDATE projects SET data=%s::jsonb WHERE id=%s", (json.dumps(data), pid))
    return True

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

class ProjectOwnershipError(Exception):
    """A project create or update collided with an existing ID or mismatched workspace."""


def save_project(pid, name, code, extra: dict, workspace_id=None):
    if isinstance(extra, dict):
        extra = dict(extra)
        extra.pop("id", None); extra.pop("name", None); extra.pop("code", None); extra.pop("workspace_id", None)
    else:
        extra = {}
    current = q("SELECT data, workspace_id FROM projects WHERE id=%s", (pid,), one=True)
    current_data = current.get("data") if current else {}
    if isinstance(current_data, str):
        try: current_data = json.loads(current_data)
        except Exception: current_data = {}
    if isinstance(current_data, dict) and "dashboard_order" in current_data and "dashboard_order" not in extra:
        extra["dashboard_order"] = current_data.get("dashboard_order")
    jdata = json.dumps(extra)
    stored_ws = str(current.get("workspace_id") or "ENG").strip().upper() if current else None
    req_ws = str(workspace_id).strip().upper() if (workspace_id is not None and str(workspace_id).strip()) else None
    if stored_ws and req_ws and req_ws != stored_ws:
        raise ProjectOwnershipError(
            "Project %r belongs to workspace %r and cannot be moved to %r" % (pid, stored_ws, req_ws)
        )
    ws_id = stored_ws or req_ws or "ENG"
    # Finding A2: metadata updates must never silently transfer a project to another workspace.
    affected = exe("""INSERT INTO projects(id,name,code,data,workspace_id) VALUES(%s,%s,%s,%s::jsonb,%s)
           ON CONFLICT(id) DO UPDATE SET name=EXCLUDED.name, code=EXCLUDED.code, data=EXCLUDED.data
           WHERE COALESCE(projects.workspace_id, 'ENG') = EXCLUDED.workspace_id""",
        (pid, name, code, jdata, ws_id))
    if affected == 0:
        raise ProjectOwnershipError(
            "Project %r does not belong to workspace %r" % (pid, ws_id)
        )


def create_project(pid, name, code, creator=None, workspace_id="ENG", client=None):
    ws_id = str(workspace_id or "ENG").strip().upper() or "ENG"
    jdata = json.dumps({"client": client}) if client else "{}"
    # Finding A2 + Finding 3: insert, seed doc types/columns/lists, and assign creator in a
    # single shared transaction so any failure rolls back cleanly without orphan projects, while
    # preserving ON CONFLICT DO NOTHING duplicate protection and offline test compatibility.
    with transaction():
        affected = exe("""INSERT INTO projects(id,name,code,data,workspace_id)
               VALUES(%s,%s,%s,%s,%s)
               ON CONFLICT (id) DO NOTHING""",
            (pid, name, code, jdata, ws_id))
        if affected == 0:
            raise ProjectOwnershipError("Project ID %r already exists" % (pid,))
        _seed_project(pid, ws_id)
        if creator:
            assign_project(creator, pid)

def delete_project(pid):
    for tbl in ("records","columns_config","doc_types","dropdown_lists","logos","user_projects","projects"):
        exe(f"DELETE FROM {tbl} WHERE {'id' if tbl=='projects' else 'project_id'}=%s", (pid,))

def _seed_hse_project(pid):
    from hse_defaults import HSE_LISTS, HSE_REGISTERS, HSE_SUBTYPES, PTW_SUBTYPE_CODES, hse_columns

    for order, (code, name, source) in enumerate(HSE_REGISTERS):
        data = {"legacy_source": source, "legacy_subtypes": HSE_SUBTYPES.get(code, {})}
        exe(
            """INSERT INTO doc_types(id,project_id,name,code,sort_order,data)
               VALUES(%s,%s,%s,%s,%s,%s::jsonb) ON CONFLICT DO NOTHING""",
            (code, pid, name, code, order, json.dumps(data, ensure_ascii=False)),
        )
        for sort_order, (key, label, col_type, list_name) in enumerate(hse_columns(code)):
            exe(
                """INSERT INTO columns_config(project_id,dt_id,col_key,label,col_type,list_name,visible,sort_order)
                   VALUES(%s,%s,%s,%s,%s,%s,true,%s) ON CONFLICT DO NOTHING""",
                (pid, code, key, label, col_type, list_name, sort_order),
            )
    for list_name, values in HSE_LISTS.items():
        for order, value in enumerate(values):
            meta = None
            if list_name == "hse_doc_status":
                meta = "pending" if value == "Open" else ("approved" if value == "Closed" else None)
            elif list_name == "hse_ptw_subtype":
                meta = PTW_SUBTYPE_CODES.get(value)
            elif list_name == "hse_extinguisher_status":
                meta = (
                    "approved" if value == "Valid"
                    else ("rejected" if value in ("Expired", "Out of Service") else "pending")
                )
            elif list_name == "hse_jsa_status":
                meta = "approved" if value == "Approved" else "pending"
            exe(
                """INSERT INTO dropdown_lists(project_id,list_name,item_value,sort_order,meta)
                   VALUES(%s,%s,%s,%s,%s) ON CONFLICT(project_id,list_name,item_value)
                   DO UPDATE SET meta=COALESCE(dropdown_lists.meta, EXCLUDED.meta)""",
                (pid, list_name, value, order, meta),
            )


def _seed_project(pid, workspace_id="ENG"):
    if str(workspace_id or "").upper() == "HSE":
        _seed_hse_project(pid)
        return
    if str(workspace_id or "").upper() == "OM":
        for code, name, order in DEFAULT_OM_DOC_TYPES:
            exe("INSERT INTO doc_types(id,project_id,name,code,sort_order) VALUES(%s,%s,%s,%s,%s) ON CONFLICT DO NOTHING",
                (code, pid, name, code, order))
            for ck, lbl, ct, ln, so, visible in DEFAULT_OM_COLS:
                exe("""INSERT INTO columns_config(project_id,dt_id,col_key,label,col_type,list_name,visible,sort_order)
                       VALUES(%s,%s,%s,%s,%s,%s,%s,%s) ON CONFLICT DO NOTHING""",
                    (pid, code, ck, lbl, ct, ln, visible, so))
        for ln, items in DEFAULT_OM_LISTS.items():
            for i, item in enumerate(items):
                meta = "pending" if item in ("Open", "In Progress", "Pending") else ("approved" if item in ("Completed", "Closed") else None)
                exe("INSERT INTO dropdown_lists(project_id,list_name,item_value,sort_order,meta)"
                    " VALUES(%s,%s,%s,%s,%s) ON CONFLICT(project_id,list_name,item_value)"
                    " DO UPDATE SET meta=COALESCE(dropdown_lists.meta, EXCLUDED.meta)",
                    (pid, ln, item, i, meta))
    else:
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
        for dt in q("SELECT id, name, code FROM doc_types WHERE project_id=%s", (pid,)):
            if _is_noc_dt(dt["code"], dt["name"]):
                _sync_noc_doc_type(pid, dt["id"], dt.get("name", "Notice of Change"))
            if _is_ltr_dt(dt["code"], dt["name"]):
                _sync_ltr_doc_type(pid, dt["id"], dt.get("name", "Letters"))

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
def project_has_doc_type(pid, code):
    """Return True when project `pid` has a doc_type configured matching `code`."""
    if not pid or not code:
        return False
    try:
        r = q(
            "SELECT 1 FROM doc_types WHERE project_id=%s AND UPPER(code)=%s LIMIT 1",
            (pid, str(code).strip().upper()),
            one=True,
        )
        return bool(r)
    except Exception:
        return False

def get_columns(pid, dt_id, visible_only=False):
    sql = "SELECT * FROM columns_config WHERE project_id=%s AND dt_id=%s"
    if visible_only: sql += " AND visible=true"
    sql += " ORDER BY sort_order"
    cols = q(sql, (pid, dt_id))

    is_noc = False
    try:
        dt = q("SELECT code FROM doc_types WHERE id=%s AND project_id=%s", (dt_id, pid), one=True)
        if dt and dt.get("code", "").upper() == "NOC":
            is_noc = True
    except Exception:
        pass

    if is_noc and project_has_doc_type(pid, "PCQ"):
        cols.append({
            "id": -1,
            "project_id": pid,
            "dt_id": dt_id,
            "col_key": "related_pcq",
            "label": "Related PCQ",
            "col_type": "text",
            "list_name": None,
            "visible": True,
            "sort_order": 999
        })

    return cols

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

def get_records(pid, dt_id, search="", search_pr_items=False):
    is_noc = False
    try:
        dt = q("SELECT code FROM doc_types WHERE id=%s AND project_id=%s", (dt_id, pid), one=True)
        if dt and dt.get("code", "").upper() == "NOC":
            is_noc = True
    except Exception:
        pass
    has_pcq = bool(is_noc and project_has_doc_type(pid, "PCQ"))

    select_clause = r"""
        SELECT records.id, 
               records.data || jsonb_build_object(
                   '_has_newer_revision', 
                   EXISTS (
                       SELECT 1 FROM records r2 
                       WHERE r2.project_id = records.project_id 
                         AND r2.dt_id = records.dt_id 
                         AND r2.id != records.id
                         AND COALESCE(r2.data->>'docNo', '') != ''
                         AND REGEXP_REPLACE(COALESCE(r2.data->>'docNo', ''), '\s*REV.*$', '', 'i') = REGEXP_REPLACE(COALESCE(records.data->>'docNo', ''), '\s*REV.*$', '', 'i')
                         AND COALESCE(r2.data->>'docNo', '') > COALESCE(records.data->>'docNo', '')
                   )
               ) AS data, 
               records.created_at
    """
    if has_pcq:
        select_clause = r"""
            SELECT records.id, 
                   records.data || jsonb_build_object('related_pcq', 
                       (SELECT p.data->>'docNo' 
                        FROM records p 
                        JOIN doc_types dt ON p.dt_id = dt.id 
                        WHERE p.project_id = records.project_id 
                          AND UPPER(dt.code) = 'PCQ' 
                          AND COALESCE(REGEXP_REPLACE(records.data->>'docNo', '\s*REV.*$', '', 'i'), '') <> ''
                          AND p.data::text ILIKE CONCAT('%%', REGEXP_REPLACE(records.data->>'docNo', '\s*REV.*$', '', 'i'), '%%')
                        LIMIT 1)
                   ) || jsonb_build_object(
                       '_has_newer_revision', 
                       EXISTS (
                           SELECT 1 FROM records r2 
                           WHERE r2.project_id = records.project_id 
                             AND r2.dt_id = records.dt_id 
                             AND r2.id != records.id
                             AND COALESCE(r2.data->>'docNo', '') != ''
                             AND REGEXP_REPLACE(COALESCE(r2.data->>'docNo', ''), '\s*REV.*$', '', 'i') = REGEXP_REPLACE(COALESCE(records.data->>'docNo', ''), '\s*REV.*$', '', 'i')
                             AND COALESCE(r2.data->>'docNo', '') > COALESCE(records.data->>'docNo', '')
                       )
                   ) AS data, 
                   records.created_at
        """

    if search:
        sq = f"%{search}%"
        if search_pr_items:
            sql = f"""
                {select_clause} 
                FROM records 
                WHERE project_id=%s AND dt_id=%s 
                AND (
                    data::text ILIKE %s
                    OR EXISTS (
                        SELECT 1 FROM pr_items 
                        WHERE pr_items.record_id = records.id 
                        AND (item_name ILIKE %s OR remarks ILIKE %s)
                    )
                )
                ORDER BY created_at
            """
            if is_noc: print(f"--- DEBUG NOC QUERY (search_pr_items) ---\n{sql}\n------------------------", flush=True)
            rows = q(sql, (pid, dt_id, sq, sq, sq))
        else:
            sql = f"""
                {select_clause} 
                FROM records 
                WHERE project_id=%s AND dt_id=%s 
                AND data::text ILIKE %s
                ORDER BY created_at
            """
            if is_noc: print(f"--- DEBUG NOC QUERY (search) ---\n{sql}\n------------------------", flush=True)
            rows = q(sql, (pid, dt_id, sq))
    else:
        sql = f"{select_clause} FROM records WHERE project_id=%s AND dt_id=%s ORDER BY created_at"
        if is_noc: print(f"--- DEBUG NOC QUERY ---\n{sql}\n------------------------", flush=True)
        rows = q(sql, (pid, dt_id))
    result = []
    for row in rows:
        d = row["data"] if isinstance(row["data"], dict) else {}
        d["_id"]      = row["id"]
        d["_created"] = str(row.get("created_at",""))
        result.append(d)
    return result

def count_records(pid, dt_id):
    r = q("SELECT COUNT(*) as c FROM records WHERE project_id=%s AND dt_id=%s", (pid, dt_id), one=True)
    return r["c"] if r else 0

def invalidate_dashboard_cache(pid=None):
    if pid:
        exe("DELETE FROM dashboard_cache WHERE project_id=%s", (pid,))
        exe("DELETE FROM dashboard_cache WHERE project_id LIKE %s", (f"\\_PR\\_ANALYTICS\\_%{pid}%",))
        exe("DELETE FROM dashboard_cache WHERE project_id='_PR_ANALYTICS_'")
    else:
        exe("DELETE FROM dashboard_cache")

class RecordOwnershipError(Exception):
    """An existing record id was addressed through a different project / document type."""


class InvalidMhrPeriodError(ValueError):
    """MHR record is missing a valid reporting year or month (1-12)."""


class DuplicateMhrPeriodError(ValueError):
    """An MHR record already exists for this project and reporting year/month."""


def normalize_mhr_period(data):
    """Extract and validate integer `(year, month)` from an MHR record payload."""
    payload = data if isinstance(data, dict) else {}
    raw_year = "" if payload.get("year") is None else str(payload.get("year")).strip()
    raw_month = "" if payload.get("month") is None else str(payload.get("month")).strip()
    if not raw_year or not raw_month:
        raise InvalidMhrPeriodError("MHR records require both reporting year and month (1-12).")
    try:
        year_val = int(raw_year)
        month_val = int(raw_month)
    except ValueError:
        raise InvalidMhrPeriodError("Reporting year and month must be integers (month 1-12).") from None
    if year_val < 1900 or year_val > 2200:
        raise InvalidMhrPeriodError("Reporting year must be between 1900 and 2200.")
    if month_val < 1 or month_val > 12:
        raise InvalidMhrPeriodError("Reporting month must be between 1 and 12.")
    return year_val, month_val


def ensure_unique_mhr_period(pid, dt_id, rec_id, data):
    """Validate and enforce unique `(project_id, year, month)` for MHR records; return `(year, month)`."""
    year_val, month_val = normalize_mhr_period(data)
    q("SELECT pg_advisory_xact_lock(hashtext(%s))", (f"mhr:{pid}:{year_val}:{month_val}",))
    rows = q(
        "SELECT id, data FROM records WHERE project_id=%s AND dt_id=%s AND id <> %s",
        (pid, dt_id, rec_id or ""),
    )
    for row in rows:
        rdata = row.get("data") if isinstance(row.get("data"), dict) else {}
        try:
            ex_year, ex_month = normalize_mhr_period(rdata)
        except InvalidMhrPeriodError:
            continue
        if ex_year == year_val and ex_month == month_val:
            raise DuplicateMhrPeriodError(
                f"Man-hours for project '{pid}' and reporting month {year_val:04d}-{month_val:02d} are already logged."
            )
    return year_val, month_val


def _upsert_record_row(pid, dt_id, rec_id, jdata):
    affected = exe("""INSERT INTO records(id,project_id,dt_id,data,updated_at) VALUES(%s,%s,%s,%s::jsonb,NOW())
           ON CONFLICT(id) DO UPDATE SET data = EXCLUDED.data ||
               CASE WHEN records.data ? 'attachments'
                    THEN jsonb_build_object('attachments', records.data->'attachments')
                    ELSE '{}'::jsonb END,
               updated_at=NOW()
           WHERE records.project_id = EXCLUDED.project_id
             AND records.dt_id = EXCLUDED.dt_id""",
        (rec_id, pid, dt_id, jdata))
    if affected == 0:
        raise RecordOwnershipError(
            "Record %r does not belong to project %r / document type %r" % (rec_id, pid, dt_id))


def save_record(pid, dt_id, rec_id, data: dict):
    # 'attachments' is server-managed (see attachments.py): never accepted from the client
    clean = {k: v for k, v in data.items() if not k.startswith("_") and k != "attachments"}
    # On update, carry over existing 'attachments' so a normal record edit can never wipe them.
    # Finding B: an existing row is only updated when it already belongs to (pid, dt_id). The WHERE
    # guard is evaluated atomically by PostgreSQL on conflict, so a record id owned by another
    # project / document type is never overwritten (and never moved): zero rows are affected.
    if str(dt_id or "").strip().upper() == "MHR":
        with transaction():
            year_val, month_val = ensure_unique_mhr_period(pid, dt_id, rec_id, clean)
            clean["year"] = year_val
            clean["month"] = month_val
            _upsert_record_row(pid, dt_id, rec_id, json.dumps(clean))
    else:
        _upsert_record_row(pid, dt_id, rec_id, json.dumps(clean))
    invalidate_dashboard_cache(pid)

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

def _get_ltr_dashboard_stats_map(project_ids=None, workspace_id=None, doc_type_scope=None):
    projects = get_projects(project_ids, workspace_id=workspace_id)
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
        if not _dt_allowed_in_scope(pid, dt["id"], doc_type_scope, dt.get("code", "")):
            continue
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
    r = q("SELECT project_id FROM records WHERE id=%s", (rec_id,), one=True)
    exe("DELETE FROM records WHERE id=%s", (rec_id,))
    if r:
        invalidate_dashboard_cache(r["project_id"])


def get_records_meta(record_ids):
    record_ids = [r for r in (record_ids or []) if r]
    if not record_ids:
        return []
    return q("""
        SELECT id, project_id, dt_id,
               COALESCE(
                   NULLIF(data->>'docNo', ''),
                   NULLIF(data->>'nocNo', ''),
                   NULLIF(data->>'prNo', ''),
                   NULLIF(data->>'prNumber', ''),
                   NULLIF(data->>'letterRef', ''),
                   ''
               ) AS doc_no
        FROM records
        WHERE id = ANY(%s)
    """, (record_ids,))


def delete_records_bulk(record_ids):
    record_ids = [r for r in (record_ids or []) if r]
    if not record_ids:
        return 0
    with DB() as cur:
        cur.execute("SELECT DISTINCT project_id FROM records WHERE id = ANY(%s)", (record_ids,))
        pids = [row["project_id"] for row in cur.fetchall()]
        cur.execute("DELETE FROM records WHERE id = ANY(%s)", (record_ids,))
        rc = cur.rowcount or 0
    for pid in pids:
        invalidate_dashboard_cache(pid)
    return rc

# ── PR Items ─────────────────────────────────────────────────
def get_pr_items(record_id):
    return q("""SELECT id, record_id, row_type, item_name, unit, quantity, remarks, sort_order, po_ref, po_qty, delivered_qty
                FROM pr_items
                WHERE record_id=%s
                ORDER BY sort_order, created_at, id""", (record_id,))

def get_pr_items_for_records(record_ids):
    if not record_ids: return {}
    record_ids = [r for r in record_ids if r]
    if not record_ids: return {}
    rows = q("""SELECT id, record_id, row_type, item_name, unit, quantity, remarks, sort_order, po_ref, po_qty, delivered_qty
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
        try:
            p_qty = float(it.get("po_qty") or 0)
        except:
            p_qty = 0
        try:
            d_qty = float(it.get("delivered_qty") or 0)
        except:
            d_qty = 0
        try:
            pr_qty = float(it.get("quantity") or 0)
        except:
            pr_qty = 0

        exe("""INSERT INTO pr_items(id,record_id,row_type,item_name,unit,quantity,remarks,sort_order,po_ref,po_qty,delivered_qty,updated_at)
               VALUES(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,NOW())""",
            (str(uuid.uuid4()), record_id, it.get("row_type","item"),
             it.get("item_name",""), it.get("unit"), pr_qty, it.get("remarks"), i,
             it.get("po_ref") or '', p_qty, d_qty))
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
    clean_meta = meta
    if str(list_name or "").strip().lower() == "hse_ptw_subtype" and meta is not None:
        from hse_defaults import get_ptw_subtype_code
        clean_meta = get_ptw_subtype_code(item_value, str(meta))
    exe("UPDATE dropdown_lists SET meta=%s"
        " WHERE project_id=%s AND list_name=%s AND item_value=%s",
        (clean_meta, pid, list_name, item_value))
    invalidate_dashboard_cache(pid)

def ensure_doc_type_subtype_column(pid, dt_id, list_name):
    """Ensure a visible 'legacySubtype' dropdown column exists for (pid, dt_id) when its first subtype is added.

    Leaves existing columns and records untouched if a subtype column is already configured.
    Returns True if a new column row was inserted, False otherwise.
    """
    pid_s = str(pid or "").strip()
    dt_s = str(dt_id or "").strip()
    ln_s = str(list_name or "").strip()
    if not pid_s or not dt_s or not ln_s:
        return False
    existing = q(
        """SELECT col_key FROM columns_config
           WHERE project_id=%s AND dt_id=%s
             AND (LOWER(col_key) IN ('legacysubtype', 'subtype')
                  OR (col_type='dropdown' AND LOWER(COALESCE(list_name, ''))=LOWER(%s)))
           LIMIT 1""",
        (pid_s, dt_s, ln_s),
        one=True,
    )
    if existing:
        return False
    exe(
        """UPDATE columns_config
           SET sort_order = sort_order + 1
           WHERE project_id=%s AND dt_id=%s AND sort_order >= 1 AND col_key != 'legacySubtype'""",
        (pid_s, dt_s),
    )
    exe(
        """INSERT INTO columns_config(project_id, dt_id, col_key, label, col_type, list_name, visible, sort_order)
           VALUES(%s, %s, 'legacySubtype', 'Subtype', 'dropdown', %s, true, 1)
           ON CONFLICT(project_id, dt_id, col_key) DO NOTHING""",
        (pid_s, dt_s, ln_s),
    )
    invalidate_dashboard_cache(pid_s)
    return True


def add_list_item(pid, list_name, item, meta=None):
    clean_item = str(item or "").strip()
    resolved_meta = str(meta).strip() if meta is not None and str(meta).strip() else None
    if str(list_name or "").strip().lower() == "hse_ptw_subtype":
        from hse_defaults import get_ptw_subtype_code, parse_ptw_list_item_input

        parsed_name, parsed_code = parse_ptw_list_item_input(clean_item)
        if parsed_name:
            clean_item = parsed_name
        explicit_code = resolved_meta or parsed_code
        resolved_meta = get_ptw_subtype_code(clean_item, explicit_code)
    if not clean_item:
        return {"item": "", "meta": resolved_meta}
    r = q("SELECT COALESCE(MAX(sort_order),0)+1 as n FROM dropdown_lists WHERE project_id=%s AND list_name=%s",
          (pid, list_name), one=True)
    exe("""INSERT INTO dropdown_lists(project_id,list_name,item_value,sort_order,meta)
           VALUES(%s,%s,%s,%s,%s)
           ON CONFLICT(project_id,list_name,item_value)
           DO UPDATE SET meta = COALESCE(EXCLUDED.meta, dropdown_lists.meta)""",
        (pid, list_name, clean_item, r["n"] if r else 0, resolved_meta))
    invalidate_dashboard_cache(pid)
    return {"item": clean_item, "meta": resolved_meta}

def rename_list_item(pid, list_name, old_item, new_item, meta=None):
    old_item = str(old_item or "").strip()
    new_item = str(new_item or "").strip()
    explicit_meta = str(meta).strip() if meta is not None and str(meta).strip() else None
    if str(list_name or "").strip().lower() == "hse_ptw_subtype":
        from hse_defaults import get_ptw_subtype_code, parse_ptw_list_item_input

        parsed_name, parsed_code = parse_ptw_list_item_input(new_item)
        if parsed_name:
            new_item = parsed_name
        if explicit_meta or parsed_code:
            explicit_meta = get_ptw_subtype_code(new_item, explicit_meta or parsed_code)
    if not old_item or not new_item or (old_item == new_item and explicit_meta is None):
        return 0
    with DB() as cur:
        if explicit_meta is not None:
            cur.execute("""
                UPDATE dropdown_lists
                SET item_value=%s, meta=%s
                WHERE project_id=%s AND list_name=%s AND item_value=%s
            """, (new_item, explicit_meta, pid, list_name, old_item))
        else:
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
        invalidate_dashboard_cache(pid)
        return renamed

def reorder_list_items(pid, list_name, ordered_items):
    for i, item in enumerate([str(v).strip() for v in (ordered_items or []) if str(v).strip()]):
        exe("""
            UPDATE dropdown_lists
            SET sort_order=%s
            WHERE project_id=%s AND list_name=%s AND item_value=%s
        """, (i, pid, list_name, item))
    invalidate_dashboard_cache(pid)

def remove_list_item(pid, list_name, item):
    exe("DELETE FROM dropdown_lists WHERE project_id=%s AND list_name=%s AND item_value=%s",
        (pid, list_name, item))
    invalidate_dashboard_cache(pid)

# ── Fast Dashboard Stats (single query) ──────────────────────
def get_dashboard_stats(project_ids=None, workspace_id=None, doc_type_scope=None):
    """One big SQL call — replaces N×M individual queries."""
    import re, json as _json
    projects = get_projects(project_ids, workspace_id=workspace_id)
    if not projects:
        return []

    pids = [p["id"] for p in projects]
    restricted_pids = {
        pid for pid in pids
        if doc_type_scope is not None and (pid not in doc_type_scope or doc_type_scope[pid] is not None)
    }

    # 1) Try fetching from DB cache (only for unrestricted projects)
    uncached_eligible = [pid for pid in pids if pid not in restricted_pids]
    cached_rows = q("SELECT project_id, cache_data FROM dashboard_cache WHERE project_id = ANY(%s)", (uncached_eligible,)) if uncached_eligible else []
    cache_map = {}
    for r in cached_rows:
        cdata = r.get("cache_data")
        if isinstance(cdata, dict):
            # Verify cache adheres to strict submittals total = approved + pending + rejected
            tot = cdata.get("total", 0)
            sub_sum = cdata.get("approved", 0) + cdata.get("pending", 0) + cdata.get("rejected", 0)
            if tot == sub_sum:
                cache_map[r["project_id"]] = cdata

    missing_pids = [pid for pid in pids if pid not in cache_map]

    result = []
    for pid in pids:
        if pid in cache_map:
            result.append(cache_map[pid])

    if not missing_pids:
        return result

    # 2) We have some missing projects, calculate ONLY for those
    missing_projects = [p for p in projects if p["id"] in missing_pids]
    ltr_stats_map = _get_ltr_dashboard_stats_map(missing_pids, doc_type_scope=doc_type_scope)
    ph   = ",".join(["%s"] * len(missing_pids))

    meta_rows = q(f"""
        SELECT project_id, item_value, meta
        FROM dropdown_lists
        WHERE project_id IN ({ph})
          AND list_name LIKE %s
          AND meta IS NOT NULL
    """, missing_pids + ["status%"])
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
    """, missing_pids)

    # All doc_types in one query
    all_dts = q(f"""
        SELECT project_id, id, code, name, sort_order
        FROM doc_types
        WHERE project_id IN ({ph})
        ORDER BY project_id, sort_order, id
    """, missing_pids)

    dt_map   = {}  # pid -> list of dts
    for dt in all_dts:
        dt_map.setdefault(dt["project_id"], []).append(dt)

    rec_map  = {}  # (pid, dt_id) -> list of data dicts
    for row in all_recs:
        key = (row["project_id"], row["dt_id"])
        d   = row["data"] if isinstance(row["data"], dict) else {}
        rec_map.setdefault(key, []).append(d)

    for p in missing_projects:
        pid   = p["id"]
        pdata = p.get("data") or {}
        if isinstance(pdata, str):
            try: pdata = _json.loads(pdata)
            except: pdata = {}

        dts      = dt_map.get(pid, [])
        dt_stats = []
        total = approved = pending = overdue_cnt = total_rj = info_closed = 0

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
            if not _dt_allowed_in_scope(pid, dt["id"], doc_type_scope, dt.get("code", "")):
                continue
            rows  = rec_map.get((pid, dt["id"]), [])
            t=ap=pe=ov=rj=ic = 0
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

            # Skip doc types without status column from submittal KPI counts
            if not dt_has_status.get(dt["id"], False):
                continue

            for base, revisions in doc_groups.items():
                revisions.sort(key=lambda x: x[0])   # sort ascending by rev
                latest_rev, d = revisions[-1]
                doc_no = d.get("docNo", "")
                status = str(d.get("status") or "").strip()
                disc   = d.get("discipline","") or "—"
                if not status:
                    continue  # Skip records without approval status

                # Resolve meta: DB first, then fallback to DEFAULT_STATUS_META
                meta = resolve_status_meta(status, pid_meta)
                if meta not in ("approved", "pending", "rejected"):
                    continue  # Skip cancelled, info_closed, or unmapped status

                is_ap = (meta == "approved")
                is_pe = (meta == "pending")
                is_rj = (meta == "rejected")

                _has_both = dt_has_exp_reply.get(dt["id"], False) and dt_has_status.get(dt["id"], False) and not dt_is_non_workflow
                status_val = d.get("status")
                action_val = d.get("action")
                dt_rule = dt_rules.get(dt["id"])
                is_ov = is_pe and is_overdue(d.get("issuedDate"), doc_no, d.get("actualReplyDate"), _has_both, rule=dt_rule, status=status_val, action=action_val, row=d)

                if is_ap: ap += 1
                elif is_pe: pe += 1
                elif is_rj: rj += 1

                if is_ov: ov += 1

                # Discipline breakdown
                if dt_has_discipline.get(dt["id"], False):
                    ds = disc_map.setdefault(disc, {"total":0,"approved":0,"pending":0,"rejected":0,"overdue":0,"info_closed":0})
                    ds["total"] += 1
                    if is_ap: ds["approved"] += 1
                    elif is_pe: ds["pending"]  += 1
                    elif is_rj: ds["rejected"] += 1
                    if is_ov: ds["overdue"]  += 1

            t = ap + pe + rj
            disc_list = [{"disc":k,"total":v["total"],"approved":v["approved"],
                          "pending":v["pending"],"rejected":v["rejected"],"overdue":v["overdue"],"info_closed":0}
                         for k,v in sorted(disc_map.items())]
            if not dt_is_pr and not dt_is_ltr:
                dt_stats.append({"id":dt["id"],"code":dt["code"],"name":dt["name"],
                                 "total":t,"approved":ap,"pending":pe,"info_closed":0,"overdue":ov,"rejected":rj,
                                 "disc_breakdown":disc_list})
            total+=t; approved+=ap; pending+=pe; info_closed+=0; overdue_cnt+=ov; total_rj+=rj

        pct = round(approved / total * 100) if total else 0
        res_p = {
            "id": pid, "name": p["name"], "code": p["code"],
            "client": pdata.get("client","") if isinstance(pdata, dict) else "",
            "total": total, "approved": approved, "pending": pending, "info_closed": 0,
            "overdue": overdue_cnt, "rejected": total_rj, "pct": pct, "dt_stats": dt_stats,
            "ltr": ltr_stats_map.get(pid, {"total": 0, "sent": 0, "received": 0, "party_stats": [], "top_parties": []})
        }
        if pid not in restricted_pids:
            exe("INSERT INTO dashboard_cache (project_id, cache_data) VALUES (%s, %s::jsonb) ON CONFLICT(project_id) DO UPDATE SET cache_data=EXCLUDED.cache_data, updated_at=NOW()", (pid, _json.dumps(res_p)))
        cache_map[pid] = res_p
    return [cache_map[pid] for pid in pids if pid in cache_map]


def get_data_quality_summary(pid=None, project_ids=None, doc_type_scope=None):
    where, params = _project_scope_clause("project_id", pid, project_ids)
    dt_rows = [
        d for d in q(f"SELECT project_id, id, code, name FROM doc_types {where}", params)
        if _dt_allowed_in_scope(d["project_id"], d["id"], doc_type_scope, d.get("code", ""))
    ]
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
        if key not in dt_meta:
            continue
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
    recs = [
        r for r in q(f"SELECT id, project_id, dt_id, data FROM records {rec_where}", rec_params)
        if _dt_allowed_in_scope(r["project_id"], r["dt_id"], doc_type_scope, (dt_meta.get((r["project_id"], r["dt_id"])) or {}).get("code", ""))
    ]
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


def get_action_required_summary(pid=None, limit=10, pending_threshold=14, project_ids=None, doc_type_scope=None):
    from utils import compute_duration

    projects = get_projects(project_ids)
    project_codes = {p["id"]: p["code"] for p in projects}

    overdue_rows = get_overdue_records(pid, project_ids=project_ids, doc_type_scope=doc_type_scope)[:limit]
    top_overdue = [
        {**r, "project_code": project_codes.get(r["project_id"], r["project_id"])}
        for r in overdue_rows
    ]

    dt_where, dt_params = _project_scope_clause("project_id", pid, project_ids)
    dt_rows = [
        d for d in q(f"SELECT project_id, id, code, name FROM doc_types {dt_where}", dt_params)
        if _dt_allowed_in_scope(d["project_id"], d["id"], doc_type_scope, d.get("code", ""))
    ]
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


def get_pr_analytics_summary(pid=None, limit=6, project_ids=None, doc_type_scope=None):
    import json
    use_cache = (doc_type_scope is None)
    if pid:
        cache_key = f"_PR_ANALYTICS_{pid}"
    elif project_ids is not None:
        cache_key = "_PR_ANALYTICS_" + "_".join(sorted(project_ids))
    else:
        cache_key = "_PR_ANALYTICS_"
    if use_cache:
        cached = q("SELECT cache_data FROM dashboard_cache WHERE project_id=%s", (cache_key,), one=True)
        if cached:
            cdata = cached.get("cache_data")
            # Invalidate cache if legacy version, or contains 'Unspecified', or unmerged 'Other'
            if isinstance(cdata, dict) and cdata.get("cache_ver") == 3:
                top_trades = cdata.get("top_trades", [])
                trade_names_lower = [str(t.get("trade") or "").lower() for t in top_trades]
                has_other_separate = "other" in trade_names_lower and "others" in trade_names_lower
                has_unspecified = any(str(t.get("trade") or "").strip().lower() == "unspecified" for t in top_trades)
                if not has_unspecified and not has_other_separate and cdata.get("top_trade_name") != "Unspecified":
                    return cdata

    projects = get_projects(project_ids)
    project_map = {p["id"]: p for p in projects}
    allowed_ids = _clean_project_ids(project_ids)
    if pid:
        if allowed_ids is not None and pid not in allowed_ids:
            pr_rows = []
        else:
            pr_rows = q("""
                SELECT r.id, r.project_id, r.dt_id, r.data, d.code, d.name
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
                SELECT r.id, r.project_id, r.dt_id, r.data, d.code, d.name
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
            SELECT r.id, r.project_id, r.dt_id, r.data, d.code, d.name
            FROM records r
            JOIN doc_types d ON d.id=r.dt_id AND d.project_id=r.project_id
            WHERE UPPER(COALESCE(d.code, ''))='PR'
               OR LOWER(COALESCE(d.name, '')) LIKE %s
               OR LOWER(COALESCE(d.name, '')) LIKE %s
        """, ("%requisition%", "%purchase request%"))

    if doc_type_scope is not None:
        pr_rows = [
            row for row in pr_rows
            if _dt_allowed_in_scope(row["project_id"], row["dt_id"], doc_type_scope, row.get("code", ""))
        ]

    total_pr_records = len(pr_rows)
    if not total_pr_records:
        return {
            "cache_ver": 3,
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
        trade = (
            str(data.get("trade") or "").strip() or
            str(data.get("Trade") or "").strip() or
            str(data.get("subTrade") or "").strip() or
            str(data.get("sub_trade") or "").strip() or
            str(data.get("SubTrade") or "").strip() or
            str(data.get("discipline") or "").strip() or
            str(data.get("Discipline") or "").strip()
        )
        if not trade or trade.lower() in ("unspecified", "none", "—", "-"):
            trade = "General / Multi-Discipline"
        elif trade.strip().lower() in ("other", "others"):
            trade = "Others"
        trade_counts[trade] = trade_counts.get(trade, 0) + 1

    top_projects = []
    # Show ALL projects with PRs so sum of projects rows equals total PRs 100%
    for row_pid, pr_count in sorted(project_counts.items(), key=lambda x: (-x[1], x[0])):
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

    res = {
        "cache_ver": 3,
        "total_pr_records": total_pr_records,
        "top_projects": top_projects,
        "top_trades": top_trades,
        "top_project_name": top_project.get("project_name", ""),
        "top_project_count": top_project.get("pr_count", 0),
        "top_trade_name": top_trade.get("trade", ""),
        "top_trade_count": top_trade.get("pr_count", 0),
        "trade_count_total": len(trade_counts),
    }
    if use_cache:
        exe("INSERT INTO dashboard_cache (project_id, cache_data) VALUES (%s, %s::jsonb) ON CONFLICT(project_id) DO UPDATE SET cache_data=EXCLUDED.cache_data, updated_at=NOW()", (cache_key, json.dumps(res)))
    return res


def get_monthly_trend(pid=None, project_ids=None, doc_type_scope=None):
    """Returns last 6 months of submission counts per month."""
    import re as _re
    where, params = _project_scope_clause("project_id", pid, project_ids)
    rows = q(f"SELECT project_id, dt_id, data FROM records {where}", params)
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
        if not _dt_allowed_in_scope(row.get("project_id"), row.get("dt_id"), doc_type_scope):
            continue
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


def get_aging_report(pid=None, project_ids=None, doc_type_scope=None):
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
    dt_where, dt_params = _project_scope_clause("project_id", pid, project_ids)
    dt_rows = q(f"SELECT id, code, name FROM doc_types {dt_where}", dt_params)
    dt_with_exp = {dt_id for dt_id in dt_has_exp_a
                   if not _is_non_workflow_dt(
                       next((d["code"] for d in dt_rows if d["id"] == dt_id), ""),
                       next((d["name"] for d in dt_rows if d["id"] == dt_id), "")
                   )}
    rows = q(f"SELECT r.project_id, r.dt_id, r.data FROM records r {where}", params)
    buckets = {"1-7": 0, "8-14": 0, "15-21": 0, ">21": 0}
    for row in rows:
        if not _dt_allowed_in_scope(row.get("project_id"), row.get("dt_id"), doc_type_scope):
            continue
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

def get_quality_report(pid=None, project_ids=None, doc_type_scope=None):
    """Returns doc quality: how many docs needed 0,1,2,3+ revisions."""
    import re as _re
    where, params = _project_scope_clause("project_id", pid, project_ids)
    rows = q(f"SELECT project_id, dt_id, data FROM records {where}", params)
    doc_revs = {}
    for row in rows:
        if not _dt_allowed_in_scope(row.get("project_id"), row.get("dt_id"), doc_type_scope):
            continue
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


def get_overdue_records(pid=None, project_ids=None, doc_type_scope=None):
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
    
    dt_where, dt_params = _project_scope_clause("project_id", pid, project_ids)
    dt_rows = q(f"SELECT id, code, name FROM doc_types {dt_where}", dt_params)
    
    dt_with_exp = {dt_id for dt_id in dt_has_exp
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
        if not _dt_allowed_in_scope(row.get("project_id"), row.get("dt_id"), doc_type_scope, row.get("dt_code", "")):
            continue
        d = row["data"] if isinstance(row["data"], dict) else {}
        if d.get("actualReplyDate"): continue
        doc_no = d.get("docNo", "") or ""
        issued = d.get("issuedDate", "") or ""
        if not issued: continue
        status_val = d.get("status")
        action_val = d.get("action")
        dt_rule = dt_rules.get(row["dt_id"])
        
        meta = resolve_status_meta(status_val, meta_map.get(row["project_id"], {}))
        if meta != "pending" and meta != "open":
            # PMO records may not have status, so meta might evaluate to "pending"
            continue

        _has_exp = row["dt_id"] in dt_with_exp
        if is_overdue(issued, doc_no, None, _has_exp, rule=dt_rule, status=status_val, action=action_val, row=d):
            try:
                from utils import compute_expected_reply, days_between_by_rule, compute_duration
                exp = compute_expected_reply(issued, doc_no, rule=dt_rule, status=status_val, action=action_val, row=d)
                import datetime
                yesterday = (datetime.date.today() - datetime.timedelta(days=1)).isoformat()
                if exp:
                    days = days_between_by_rule(exp, yesterday) or 0
                else:
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
def parse_user_agent(ua_str):
    """Extract a friendly, concise OS / Browser string from User-Agent."""
    if not ua_str:
        return "Unknown"
    ua = str(ua_str).lower()
    
    # Detect OS
    if "windows nt 10.0" in ua or "windows nt 11.0" in ua:
        os_name = "Windows 10/11"
    elif "windows" in ua:
        os_name = "Windows"
    elif "iphone" in ua:
        os_name = "iPhone (iOS)"
    elif "ipad" in ua:
        os_name = "iPad (iOS)"
    elif "android" in ua:
        os_name = "Android"
    elif "macintosh" in ua or "mac os" in ua:
        os_name = "macOS"
    elif "linux" in ua:
        os_name = "Linux"
    else:
        os_name = "Other"

    # Detect Browser
    if "edg/" in ua or "edge/" in ua:
        browser = "Edge"
    elif "opr/" in ua or "opera" in ua:
        browser = "Opera"
    elif "chrome/" in ua:
        browser = "Chrome"
    elif "firefox/" in ua:
        browser = "Firefox"
    elif "safari/" in ua and "chrome" not in ua:
        browser = "Safari"
    else:
        browser = "Browser"

    return f"{os_name} • {browser}"

def ensure_audit_schema():
    """Idempotent migration to ensure audit columns and indexes exist."""
    try:
        exe("ALTER TABLE audit_log ADD COLUMN IF NOT EXISTS ip_address TEXT DEFAULT NULL;")
        exe("ALTER TABLE audit_log ADD COLUMN IF NOT EXISTS user_agent TEXT DEFAULT NULL;")
        exe("ALTER TABLE audit_log ADD COLUMN IF NOT EXISTS device TEXT DEFAULT NULL;")
        exe("CREATE INDEX IF NOT EXISTS idx_audit_user ON audit_log(username, ts DESC);")
        exe("CREATE INDEX IF NOT EXISTS idx_audit_action ON audit_log(action, ts DESC);")
    except Exception as e:
        logger.warning("ensure_audit_schema warning: %s", e)

def log_action(username, action, project_id=None, dt_id=None, record_id=None,
               doc_no=None, field_name=None, old_value=None, new_value=None, detail=None,
               ip_address=None, user_agent=None, device=None):
    """Write one audit entry with automatic IP & device extraction. Never raises — silently ignores errors."""
    try:
        if ip_address is None and user_agent is None:
            try:
                from flask import request, has_request_context
                if has_request_context():
                    cf_ip = request.headers.get("CF-Connecting-IP")
                    x_fwd = request.headers.get("X-Forwarded-For")
                    if cf_ip:
                        ip_address = cf_ip.strip()
                    elif x_fwd:
                        ip_address = x_fwd.split(",")[0].strip()
                    else:
                        ip_address = request.remote_addr or ""
                    
                    user_agent = request.headers.get("User-Agent", "")[:300]
                    device = parse_user_agent(user_agent)
            except Exception:
                pass

        exe("""INSERT INTO audit_log(username,action,project_id,dt_id,record_id,
               doc_no,field_name,old_value,new_value,detail,ip_address,user_agent,device)
               VALUES(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)""",
            (username, action, project_id, dt_id, record_id,
             doc_no, field_name,
             str(old_value)[:500] if old_value is not None else None,
             str(new_value)[:500] if new_value is not None else None,
             str(detail)[:500] if detail else None,
             str(ip_address)[:100] if ip_address else None,
             str(user_agent)[:300] if user_agent else None,
             str(device)[:100] if device else None))
    except Exception as e:
        logger.error("audit_log_failed action=%s project_id=%s dt_id=%s record_id=%s error=%s",
                     action, project_id, dt_id, record_id, e)

def get_audit_log(project_id=None, username=None, action=None,
                  limit=200, offset=0, allowed_project_ids=None, allowed_usernames=None):
    """Fetch audit entries with optional filters."""
    conditions = []
    params = []
    if project_id:
        conditions.append("project_id=%s"); params.append(project_id)
    elif allowed_project_ids is not None:
        if not allowed_project_ids:
            return []
        conditions.append("project_id = ANY(%s)"); params.append(list(allowed_project_ids))

    if username:
        conditions.append("username=%s"); params.append(username)
    elif allowed_usernames is not None:
        if not allowed_usernames:
            return []
        conditions.append("username = ANY(%s)"); params.append(list(allowed_usernames))

    if action:
        conditions.append("action=%s"); params.append(action)
    where = ("WHERE " + " AND ".join(conditions)) if conditions else ""
    params += [limit, offset]
    return q(f"""SELECT id,ts,username,action,project_id,dt_id,
                        doc_no,field_name,old_value,new_value,detail,
                        ip_address,user_agent,device
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

def update_record_link(doc_no, link, sync_links=None):
    """Updates the fileLocation (Drive Link) of a record by its docNo."""
    import json
    rows = q("SELECT id, data FROM records WHERE data->>'docNo' = %s", (doc_no,))
    for r in rows:
        d = r["data"] if isinstance(r["data"], dict) else {}
        d["fileLocation"] = link
        if sync_links:  # additive: merge Submit/Replied links without overwriting DCR uploads
            import attachments as _att
            d = _att.merge_sync_links(d, sync_links)
        exe("UPDATE records SET data=%s WHERE id=%s", (json.dumps(d), r["id"]))


# ── Dual attachments (Submission / Approval) - additive, see attachments.py ──
def get_record_attachments(record_id):
    """Return {'submission':..., 'approval':...} for a record (legacy fileLocation = submission)."""
    import attachments as _att
    row = q("SELECT data FROM records WHERE id=%s", (record_id,), one=True)
    if not row:
        return None
    d = row["data"] if isinstance(row["data"], dict) else {}
    return _att.normalize_attachments(d)

def set_record_attachment(record_id, slot, meta):
    """Store attachment metadata in records.data['attachments'][slot]. fileLocation is untouched."""
    import attachments as _att
    if slot not in _att.SLOTS:
        raise ValueError("Invalid attachment slot: %r" % (slot,))
    # Single atomic statement: no read-modify-write race; other keys untouched.
    rows = q("""UPDATE records SET data = data || jsonb_build_object('attachments',
                    COALESCE(data->'attachments','{}'::jsonb) || jsonb_build_object(%s::text, %s::jsonb)),
                    updated_at = NOW()
                WHERE id=%s RETURNING id""",
             (slot, json.dumps(meta), record_id))
    return bool(rows)

def clear_record_attachment(record_id, slot):
    """Remove one attachment slot from a record. fileLocation is untouched."""
    import attachments as _att
    if slot not in _att.SLOTS:
        raise ValueError("Invalid attachment slot: %r" % (slot,))
    rows = q("""UPDATE records SET data = jsonb_set(data, '{attachments}',
                    COALESCE(data->'attachments','{}'::jsonb) - %s::text),
                    updated_at = NOW()
                WHERE id=%s RETURNING id""", (slot, record_id))
    return bool(rows)


# ── Drive folder routing cache (doc_types.data) - additive, see drive_routing.py ──
def set_doc_type_drive_cache(pid, dt_id, patch):
    """Merge Drive routing keys into doc_types.data atomically; other keys are untouched."""
    allowed = {"drive_folder_id", "drive_folder_name", "drive_slot_folders", "folder_synced_at"}
    patch = {k: v for k, v in (patch or {}).items() if k in allowed}
    if not patch:
        return False
    rows = q("""UPDATE doc_types SET data = COALESCE(data,'{}'::jsonb) || %s::jsonb
                WHERE id=%s AND project_id=%s RETURNING id""",
             (json.dumps(patch), dt_id, pid))
    return bool(rows)
