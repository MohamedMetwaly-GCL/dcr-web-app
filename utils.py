"""utils.py - Date calculations and document numbering"""
import datetime, re

# Default Egyptian holidays — overridable from app_settings in DB
DEFAULT_HOLIDAYS = {
    "2025-01-07","2025-01-25","2025-03-30","2025-03-31","2025-04-01","2025-04-02",
    "2025-04-21","2025-04-25","2025-05-01","2025-06-06","2025-06-07","2025-06-08",
    "2025-06-09","2025-06-30","2025-07-23","2025-09-04","2025-10-06",
    "2026-01-07","2026-01-25","2026-03-20","2026-03-21","2026-03-22","2026-03-23",
    "2026-04-25","2026-05-01","2026-05-27","2026-05-28","2026-05-29","2026-05-30",
    "2026-06-30","2026-07-23","2026-10-06",
}

_holidays_cache = None

def get_holidays():
    """Load holidays from DB, fallback to defaults."""
    global _holidays_cache
    if _holidays_cache is not None:
        return _holidays_cache
    try:
        import db as _db
        h = _db.get_setting("holidays", None)
        if isinstance(h, list) and h:
            _holidays_cache = set(h)
            return _holidays_cache
    except Exception:
        pass
    _holidays_cache = set(DEFAULT_HOLIDAYS)
    return _holidays_cache

def invalidate_holidays_cache():
    global _holidays_cache
    _holidays_cache = None

def is_working_day(d):
    if isinstance(d, str): d = datetime.date.fromisoformat(d)
    return d.weekday() != 4 and d.strftime("%Y-%m-%d") not in get_holidays()

def add_working_days(start, days):
    if isinstance(start, str): start = datetime.date.fromisoformat(start)
    d, added = start, 0
    while added < days:
        d += datetime.timedelta(days=1)
        if is_working_day(d): added += 1
    return d.strftime("%Y-%m-%d")

def working_days_between(start, end):
    if not start or not end: return None
    try:
        if isinstance(start, str): start = datetime.date.fromisoformat(start)
        if isinstance(end, str):   end   = datetime.date.fromisoformat(end)
    except: return None
    if end <= start: return 0
    count, cur = 0, start + datetime.timedelta(days=1)
    while cur <= end:
        if is_working_day(cur): count += 1
        cur += datetime.timedelta(days=1)
    return count

def working_days_until_today(start):
    """Working days from start date until today (for open/unreplied docs)."""
    return working_days_between(start, datetime.date.today().isoformat())

def extract_rev(doc_no):
    m = re.search(r'REV(\d+)', doc_no or '', re.IGNORECASE)
    return int(m.group(1)) if m else 0

def extract_seq(doc_no, prefix):
    m = re.match(rf'^{re.escape(prefix)}-(\d+)\s+REV(\d+)$', doc_no or '', re.IGNORECASE)
    return (int(m.group(1)), int(m.group(2))) if m else None

def build_doc_no(prefix, num, rev):
    return f"{prefix}-{num:03d} REV{rev:02d}"

def get_next_doc_no(prefix, records):
    """Find last full doc_no used and suggest next sequential one."""
    max_num = 0
    for r in records:
        p = extract_seq(r.get('docNo',''), prefix)
        if p: max_num = max(max_num, p[0])
    if max_num == 0:
        return ""   # No existing docs — return empty, let user type freely
    return build_doc_no(prefix, max_num + 1, 0)

def compute_expected_reply(issued_date, doc_no):
    if not issued_date: return None
    return add_working_days(issued_date, 14 if extract_rev(doc_no) == 0 else 7)

def compute_duration(issued_date, actual_reply):
    return working_days_between(issued_date, actual_reply)

def is_overdue(issued_date, doc_no, actual_reply):
    if actual_reply: return False
    exp = compute_expected_reply(issued_date, doc_no)
    if not exp: return False
    return datetime.date.fromisoformat(exp) < datetime.date.today()

def format_date(d):
    if not d: return ''
    try: return datetime.date.fromisoformat(str(d)[:10]).strftime("%d-%m-%Y")
    except: return str(d) or ''

STATUS_COLORS = {
    'A - Approved':              ('bbf7d0','166534'),
    'B - Approved As Noted':     ('dcfce7','14532d'),
    'B,C - Approved & Resubmit': ('fed7aa','7c2d12'),
    'C - Revise & Resubmit':     ('fce7f3','831843'),
    'D - Review not Required':   ('fecaca','7f1d1d'),
    'Under Review':              ('fef9c3','713f12'),
    'Cancelled':                 ('ef4444','ffffff'),
    'Open':                      ('fed7aa','7c2d12'),
    'Closed':                    ('bfdbfe','1e3a5f'),
    'Replied':                   ('d1fae5','064e3b'),
    'Pending':                   ('e0e7ff','312e81'),
}

REJECTED_STATUSES = {'C - Revise & Resubmit', 'D - Review not Required', 'Cancelled'}
