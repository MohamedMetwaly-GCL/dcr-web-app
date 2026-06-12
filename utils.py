"""utils.py - Date calculations and document numbering"""
import datetime, re

NON_REV_DOC_TYPE_CODES = {"SI"}

DEFAULT_HOLIDAYS = {
    "2025-01-07","2025-01-25","2025-03-30","2025-03-31","2025-04-01","2025-04-02",
    "2025-04-21","2025-04-25","2025-05-01","2025-06-06","2025-06-07","2025-06-08",
    "2025-06-09","2025-06-30","2025-07-23","2025-09-04","2025-10-06",
    "2026-01-07","2026-01-25","2026-03-20","2026-03-21","2026-03-22","2026-03-23",
    "2026-04-25","2026-05-01","2026-05-27","2026-05-28","2026-05-29","2026-05-30",
    "2026-06-30","2026-07-23","2026-10-06",
}

_holidays_cache = None

DEFAULT_EXPECTED_REPLY_RULE = {
    "rev0_reply_days": 14,
    "rev_reply_days": 7,
    "calculation_mode": "working_days",
    "weekend_mode": "friday_only",
    "exclude_official_holidays": True,
}

DEFAULT_DOC_TYPE_EXPECTED_REPLY_OVERRIDE = {
    "use_expected_reply_override": False,
    "rev0_reply_days_override": 14,
    "rev_reply_days_override": 7,
}

_WEEKEND_DAYS = {
    "none": set(),
    "friday_only": {4},
    "friday_saturday": {4, 5},
}

def normalize_expected_reply_rule(rule=None):
    """Return a safe project expected-reply rule, preserving legacy defaults."""
    src = rule if isinstance(rule, dict) else {}
    out = dict(DEFAULT_EXPECTED_REPLY_RULE)
    for key in ("rev0_reply_days", "rev_reply_days"):
        try:
            val = int(src.get(key, out[key]))
            out[key] = max(0, val)
        except Exception:
            pass
    calc = str(src.get("calculation_mode", out["calculation_mode"]) or "").strip().lower()
    out["calculation_mode"] = calc if calc in ("calendar_days", "working_days") else "working_days"
    weekend = str(src.get("weekend_mode", out["weekend_mode"]) or "").strip().lower()
    out["weekend_mode"] = weekend if weekend in _WEEKEND_DAYS else "friday_only"
    raw_exclude = src.get("exclude_official_holidays", out["exclude_official_holidays"])
    if isinstance(raw_exclude, str):
        out["exclude_official_holidays"] = raw_exclude.strip().lower() in ("1", "true", "yes", "on")
    else:
        out["exclude_official_holidays"] = bool(raw_exclude)
    return out

def normalize_doc_type_expected_reply_override(override=None):
    """Return safe optional document-type day overrides."""
    src = override if isinstance(override, dict) else {}
    out = dict(DEFAULT_DOC_TYPE_EXPECTED_REPLY_OVERRIDE)
    raw_enabled = src.get("use_expected_reply_override", src.get("enabled", out["use_expected_reply_override"]))
    if isinstance(raw_enabled, str):
        out["use_expected_reply_override"] = raw_enabled.strip().lower() in ("1", "true", "yes", "on")
    else:
        out["use_expected_reply_override"] = bool(raw_enabled)
    for key in ("rev0_reply_days_override", "rev_reply_days_override"):
        try:
            val = int(src.get(key, out[key]))
            out[key] = max(0, val)
        except Exception:
            pass
    return out

def apply_doc_type_expected_reply_override(rule=None, override=None):
    """
    Merge optional doc-type day counts into a project rule.
    Calculation mode, weekend rule, and holiday exclusion always remain project-level.
    """
    out = normalize_expected_reply_rule(rule)
    dt = normalize_doc_type_expected_reply_override(override)
    if dt["use_expected_reply_override"]:
        out["rev0_reply_days"] = dt["rev0_reply_days_override"]
        out["rev_reply_days"] = dt["rev_reply_days_override"]
    return out

def get_holidays():
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

def parse_any_date(d_str):
    if not d_str: return None
    if isinstance(d_str, datetime.date): return d_str
    if isinstance(d_str, datetime.datetime): return d_str.date()
    d_str = str(d_str).strip()[:10]
    try:
        return datetime.date.fromisoformat(d_str)
    except ValueError:
        pass
    import re
    m = re.match(r'^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})$', d_str)
    if m:
        try:
            return datetime.date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
        except ValueError:
            pass
    m2 = re.match(r'^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})$', d_str)
    if m2:
        try:
            return datetime.date(int(m2.group(1)), int(m2.group(2)), int(m2.group(3)))
        except ValueError:
            pass
    return None

def is_working_day(d, weekend_mode="friday_only", exclude_official_holidays=True):
    d = parse_any_date(d)
    if not d: return False
    weekend_days = _WEEKEND_DAYS.get(weekend_mode, _WEEKEND_DAYS["friday_only"])
    if d.weekday() in weekend_days:
        return False
    if exclude_official_holidays and d.strftime("%Y-%m-%d") in get_holidays():
        return False
    return True

import functools

@functools.lru_cache(maxsize=8192)
def add_working_days(start, days, weekend_mode="friday_only", exclude_official_holidays=True):
    if not start: return None
    start = parse_any_date(start)
    if not start: return None
    d, added = start, 0
    while added < days:
        d += datetime.timedelta(days=1)
        if is_working_day(d, weekend_mode, exclude_official_holidays): added += 1
    return d.strftime("%Y-%m-%d")

@functools.lru_cache(maxsize=8192)
def add_calendar_days(start, days):
    if not start: return None
    start = parse_any_date(start)
    if not start: return None
    return (start + datetime.timedelta(days=int(days or 0))).strftime("%Y-%m-%d")

def working_days_between(start, end, weekend_mode="friday_only", exclude_official_holidays=True):
    """
    Count working days between two dates (exclusive of start, inclusive of end).
    Rules:
      - If either date missing → None
      - If start == end → 0
      - If end < start → 0  (no negatives)
      - Excludes Fridays and holidays
    """
    if not start or not end:
        return None
    start = parse_any_date(start)
    end = parse_any_date(end)
    if not start or not end:
        return None
    if end <= start:
        return 0
    count, cur = 0, start + datetime.timedelta(days=1)
    while cur <= end:
        if is_working_day(cur, weekend_mode, exclude_official_holidays): count += 1
        cur += datetime.timedelta(days=1)
    return count

def days_between_by_rule(start, end, rule=None):
    """
    Count days between two dates using the project expected-reply rule.
    Keeps the existing Duration endpoint semantics: exclusive of start, inclusive of end.
    """
    if not start or not end:
        return None
    start = parse_any_date(start)
    end = parse_any_date(end)
    if not start or not end:
        return None
    if end <= start:
        return 0
    cfg = normalize_expected_reply_rule(rule)
    if cfg["calculation_mode"] == "calendar_days":
        return (end - start).days
    return working_days_between(
        start,
        end,
        cfg["weekend_mode"],
        cfg["exclude_official_holidays"],
    )

def compute_duration(issued_date, actual_reply, rule=None, status=None, action=None):
    """
    Working days between issued and reply dates.
    If actual_reply is absent → compute from issued to Yesterday (today-1).
    If issued_date is absent → return None.
    Never returns negative.
    """
    chk = f"{status or ''} {action or ''}".upper()
    if "FOR INFORMATION" in chk or "FI" in re.findall(r'[A-Z]+', chk):
        return None
    if not issued_date:
        return None
    if actual_reply and str(actual_reply).strip().lower() not in ['none', 'null', '']:
        result = days_between_by_rule(issued_date, actual_reply, rule)
        return max(0, result) if result is not None else 0
    else:
        yesterday = (datetime.date.today() - datetime.timedelta(days=1)).isoformat()
        result = days_between_by_rule(issued_date, yesterday, rule)
        return max(0, result) if result is not None else 0

def extract_rev(doc_no):
    m = re.search(r'REV(\d+)', doc_no or '', re.IGNORECASE)
    return int(m.group(1)) if m else 0

def extract_seq(doc_no, prefix):
    m = re.match(rf'^{re.escape(prefix)}-(\d+)\s+REV(\d+)$', doc_no or '', re.IGNORECASE)
    return (int(m.group(1)), int(m.group(2))) if m else None

def build_doc_no(prefix, num, rev):
    return f"{prefix}-{num:03d} REV{rev:02d}"

def get_next_doc_no(prefix, records):
    """Find highest existing base number and return next one."""
    max_num = 0
    for r in records:
        p = extract_seq(r.get('docNo',''), prefix)
        if p: max_num = max(max_num, p[0])
    if max_num == 0:
        return build_doc_no(prefix, 1, 0)   # first doc
    return build_doc_no(prefix, max_num + 1, 0)

def extract_plain_seq(doc_no, prefix):
    m = re.match(rf'^{re.escape(prefix)}-(\d+)$', doc_no or '', re.IGNORECASE)
    return int(m.group(1)) if m else None

def build_plain_doc_no(prefix, num, width=3):
    return f"{prefix}-{num:0{width}d}"

def get_next_plain_doc_no(prefix, records):
    """Find highest existing plain numeric suffix and return next one without REV."""
    max_num = 0
    max_width = 3
    tail_re = re.compile(rf'^{re.escape(prefix)}-(\d+)$', re.IGNORECASE)
    for r in records:
        doc_no = str(r.get('docNo', '') or '').strip()
        m = tail_re.match(doc_no)
        if not m:
            continue
        num_txt = m.group(1)
        max_num = max(max_num, int(num_txt))
        max_width = max(max_width, len(num_txt))
    if max_num == 0:
        return build_plain_doc_no(prefix, 1, max_width)
    return build_plain_doc_no(prefix, max_num + 1, max_width)

def doc_type_uses_revision(dt_code, records=None):
    code = str(dt_code or "").strip().upper()
    if code in NON_REV_DOC_TYPE_CODES:
        return False
    rows = records or []
    if any("REV" in str(r.get("docNo", "") or "").upper() for r in rows):
        return True
    if rows:
        return False
    return True

def get_pmo_dates(row):
    if not isinstance(row, dict): return None
    d1, d2, d3, d4 = None, None, None, None
    has_pmo = False
    for k, v in row.items():
        lk = k.lower()
        if "rec" in lk and "scas" in lk: has_pmo = True
        if "rec" in lk and "style" in lk: has_pmo = True
        
        if not v or str(v).strip() == "": continue
        if "rec" in lk and "from" in lk and "scas" in lk: d1 = str(v).strip()
        elif "rec" in lk and "by" in lk and "style" in lk: d2 = str(v).strip()
        elif "rec" in lk and "from" in lk and "style" in lk: d3 = str(v).strip()
        elif "rec" in lk and "by" in lk and "scas" in lk: d4 = str(v).strip()
        
    if not has_pmo: return None
    return d1, d2, d3, d4

def compute_expected_reply(issued_date, doc_no, rule=None, status=None, action=None, row=None):
    chk = f"{status or ''} {action or ''}".upper()
    if "FOR INFORMATION" in chk or "FI" in re.findall(r'[A-Z]+', chk):
        return None

    pmo_dates = get_pmo_dates(row) if row else None
    if pmo_dates is not None:
        d1, d2, d3, d4 = pmo_dates
        if d4: return None # State 4: Closed
        cfg = normalize_expected_reply_rule(rule)
        rev = 0
        try: rev = extract_rev(doc_no)
        except Exception: pass
        
        if d3 and not d4: # State 3
            return add_working_days(d3, 1, cfg["weekend_mode"], cfg["exclude_official_holidays"])
        elif d2 and not d3: # State 2
            days = cfg["rev0_reply_days"] if rev == 0 else cfg["rev_reply_days"]
            return add_working_days(d2, days, cfg["weekend_mode"], cfg["exclude_official_holidays"])
        elif d1 and not d2: # State 1
            return add_working_days(d1, 2, cfg["weekend_mode"], cfg["exclude_official_holidays"])
        
        if not d1: return None # Not even started

    if not issued_date: return None
    try:
        rev = extract_rev(doc_no)
    except Exception:
        rev = 0
    try:
        cfg = normalize_expected_reply_rule(rule)
        days = cfg["rev0_reply_days"] if rev == 0 else cfg["rev_reply_days"]
        if cfg["calculation_mode"] == "calendar_days":
            return add_calendar_days(issued_date, days)
        return add_working_days(
            issued_date,
            days,
            cfg["weekend_mode"],
            cfg["exclude_official_holidays"],
        )
    except Exception:
        return None

def is_overdue(issued_date, doc_no, actual_reply, has_expected_reply_col=True, rule=None, status=None, action=None, row=None):
    """Returns True only if the doc type has an Expected Reply column and is past due."""
    pmo_dates = get_pmo_dates(row) if row else None
    if pmo_dates is not None:
        d1, d2, d3, d4 = pmo_dates
        if d4: return False # State 4: Closed
        exp = compute_expected_reply(issued_date, doc_no, rule, status, action, row)
        if not exp: return False
        exp_d = parse_any_date(exp)
        if not exp_d: return False
        return exp_d < datetime.date.today()

    if not has_expected_reply_col: return False

    if actual_reply and str(actual_reply).strip().lower() not in ['none', 'null', '']: return False
    exp = compute_expected_reply(issued_date, doc_no, rule, status, action, row)
    if not exp: return False
    exp_d = parse_any_date(exp)
    if not exp_d: return False
    return exp_d < datetime.date.today()

def format_date(d):
    """Format as DD-MM-YYYY (LTR)."""
    if not d: return ''
    d_obj = parse_any_date(d)
    if not d_obj: return str(d) or ''
    return d_obj.strftime("%d-%m-%Y")

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
