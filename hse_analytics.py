"""Pure HSE analytics calculations and RBAC-scoped HSE dashboard assembly."""

from datetime import date, datetime, timedelta
import json

from hse_defaults import HSE_REGISTERS

REGISTER_ORDER = [code for code, _, _ in HSE_REGISTERS]
REGISTER_NAMES = {code: name for code, name, _ in HSE_REGISTERS}

STALE_OBS_DAYS = 30
DEFAULT_WARN_DAYS = 30
HOURS_PER_LOST_DAY = 8

INJURY_LTI_CLASSES = {"LTI", "FATALITY"}
INJURY_ANY_CLASSES = {"FAC", "MTC", "LTI", "FATALITY"}


def parse_hse_date(val):
    """Parse date, datetime, or ISO/DMY string into datetime.date, else None."""
    if val is None or isinstance(val, bool):
        return None
    if isinstance(val, datetime):
        return val.date()
    if isinstance(val, date):
        return val
    text = str(val).strip()
    if not text:
        return None
    head = text[:10]
    try:
        return date.fromisoformat(head)
    except ValueError:
        pass
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d"):
        try:
            return datetime.strptime(head, fmt).date()
        except ValueError:
            continue
    return None


def normalize_hse_status(val):
    """Normalize English/legacy HSE document lifecycle status to 'Open' or 'Closed'."""
    raw = str(val or "").strip()
    if not raw:
        return ""
    low = raw.lower()
    if low in ("open", "pending", "in progress", "active") or raw == "مفتوح":
        return "Open"
    if "not closed" not in low and (
        any(tok in low for tok in ("closed", "close", "complete", "approved")) or raw == "مغلق"
    ):
        return "Closed"
    return raw


def _parse_int(val, default=0, minimum=None):
    try:
        num = int(float(val))
    except (TypeError, ValueError):
        return default
    return max(minimum, num) if minimum is not None else num


def _parse_float(val, default=0.0, minimum=None):
    try:
        num = float(val)
    except (TypeError, ValueError):
        return default
    return max(minimum, num) if minimum is not None else num


def _first_present(mapping, *keys):
    for key in keys:
        val = mapping.get(key)
        if val not in (None, ""):
            return val
    return None


def _normalize_injury_class(val):
    raw = str(val or "").strip().upper()
    aliases = {
        "FIRST AID": "FAC",
        "FIRST AID CASE": "FAC",
        "MEDICAL TREATMENT": "MTC",
        "MEDICAL TREATMENT CASE": "MTC",
        "LOST TIME INJURY": "LTI",
        "NEAR MISS": "NEAR MISS",
    }
    return aliases.get(raw, raw)


def normalize_hse_record(raw, dt_code_lookup=None, project_meta=None):
    """Normalize a PostgreSQL records row or flat dict into an analytics record dict."""
    if not isinstance(raw, dict):
        return None
    if raw.get("_normalized") is True:
        return raw
    payload = raw.get("data")
    if isinstance(payload, str):
        try:
            payload = json.loads(payload)
        except Exception:
            payload = {}
    merged = {**(payload if isinstance(payload, dict) else {}), **raw}
    merged.pop("data", None)

    pid = str(merged.get("project_id") or merged.get("_project_id") or "").strip()
    dt_id = str(merged.get("dt_id") or merged.get("_dt_id") or merged.get("dt_code") or "").strip()
    code_hint = merged.get("dt_code") or (dt_code_lookup or {}).get((pid, dt_id.upper())) or dt_id
    dt_code = str(code_hint or "").strip().upper()

    year = _parse_int(_first_present(merged, "year"), default=None)
    month = _parse_int(_first_present(merged, "month"), default=None)
    doc_date = parse_hse_date(
        _first_present(merged, "issuedDate", "issued_date", "incidentDate", "incident_date", "doc_date", "hireDate", "date")
    )
    if doc_date is None and dt_code == "MHR" and year and month and 1 <= month <= 12:
        doc_date = date(year, month, 1)
    if doc_date is None and dt_code != "FEI":
        doc_date = parse_hse_date(merged.get("created_at"))
    if year is None and doc_date is not None:
        year = doc_date.year
    if month is None and doc_date is not None:
        month = doc_date.month

    injury_class = _normalize_injury_class(
        _first_present(merged, "injuryClass", "injury_class", "injuryType", "injury_type")
    )
    lost_days = (
        _parse_int(_first_present(merged, "lostDays", "lost_days"), default=0, minimum=0)
        if injury_class in INJURY_LTI_CLASSES else 0
    )

    valid_until = parse_hse_date(_first_present(merged, "validUntil", "valid_until"))
    expiry_date = parse_hse_date(_first_present(merged, "expiryDate", "expiry_date", "validUntil", "valid_until"))
    man_hours = _parse_float(_first_present(merged, "hours", "man_hours", "worked_hours"), default=0.0, minimum=0.0)
    proj_info = (project_meta or {}).get(pid, {})

    return {
        "id": str(merged.get("id") or ""),
        "project_id": pid,
        "project_code": str(merged.get("project_code") or proj_info.get("code") or pid),
        "project_name": str(merged.get("project_name") or proj_info.get("name") or pid),
        "dt_id": dt_id or dt_code,
        "dt_code": dt_code,
        "doc_no": str(_first_present(merged, "docNo", "doc_no") or ""),
        "title": str(_first_present(merged, "title", "equipmentRef") or ""),
        "subtype": str(_first_present(merged, "legacySubtype", "subtype", "category") or ""),
        "location": str(merged.get("location") or "").strip(),
        "responsible": str(_first_present(merged, "responsible", "inspector", "evaluator") or "").strip(),
        "status": normalize_hse_status(merged.get("status")),
        "date": doc_date,
        "year": year,
        "month": month,
        "valid_until": valid_until,
        "expiry_date": expiry_date,
        "injury_class": injury_class,
        "lost_days": lost_days,
        "man_hours": man_hours,
        "closed_date": parse_hse_date(_first_present(merged, "closedDate", "closed_date", "closedAt", "closed_at")),
        "_normalized": True,
    }


def normalize_hse_records(rows, dt_code_lookup=None, project_meta=None):
    """Normalize a sequence of raw records into clean HSE analytics dicts."""
    out = []
    for row in rows or ():
        norm = normalize_hse_record(row, dt_code_lookup=dt_code_lookup, project_meta=project_meta)
        if norm is not None:
            out.append(norm)
    return out


def _compute_single_streak(event_dates, today, baseline_date=None, first_date=None):
    base = baseline_date if (baseline_date and baseline_date <= today) else None
    start = base or first_date
    if start is None or start > today:
        return {"days": None, "best_days": None, "basis": "none", "ref_date": None, "events": 0}
    events = sorted(d for d in event_dates if start <= d <= today)
    if events:
        ref, basis = events[-1], "event"
    elif base is not None:
        ref, basis = base, "baseline"
    else:
        ref, basis = first_date, "first"
    boundaries = [start] + events + [today]
    best = max((b - a).days for a, b in zip(boundaries, boundaries[1:]))
    return {
        "days": (today - ref).days,
        "best_days": best,
        "basis": basis,
        "ref_date": ref.isoformat(),
        "events": len(events),
    }


def compute_streaks(records, today=None, baseline_date=None):
    """Compute LTI-Free Days and Any-Injury-Free Days streaks from normalized/raw records."""
    ref_day = parse_hse_date(today) or date.today()
    base = parse_hse_date(baseline_date)
    docs = normalize_hse_records(records)

    all_dates = [d["date"] for d in docs if d["date"] and d["date"] <= ref_day]
    first_date = min(all_dates) if all_dates else None

    lti_dates = [
        d["date"] for d in docs
        if d["dt_code"] == "INC" and d["date"] and d["injury_class"] in INJURY_LTI_CLASSES
    ]
    any_dates = [
        d["date"] for d in docs
        if d["dt_code"] == "INC" and d["date"] and d["injury_class"] in INJURY_ANY_CLASSES
    ]
    return {
        "lti_free": _compute_single_streak(lti_dates, ref_day, baseline_date=base, first_date=first_date),
        "any_injury_free": _compute_single_streak(any_dates, ref_day, baseline_date=base, first_date=first_date),
    }


def compute_frequency_rates(lti_count, mtc_count, lost_days, man_hours):
    """Compute LTIF, TRIR, and Severity Rate per 1,000,000 worked man-hours safely."""
    hours = _parse_float(man_hours, default=0.0, minimum=0.0)
    if hours <= 0:
        return {
            "man_hours": 0.0,
            "has_hours": False,
            "ltif": None,
            "trir": None,
            "severity_rate": None,
        }
    factor = 1_000_000.0 / hours
    lti = max(0, int(lti_count or 0))
    mtc = max(0, int(mtc_count or 0))
    lost = max(0, int(lost_days or 0))
    return {
        "man_hours": round(hours, 2),
        "has_hours": True,
        "ltif": round(lti * factor, 2),
        "trir": round((lti + mtc) * factor, 2),
        "severity_rate": round(lost * factor, 2),
    }


def _year_matches(doc, year):
    if year is None:
        return True
    if doc["year"] is not None:
        return doc["year"] == year
    return doc["date"] is not None and doc["date"].year == year


def compute_hse_kpis(records, today=None, year=None, warn_days=DEFAULT_WARN_DAYS, baseline_date=None):
    """Compute HSE KPI summary metrics (streaks, rates, OBS, PTW, FEI, register counts)."""
    ref_day = parse_hse_date(today) or date.today()
    warn_window = max(0, _parse_int(warn_days, default=DEFAULT_WARN_DAYS, minimum=0))
    warn_cutoff = ref_day + timedelta(days=warn_window)
    target_year = _parse_int(year, default=None) if year not in (None, "", "all") else None

    docs = normalize_hse_records(records)
    available_years = sorted(
        {y for d in docs for y in (d["year"], d["date"].year if d["date"] else None) if y},
        reverse=True,
    )

    man_hours = sum(d["man_hours"] for d in docs if d["dt_code"] == "MHR" and _year_matches(d, target_year))

    fac = mtc = lti = fatalities = near_miss_inc = lost_days = 0
    for d in docs:
        if d["dt_code"] != "INC" or not _year_matches(d, target_year):
            continue
        ic = d["injury_class"]
        if ic == "FAC":
            fac += 1
        elif ic == "MTC":
            mtc += 1
        elif ic == "LTI":
            lti += 1
            lost_days += d["lost_days"]
        elif ic == "FATALITY":
            fatalities += 1
            lost_days += d["lost_days"]
        elif ic == "NEAR MISS":
            near_miss_inc += 1

    streaks = compute_streaks(docs, today=ref_day, baseline_date=baseline_date)
    rates = compute_frequency_rates(lti + fatalities, mtc, lost_days, man_hours)

    obs_docs = [d for d in docs if d["dt_code"] == "OBS"]
    obs_open = [d for d in obs_docs if d["status"] == "Open"]
    obs_closed = [d for d in obs_docs if d["status"] == "Closed"]
    obs_stale = [d for d in obs_open if d["date"] and (ref_day - d["date"]).days > STALE_OBS_DAYS]

    ptw_docs = [d for d in docs if d["dt_code"] == "PTW"]
    ptw_open = [d for d in ptw_docs if d["status"] == "Open"]
    ptw_closed = [d for d in ptw_docs if d["status"] == "Closed"]
    ptw_overdue = [d for d in ptw_open if d["valid_until"] and d["valid_until"] < ref_day]
    ptw_expiring = [d for d in ptw_open if d["valid_until"] and ref_day <= d["valid_until"] <= warn_cutoff]

    fei_docs = [d for d in docs if d["dt_code"] == "FEI"]
    fei_expired = [d for d in fei_docs if d["expiry_date"] and d["expiry_date"] < ref_day]
    fei_expiring = [d for d in fei_docs if d["expiry_date"] and ref_day <= d["expiry_date"] <= warn_cutoff]
    fei_valid = [d for d in fei_docs if d["expiry_date"] and d["expiry_date"] > warn_cutoff]
    fei_no_date = [d for d in fei_docs if d["expiry_date"] is None]

    completed_records = sum(1 for d in docs if d["status"] == "Closed")
    open_records = sum(1 for d in docs if d["status"] == "Open")

    return {
        "as_of": ref_day.isoformat(),
        "year": target_year,
        "warn_days": warn_window,
        "available_years": available_years,
        "man_hours": rates["man_hours"],
        "incidents": {
            "fac": fac,
            "mtc": mtc,
            "lti": lti,
            "fatalities": fatalities,
            "near_miss": near_miss_inc,
            "total_injuries": fac + mtc + lti + fatalities,
            "recordable": lti + fatalities + mtc,
            "lost_days": lost_days,
            "lost_hours": lost_days * HOURS_PER_LOST_DAY,
        },
        "rates": rates,
        "streaks": streaks,
        "observations": {
            "total": len(obs_docs),
            "open": len(obs_open),
            "closed": len(obs_closed),
            "stale": len(obs_stale),
            "stale_days": STALE_OBS_DAYS,
        },
        "permits": {
            "total": len(ptw_docs),
            "open": len(ptw_open),
            "closed": len(ptw_closed),
            "overdue": len(ptw_overdue),
            "expiring": len(ptw_expiring),
            "valid_active": max(0, len(ptw_open) - len(ptw_overdue) - len(ptw_expiring)),
        },
        "equipment": {
            "total": len(fei_docs),
            "expired": len(fei_expired),
            "expiring": len(fei_expiring),
            "actionable": len(fei_expired) + len(fei_expiring),
            "valid": len(fei_valid),
            "no_date": len(fei_no_date),
        },
        "summary": {
            "total_records": len(docs),
            "completed_records": completed_records,
            "open_records": open_records,
        },
    }


def _alert_item(doc, alert_type, severity, target_date, days_delta, detail):
    return {
        "id": doc["id"],
        "project_id": doc["project_id"],
        "project_code": doc["project_code"],
        "project_name": doc["project_name"],
        "dt_id": doc["dt_id"],
        "dt_code": doc["dt_code"],
        "doc_no": doc["doc_no"] or doc["id"],
        "title": doc["title"] or REGISTER_NAMES.get(doc["dt_code"], doc["dt_code"]),
        "subtype": doc["subtype"],
        "location": doc["location"],
        "responsible": doc["responsible"],
        "status": doc["status"],
        "alert_type": alert_type,
        "severity": severity,
        "target_date": target_date.isoformat() if target_date else None,
        "days_delta": days_delta,
        "detail": detail,
    }


def compute_hse_alerts(records, today=None, warn_days=DEFAULT_WARN_DAYS, alert_filter=None):
    """Build actionable PTW overdue/expiring, FEI expired/expiring, and stale OBS alerts."""
    ref_day = parse_hse_date(today) or date.today()
    warn_window = max(0, _parse_int(warn_days, default=DEFAULT_WARN_DAYS, minimum=0))
    warn_cutoff = ref_day + timedelta(days=warn_window)
    docs = normalize_hse_records(records)

    alerts = []
    counts = {"ptw_overdue": 0, "ptw_expiring": 0, "fei_expired": 0, "fei_expiring": 0, "obs_stale": 0}
    filter_key = str(alert_filter or "").strip().lower()

    for d in docs:
        code = d["dt_code"]
        if code == "PTW" and d["status"] == "Open" and d["valid_until"]:
            if d["valid_until"] < ref_day:
                overdue_days = (ref_day - d["valid_until"]).days
                counts["ptw_overdue"] += 1
                if not filter_key or filter_key in ("ptw", "overdue"):
                    alerts.append(_alert_item(
                        d, "ptw_overdue", "danger", d["valid_until"], overdue_days,
                        f"Overdue permit ({overdue_days}d past Valid Until)",
                    ))
            elif d["valid_until"] <= warn_cutoff:
                rem = (d["valid_until"] - ref_day).days
                counts["ptw_expiring"] += 1
                if not filter_key or filter_key in ("ptw", "expiring"):
                    alerts.append(_alert_item(
                        d, "ptw_expiring", "warning", d["valid_until"], rem,
                        "Expires today" if rem == 0 else f"Expires in {rem}d",
                    ))
        elif code == "FEI" and d["expiry_date"]:
            if d["expiry_date"] < ref_day:
                overdue_days = (ref_day - d["expiry_date"]).days
                counts["fei_expired"] += 1
                if not filter_key or filter_key in ("fei", "overdue", "expired"):
                    alerts.append(_alert_item(
                        d, "fei_expired", "danger", d["expiry_date"], overdue_days,
                        f"Inspection/equipment expired ({overdue_days}d ago)",
                    ))
            elif d["expiry_date"] <= warn_cutoff:
                rem = (d["expiry_date"] - ref_day).days
                counts["fei_expiring"] += 1
                if not filter_key or filter_key in ("fei", "expiring"):
                    alerts.append(_alert_item(
                        d, "fei_expiring", "warning", d["expiry_date"], rem,
                        "Inspection due today" if rem == 0 else f"Inspection due in {rem}d",
                    ))
        elif code == "OBS" and d["status"] == "Open" and d["date"]:
            age = (ref_day - d["date"]).days
            if age > STALE_OBS_DAYS:
                counts["obs_stale"] += 1
                if not filter_key or filter_key in ("obs", "stale"):
                    alerts.append(_alert_item(
                        d, "obs_stale", "warning", d["date"], age,
                        f"Observation open for {age}d (> {STALE_OBS_DAYS}d SLA)",
                    ))

    sev_rank = {"danger": 0, "warning": 1, "info": 2}
    alerts.sort(key=lambda item: (sev_rank.get(item["severity"], 9), item["target_date"] or "9999-12-31", item["doc_no"]))
    counts["total"] = sum(counts.values())
    return {"alerts": alerts, "counts": counts, "warn_days": warn_window, "as_of": ref_day.isoformat()}


def _month_slots(ref_day, target_year, months_count):
    if target_year is not None:
        return [(target_year, m) for m in range(1, 13)]
    total = max(1, min(24, _parse_int(months_count, default=12, minimum=1)))
    ym = ref_day.year * 12 + (ref_day.month - 1)
    slots = []
    for offset in range(total - 1, -1, -1):
        val = ym - offset
        slots.append((val // 12, (val % 12) + 1))
    return slots


def compute_hse_trends(records, today=None, year=None, months=12, warn_days=DEFAULT_WARN_DAYS):
    """Compute monthly incident/observation trends and register status breakdown."""
    ref_day = parse_hse_date(today) or date.today()
    warn_window = max(0, _parse_int(warn_days, default=DEFAULT_WARN_DAYS, minimum=0))
    warn_cutoff = ref_day + timedelta(days=warn_window)
    target_year = _parse_int(year, default=None) if year not in (None, "", "all") else None

    docs = normalize_hse_records(records)
    slots = _month_slots(ref_day, target_year, months)
    buckets = {}
    monthly_list = []
    for yr, mo in slots:
        key = f"{yr:04d}-{mo:02d}"
        entry = {
            "period": key,
            "label": date(yr, mo, 1).strftime("%b %y"),
            "incidents_fac": 0,
            "incidents_mtc": 0,
            "incidents_lti": 0,
            "incidents_total": 0,
            "observations_open": 0,
            "observations_closed": 0,
            "observations_total": 0,
            "permits_total": 0,
            "man_hours": 0.0,
        }
        buckets[(yr, mo)] = entry
        monthly_list.append(entry)

    reg_map = {
        code: {
            "code": code, "name": REGISTER_NAMES[code],
            "total": 0, "open": 0, "closed": 0, "other": 0,
            "sub_counts": {}, "subtitle": "",
        }
        for code in REGISTER_ORDER
    }
    ptw_status = {"valid_active": 0, "expiring": 0, "overdue": 0, "closed": 0}
    inc_sub = {"lti": 0, "mtc": 0, "fac": 0, "near_miss": 0}
    fei_sub = {"valid": 0, "expiring": 0, "expired": 0, "no_date": 0}
    jsa_sub = {"approved": 0, "pending": 0}
    mhr_periods = set()
    mhr_hours = 0.0

    for d in docs:
        code = d["dt_code"]
        if code in reg_map:
            reg_map[code]["total"] += 1
            if d["status"] == "Open":
                reg_map[code]["open"] += 1
            elif d["status"] == "Closed":
                reg_map[code]["closed"] += 1
            else:
                reg_map[code]["other"] += 1

        if code == "PTW":
            if d["status"] == "Closed":
                ptw_status["closed"] += 1
            elif d["status"] == "Open":
                if d["valid_until"] and d["valid_until"] < ref_day:
                    ptw_status["overdue"] += 1
                elif d["valid_until"] and d["valid_until"] <= warn_cutoff:
                    ptw_status["expiring"] += 1
                else:
                    ptw_status["valid_active"] += 1
        elif code == "INC":
            ic = d["injury_class"]
            if ic in INJURY_LTI_CLASSES:
                inc_sub["lti"] += 1
            elif ic == "MTC":
                inc_sub["mtc"] += 1
            elif ic == "FAC":
                inc_sub["fac"] += 1
            elif ic == "NEAR MISS":
                inc_sub["near_miss"] += 1
        elif code == "FEI":
            if d["expiry_date"] is None:
                fei_sub["no_date"] += 1
            elif d["expiry_date"] < ref_day:
                fei_sub["expired"] += 1
            elif d["expiry_date"] <= warn_cutoff:
                fei_sub["expiring"] += 1
            else:
                fei_sub["valid"] += 1
        elif code == "JSA":
            st_raw = str(d.get("status") or "").strip().lower()
            if d["status"] == "Closed" or "معتمد" in st_raw or "approv" in st_raw:
                jsa_sub["approved"] += 1
            else:
                jsa_sub["pending"] += 1
        elif code == "MHR":
            if d["year"] and d["month"]:
                mhr_periods.add((d["year"], d["month"]))
            mhr_hours += d["man_hours"]

        ym_key = (d["date"].year, d["date"].month) if d["date"] else (d["year"], d["month"])
        bucket = buckets.get(ym_key)
        if bucket is None:
            continue
        if code == "INC":
            ic = d["injury_class"]
            if ic == "FAC":
                bucket["incidents_fac"] += 1
                bucket["incidents_total"] += 1
            elif ic == "MTC":
                bucket["incidents_mtc"] += 1
                bucket["incidents_total"] += 1
            elif ic in INJURY_LTI_CLASSES:
                bucket["incidents_lti"] += 1
                bucket["incidents_total"] += 1
        elif code == "OBS":
            bucket["observations_total"] += 1
            if d["status"] == "Closed":
                bucket["observations_closed"] += 1
            elif d["status"] == "Open":
                bucket["observations_open"] += 1
        elif code == "PTW":
            bucket["permits_total"] += 1
        elif code == "MHR":
            bucket["man_hours"] = round(bucket["man_hours"] + d["man_hours"], 2)

    for code, entry in reg_map.items():
        if code in ("PTW", "OBS", "SVP"):
            entry["sub_counts"] = {"open": entry["open"], "closed": entry["closed"]}
            entry["subtitle"] = f"Open: {entry['open']} • Closed: {entry['closed']}"
        elif code == "INC":
            entry["sub_counts"] = dict(inc_sub)
            entry["subtitle"] = f"LTI: {inc_sub['lti']} • MTC: {inc_sub['mtc']} • FAC: {inc_sub['fac']}"
        elif code == "FEI":
            entry["sub_counts"] = dict(fei_sub)
            entry["subtitle"] = f"Valid: {fei_sub['valid']} • Expiring: {fei_sub['expiring']} • Expired: {fei_sub['expired']}"
        elif code == "MHR":
            logged = len(mhr_periods) if mhr_periods else entry["total"]
            entry["sub_counts"] = {"logged_months": logged, "total_hours": round(mhr_hours, 2)}
            entry["subtitle"] = f"Logged Months: {logged} • {mhr_hours:,.0f} hrs"
        elif code == "WKR":
            entry["sub_counts"] = {"workers": entry["total"]}
            entry["subtitle"] = f"Registered Workers: {entry['total']}"
        elif code == "JSA":
            entry["sub_counts"] = dict(jsa_sub)
            entry["subtitle"] = f"Approved: {jsa_sub['approved']} • Pending/Draft: {jsa_sub['pending']}"
        elif code == "TBT":
            entry["sub_counts"] = {"talks": entry["total"]}
            entry["subtitle"] = f"Logged Talks: {entry['total']}"
        elif code == "CEV":
            entry["sub_counts"] = {"evaluations": entry["total"]}
            entry["subtitle"] = f"Evaluations: {entry['total']}"
        elif code == "EVAC":
            entry["sub_counts"] = {"drills": entry["total"]}
            entry["subtitle"] = f"Logged Drills: {entry['total']}"

    return {
        "monthly_trends": monthly_list,
        "ptw_status": ptw_status,
        "register_breakdown": [reg_map[c] for c in REGISTER_ORDER],
    }


def build_hse_overview(records, projects=None, today=None, year=None, warn_days=DEFAULT_WARN_DAYS, months=12, alert_filter=None):
    """Assemble full pure-Python HSE dashboard payload from already-scoped records."""
    docs = normalize_hse_records(records)
    kpis = compute_hse_kpis(docs, today=today, year=year, warn_days=warn_days)
    alerts_data = compute_hse_alerts(docs, today=today, warn_days=warn_days, alert_filter=alert_filter)
    trends_data = compute_hse_trends(docs, today=today, year=year, months=months, warn_days=warn_days)
    return {
        "kpis": kpis,
        "alerts": alerts_data["alerts"],
        "alerts_counts": alerts_data["counts"],
        "monthly_trends": trends_data["monthly_trends"],
        "ptw_status": trends_data["ptw_status"],
        "register_breakdown": trends_data["register_breakdown"],
        "projects": list(projects or []),
        "available_years": kpis["available_years"],
        "warn_days": kpis["warn_days"],
    }


def get_scoped_hse_analytics(user, pid=None, year=None, warn_days=DEFAULT_WARN_DAYS, months=12, alert_filter=None, today=None):
    """Load HSE records from DB strictly within user's RBAC project & doc-type scope.

    Returns (payload, None) or (None, (error_message, status_code)).
    """
    import db
    from auth import get_allowed_project_ids, get_user_doc_type_scope

    if not user:
        return None, ("LOGIN_REQUIRED", 403)

    allowed_ids = get_allowed_project_ids(user, workspace_id="HSE")
    raw_pid = str(pid or "").strip()
    if raw_pid and raw_pid.lower() != "all":
        proj = db.resolve_project_alias(raw_pid)
        if not proj or str(proj.get("workspace_id") or "").strip().upper() != "HSE" or proj["id"] not in allowed_ids:
            return None, ("Forbidden", 403)
        scoped_pids = [proj["id"]]
    else:
        scoped_pids = list(allowed_ids)

    all_hse_projects = db.get_projects(allowed_ids, workspace_id="HSE") if allowed_ids else []
    proj_meta = {p["id"]: p for p in all_hse_projects}
    proj_summaries = [
        {"id": p["id"], "code": p.get("code") or p["id"], "name": p.get("name") or p["id"]}
        for p in all_hse_projects
    ]

    if not scoped_pids:
        empty = build_hse_overview(
            [], projects=proj_summaries, today=today, year=year,
            warn_days=warn_days, months=months, alert_filter=alert_filter,
        )
        return empty, None

    dt_scope = get_user_doc_type_scope(user, scoped_pids)
    dt_rows = db.q(
        "SELECT project_id, id, code FROM doc_types WHERE project_id = ANY(%s)",
        (scoped_pids,),
    ) or []
    dt_code_lookup = {}
    for row in dt_rows:
        p_id = row.get("project_id")
        d_id = str(row.get("id") or "").strip().upper()
        d_code = str(row.get("code") or row.get("id") or "").strip().upper()
        dt_code_lookup[(p_id, d_id)] = d_code

    rec_rows = db.q(
        "SELECT id, project_id, dt_id, data, created_at, updated_at FROM records WHERE project_id = ANY(%s)",
        (scoped_pids,),
    ) or []

    scoped_rows = []
    for row in rec_rows:
        r_pid = row.get("project_id")
        r_dt = row.get("dt_id")
        dt_code_hint = dt_code_lookup.get((r_pid, str(r_dt or "").strip().upper())) or str(r_dt or "").strip().upper()
        if not db._dt_allowed_in_scope(r_pid, r_dt, dt_scope, dt_code_hint):
            continue
        scoped_rows.append(row)

    normalized = normalize_hse_records(scoped_rows, dt_code_lookup=dt_code_lookup, project_meta=proj_meta)
    payload = build_hse_overview(
        normalized,
        projects=proj_summaries,
        today=today,
        year=year,
        warn_days=warn_days,
        months=months,
        alert_filter=alert_filter,
    )
    return payload, None
