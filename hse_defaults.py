"""Defaults for new HSE projects. See docs/hse/HSE_DECISION_LOG.md for provenance."""

import json
import re
import zlib

HSE_REGISTERS = (
    ("PTW", "Permit to Work", "DOC_CATEGORIES.permits"),
    ("OBS", "Safety Observations", "DOC_CATEGORIES.OBS_KEYS"),
    ("INC", "Incidents and First Aid", "DOC_CATEGORIES.first_aid / injuries"),
    ("SVP", "Safety Violations and Penalties", "DOC_CATEGORIES.exclusion / financial_violation"),
    ("TBT", "Toolbox Talks", "DOC_CATEGORIES.toolbox / Tool Box Talk TBT.xlsx"),
    ("FEI", "Fire and Equipment Inspections", "TYPES.extinguisher / FORAM inspection forms"),
    ("JSA", "Job Safety Analysis and Risk Assessments", "risk_assessments"),
    ("CEV", "Contractor Evaluations", "contractor_eval_form"),
    ("EVAC", "Evacuation Drills", "evacuation_drills"),
    ("MHR", "Man-hours", "man_hours"),
    ("WKR", "Worker Register", "TYPES.person"),
)

# Ordered legacy category keys and the labels shown in the DCR subtype dropdown.
HSE_SUBTYPES = {
    "PTW": {
        "hot_permit": "Hot Work Permit",
        "cold_permit": "Cold Work Permit",
        "height_permit": "Work at Height Permit",
        "confined_permit": "Confined Space Permit",
        "excavation_permit": "Excavation Permit",
        "electrical_permit": "Electrical / Isolation Permit",
        "lifting_permit": "Lifting Operations Permit",
        "radiography_permit": "Radiography Permit",
    },
    "OBS": {
        "unsafe_act": "Unsafe Act",
        "unsafe_condition": "Unsafe Condition",
        "near_miss": "Near Miss",
    },
    "INC": {
        "first_aid": "First Aid Report",
        "injuries": "Injury Register",
    },
    "SVP": {
        "exclusion": "Exclusion Report",
        "financial_violation": "Financial Violation Report",
        "HSE Violation Warning.docx": "HSE Violation Warning",
    },
    "FEI": {
        "extinguisher": "Fire Extinguisher Inspection",
        "hand_tools_inspection": "Hand Tools Inspection",
        "electrical_panel_inspection": "Electrical Panel Inspection",
        "safety_harness_inspection": "Safety Harness Inspection",
        "welding_machine_inspection": "Welding Machine Inspection",
    },
}

HSE_LISTS = {
    "hse_doc_status": ("Open", "Closed"),
    "hse_injury_class": ("FAC", "MTC", "LTI", "Near Miss", "Fatality"),
    "hse_body_part": (
        "Head / Face",
        "Neck",
        "Chest",
        "Abdomen",
        "Upper Back",
        "Lower Back",
        "Left Shoulder",
        "Right Shoulder",
        "Left Arm",
        "Right Arm",
        "Left Hand",
        "Right Hand",
        "Pelvis / Hip",
        "Left Thigh",
        "Right Thigh",
        "Left Knee",
        "Right Knee",
        "Left Leg",
        "Right Leg",
        "Left Foot",
        "Right Foot",
    ),
    "hse_extinguisher_type": (
        "CO2",
        "Dry Powder",
        "Foam",
        "Water",
        "Wet Chemical",
        "Fire Hose Reel",
    ),
    "hse_extinguisher_status": (
        "Valid",
        "Needs Maintenance",
        "Expired",
        "Under Refill",
        "Out of Service",
    ),
    "hse_jsa_status": ("Approved", "Under Review", "Draft", "Revised"),
    "hse_evac_scenario": (
        "Fire Emergency Drill",
        "Gas / Chemical Leak",
        "Structural Collapse / Earthquake",
        "General Site Evacuation",
        "Power Failure / Trapped Personnel",
    ),
    "hse_evac_grade": ("Excellent", "Very Good", "Satisfactory", "Needs Retraining"),
    "hse_alarm_check": ("Yes", "Partial", "No"),
    "hse_yes_no": ("Yes", "No"),
    "hse_worker_affiliation": ("Contractor", "Company Staff"),
    "hse_worker_role": (
        "Staff",
        "Worker",
        "Storekeeper",
        "Safety Officer",
        "Supervisor",
        "Engineer",
        "Visitor",
    ),
}
HSE_LISTS.update({f"hse_{code.lower()}_subtype": tuple(values.values()) for code, values in HSE_SUBTYPES.items()})


# Column key, label, DCR type, optional project-local dropdown name.
DOC_NO = ("docNo", "Register No.", "docno", None)
TITLE = ("title", "Subject", "text", None)
LOCATION = ("location", "Location", "text", None)
RESPONSIBLE = ("responsible", "Responsible Person", "text", None)
NOTES = ("remarks", "Notes", "long_text", None)
DOC_STATUS = ("status", "Status", "dropdown", "hse_doc_status")
YEAR = ("year", "Year", "number", None)
MONTH = ("month", "Month (1–12)", "number", None)
CORRECTIVE_ACTION = ("correctiveAction", "Corrective Action", "long_text", None)
FILE_LOCATION = ("fileLocation", "Drive Link", "url", None)

HSE_COLUMNS = {
    "PTW": (
        TITLE, ("issuedDate", "Permit Date", "date", None),
        ("validUntil", "Valid Until", "date", None), LOCATION, RESPONSIBLE, DOC_STATUS, NOTES,
    ),
    "OBS": (TITLE, ("issuedDate", "Observation Date", "date", None), LOCATION, RESPONSIBLE, DOC_STATUS, NOTES),
    "INC": (
        TITLE, ("issuedDate", "Incident Date", "date", None),
        ("incidentTime", "Incident Time", "text", None), LOCATION,
        ("injuryClass", "Incident Classification", "dropdown", "hse_injury_class"),
        ("bodyPart", "Injured Body Part", "dropdown", "hse_body_part"),
        ("lostDays", "Lost Days", "number", None),
        ("description", "Incident Description", "long_text", None),
        ("rootCause", "Root Cause", "long_text", None), CORRECTIVE_ACTION, NOTES,
    ),
    "SVP": (TITLE, ("issuedDate", "Report Date", "date", None), LOCATION, RESPONSIBLE, DOC_STATUS, NOTES),
    "TBT": (("title", "Talk Topic", "text", None), ("issuedDate", "Talk Date", "date", None), LOCATION, RESPONSIBLE, NOTES),
    "FEI": (
        TITLE, ("equipmentRef", "Equipment / Extinguisher No.", "text", None),
        ("issuedDate", "Inspection Date", "date", None), LOCATION,
        ("inspector", "Inspector", "text", None),
        ("extinguisherType", "Extinguisher Type", "dropdown", "hse_extinguisher_type"),
        ("capacity", "Extinguisher Capacity", "text", None),
        ("manufactureDate", "Manufacture Date", "date", None),
        ("expiryDate", "Expiry / Next Inspection Date", "date", None),
        ("extinguisherStatus", "Extinguisher Status", "dropdown", "hse_extinguisher_status"),
        ("conditionNotes", "Condition Notes", "long_text", None), NOTES,
    ),
    "JSA": (
        TITLE, ("issuedDate", "Assessment Date", "date", None),
        ("reviewDate", "Review Date", "date", None), LOCATION,
        ("activityName", "Activity", "text", None), ("department", "Department", "text", None),
        ("assessorName", "Assessor", "text", None), ("approverName", "Approver Name", "text", None),
        ("status", "Assessment Status", "dropdown", "hse_jsa_status"), NOTES,
    ),
    "CEV": (
        ("title", "Contractor", "text", None), YEAR, MONTH, ("issuedDate", "Evaluation Date", "date", None),
        ("evaluator", "Evaluator", "text", None), ("positiveNotes", "Strengths", "long_text", None),
        ("negativeNotes", "Observations and Deficiencies", "long_text", None),
        ("actionPlan", "Corrective Action Plan", "long_text", None),
    ),
    "EVAC": (
        TITLE, ("issuedDate", "Drill Date", "date", None),
        ("drillTime", "Drill Time", "text", None), ("scenario", "Scenario", "dropdown", "hse_evac_scenario"),
        LOCATION, ("assemblyPoint", "Assembly Point", "text", None), ("leaderName", "Drill Leader", "text", None),
        ("targetDurationSec", "Target Duration (seconds)", "number", None),
        ("actualDurationSec", "Actual Duration (seconds)", "number", None),
        ("evaluationGrade", "Recorded Evaluation", "dropdown", "hse_evac_grade"),
        ("alarmWorking", "Alarm Working", "dropdown", "hse_alarm_check"),
        ("exitRoutesClear", "Exit Routes Clear", "dropdown", "hse_yes_no"),
        ("assemblyAdequate", "Assembly Area Adequate", "dropdown", "hse_yes_no"),
        CORRECTIVE_ACTION, NOTES,
    ),
    "MHR": (YEAR, MONTH, ("hours", "Worked Hours", "number", None)),
    "WKR": (
        ("title", "Worker Name", "text", None),
        ("affiliation", "Affiliation", "dropdown", "hse_worker_affiliation"),
        ("company", "Company / Contractor", "text", None),
        ("personType", "Person Type", "dropdown", "hse_worker_role"),
        ("jobTitle", "Job Title", "text", None), ("hireDate", "Start Date", "date", None), NOTES,
    ),
}


PTW_SUBTYPE_CODES = {
    "Hot Work Permit": "HWP",
    "Cold Work Permit": "CWP",
    "Work at Height Permit": "WAH",
    "Confined Space Permit": "CSP",
    "Excavation Permit": "EXP",
    "Electrical / Isolation Permit": "EIP",
    "Lifting Operations Permit": "LOP",
    "Radiography Permit": "RGP",
}

_STATUS_META_TOKENS = frozenset({"PENDING", "APPROVED", "REJECTED", "CANCELLED", "INFOCLOSED", "INFO_CLOSED"})
_PTW_LIST_ITEM_BRACKET_RE = re.compile(r"^(.*?)\s*\[\s*([A-Za-z0-9_-]{2,8})\s*\]\s*$")
_PTW_LIST_ITEM_PIPE_RE = re.compile(r"^(.*?)\s*\|\s*([A-Za-z0-9_-]{2,8})\s*$")


def normalize_ptw_subtype_value(val) -> str:
    """Normalize legacy string/list/CSV subtype values into a single stripped string."""
    if val is None:
        return ""
    if isinstance(val, (list, tuple)):
        for item in val:
            cleaned = normalize_ptw_subtype_value(item)
            if cleaned:
                return cleaned
        return ""
    s = str(val).strip()
    if not s:
        return ""
    if s.startswith("["):
        try:
            parsed = json.loads(s)
            if isinstance(parsed, list):
                return normalize_ptw_subtype_value(parsed)
        except Exception:
            pass
    if "," in s:
        if s in PTW_SUBTYPE_CODES or any(s.lower() == k.lower() for k in PTW_SUBTYPE_CODES):
            return s
        for part in s.split(","):
            part_s = part.strip()
            if part_s:
                return part_s
        return ""
    return s


def parse_ptw_list_item_input(raw_item: str) -> tuple[str, str | None]:
    """Extract optional explicit code suffix from 'Subtype Name [CODE]' or 'Subtype Name | CODE'."""
    s = str(raw_item or "").strip()
    if not s:
        return "", None
    for pat in (_PTW_LIST_ITEM_BRACKET_RE, _PTW_LIST_ITEM_PIPE_RE):
        m = pat.match(s)
        if m:
            name = m.group(1).strip()
            code = re.sub(r"[^A-Z0-9]", "", m.group(2).upper())
            if name and code:
                return name, code
    return s, None


def _extract_code_from_meta(raw_meta) -> str:
    if raw_meta is None:
        return ""
    if isinstance(raw_meta, dict):
        cand = raw_meta.get("code") or raw_meta.get("subtype_code") or ""
        return _extract_code_from_meta(cand)
    s = str(raw_meta).strip()
    if not s:
        return ""
    if s.startswith("{"):
        try:
            parsed = json.loads(s)
            if isinstance(parsed, dict):
                return _extract_code_from_meta(parsed)
        except Exception:
            pass
    if s.upper() in _STATUS_META_TOKENS:
        return ""
    cleaned = re.sub(r"[^A-Z0-9]", "", s.upper())
    if not cleaned or cleaned in _STATUS_META_TOKENS:
        return ""
    return cleaned[:8]


def get_ptw_subtype_code(subtype_name: str, custom_meta=None) -> str:
    """Resolve the configurable Doc No. code for a PTW subtype name."""
    name = normalize_ptw_subtype_value(subtype_name)
    if not name:
        return ""
    if isinstance(custom_meta, dict):
        if name in custom_meta:
            c = _extract_code_from_meta(custom_meta[name])
            if c:
                return c
        low = name.lower()
        for k, v in custom_meta.items():
            if str(k).strip().lower() == low:
                c = _extract_code_from_meta(v)
                if c:
                    return c
    elif custom_meta is not None:
        c = _extract_code_from_meta(custom_meta)
        if c:
            return c

    if name in PTW_SUBTYPE_CODES:
        return PTW_SUBTYPE_CODES[name]
    low = name.lower()
    for k, v in PTW_SUBTYPE_CODES.items():
        if k.lower() == low:
            return v

    # Deterministic fallback code for custom subtypes added via 'Lists' without explicit code metadata.
    ascii_words = re.findall(r"[A-Za-z0-9]+", name)
    if ascii_words:
        if len(ascii_words) == 1 and 2 <= len(ascii_words[0]) <= 5 and ascii_words[0].isupper():
            return ascii_words[0]
        has_permit = ascii_words[-1].lower() in ("permit", "permits") and len(ascii_words) > 1
        base_words = ascii_words[:-1] if has_permit else ascii_words
        sig_words = [w for w in base_words if w.lower() not in {"of", "and", "the", "for", "at", "in", "or", "on", "to"}] or base_words
        if has_permit:
            if len(sig_words) >= 2:
                return (sig_words[0][0] + sig_words[1][0] + "P").upper()
            w = sig_words[0].upper()
            return (w[:2] + "P") if len(w) >= 2 else (w + "WP")
        if len(sig_words) >= 3:
            return "".join(w[0] for w in sig_words[:3]).upper()
        if len(sig_words) == 2:
            return (sig_words[0][:2] + sig_words[1][0]).upper()
        w = sig_words[0].upper()
        return w[:3] if len(w) >= 3 else (w + "P")

    digest = zlib.crc32(name.encode("utf-8")) % 900 + 100
    return f"P{digest}"


def resolve_ptw_project_prefix(pid: str, project: dict | None = None, records: list[dict] | None = None) -> str:
    """Resolve the active project's PTW Doc No. prefix (e.g., PEM058 for HSE-CFC-P2)."""
    proj = project if isinstance(project, dict) else {}
    for key in ("ptw_prefix", "doc_prefix"):
        explicit = str(proj.get(key) or "").strip()
        if explicit:
            return re.sub(r"[^A-Za-z0-9_-]", "", explicit).upper()

    pid_up = str(pid or "").strip().upper()
    code_up = str(proj.get("code") or "").strip().upper()
    if pid_up == "HSE-CFC-P2" or code_up == "HSE-CFC-P2":
        return "PEM058"

    for rec in records or []:
        doc_no = str((rec or {}).get("docNo") or "").strip()
        m = re.match(r"^([A-Za-z]{2,8})-?(\d{2,4})-", doc_no)
        if m:
            return f"{m.group(1).upper()}{m.group(2)}"

    raw = code_up or pid_up
    stripped = re.sub(r"^HSE[-_]+", "", raw)
    if stripped and re.search(r"[A-Z]", stripped):
        cleaned = re.sub(r"[^A-Z0-9]", "", stripped)
    else:
        cleaned = re.sub(r"[^A-Z0-9]", "", raw)
    return cleaned or "PTW"


def generate_next_ptw_doc_no(
    pid: str,
    subtype: str,
    project: dict | None = None,
    records: list[dict] | None = None,
    subtype_meta_map: dict | None = None,
) -> dict:
    """Generate the next sequential PTW Doc No. independently per (project, subtype)."""
    clean_subtype = normalize_ptw_subtype_value(subtype)
    prefix = resolve_ptw_project_prefix(pid, project, records)
    code = get_ptw_subtype_code(clean_subtype, subtype_meta_map)
    if not clean_subtype or not code:
        return {"next": "", "prefix": prefix, "subtype": clean_subtype, "subtype_code": "", "seq": 0}

    pat = re.compile(rf"^(?:[A-Z0-9_-]+-)?{re.escape(code)}-(\d+)(?:\s+REV\d+)?$", re.IGNORECASE)
    max_seq = 0
    for rec in records or []:
        doc_no = str((rec or {}).get("docNo") or "").strip()
        if not doc_no:
            continue
        m = pat.match(doc_no)
        if m:
            max_seq = max(max_seq, int(m.group(1)))

    next_seq = max_seq + 1
    return {
        "next": f"{prefix}-{code}-{next_seq:03d}",
        "prefix": prefix,
        "subtype": clean_subtype,
        "subtype_code": code,
        "seq": next_seq,
    }


def hse_columns(code):
    subtype = (
        (("legacySubtype", "Subtype", "dropdown", f"hse_{code.lower()}_subtype"),)
        if code in HSE_SUBTYPES else ()
    )
    return (DOC_NO, *subtype, *HSE_COLUMNS[code], FILE_LOCATION)

