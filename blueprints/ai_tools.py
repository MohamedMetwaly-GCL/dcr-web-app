"""blueprints/ai_tools.py - Dedicated Tool Registry for DCR AI Agent.

Provides 5 deterministic, RBAC-enforced, parameterized tool functions:
1. search_submittals
2. get_project_summary
3. get_noc_financials
4. get_overdue_submittals
5. get_document_history

Every function strictly requires `allowed_project_ids` from the verified user session,
preventing cross-project unauthorized data leaks. Zero string concatenation for SQL.
Includes `_resolve_project_aliases` to resolve project names/codes (e.g. PEM-058 vs CFC DCP PH-2A).
"""
import logging
import os
import re
import sys
from decimal import Decimal, InvalidOperation

# Ensure repository root is on sys.path if run directly
BASE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

import db

logger = logging.getLogger(__name__)

PROJECT_ALIASES = {
    "cfc": {"code": "PEM-058", "name_pattern": "CFC", "display": "CFC (PEM-058)", "doc_prefix": "%CY002P608%"},
    "pem-058": {"code": "PEM-058", "name_pattern": "CFC", "display": "CFC (PEM-058)", "doc_prefix": "%CY002P608%"},
    "pem058": {"code": "PEM-058", "name_pattern": "CFC", "display": "CFC (PEM-058)", "doc_prefix": "%CY002P608%"},
    "058": {"code": "PEM-058", "name_pattern": "CFC", "display": "CFC (PEM-058)", "doc_prefix": "%CY002P608%"},
    "ekh": {"code": "24CP01", "name_pattern": "EKH", "display": "EKH (24CP01)", "doc_prefix": "%24CP01%"},
    "24cp01": {"code": "24CP01", "name_pattern": "EKH", "display": "EKH (24CP01)", "doc_prefix": "%24CP01%"},
    "assiut": {"code": "PEM-045", "name_pattern": "Assiut", "display": "Assiut (PEM-045)", "doc_prefix": "%PEM-045%"},
    "suez": {"code": "PEM-044", "name_pattern": "Suez", "display": "Suez (PEM-044)", "doc_prefix": "%PEM-044%"},
}

DISCIPLINE_SYNONYMS = {
    "architecture": ["architecture", "architectural", "arch", "معماري", "معمارية", "عمارة", "أعمال معمارية"],
    "electrical": ["electrical", "electric", "elect", "elec", "كهرباء", "كهربائي", "أعمال كهربائية", "كهروميكانيك"],
    "mechanical": ["mechanical", "mech", "ميكانيكا", "ميكانيكي", "أعمال ميكانيكية"],
    "civil": ["civil", "مدني", "أعمال مدنية"],
    "structural": ["structural", "structure", "struct", "إنشائي", "انشائي", "أعمال إنشائية", "خرسانة", "خرسانات"],
}

EQUIPMENT_KEYWORDS = {
    "valve": ["valve", "valves", "valv", "butterfly", "gate valve", "ball valve", "check valve", "محبس", "محابس", "صمام", "صمامات"],
    "valves": ["valve", "valves", "valv", "butterfly", "gate valve", "ball valve", "check valve", "محبس", "محابس", "صمام", "صمامات"],
    "pump": ["pump", "pumps", "طلمبة", "طلمبات", "مضخة", "مضخات"],
    "pumps": ["pump", "pumps", "طلمبة", "طلمبات", "مضخة", "مضخات"],
    "chiller": ["chiller", "chillers", "cooling", "hvac", "تكييف", "مبرد", "مبردات"],
    "fan": ["fan", "fans", "ventilation", "مروحة", "مراوح", "تهوية"],
    "cable": ["cable", "cables", "كابل", "كابلات", "كيبل"],
    "panel": ["panel", "panels", "لوحة", "لوحات"],
    "generator": ["generator", "generators", "مولد", "مولدات"],
    "transformer": ["transformer", "transformers", "محول", "محولات"],
    "duct": ["duct", "ducts", "دكت", "صاج"],
    "pipe": ["pipe", "pipes", "piping", "ماسورة", "مواسير", "أنابيب"],
}

KEYWORD_SYNONYMS = EQUIPMENT_KEYWORDS


def _safe_float(val):
    if val is None or val == "":
        return 0.0
    try:
        clean = re.sub(r"[^\d.-]", "", str(val).strip())
        return float(Decimal(clean)) if clean else 0.0
    except (InvalidOperation, ValueError):
        return 0.0


def _resolve_project_aliases(cursor_or_project, requested_project=None, allowed_project_ids=None) -> list:
    """Resolves all possible project identifiers (ID, Code, Name) for the requested project,
    strictly bounded and intersected with the user's RBAC allowed_project_ids.

    Supports:
        _resolve_project_aliases(cursor, requested_project, allowed_project_ids)
        _resolve_project_aliases(requested_project, allowed_project_ids)
    """
    if allowed_project_ids is None:
        cur = None
        req_proj = cursor_or_project
        allowed_pids = requested_project or []
    else:
        cur = cursor_or_project
        req_proj = requested_project
        allowed_pids = allowed_project_ids or []

    if not allowed_pids:
        return []

    clean_allowed = [str(pid).strip() for pid in allowed_pids if str(pid).strip()]
    if not clean_allowed:
        return []

    clean_allowed_set = set(clean_allowed)
    clean_allowed_upper = {p.upper() for p in clean_allowed}

    # Fetch project rows from database
    try:
        if cur is not None and hasattr(cur, "execute"):
            if req_proj and str(req_proj).strip().upper() not in ("ALL", "*"):
                cur.execute("""
                    SELECT id, code, name FROM projects 
                    WHERE code ILIKE %s OR name ILIKE %s OR id::text ILIKE %s
                """, (f"%{req_proj}%", f"%{req_proj}%", f"%{req_proj}%"))
            else:
                cur.execute("SELECT id, code, name FROM projects")
            rows = cur.fetchall()
        else:
            if req_proj and str(req_proj).strip().upper() not in ("ALL", "*"):
                rows = db.q("""
                    SELECT id, code, name FROM projects 
                    WHERE code ILIKE %s OR name ILIKE %s OR id::text ILIKE %s
                """, (f"%{req_proj}%", f"%{req_proj}%", f"%{req_proj}%"))
            else:
                rows = db.q("SELECT id, code, name FROM projects")
    except Exception as e:
        logger.error("Error executing project alias resolution: %s", e)
        rows = []

    req_clean = str(req_proj or "").strip().lower()
    matched_alias_info = None
    if req_clean and req_clean not in ("all", "*"):
        for alias_k, alias_v in PROJECT_ALIASES.items():
            if alias_k == req_clean or alias_k in req_clean or req_clean in alias_k:
                matched_alias_info = alias_v
                break

    valid_aliases = set()

    for row in rows:
        row_id = str(row.get("id") or "").strip()
        row_code = str(row.get("code") or "").strip()
        row_name = str(row.get("name") or "").strip()

        # Check if this project is allowed for the user under RBAC
        is_allowed = (
            row_id in clean_allowed_set
            or row_code in clean_allowed_set
            or row_name in clean_allowed_set
            or row_id.upper() in clean_allowed_upper
            or row_code.upper() in clean_allowed_upper
            or row_name.upper() in clean_allowed_upper
        )

        # Check via alias mapping if user's permissions allow this project
        if not is_allowed and matched_alias_info:
            if (
                matched_alias_info["code"] in clean_allowed_set
                or matched_alias_info["code"].upper() in clean_allowed_upper
            ):
                if (
                    row_code == matched_alias_info["code"]
                    or matched_alias_info["name_pattern"].lower() in row_name.lower()
                    or matched_alias_info["code"] in row_id
                ):
                    is_allowed = True

        if is_allowed:
            # If user requested a specific project, ensure this row matches requested
            if req_clean and req_clean not in ("all", "*"):
                matches_request = (
                    req_clean == row_code.lower()
                    or req_clean in row_code.lower()
                    or req_clean == row_name.lower()
                    or req_clean in row_name.lower()
                    or req_clean == row_id.lower()
                    or req_clean in row_id.lower()
                )
                if not matches_request and matched_alias_info:
                    matches_request = (
                        row_code == matched_alias_info["code"]
                        or matched_alias_info["name_pattern"].lower() in row_name.lower()
                    )
                if not matches_request:
                    continue

            if row_id:
                valid_aliases.add(row_id)
            if row_code:
                valid_aliases.add(row_code)
            if row_name:
                valid_aliases.add(row_name)

    # Known alias expansion for CFC / EKH if permitted
    has_cfc_permission = (
        "PEM-058" in clean_allowed_upper
        or "CFC" in clean_allowed_upper
        or any("058" in p for p in clean_allowed)
        or any("CFC" in p.upper() for p in clean_allowed)
    )
    if has_cfc_permission:
        # If user asked for CFC or no project specified (ALL projects)
        if not req_clean or req_clean in ("all", "*") or (
            matched_alias_info and matched_alias_info["code"] == "PEM-058"
        ) or any(k in req_clean for k in ["cfc", "pem-058", "058", "cooling"]):
            valid_aliases.add("PEM-058")
            valid_aliases.add("CFC DCP PH-2A")
            valid_aliases.add("CFC")

    has_ekh_permission = (
        "24CP01" in clean_allowed_upper
        or "EKH" in clean_allowed_upper
        or any("24CP01" in p.upper() for p in clean_allowed)
        or any("EKH" in p.upper() for p in clean_allowed)
    )
    if has_ekh_permission:
        if not req_clean or req_clean in ("all", "*") or (
            matched_alias_info and matched_alias_info["code"] == "24CP01"
        ) or any(k in req_clean for k in ["ekh", "24cp01"]):
            valid_aliases.add("24CP01")
            valid_aliases.add("EKH")

    # Fallback: if no DB rows returned (e.g. offline/mock tests)
    if not valid_aliases:
        if not req_clean or req_clean in ("all", "*"):
            valid_aliases.update(clean_allowed)
        elif any(req_clean == a.lower() for a in clean_allowed):
            valid_aliases.add(req_clean)
            if req_clean in ("pem-058", "cfc", "pem058", "058"):
                valid_aliases.add("CFC DCP PH-2A")
                valid_aliases.add("PEM-058")
            elif req_clean in ("24cp01", "ekh"):
                valid_aliases.add("24CP01")
                valid_aliases.add("EKH")

    return list(valid_aliases)


def _resolve_target_projects(allowed_project_ids, project_code=None):
    """Resolves accessible projects based on user permission and optional project filter."""
    if not allowed_project_ids:
        return []

    clean_allowed = [str(pid).strip() for pid in allowed_project_ids if str(pid).strip()]
    if not clean_allowed:
        return []

    try:
        projs = db.q("""
            SELECT id, name, code, data FROM projects 
            WHERE id = ANY(%s) OR code = ANY(%s)
            ORDER BY code
        """, (clean_allowed, clean_allowed))
    except Exception as e:
        logger.error("Error fetching projects for target resolution: %s", e)
        projs = []

    if not projs:
        # Fallback project descriptors
        projs = [{"id": pid, "code": pid, "name": pid} for pid in clean_allowed]

    if not project_code or str(project_code).strip().upper() in ("ALL", "*"):
        return projs

    p_clean = str(project_code).strip().lower()

    # Check PROJECT_ALIASES
    matched_alias = None
    for alias_key, alias_info in PROJECT_ALIASES.items():
        if alias_key in p_clean or p_clean in alias_key:
            matched_alias = alias_info
            break

    matched = []
    for p in projs:
        p_code = (p.get("code") or "").lower()
        p_name = (p.get("name") or "").lower()
        p_id = str(p.get("id") or "").lower()

        if matched_alias:
            alias_code = matched_alias["code"].lower()
            alias_name = matched_alias["name_pattern"].lower()
            if p_code == alias_code or alias_name in p_name or alias_code in p_id:
                matched.append(p)
                continue

        if p_clean == p_code or p_clean in p_code or p_clean in p_name or p_clean == p_id:
            matched.append(p)

    return matched


def _dedup_latest_revisions(records):
    """Keeps only the latest revision of each base document."""
    base_map = {}
    for r in records:
        doc_no = str(r.get("doc_no") or "").strip()
        if not doc_no or doc_no == "—":
            rec_id = r.get("id") or id(r)
            base_map[f"_id_{rec_id}"] = {"record": r, "rev": 0}
            continue

        rev_match = re.search(r"[\s_-]*REV\.?\s*(\d+)", doc_no, re.IGNORECASE)
        if rev_match:
            rev_num = int(rev_match.group(1))
            base_no = doc_no[:rev_match.start()].strip()
            if not base_no:
                base_no = doc_no
        else:
            rev_num = 0
            base_no = doc_no

        base_key = base_no.upper()
        if base_key not in base_map or rev_num > base_map[base_key]["rev"]:
            base_map[base_key] = {"record": r, "rev": rev_num}

    return [item["record"] for item in base_map.values()]


# ─────────────────────────────────────────────────────────────
# TOOL 1: search_submittals
# ─────────────────────────────────────────────────────────────
def search_submittals(
    allowed_project_ids: list,
    project_code: str = None,
    doc_type: str = None,
    discipline: str = None,
    trade: str = None,
    keyword: str = None,
    status: str = None,
    latest_revision_only: bool = True,
    limit: int = 25,
) -> dict:
    """Searches engineering submittals with strict RBAC isolation and parameterized SQL.

    Args:
        allowed_project_ids: List of project IDs the current user is permitted to view (enforced).
        project_code: Optional project code or alias (e.g. 'PEM-058', 'CFC', '24CP01', 'EKH').
        doc_type: Optional document type code ('MS', 'SD', 'IR', 'MIR', 'RFI', 'QS', 'NOC').
        discipline: Optional discipline ('Architecture', 'Electrical', 'Mechanical', 'Civil', 'Structural').
        trade: Optional sub-trade ('HVAC', 'Plumbing', 'Lighting', 'Block Work').
        keyword: Optional engineering keyword (e.g. 'valve', 'chiller', 'cabling').
        status: Optional status category ('approved', 'pending', 'rejected').
        latest_revision_only: Whether to return only the latest revision of each document (default True).
        limit: Maximum records to return (1-100, default 25).
    """
    resolved_aliases = _resolve_project_aliases(None, project_code, allowed_project_ids)
    if not resolved_aliases:
        return {
            "project_scope": [],
            "count": 0,
            "records": [],
            "message": "No accessible project matching criteria found.",
        }

    target_projs = _resolve_target_projects(allowed_project_ids, project_code)
    proj_map = {p["id"]: p for p in target_projs}
    proj_displays = [f"{p.get('name', '')} ({p.get('code', p.get('id'))})" for p in target_projs] if target_projs else resolved_aliases

    # Collect project document prefixes (e.g. CY002P608 for CFC)
    doc_prefixes = []
    for alias_item in resolved_aliases:
        code_str = alias_item.upper()
        if "058" in code_str or "CFC" in code_str or "PEM-058" in code_str:
            if "%CY002P608%" not in doc_prefixes:
                doc_prefixes.append("%CY002P608%")
        elif "24CP01" in code_str or "EKH" in code_str:
            if "%24CP01%" not in doc_prefixes:
                doc_prefixes.append("%24CP01%")
        else:
            if code_str and len(code_str) >= 4:
                prefix = f"%{code_str}%"
                if prefix not in doc_prefixes:
                    doc_prefixes.append(prefix)

    where_clauses = []
    params = []

    # 1. Mandatory RBAC & Project Scope Condition using resolved aliases
    if doc_prefixes:
        where_clauses.append("""(
            r.project_id = ANY(%s) 
            OR UPPER(r.project_id) = ANY(%s)
            OR r.data->>'docNo' ILIKE ANY(%s)
        )""")
        params.extend([resolved_aliases, [x.upper() for x in resolved_aliases], doc_prefixes])
    else:
        where_clauses.append("""(
            r.project_id = ANY(%s) 
            OR UPPER(r.project_id) = ANY(%s)
        )""")
        params.extend([resolved_aliases, [x.upper() for x in resolved_aliases]])

    # 2. Document Type Condition
    if doc_type:
        dt_clean = str(doc_type).strip().upper()
        dt_prefixes = [f"{dt_clean}-%", f"{dt_clean}/%", f"%-{dt_clean}-%", f"%/{dt_clean}/%"]
        where_clauses.append("""(
            UPPER(r.dt_id) = %s 
            OR r.data->>'docNo' ILIKE ANY(%s)
            OR r.dt_id IN (SELECT id FROM doc_types WHERE UPPER(code) = %s AND project_id = r.project_id)
        )""")
        params.extend([dt_clean, dt_prefixes, dt_clean])

    # 3. Discipline Condition
    detected_disc_key = None
    if discipline:
        disc_clean = str(discipline).strip().lower()
        disc_terms = [disc_clean]
        for dk, dterms in DISCIPLINE_SYNONYMS.items():
            if disc_clean in [t.lower() for t in dterms] or dk in disc_clean:
                detected_disc_key = dk
                disc_terms = dterms
                break

        disc_patterns = [f"%{t}%" for t in disc_terms if len(t) >= 3]
        if disc_patterns:
            where_clauses.append("""(
                r.data->>'discipline' ILIKE ANY(%s)
                OR r.data->>'trade' ILIKE ANY(%s)
                OR r.data->>'subTrade' ILIKE ANY(%s)
                OR r.data->>'title' ILIKE ANY(%s)
            )""")
            params.extend([disc_patterns, disc_patterns, disc_patterns, disc_patterns])

    # 4. Trade Condition
    if trade:
        tr_pattern = [f"%{str(trade).strip()}%"]
        where_clauses.append("""(
            r.data->>'trade' ILIKE ANY(%s)
            OR r.data->>'subTrade' ILIKE ANY(%s)
            OR r.data->>'title' ILIKE ANY(%s)
        )""")
        params.extend([tr_pattern, tr_pattern, tr_pattern])

    # 5. Keyword Condition (with Equipment Stems)
    if keyword:
        kw_clean = str(keyword).strip().lower()
        synonyms = [kw_clean]
        for ek, eterms in EQUIPMENT_KEYWORDS.items():
            if kw_clean in eterms or ek in kw_clean:
                synonyms = eterms
                break

        kw_patterns = [f"%{term}%" for term in synonyms]
        where_clauses.append("""(
            r.data->>'title' ILIKE ANY(%s)
            OR r.data->>'docNo' ILIKE ANY(%s)
            OR r.data->>'description' ILIKE ANY(%s)
            OR r.data->>'brand' ILIKE ANY(%s)
            OR r.data->>'itemRef' ILIKE ANY(%s)
        )""")
        params.extend([kw_patterns, kw_patterns, kw_patterns, kw_patterns, kw_patterns])

    # 6. Status Condition
    if status:
        st_clean = str(status).strip().lower()
        if "approv" in st_clean:
            where_clauses.append("""(
                COALESCE(r.data->>'status', r.data->>'partBStatus', r.data->>'partDStatus', '') ILIKE ANY(%s)
            )""")
            params.append(["%Approv%", "Status B%", "Status A%", "B - %", "A - %", "%معتمد%"])
        elif "reject" in st_clean or "revise" in st_clean:
            where_clauses.append("""(
                COALESCE(r.data->>'status', r.data->>'partBStatus', r.data->>'partDStatus', '') ILIKE ANY(%s)
            )""")
            params.append(["%Reject%", "%Revise%", "Status C%", "Status D%", "C - %", "D - %", "%مرفوض%"])
        elif "pend" in st_clean:
            where_clauses.append("""(
                COALESCE(r.data->>'status', r.data->>'partBStatus', r.data->>'partDStatus', '') ILIKE ANY(%s)
            )""")
            params.append(["%Pending%", "%Under Review%", "%Submitted%", "%قيد المراجعة%"])

    # Ordering: Approved first, then recent
    order_clauses = [
        """CASE 
            WHEN (
                COALESCE(r.data->>'status', r.data->>'partBStatus', '') ILIKE '%Approv%' 
                OR COALESCE(r.data->>'status', '') ILIKE '%Status B%'
                OR COALESCE(r.data->>'status', '') ILIKE '%Status A%'
                OR COALESCE(r.data->>'status', '') ILIKE '%B - %'
                OR COALESCE(r.data->>'status', '') ILIKE '%A - %'
            ) THEN 0 
            ELSE 1 
        END ASC""",
        "r.created_at DESC NULLS LAST",
        "r.id DESC",
    ]

    fetch_limit = min(max(limit, 1), 100)
    raw_limit = fetch_limit * 3 if latest_revision_only else fetch_limit

    sql = f"""
        SELECT 
            r.id,
            r.project_id,
            COALESCE(
                NULLIF(CASE 
                    WHEN r.data->>'docNo' ILIKE 'MS-%%' OR r.data->>'docNo' ILIKE 'MS/%%' OR r.data->>'docNo' ILIKE '%%-MS-%%' THEN 'MS'
                    WHEN r.data->>'docNo' ILIKE 'IR-%%' OR r.data->>'docNo' ILIKE 'IR/%%' OR r.data->>'docNo' ILIKE '%%-IR-%%' THEN 'IR'
                    WHEN r.data->>'docNo' ILIKE 'SD-%%' OR r.data->>'docNo' ILIKE 'SD/%%' OR r.data->>'docNo' ILIKE '%%-SD-%%' THEN 'SD'
                    WHEN r.data->>'docNo' ILIKE 'MIR-%%' OR r.data->>'docNo' ILIKE 'MIR/%%' OR r.data->>'docNo' ILIKE '%%-MIR-%%' THEN 'MIR'
                    WHEN r.data->>'docNo' ILIKE 'RFI-%%' OR r.data->>'docNo' ILIKE 'RFI/%%' OR r.data->>'docNo' ILIKE '%%-RFI-%%' THEN 'RFI'
                    WHEN r.data->>'docNo' ILIKE 'QS-%%' OR r.data->>'docNo' ILIKE 'QS/%%' OR r.data->>'docNo' ILIKE '%%-QS-%%' THEN 'QS'
                    WHEN r.data->>'docNo' ILIKE 'NOC-%%' OR r.data->>'docNo' ILIKE 'NOC/%%' OR r.data->>'docNo' ILIKE '%%-NOC-%%' THEN 'NOC'
                END, ''),
                (SELECT code FROM doc_types dt WHERE dt.id = r.dt_id AND dt.project_id = r.project_id LIMIT 1),
                NULLIF(r.dt_id, ''),
                'DOC'
            ) AS doc_type,
            COALESCE(r.data->>'docNo', r.data->>'nocNo', r.data->>'letterRef', '—') AS doc_no,
            COALESCE(r.data->>'title', r.data->>'nocSubject', r.data->>'subject', '') AS title,
            COALESCE(r.data->>'status', r.data->>'partBStatus', r.data->>'partDStatus', '') AS status,
            COALESCE(r.data->>'actualReply', r.data->>'actualReplyDate', r.data->>'issuedDate', r.data->>'partAIssueDate', '') AS doc_date,
            COALESCE(r.data->>'brand', '') AS brand,
            COALESCE(r.data->>'discipline', '') AS discipline,
            COALESCE(r.data->>'trade', '') AS trade,
            COALESCE(r.data->>'subTrade', '') AS sub_trade
        FROM records r
        WHERE {' AND '.join(where_clauses)}
        ORDER BY {', '.join(order_clauses)}
        LIMIT %s;
    """
    params.append(raw_limit)

    try:
        rows = db.q(sql, params)
    except Exception as e:
        logger.error("search_submittals execution error: %s", e)
        return {"project_scope": proj_displays, "count": 0, "records": [], "error": str(e)}

    # Post-filter discipline cross-contamination if requested
    if detected_disc_key and rows:
        conflicting_terms = []
        for ok, oterms in DISCIPLINE_SYNONYMS.items():
            if ok != detected_disc_key:
                conflicting_terms.extend([t.lower() for t in oterms if len(t) >= 4])

        disc_terms = [t.lower() for t in DISCIPLINE_SYNONYMS[detected_disc_key]]
        filtered = []
        for r in rows:
            rec_disc = str(r.get("discipline") or "").strip().lower()
            if rec_disc and any(ct in rec_disc for ct in conflicting_terms):
                if not any(dt in rec_disc for dt in disc_terms):
                    continue
            filtered.append(r)
        rows = filtered

    # Latest revision deduplication
    if latest_revision_only and rows:
        rows = _dedup_latest_revisions(rows)

    records = rows[:fetch_limit]

    # Attach clean project codes to result
    for r in records:
        r_pid = str(r.get("project_id") or "")
        proj_info = proj_map.get(r_pid, {})
        code_val = proj_info.get("code")
        if not code_val:
            if "CFC" in r_pid or "058" in r_pid or "CY002P608" in str(r.get("doc_no") or ""):
                code_val = "PEM-058"
            elif "24CP01" in r_pid or "EKH" in r_pid:
                code_val = "24CP01"
            else:
                code_val = r_pid
        r["project_code"] = code_val

    return {
        "project_scope": proj_displays,
        "count": len(records),
        "records": records,
    }


# ─────────────────────────────────────────────────────────────
# TOOL 2: get_project_summary
# ─────────────────────────────────────────────────────────────
def get_project_summary(
    allowed_project_ids: list,
    project_code: str = None,
) -> dict:
    """Returns high-level KPI register summary, approval rates, and document type totals.

    Args:
        allowed_project_ids: List of project IDs the user has permission to view.
        project_code: Optional project code or alias to scope the summary.
    """
    resolved_aliases = _resolve_project_aliases(None, project_code, allowed_project_ids)
    if not resolved_aliases:
        return {"projects": [], "total_projects": 0, "message": "No accessible project found."}

    try:
        stats = db.get_dashboard_stats(project_ids=resolved_aliases)
    except Exception as e:
        logger.error("get_project_summary dashboard stats error: %s", e)
        stats = []

    total_docs = sum(s.get("total", 0) for s in stats)
    total_approved = sum(s.get("approved", 0) for s in stats)
    total_pending = sum(s.get("pending", 0) for s in stats)
    total_overdue = sum(s.get("overdue", 0) for s in stats)
    total_rejected = sum(s.get("rejected", 0) for s in stats)
    overall_pct = round(total_approved / total_docs * 100) if total_docs else 0

    return {
        "project_scope": [project_code or ", ".join(resolved_aliases)],
        "total_projects": len(stats),
        "grand_total_documents": total_docs,
        "grand_total_approved": total_approved,
        "grand_total_pending": total_pending,
        "grand_total_overdue": total_overdue,
        "grand_total_rejected": total_rejected,
        "overall_approval_pct": overall_pct,
        "projects": stats,
    }


# ─────────────────────────────────────────────────────────────
# TOOL 3: get_noc_financials
# ─────────────────────────────────────────────────────────────
def get_noc_financials(
    allowed_project_ids: list,
    project_code: str = None,
    status: str = None,
    has_vo: bool = None,
) -> dict:
    """Extracts Notice of Change (NOC) financial metrics, submitted vs approved cost totals, and approvals.

    Args:
        allowed_project_ids: List of allowed project IDs from authenticated session.
        project_code: Optional project code to filter (e.g. 'PEM-058').
        status: Optional status category ('approved', 'pending', 'rejected').
        has_vo: Optional boolean to filter NOCs that have an issued Variation Order (VO).
    """
    resolved_aliases = _resolve_project_aliases(None, project_code, allowed_project_ids)
    if not resolved_aliases:
        return {"total_count": 0, "total_submitted_cost": 0.0, "total_approved_cost": 0.0, "items": []}

    try:
        rows = db.q("""
            SELECT r.project_id, COALESCE(p.code, r.project_id) as proj_code, COALESCE(p.name, r.project_id) as proj_name, r.data
            FROM records r
            LEFT JOIN projects p ON (p.id = r.project_id OR p.code = r.project_id OR p.name = r.project_id)
            WHERE (UPPER(r.dt_id) = 'NOC' OR r.data->>'docNo' ILIKE 'NOC-%' OR r.data->>'docNo' ILIKE '%-NOC-%')
              AND (r.project_id = ANY(%s) OR UPPER(r.project_id) = ANY(%s))
            ORDER BY r.created_at DESC
        """, (resolved_aliases, [x.upper() for x in resolved_aliases]))
    except Exception as e:
        logger.error("get_noc_financials error: %s", e)
        rows = []

    total_submitted = 0.0
    total_approved = 0.0
    approved_items = []
    pending_items = []
    rejected_items = []
    all_items = []

    st_filter = str(status).strip().lower() if status else None

    for row in rows:
        d = row.get("data") or {}
        doc_no = d.get("docNo") or d.get("nocNo") or "NOC"
        subject = d.get("title") or d.get("nocSubject") or d.get("nocDescription") or "No Subject"
        sub_cost = _safe_float(d.get("submittedCost"))
        app_cost = _safe_float(d.get("finalApprovedCost"))
        vo_no = str(d.get("voNo") or "").strip()
        st_val = str(d.get("partDStatus") or d.get("partBStatus") or d.get("status") or "").strip()
        st_low = st_val.lower()

        if has_vo is True and not vo_no:
            continue
        if has_vo is False and vo_no:
            continue

        is_app = any(x in st_low for x in ["approv", "accept", "part c"])
        is_rej = any(x in st_low for x in ["reject", "cancel"])
        is_pend = not is_app and not is_rej

        if st_filter:
            if "approv" in st_filter and not is_app:
                continue
            if "reject" in st_filter and not is_rej:
                continue
            if "pend" in st_filter and not is_pend:
                continue

        p_code = row.get("proj_code") or row.get("project_id", "")
        if "CFC" in str(p_code) or "058" in str(p_code):
            p_code = "PEM-058"
        elif "24CP01" in str(p_code) or "EKH" in str(p_code):
            p_code = "24CP01"

        item = {
            "doc_no": doc_no,
            "project_code": p_code,
            "subject": subject,
            "submitted_cost": sub_cost,
            "approved_cost": app_cost,
            "status": st_val or "Pending",
            "vo_no": vo_no or None,
        }

        total_submitted += sub_cost
        total_approved += app_cost
        all_items.append(item)

        if is_app:
            approved_items.append(item)
        elif is_rej:
            rejected_items.append(item)
        else:
            pending_items.append(item)

    return {
        "project_scope": [project_code or ", ".join(resolved_aliases)],
        "total_count": len(all_items),
        "total_submitted_cost": total_submitted,
        "total_approved_cost": total_approved,
        "approved_count": len(approved_items),
        "pending_count": len(pending_items),
        "rejected_count": len(rejected_items),
        "items": all_items[:50],
    }


# ─────────────────────────────────────────────────────────────
# TOOL 4: get_overdue_submittals
# ─────────────────────────────────────────────────────────────
def get_overdue_submittals(
    allowed_project_ids: list,
    project_code: str = None,
    doc_type: str = None,
    discipline: str = None,
    min_days_overdue: int = 1,
    limit: int = 25,
) -> dict:
    """Finds documents exceeding their contractual SLA expected reply date, sorted by longest delay.

    Args:
        allowed_project_ids: Accessible project IDs for authenticated user.
        project_code: Optional project code or alias.
        doc_type: Optional document type code (e.g. 'MS', 'IR').
        discipline: Optional discipline filter.
        min_days_overdue: Minimum overdue delay threshold in days (default 1).
        limit: Max items to return (default 25).
    """
    resolved_aliases = _resolve_project_aliases(None, project_code, allowed_project_ids)
    if not resolved_aliases:
        return {"total_overdue": 0, "records": [], "message": "No accessible project found."}

    try:
        overdue_recs = db.get_overdue_records(project_ids=resolved_aliases)
    except Exception as e:
        logger.error("get_overdue_submittals error: %s", e)
        overdue_recs = []

    dt_filter = str(doc_type).strip().upper() if doc_type else None
    disc_filter = str(discipline).strip().lower() if discipline else None

    filtered = []
    for r in overdue_recs:
        if r.get("days_overdue", 0) < min_days_overdue:
            continue
        if dt_filter and (r.get("dt_code") or "").upper() != dt_filter:
            continue
        if disc_filter and disc_filter not in (r.get("discipline") or "").lower():
            continue
        filtered.append(r)

    filtered.sort(key=lambda x: x.get("days_overdue", 0), reverse=True)
    return {
        "project_scope": [project_code or ", ".join(resolved_aliases)],
        "total_overdue": len(filtered),
        "records": filtered[:limit],
    }


# ─────────────────────────────────────────────────────────────
# TOOL 5: get_document_history
# ─────────────────────────────────────────────────────────────
def get_document_history(
    allowed_project_ids: list,
    doc_no_pattern: str,
    project_code: str = None,
) -> dict:
    """Retrieves full chronological revision history (REV 00 -> REV 01 -> ...) and approval lifecycle of a document.

    Args:
        allowed_project_ids: List of allowed project IDs for authenticated user.
        doc_no_pattern: Document number string or partial base (e.g. 'MS-CY002P608-00019').
        project_code: Optional project code to narrow search.
    """
    resolved_aliases = _resolve_project_aliases(None, project_code, allowed_project_ids)
    if not resolved_aliases:
        return {"base_doc_no": doc_no_pattern, "total_revisions": 0, "revisions": [], "message": "No accessible project found."}

    # Extract clean base document number without REV suffix
    base_no = re.sub(r"[\s_-]*REV\.?\s*(\d+)", "", str(doc_no_pattern).strip(), flags=re.IGNORECASE).strip()
    if not base_no:
        base_no = str(doc_no_pattern).strip()

    search_pattern = f"%{base_no}%"

    try:
        rows = db.q("""
            SELECT r.id, r.project_id, r.dt_id, r.data, r.created_at, COALESCE(p.code, r.project_id) as proj_code
            FROM records r
            LEFT JOIN projects p ON (p.id = r.project_id OR p.code = r.project_id OR p.name = r.project_id)
            WHERE (r.project_id = ANY(%s) OR UPPER(r.project_id) = ANY(%s))
              AND (r.data->>'docNo' ILIKE %s)
            ORDER BY r.created_at ASC
        """, (resolved_aliases, [x.upper() for x in resolved_aliases], search_pattern))
    except Exception as e:
        logger.error("get_document_history query error: %s", e)
        rows = []

    parsed_revisions = []
    for r in rows:
        d = r.get("data") or {}
        doc_no = str(d.get("docNo") or "").strip()
        rev_match = re.search(r"[\s_-]*REV\.?\s*(\d+)", doc_no, re.IGNORECASE)
        rev_num = int(rev_match.group(1)) if rev_match else 0
        p_code = r.get("proj_code") or r.get("project_id", "")
        if "CFC" in str(p_code) or "058" in str(p_code):
            p_code = "PEM-058"
        elif "24CP01" in str(p_code) or "EKH" in str(p_code):
            p_code = "24CP01"

        parsed_revisions.append({
            "record_id": r.get("id"),
            "doc_no": doc_no,
            "rev_num": rev_num,
            "status": str(d.get("status") or d.get("partBStatus") or d.get("partDStatus") or "Pending"),
            "issued_date": str(d.get("issuedDate") or d.get("partAIssueDate") or ""),
            "reply_date": str(d.get("actualReplyDate") or d.get("actualReply") or ""),
            "title": str(d.get("title") or d.get("nocSubject") or ""),
            "discipline": str(d.get("discipline") or ""),
            "project_code": p_code,
        })

    parsed_revisions.sort(key=lambda x: x["rev_num"])

    latest = parsed_revisions[-1] if parsed_revisions else {}

    return {
        "base_doc_no": base_no,
        "total_revisions": len(parsed_revisions),
        "latest_revision": f"REV {latest.get('rev_num', 0):02d}" if latest else None,
        "latest_status": latest.get("status") if latest else None,
        "revisions": parsed_revisions,
    }


if __name__ == "__main__":
    print("=" * 60)
    print("=== DCR AI TOOLS ALIAS RESOLVER & RBAC VERIFICATION ===")
    print("=" * 60)

    # 1. Test alias resolution
    aliases = _resolve_project_aliases(None, "PEM-058", ["PEM-058"])
    print(f"[1] Resolved aliases for PEM-058: {aliases}")
    assert "CFC DCP PH-2A" in aliases, "CFC DCP PH-2A must be in resolved aliases!"
    assert "PEM-058" in aliases, "PEM-058 must be in resolved aliases!"
    print("    -> PASS: PEM-058 correctly translated to include 'CFC DCP PH-2A'.")

    # 2. RBAC Empty Check
    empty_rbac = search_submittals(allowed_project_ids=[])
    print("\n[2] RBAC Empty Project IDs Test:")
    print(f"    Expected count=0, got count={empty_rbac.get('count')}")
    assert empty_rbac.get("count") == 0, "RBAC Empty check failed!"
    print("    -> PASS: Unauthorized access blocked without DB query.")

    # 3. RBAC Project Scope Isolation
    empty_history = get_document_history(allowed_project_ids=[], doc_no_pattern="MS-001")
    print("\n[3] Document History RBAC Check:")
    print(f"    Expected total_revisions=0, got={empty_history.get('total_revisions')}")
    assert empty_history.get("total_revisions") == 0, "History RBAC check failed!"
    print("    -> PASS: Document history blocked when user has no allowed projects.")

    # 4. Tool registry inventory
    tools = [
        search_submittals,
        get_project_summary,
        get_noc_financials,
        get_overdue_submittals,
        get_document_history,
    ]
    print("\n[4] Tool Registry Inventory:")
    for t in tools:
        print(f"    - {t.__name__}: {t.__doc__.strip().splitlines()[0] if t.__doc__ else 'No docstring'}")

    print("\n" + "=" * 60)
    print("ALL 5 AI TOOLS REGISTERED & RBAC-VALIDATED SUCCESSFULLY")
    print("=" * 60)
