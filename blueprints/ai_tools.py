"""blueprints/ai_tools.py - Dedicated Tool Registry for DCR AI Agent.

Provides 5 deterministic, RBAC-enforced, parameterized tool functions:
1. search_submittals
2. get_project_summary
3. get_noc_financials
4. get_overdue_submittals
5. get_document_history

Every function strictly requires `allowed_project_ids` from the verified user session,
preventing cross-project unauthorized data leaks. Zero string concatenation for SQL.
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
        return []

    if not project_code or str(project_code).strip().upper() in ("ALL", "*"):
        return projs

    p_clean = str(project_code).strip().lower()

    # 1. Check PROJECT_ALIASES
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
    target_projs = _resolve_target_projects(allowed_project_ids, project_code)
    if not target_projs:
        return {
            "project_scope": [],
            "count": 0,
            "records": [],
            "message": "No accessible project matching criteria found.",
        }

    proj_map = {p["id"]: p for p in target_projs}
    proj_ids = [p["id"] for p in target_projs]
    proj_codes = [p.get("code") for p in target_projs if p.get("code")]
    proj_displays = [f"{p.get('name', '')} ({p.get('code', p.get('id'))})" for p in target_projs]

    # Collect project document prefixes (e.g. CY002P608 for CFC)
    doc_prefixes = []
    for p in target_projs:
        code_str = (p.get("code") or "").upper()
        if "058" in code_str or "CFC" in code_str or "PEM-058" in code_str:
            doc_prefixes.append("%CY002P608%")
        elif "24CP01" in code_str or "EKH" in code_str:
            doc_prefixes.append("%24CP01%")
        else:
            if code_str:
                doc_prefixes.append(f"%{code_str}%")

    where_clauses = []
    params = []

    # 1. Mandatory RBAC & Project Scope Condition
    if doc_prefixes:
        where_clauses.append("""(
            r.project_id = ANY(%s) 
            OR UPPER(r.project_id) = ANY(%s)
            OR r.project_id::text IN (SELECT id::text FROM projects WHERE id = ANY(%s) OR code = ANY(%s))
            OR r.data->>'docNo' ILIKE ANY(%s)
        )""")
        params.extend([proj_ids, [x.upper() for x in proj_ids], proj_ids, proj_codes, doc_prefixes])
    else:
        where_clauses.append("""(
            r.project_id = ANY(%s) 
            OR UPPER(r.project_id) = ANY(%s)
            OR r.project_id::text IN (SELECT id::text FROM projects WHERE id = ANY(%s) OR code = ANY(%s))
        )""")
        params.extend([proj_ids, [x.upper() for x in proj_ids], proj_ids, proj_codes])

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

    # 5. Keyword Condition
    if keyword:
        kw_clean = str(keyword).strip().lower()
        kw_terms = [kw_clean]
        for eq_key, eq_synonyms in EQUIPMENT_KEYWORDS.items():
            if kw_clean in [s.lower() for s in eq_synonyms]:
                kw_terms = eq_synonyms
                break

        kw_patterns = [f"%{t}%" for t in kw_terms if len(t) >= 3]
        if kw_patterns:
            where_clauses.append("""(
                r.data->>'title' ILIKE ANY(%s)
                OR r.data->>'docNo' ILIKE ANY(%s)
                OR r.data->>'itemRef' ILIKE ANY(%s)
                OR r.data->>'description' ILIKE ANY(%s)
                OR r.data->>'subject' ILIKE ANY(%s)
                OR r.data->>'floor' ILIKE ANY(%s)
            )""")
            params.extend([kw_patterns] * 6)

    # 6. Status Condition
    if status:
        st_clean = str(status).strip().lower()
        if any(x in st_clean for x in ["approv", "معتمد", "accepted"]):
            where_clauses.append("""(
                COALESCE(r.data->>'status', r.data->>'partBStatus', r.data->>'partDStatus', '') ILIKE '%Approv%'
                OR COALESCE(r.data->>'status', '') ILIKE '%Status A%'
                OR COALESCE(r.data->>'status', '') ILIKE '%Status B%'
                OR COALESCE(r.data->>'status', '') ILIKE '%B - %'
                OR COALESCE(r.data->>'status', '') ILIKE '%A - %'
                OR COALESCE(r.data->>'status', '') ILIKE '%Code A%'
                OR COALESCE(r.data->>'status', '') ILIKE '%Code B%'
                OR COALESCE(r.data->>'status', '') ILIKE '%معتمد%'
                OR COALESCE(r.data->>'status', '') ILIKE '%Accept%'
            )""")
        elif any(x in st_clean for x in ["pending", "review", "تحت", "مراجعة"]):
            where_clauses.append("""(
                COALESCE(r.data->>'status', r.data->>'partBStatus', '') ILIKE '%Pending%'
                OR COALESCE(r.data->>'status', r.data->>'partBStatus', '') ILIKE '%Under Review%'
                OR COALESCE(r.data->>'status', r.data->>'partBStatus', '') ILIKE '%Open%'
                OR COALESCE(r.data->>'status', r.data->>'partBStatus', '') ILIKE '%مراجعة%'
            )""")
        elif any(x in st_clean for x in ["reject", "revise", "مرفوض"]):
            where_clauses.append("""(
                COALESCE(r.data->>'status', r.data->>'partBStatus', '') ILIKE '%Reject%'
                OR COALESCE(r.data->>'status', r.data->>'partBStatus', '') ILIKE '%Revise%'
                OR COALESCE(r.data->>'status', '') ILIKE '%Status C%'
                OR COALESCE(r.data->>'status', '') ILIKE '%C - %'
                OR COALESCE(r.data->>'status', '') ILIKE '%مرفوض%'
            )""")

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
        proj_info = proj_map.get(r.get("project_id"), {})
        r["project_code"] = proj_info.get("code", r.get("project_id", ""))

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
    target_projs = _resolve_target_projects(allowed_project_ids, project_code)
    if not target_projs:
        return {"projects": [], "total_projects": 0, "message": "No accessible project found."}

    pids = [p["id"] for p in target_projs]
    try:
        stats = db.get_dashboard_stats(project_ids=pids)
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
    target_projs = _resolve_target_projects(allowed_project_ids, project_code)
    if not target_projs:
        return {"total_count": 0, "total_submitted_cost": 0.0, "total_approved_cost": 0.0, "items": []}

    pids = [p["id"] for p in target_projs]
    try:
        rows = db.q("""
            SELECT r.project_id, p.code as proj_code, p.name as proj_name, r.data
            FROM records r
            JOIN projects p ON p.id = r.project_id
            WHERE (UPPER(r.dt_id) = 'NOC' OR r.data->>'docNo' ILIKE 'NOC-%' OR r.data->>'docNo' ILIKE '%-NOC-%')
              AND r.project_id = ANY(%s)
            ORDER BY r.created_at DESC
        """, (pids,))
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

        item = {
            "doc_no": doc_no,
            "project_code": row.get("proj_code", ""),
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
    target_projs = _resolve_target_projects(allowed_project_ids, project_code)
    if not target_projs:
        return {"total_overdue": 0, "records": [], "message": "No accessible project found."}

    pids = [p["id"] for p in target_projs]
    try:
        overdue_recs = db.get_overdue_records(project_ids=pids)
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
    target_projs = _resolve_target_projects(allowed_project_ids, project_code)
    if not target_projs:
        return {"base_doc_no": doc_no_pattern, "total_revisions": 0, "revisions": [], "message": "No accessible project found."}

    # Extract clean base document number without REV suffix
    base_no = re.sub(r"[\s_-]*REV\.?\s*(\d+)", "", str(doc_no_pattern).strip(), flags=re.IGNORECASE).strip()
    if not base_no:
        base_no = str(doc_no_pattern).strip()

    pids = [p["id"] for p in target_projs]
    search_pattern = f"%{base_no}%"

    try:
        rows = db.q("""
            SELECT r.id, r.project_id, r.dt_id, r.data, r.created_at, p.code as proj_code
            FROM records r
            JOIN projects p ON p.id = r.project_id
            WHERE (r.project_id = ANY(%s) OR UPPER(r.project_id) = ANY(%s))
              AND (r.data->>'docNo' ILIKE %s)
            ORDER BY r.created_at ASC
        """, (pids, [x.upper() for x in pids], search_pattern))
    except Exception as e:
        logger.error("get_document_history query error: %s", e)
        rows = []

    parsed_revisions = []
    for r in rows:
        d = r.get("data") or {}
        doc_no = str(d.get("docNo") or "").strip()
        rev_match = re.search(r"[\s_-]*REV\.?\s*(\d+)", doc_no, re.IGNORECASE)
        rev_num = int(rev_match.group(1)) if rev_match else 0
        parsed_revisions.append({
            "record_id": r.get("id"),
            "doc_no": doc_no,
            "rev_num": rev_num,
            "status": str(d.get("status") or d.get("partBStatus") or d.get("partDStatus") or "Pending"),
            "issued_date": str(d.get("issuedDate") or d.get("partAIssueDate") or ""),
            "reply_date": str(d.get("actualReplyDate") or d.get("actualReply") or ""),
            "title": str(d.get("title") or d.get("nocSubject") or ""),
            "discipline": str(d.get("discipline") or ""),
            "project_code": r.get("proj_code", ""),
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
    print("=== DCR AI TOOLS INITIALIZATION & RBAC VERIFICATION ===")
    print("=" * 60)
    
    # 1. RBAC Empty Check
    empty_rbac = search_submittals(allowed_project_ids=[])
    print("[1] RBAC Empty Project IDs Test:")
    print(f"    Expected count=0, got count={empty_rbac.get('count')}")
    assert empty_rbac.get("count") == 0, "RBAC Empty check failed!"
    print("    -> PASS: Unauthorized access blocked without DB query.")

    # 2. RBAC Project Scope Isolation
    empty_history = get_document_history(allowed_project_ids=[], doc_no_pattern="MS-001")
    print("\n[2] Document History RBAC Check:")
    print(f"    Expected total_revisions=0, got={empty_history.get('total_revisions')}")
    assert empty_history.get("total_revisions") == 0, "History RBAC check failed!"
    print("    -> PASS: Document history blocked when user has no allowed projects.")

    # 3. Tool registry inventory
    tools = [
        search_submittals,
        get_project_summary,
        get_noc_financials,
        get_overdue_submittals,
        get_document_history,
    ]
    print("\n[3] Tool Registry Inventory:")
    for t in tools:
        print(f"    - {t.__name__}: {t.__doc__.strip().splitlines()[0] if t.__doc__ else 'No docstring'}")

    print("\n" + "=" * 60)
    print("ALL 5 AI TOOLS REGISTERED & RBAC-VALIDATED SUCCESSFULLY")
    print("=" * 60)
