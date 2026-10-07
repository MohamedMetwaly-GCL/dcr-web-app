"""blueprints/ai.py - DCR Native AI Assistant API (Gemini Flash).

Provides read-only AI Assistant capabilities using Google Gemini Flash API.
Strictly RBAC-aware: users can only query records and summaries from projects
they have explicit permission to view.
"""
import logging
import os
import re
import time
from decimal import Decimal, InvalidOperation
from flask import Blueprint, jsonify, request
from google import genai
from google.genai import types

import db
from auth import current_user, can_view_project, get_allowed_project_ids

ai_bp = Blueprint("ai", __name__)
logger = logging.getLogger(__name__)

SYSTEM_INSTRUCTION = (
    "You are the DCR Engineering AI Assistant for Gas Chill contracting projects. "
    "In the project document registers, abbreviations are defined as follows: "
    "- 'MS': Material Submittal (اعتماد مواد / عينات / موردين) or Method Statement depending on register. "
    "- 'SD': Shop Drawing (مخططات تنفيذية / شوب درونج). "
    "- 'MIR': Material Inspection Request (طلب فحص وتوريد مواد ومهمات). "
    "- 'RFI': Request for Information (استفسارات وتوضيحات فنية). "
    "- 'IR' / 'WIR': Inspection Request / Work Inspection Request (طلبات استلام أعمال). "
    "- 'NOC': Notice of Change (أوامر التغيير / مطالبات مالية). "
    "- 'PR': Purchase Request / Procurement (طلبات الشراء والتوريد). "
    "You provide accurate, concise, factual answers about document statuses, "
    "overdue submittals, and cost approvals using ONLY the provided project register context. "
    "Never modify or fabricate data."
)


KEYWORD_MAP = {
    "محابس": ["valve", "valves", "valv", "butterfly", "check valve", "fittings"],
    "محبس": ["valve", "valves", "valv"],
    "فالف": ["valve", "valves", "valv"],
    "فالفات": ["valve", "valves", "valv"],
    "طلمبات": ["pump", "pumps"],
    "طلمبة": ["pump", "pumps"],
    "مضخات": ["pump", "pumps"],
    "مضخة": ["pump", "pumps"],
    "بمب": ["pump", "pumps"],
    "مواسير": ["pipe", "pipes", "piping"],
    "ماسورة": ["pipe", "pipes", "piping"],
    "أنابيب": ["pipe", "pipes", "piping"],
    "انابيب": ["pipe", "pipes", "piping"],
    "عزل": ["insulation", "insulated"],
    "عوازل": ["insulation", "insulated"],
    "مخططات": ["drawing", "drawings", "dwg"],
    "مخطط": ["drawing", "drawings", "dwg"],
    "رسومات": ["drawing", "drawings", "dwg"],
    "رسم": ["drawing", "drawings", "dwg"],
    "شوب درونج": ["shop drawing", "drawing", "dwg"],
    "شوب دروينج": ["shop drawing", "drawing", "dwg"],
    "شوبدروينج": ["shop drawing", "drawing", "dwg"],
    "تكييف": ["chiller", "cooling", "hvac"],
    "تكييفات": ["chiller", "cooling", "hvac"],
    "مبردات": ["chiller", "chillers", "cooling"],
    "مبرد": ["chiller", "cooling"],
    "تشيلر": ["chiller", "chillers"],
    "تشيلرات": ["chiller", "chillers"],
    "فلاتر": ["filter", "filters"],
    "فلتر": ["filter", "filters"],
    "تهوية": ["fan", "fans", "ventilation"],
    "مراوح": ["fan", "fans", "ventilation"],
    "مروحة": ["fan", "fans", "ventilation"],
    "لوحات": ["panel", "panels"],
    "لوحة": ["panel", "panels"],
    "كابلات": ["cable", "cables"],
    "كيبلات": ["cable", "cables"],
    "كابل": ["cable", "cables"],
    "محولات": ["transformer", "transformers"],
    "محول": ["transformer", "transformers"],
    "مولدات": ["generator", "generators"],
    "مولد": ["generator", "generators"],
    "إنذار": ["alarm", "alarms"],
    "انذار": ["alarm", "alarms"],
    "حريق": ["fire"],
    "إطفاء": ["fire", "fighting"],
    "اطفاء": ["fire", "fighting"],
    "دكت": ["duct", "ducts"],
    "صاج": ["duct", "ducts"],
}

ENGLISH_KEYWORD_MAP = {
    "valve": ["valve", "valves", "valv", "butterfly", "check valve", "fittings"],
    "valves": ["valve", "valves", "valv", "butterfly", "check valve", "fittings"],
    "pump": ["pump", "pumps"],
    "pumps": ["pump", "pumps"],
    "pipe": ["pipe", "pipes", "piping"],
    "pipes": ["pipe", "pipes", "piping"],
    "piping": ["pipe", "pipes", "piping"],
    "chiller": ["chiller", "chillers", "cooling", "hvac"],
    "chillers": ["chiller", "chillers", "cooling", "hvac"],
    "filter": ["filter", "filters"],
    "filters": ["filter", "filters"],
    "fan": ["fan", "fans", "ventilation"],
    "fans": ["fan", "fans", "ventilation"],
    "duct": ["duct", "ducts"],
    "ducts": ["duct", "ducts"],
    "panel": ["panel", "panels"],
    "panels": ["panel", "panels"],
    "cable": ["cable", "cables"],
    "cables": ["cable", "cables"],
    "insulation": ["insulation", "insulated"],
    "drawing": ["drawing", "drawings", "dwg"],
    "drawings": ["drawing", "drawings", "dwg"],
}

KNOWN_DOC_TYPES = ["MS", "SD", "MIR", "RFI", "IR", "NOC", "PR", "NCR", "ITP", "PQ", "WIR", "MAR", "MOM"]


def _safe_float(val):
    if val is None or val == "":
        return 0.0
    try:
        clean = re.sub(r"[^\d.-]", "", str(val).strip())
        return float(Decimal(clean)) if clean else 0.0
    except (InvalidOperation, ValueError):
        return 0.0


def _build_ai_context(target_pids, user_prompt="", active_tab=None, history=None):
    """Generates concise, factual register context for the active project scope."""
    if not target_pids:
        return "No accessible projects found for the current user."

    context_lines = []
    p_lower = user_prompt.lower()

    # 1. Projects Metadata
    try:
        projs = db.q("SELECT id, name, code FROM projects WHERE id = ANY(%s) ORDER BY code", (target_pids,))
    except Exception as e:
        logger.warning("Error fetching projects for AI context: %s", e)
        projs = []

    proj_map = {p["id"]: p for p in projs}
    context_lines.append("### 1. ACTIVE PROJECT(S) SCOPE:")
    for p in projs:
        context_lines.append(f"- Project Code: {p['code']} | Name: {p['name']} (ID: {p['id']})")
    context_lines.append("")

    # Project Scope Detection from Prompt (e.g. CFC, Assiut, Suez, PEM-058, etc.)
    scoped_search_pids = list(target_pids)
    matched_pids = []
    for p in projs:
        p_code = (p.get("code") or "").lower()
        p_name = (p.get("name") or "").lower()
        if p_code and (p_code in p_lower or p_code.replace("pem-", "") in p_lower):
            matched_pids.append(p["id"])
            continue
        name_words = [w for w in re.findall(r"[a-zA-Z0-9\u0600-\u06FF]{3,}", p_name) if len(w) >= 3]
        if any(w in p_lower for w in name_words):
            matched_pids.append(p["id"])
    if matched_pids:
        scoped_search_pids = matched_pids

    # Detect Targeted Document Types (MS, SD, MIR, RFI, IR, NOC, PR, etc.)
    detected_doc_types = []
    for dt in KNOWN_DOC_TYPES:
        if re.search(r"\b" + re.escape(dt) + r"\b", user_prompt, re.IGNORECASE):
            detected_doc_types.append(dt.upper())

    if any(w in p_lower for w in ["شوب درونج", "شوب دروينج", "شوبدروينج", "shop drawing", "shop drawings", "drawings", "مخططات"]):
        if "SD" not in detected_doc_types:
            detected_doc_types.append("SD")
    if any(w in p_lower for w in ["ماتريال", "اعتماد مواد", "اعتمادات مواد", "material submittal", "material submittals", "عينة", "عينات", "مورد"]):
        if "MS" not in detected_doc_types:
            detected_doc_types.append("MS")
    if any(w in p_lower for w in ["فحص مواد", "توريد مواد", "material inspection"]):
        if "MIR" not in detected_doc_types:
            detected_doc_types.append("MIR")

    if active_tab and active_tab.upper() in KNOWN_DOC_TYPES and not detected_doc_types:
        detected_doc_types.append(active_tab.upper())

    # Bilingual Keyword Extraction (Arabic to English mapping + raw keywords)
    search_terms = []
    for ar_kw, en_terms in KEYWORD_MAP.items():
        if ar_kw in p_lower:
            terms_list = en_terms if isinstance(en_terms, list) else [en_terms]
            for term in terms_list:
                if term not in search_terms:
                    search_terms.append(term)

    for en_word, en_terms in ENGLISH_KEYWORD_MAP.items():
        if re.search(r"\b" + re.escape(en_word) + r"\b", p_lower):
            for term in en_terms:
                if term not in search_terms:
                    search_terms.append(term)

    # General English/Alphanumeric tokens
    stop_words = {
        "and", "the", "for", "with", "all", "what", "show", "list", "give", "from",
        "find", "submittal", "submittals", "document", "documents", "project", "projects",
        "status", "approved", "noted", "date", "please", "cfc", "pem", "any", "our",
        "help", "query", "record", "records", "material", "materials", "method", "statement",
        "cost", "costs", "price", "change", "financial", "total", "summary", "kpi"
    }
    raw_tokens = re.findall(r"[A-Za-z0-9_-]{3,}", user_prompt)
    for tok in raw_tokens:
        tok_low = tok.lower()
        if (
            tok.upper() not in KNOWN_DOC_TYPES
            and tok_low not in stop_words
            and not tok.isdigit()
            and tok_low not in [s.lower() for s in search_terms]
        ):
            search_terms.append(tok)
            if len(search_terms) >= 12:
                break

    # ── Conversational Context & Clarification Recovery from Chat History ──
    has_inherited_equipment = False
    if history:
        # 1. Recover Project Scope if missing from current prompt
        if not matched_pids:
            for msg in reversed(history):
                if msg.get("role") == "user":
                    prev_text = (msg.get("text") or "").lower()
                    for p in projs:
                        p_code = (p.get("code") or "").lower()
                        p_name = (p.get("name") or "").lower()
                        if p_code and (p_code in prev_text or p_code.replace("pem-", "") in prev_text):
                            matched_pids.append(p["id"])
                            break
                        name_words = [w for w in re.findall(r"[a-zA-Z0-9\u0600-\u06FF]{3,}", p_name) if len(w) >= 3]
                        if any(w in prev_text for w in name_words):
                            matched_pids.append(p["id"])
                            break
                    if matched_pids:
                        scoped_search_pids = matched_pids
                        break

        # 2. Recover Equipment Keywords if current prompt is a clarification or lacks equipment keywords
        if not search_terms:
            for msg in reversed(history):
                if msg.get("role") == "user":
                    prev_text = (msg.get("text") or "").lower()
                    for ar_kw, en_terms in KEYWORD_MAP.items():
                        if ar_kw in prev_text:
                            terms_list = en_terms if isinstance(en_terms, list) else [en_terms]
                            for term in terms_list:
                                if term not in search_terms:
                                    search_terms.append(term)
                                    has_inherited_equipment = True
                    for en_word, en_terms in ENGLISH_KEYWORD_MAP.items():
                        if re.search(r"\b" + re.escape(en_word) + r"\b", prev_text):
                            for term in en_terms:
                                if term not in search_terms:
                                    search_terms.append(term)
                                    has_inherited_equipment = True
                    if search_terms:
                        break

        # 3. Recover Doc Types if current prompt has no doc types
        if not detected_doc_types:
            for msg in reversed(history):
                if msg.get("role") == "user":
                    prev_text = (msg.get("text") or "").lower()
                    for dt in KNOWN_DOC_TYPES:
                        if re.search(r"\b" + re.escape(dt) + r"\b", msg.get("text") or "", re.IGNORECASE):
                            if dt.upper() not in detected_doc_types:
                                detected_doc_types.append(dt.upper())
                    if any(w in prev_text for w in ["شوب درونج", "شوب دروينج", "شوبدروينج", "shop drawing", "shop drawings", "drawings", "مخططات"]):
                        if "SD" not in detected_doc_types:
                            detected_doc_types.append("SD")
                    if any(w in prev_text for w in ["ماتريال", "اعتماد مواد", "اعتمادات مواد", "material submittal", "material submittals"]):
                        if "MS" not in detected_doc_types:
                            detected_doc_types.append("MS")
                    if detected_doc_types:
                        break

    # Determine query intent to slim down context payload
    has_submittal_keyword = (
        any(ar_kw in p_lower for ar_kw in KEYWORD_MAP)
        or any(re.search(r"\b" + re.escape(en_k) + r"\b", p_lower) for en_k in ENGLISH_KEYWORD_MAP)
        or has_inherited_equipment
    )
    has_submittal_doctype = any(dt in detected_doc_types for dt in ["MS", "SD", "MIR", "RFI", "IR", "PR", "NCR", "ITP", "PQ", "WIR", "MAR"])

    is_specific_submittal_query = has_submittal_keyword or has_submittal_doctype
    is_noc_requested = "NOC" in detected_doc_types or any(w in p_lower for w in ["noc", "change", "تغيير", "تكلفة", "cost", "financial", "مالي", "أمر تغيير"]) or (active_tab and active_tab.upper() == "NOC")
    is_overdue_requested = any(w in p_lower for w in ["overdue", "متأخر", "متأخرات", "delay", "تأخير"])

    # 2. Project Dashboard KPIs (Statuses, Totals, Overdues)
    include_kpis = not is_specific_submittal_query or any(w in p_lower for w in ["kpi", "موقف", "تقرير", "إجمالي", "total", "summary", "لخص"])
    if include_kpis:
        try:
            stats = db.get_dashboard_stats(project_ids=target_pids)
        except Exception as e:
            logger.warning("Error getting dashboard stats for AI context: %s", e)
            stats = []

        context_lines.append("### 2. REGISTER STATUS SUMMARY & KPIS:")
        if stats:
            for s in stats:
                p_code = s.get("code") or s.get("name") or "PRJ"
                p_name = s.get("name") or ""
                total = s.get("total", 0)
                appr = s.get("approved", 0)
                pend = s.get("pending", 0)
                rej = s.get("rejected", 0)
                over = s.get("overdue", 0)
                pct = s.get("pct", 0)
                context_lines.append(
                    f"- Project [{p_code}] {p_name}: Total Documents={total}, "
                    f"Approved={appr} ({pct}%), Under Review/Pending={pend}, "
                    f"Overdue={over}, Rejected/Revise={rej}"
                )
        else:
            context_lines.append("- No register statistics available.")
        context_lines.append("")

    # 3. Notice of Change (NOC) Cost & Approval Totals (only when requested or general overview)
    if is_noc_requested or not is_specific_submittal_query:
        try:
            noc_rows = db.q("""
                SELECT r.project_id, p.code as proj_code, p.name as proj_name, r.data
                FROM records r
                JOIN projects p ON p.id = r.project_id
                WHERE UPPER(r.dt_id) = 'NOC' AND r.project_id = ANY(%s)
                ORDER BY r.project_id, r.created_at DESC
            """, (target_pids,))
        except Exception as e:
            logger.warning("Error getting NOC rows for AI context: %s", e)
            noc_rows = []

        context_lines.append("### 3. NOTICE OF CHANGE (NOC) FINANCIAL & APPROVAL SUMMARY:")
        if noc_rows:
            total_nocs = len(noc_rows)
            total_submitted_cost = 0.0
            total_approved_cost = 0.0
            approved_nocs = []
            pending_nocs = []
            rejected_nocs = []

            for row in noc_rows:
                d = row.get("data") or {}
                doc_no = d.get("docNo") or d.get("nocNo") or "NOC"
                subject = d.get("title") or d.get("nocSubject") or d.get("nocDescription") or "No Subject"
                sub_cost = _safe_float(d.get("submittedCost"))
                app_cost = _safe_float(d.get("finalApprovedCost"))
                status = str(d.get("partDStatus") or d.get("partBStatus") or d.get("status") or "").strip()

                total_submitted_cost += sub_cost
                total_approved_cost += app_cost

                st_lower = status.lower()
                noc_entry = (
                    f"{doc_no} (Proj: {row.get('proj_code')}): '{subject}' | "
                    f"Submitted: {sub_cost:,.2f} EGP | Approved: {app_cost:,.2f} EGP | Status: {status or 'Pending'}"
                )

                if "approv" in st_lower or "accept" in st_lower or "part c" in st_lower:
                    approved_nocs.append(noc_entry)
                elif "reject" in st_lower or "cancel" in st_lower:
                    rejected_nocs.append(noc_entry)
                else:
                    pending_nocs.append(noc_entry)

            context_lines.append(f"- Total NOCs: {total_nocs}")
            context_lines.append(f"- Total Submitted Cost: {total_submitted_cost:,.2f} EGP")
            context_lines.append(f"- Total Final Approved Cost: {total_approved_cost:,.2f} EGP")
            context_lines.append(f"- Approved/Accepted NOCs Count: {len(approved_nocs)}")
            context_lines.append(f"- Pending/Under Review NOCs Count: {len(pending_nocs)}")
            context_lines.append(f"- Rejected/Cancelled NOCs Count: {len(rejected_nocs)}")
            context_lines.append("")
            context_lines.append("Key NOC Details (Top items):")
            for item in (approved_nocs[:15] + pending_nocs[:10]):
                context_lines.append(f"  * {item}")
        else:
            context_lines.append("- No NOC (Notice of Change) records found in this scope.")
        context_lines.append("")

    # 4. Overdue Submittals (only when requested or general overview)
    if is_overdue_requested or not is_specific_submittal_query:
        try:
            overdue_recs = db.get_overdue_records(project_ids=target_pids)
        except Exception as e:
            logger.warning("Error getting overdue records for AI context: %s", e)
            overdue_recs = []

        context_lines.append("### 4. OVERDUE SUBMITTALS (ACTION REQUIRED):")
        if overdue_recs:
            context_lines.append(f"- Total Overdue Submittals: {len(overdue_recs)}")
            context_lines.append("Top Overdue Submittals (Sorted by longest delay):")
            for r in overdue_recs[:20]:
                p_code = proj_map.get(r.get("project_id"), {}).get("code", r.get("project_id", ""))
                doc_no = r.get("docNo", "—")
                dt = r.get("dt_code", "DOC")
                title = r.get("title", "")
                days = r.get("days_overdue", 0)
                st = r.get("status", "Pending")
                issued = r.get("issuedDate", "")
                context_lines.append(
                    f"  * [{p_code}] {doc_no} ({dt}): '{title}' | Overdue by: {days} days | Issued: {issued} | Status: {st}"
                )
        else:
            context_lines.append("- Zero (0) overdue submittals in this scope. All documents on track.")
        context_lines.append("")

    # 5. Targeted Specific Search & Multi-condition Strict Semantic Matching
    approval_indicators = [
        "معتمد", "معتمدة", "معتمدين", "موافقة", "موافق", "مقبول",
        "approved", "approval", "status a", "status b", "code a", "code b"
    ]
    is_approved_query = any(ind in p_lower for ind in approval_indicators)
    if not is_approved_query and history:
        for msg in reversed(history):
            if msg.get("role") == "user":
                prev_text = (msg.get("text") or "").lower()
                if any(ind in prev_text for ind in approval_indicators):
                    is_approved_query = True
                    break

    if search_terms or detected_doc_types:
        where_clauses = ["r.project_id = ANY(%s)"]
        params = [scoped_search_pids]

        # Enforce doc type filter if detected
        if detected_doc_types:
            where_clauses.append("UPPER(r.dt_id) = ANY(%s)")
            params.append([dt.upper() for dt in detected_doc_types])

        # Enforce equipment/keyword matching with ANY(keywords_array)
        if search_terms:
            kw_patterns = [f"%{term}%" for term in search_terms[:12]]
            where_clauses.append("(r.data->>'title' ILIKE ANY(%s) OR r.data->>'docNo' ILIKE ANY(%s))")
            params.extend([kw_patterns, kw_patterns])

        order_clauses = []
        if is_approved_query:
            order_clauses.append("""
                CASE 
                    WHEN (
                        r.data->>'status' ILIKE '%Approv%' 
                        OR r.data->>'status' ILIKE 'Status A%' 
                        OR r.data->>'status' ILIKE 'Status B%'
                        OR r.data->>'status' ILIKE '%Code A%'
                        OR r.data->>'status' ILIKE '%Code B%'
                        OR r.data->>'status' ILIKE '%معتمد%'
                        OR r.data->>'partBStatus' ILIKE '%Approv%'
                        OR r.data->>'partDStatus' ILIKE '%Approv%'
                    ) THEN 0 
                    ELSE 1 
                END ASC
            """)
        order_clauses.append("r.created_at DESC NULLS LAST")

        where_sql = " AND ".join(where_clauses)
        order_sql = ", ".join(order_clauses)

        search_sql = f"""
            SELECT 
                r.id,
                r.project_id,
                COALESCE(r.dt_id, '') AS doc_type,
                COALESCE(r.data->>'docNo', r.data->>'nocNo', r.data->>'letterRef', '—') AS doc_no,
                COALESCE(r.data->>'title', r.data->>'nocSubject', r.data->>'subject', '') AS title,
                COALESCE(r.data->>'status', r.data->>'partBStatus', r.data->>'partDStatus', '') AS status,
                COALESCE(r.data->>'issuedDate', r.data->>'partAIssueDate', '') AS issued_date,
                COALESCE(r.data->>'actualReplyDate', r.data->>'actualReply', r.data->>'partDReturnDate', '') AS actual_reply
            FROM records r
            WHERE {where_sql}
            ORDER BY {order_sql}
            LIMIT 10;
        """

        try:
            matched_records = db.q(search_sql, params)
        except Exception as e_search:
            logger.warning("Error searching records in AI context: %s", e_search)
            matched_records = []

        if matched_records:
            context_lines.append("### 5. MATCHING SUBMITTALS IN REGISTER:")
            context_lines.append(f"Found {len(matched_records)} matching submittals based on query keywords and filters:")
            for rec in matched_records:
                doc_no = rec.get("doc_no") or "—"
                title = rec.get("title") or "No Title"
                st = rec.get("status") or "Pending"
                dt = rec.get("doc_type") or "DOC"
                p_code = proj_map.get(rec.get("project_id"), {}).get("code", rec.get("project_id", ""))
                date_val = rec.get("actual_reply") or rec.get("issued_date") or "—"
                context_lines.append(
                    f"- DocNo: {doc_no} | Title: {title} | DocType: {dt} | Status: {st} | Date: {date_val} | Project: {p_code}"
                )
            context_lines.append("")
        else:
            context_lines.append("### 5. MATCHING SUBMITTALS IN REGISTER:")
            context_lines.append("- No matching submittals found for the requested keywords/filters in the scoped project(s).")
            context_lines.append("")

    return "\n".join(context_lines)


_working_gemini_model = None


def _call_gemini_api(prompt, context_text, custom_instruction=None, history=None):
    """Calls Gemini Flash API with standard client, active models cascade, 503 retry, and history context."""
    global _working_gemini_model

    api_key = os.environ.get("GEMINI_API_KEY", "").strip()
    if not api_key:
        return None, "NO_API_KEY"

    history_lines = []
    if history:
        for msg in history[-6:]:
            r_label = "User" if msg.get("role") == "user" else "Assistant"
            m_text = (msg.get("text") or "").strip()
            if m_text:
                history_lines.append(f"{r_label}: {m_text}")
    history_str = "\n".join(history_lines) if history_lines else "None"

    full_contents = (
        f"### CONTEXT REGISTER DATA (READ-ONLY):\n{context_text}\n\n"
        f"### RECENT CONVERSATION HISTORY:\n{history_str}\n\n"
        f"### USER CURRENT QUESTION / CLARIFICATION:\n{prompt}\n\n"
        "Provide a structured, helpful, professional engineering response in Markdown format. "
        "IMPORTANT: In this engineering system, 'MS' refers to Material Submittal (اعتماد مواد / عينات / موردين). "
        "Directly answer based on the matching records in the register context. "
        "If the user asks in Arabic, answer in clear, professional Arabic (باللغة العربية الهندسية). "
        "If the user asks in English, answer in English. "
        "Highlight document numbers, statuses, titles, and approval codes clearly. "
        "Use bullet points, bold key figures, and tables where appropriate."
    )

    models_to_try = []
    if _working_gemini_model:
        models_to_try.append(_working_gemini_model)
    for m in ["gemini-2.5-flash", "gemini-2.5-pro"]:
        if m not in models_to_try:
            models_to_try.append(m)

    system_inst = custom_instruction or (
        "You are the DCR Engineering AI Assistant for Gas Chill contracting projects. "
        "In this project register, document abbreviations are: "
        "- MS = Material Submittal (اعتماد مواد / عينات / موردين) or Method Statement depending on register context. "
        "- SD = Shop Drawing (شوب درونج / مخططات تنفيذية). "
        "- MIR = Material Inspection Request (طلب فحص وتوريد مواد). "
        "- RFI = Request for Information (استفسارات وتوضيحات فنية). "
        "- NOC = Notice of Change (أوامر التغيير / مطالبات مالية). "
        "Directly list matching documents from the provided context in a compact Markdown table or bullet points. "
        "Never invent or assume records not present in the context. "
        "Avoid long introductions or filler text."
    )

    last_err = None
    try:
        client = genai.Client(api_key=api_key)
        for model_name in models_to_try:
            clean_model = model_name.replace("models/", "").strip()
            try:
                resp = client.models.generate_content(
                    model=clean_model,
                    contents=full_contents,
                    config=types.GenerateContentConfig(
                        system_instruction=system_inst,
                        temperature=0.2,
                        max_output_tokens=1000,
                    ),
                )
                if resp and resp.text:
                    _working_gemini_model = clean_model
                    return resp.text, None
            except Exception as e_model:
                last_err = e_model
                if _working_gemini_model == clean_model:
                    _working_gemini_model = None
                logger.warning("Gemini model %s failed: %s", clean_model, e_model)
                continue
    except Exception as e_client:
        last_err = e_client
        logger.error("Gemini client initialization error: %s", e_client)

    logger.error("[AI Assistant Error] Failed to generate: %s", last_err)
    return None, f"Gemini API Error: {str(last_err)}"


def _format_fallback_response_from_context(prompt, context_text):
    """Generates a structured, factual response directly from the extracted SQL context

    when Gemini API encounters high demand (503) or is temporarily unavailable.
    """
    if not context_text or not context_text.strip():
        return None

    # 1. Matching submittals section
    if "### 5. MATCHING SUBMITTALS IN REGISTER:" in context_text:
        recs = []
        for line in context_text.splitlines():
            line_str = line.strip()
            if line_str.startswith("- DocNo:"):
                parts = {}
                for segment in line_str[2:].split(" | "):
                    if ":" in segment:
                        k, v = segment.split(":", 1)
                        parts[k.strip()] = v.strip()
                if parts:
                    recs.append(parts)

        if recs:
            rows = []
            for r in recs:
                st = r.get("Status", "—")
                st_low = st.lower()
                st_badge = st
                if any(x in st_low for x in ["comment", "status b", "code b", "معتمد بملاحظات"]):
                    st_badge = f"🟡 **{st}**"
                elif any(x in st_low for x in ["reject", "revise", "code c", "status c", "مرفوض"]):
                    st_badge = f"❌ **{st}**"
                elif any(x in st_low for x in ["approv", "status a", "code a", "معتمد"]):
                    st_badge = f"✅ **{st}**"

                rows.append(
                    f"| **{r.get('DocNo', '—')}** | {r.get('Title', '—')} | {r.get('DocType', '—')} | {st_badge} | {r.get('Date', '—')} | {r.get('Project', '—')} |"
                )

            table_md = "\n".join(rows)
            return (
                "> 💡 **ملاحظة تشغيلية:** نظراً لضغط مؤقت على خوادم المعالجة التوليدية من Google (503 High Demand)، "
                "تم استخراج واعتماد النتائج التالية **مباشرة وفورياً من سجلات المشروع**:\n\n"
                f"### 📋 السجلات المطابقة لبحثك ({len(recs)} وثيقة):\n\n"
                "| رقم الوثيقة (Doc No) | الوصف / العنوان | النوع | الحالة (Status) | التاريخ | المشروع |\n"
                "| :--- | :--- | :--- | :--- | :--- | :--- |\n"
                f"{table_md}\n\n"
                "📌 *البيانات أعلاه مستخرجة ومطابقة 100% لسجلات النظام الرسمية.*"
            )
        elif "No matching submittals found" in context_text:
            return (
                "> 💡 **ملاحظة تشغيلية:** تم البحث مباشرة في سجلات النظام في قاعدة البيانات:\n\n"
                "⚠️ لم يتم العثور على وثائق مطابقة لكلمات البحث المحددة في نطاق المشروع. "
                "يرجى مراجعة الكلمات الدلالية أو استعراض جدول السجل المباشر."
            )

    # 2. NOC summary section
    p_low = prompt.lower()
    if "### 3. NOTICE OF CHANGE" in context_text and any(x in p_low for x in ["noc", "تغيير", "أوامر"]):
        lines = []
        for line in context_text.splitlines():
            if line.startswith("- Total") or line.startswith("  * "):
                lines.append(line)
        if lines:
            return (
                "> 💡 **ملاحظة تشغيلية:** تم استخراج ملخص أوامر التغيير (NOC) مباشرة من سجلات المشروع:\n\n"
                "### 💰 ملخص أوامر التغيير (Notice of Change):\n\n"
                + "\n".join(lines)
            )

    # 3. Overdue submittals section
    if "### 4. OVERDUE SUBMITTALS" in context_text and any(x in p_low for x in ["overdue", "متأخر", "delay"]):
        lines = [l for l in context_text.splitlines() if l.startswith("  * ") or l.startswith("- Total Overdue")]
        if lines:
            return (
                "> 💡 **ملاحظة تشغيلية:** تم استخراج الوثائق المتأخرة مباشرة من سجلات النظام:\n\n"
                "### ⏰ الوثائق المتأخرة (Overdue Submittals):\n\n"
                + "\n".join(lines)
            )

    # 4. General register summary
    if "### 2. REGISTER SUMMARY BY PROJECT:" in context_text:
        stat_lines = [l for l in context_text.splitlines() if l.startswith("- Project [")]
        if stat_lines:
            return (
                "> 💡 **ملاحظة تشغيلية:** تم استخراج إحصائيات السجلات مباشرة من النظام:\n\n"
                "### 📊 ملخص السجلات والمستندات:\n\n"
                + "\n".join(stat_lines)
            )

    return None


@ai_bp.route("/query", methods=["POST"])
@ai_bp.route("/api/ai/query", methods=["POST"])
def api_ai_query():
    """Main AI Assistant Query Endpoint."""
    try:
        u = current_user()
        if not u:
            return jsonify(error="LOGIN_REQUIRED"), 403

        api_key = os.environ.get("GEMINI_API_KEY", "").strip()
        if not api_key:
            return jsonify(error="AI Assistant is not configured on this instance."), 503

        data = request.get_json(silent=True) or {}
        prompt = str(data.get("prompt") or "").strip()
        if not prompt:
            return jsonify(error="Prompt is required"), 400

        project_id = data.get("project_id")
        tab = data.get("tab")
        history = data.get("history") or []

        # RBAC Enforcement
        target_pids = []
        if project_id and str(project_id).strip().upper() not in ("ALL", ""):
            pid = str(project_id).strip()
            if not can_view_project(pid, u):
                return jsonify(error="Forbidden"), 403
            target_pids = [pid]
        else:
            target_pids = get_allowed_project_ids(u)

        if not target_pids:
            return jsonify(reply="You do not have access to any projects in the system."), 200

        # Short-circuit for simple greetings/generic conversation
        p_clean = re.sub(r"[^\w\s]", "", prompt).strip().lower()
        greeting_words = [
            "ازيك", "ازىك", "عامل ايه", "مرحبا", "أهلا", "اهلا", 
            "سلام", "السلام عليكم", "صباح الخير", "مساء الخير", "hello", "hi", "hey", "هاي"
        ]
        is_greeting = any(g == p_clean or p_clean.startswith(g) for g in greeting_words) and len(prompt.split()) <= 4

        if is_greeting:
            reply = (
                "أهلاً بك يا باشمهندس! 👋 أنا مساعدك الذكي لنظام مراقبة وثائق ومشاريع جازشيل (DCR).\n\n"
                "أنا جاهز لمساعدتك فوراً في:\n"
                "- 📋 استعراض واعتمادات سجلات الـ (MS) والشوب درونج (SD).\n"
                "- ⏰ متابعة الوثائق المتأخرة (Overdue Submittals).\n"
                "- 💰 مراجعة أوامر التغيير والمطالبات (NOC).\n\n"
                "تفضل بسؤالي عن أي مشروع أو معدة أو مستند!"
            )
            return jsonify(reply=reply), 200

        context_text = _build_ai_context(target_pids, user_prompt=prompt, active_tab=tab, history=history)
        reply, err = _call_gemini_api(prompt, context_text, history=history)

        if err == "NO_API_KEY":
            return jsonify(error="AI Assistant is not configured on this instance."), 503

        if not reply:
            logger.warning(
                "[AI Assistant Gemini Unavailable] Falling back to direct context formatting. Gemini error: %s",
                err,
            )
            fallback_reply = _format_fallback_response_from_context(prompt, context_text)
            if fallback_reply:
                return jsonify(reply=fallback_reply), 200

            # If even direct context formatting didn't produce an answer, return a clear engineering message
            return jsonify(
                error=(
                    "خوادم المعالجة التوليدية من Google تواجه ضغطاً مؤقتاً (503 High Demand). "
                    "يرجى إعادة المحاولة بعد بضع ثوانٍ."
                )
            ), 503

        return jsonify(reply=reply), 200

    except Exception as e:
        logger.exception("AI assistant query exception: %s", e)
        return jsonify(error=f"Internal Server Error: {str(e)}"), 500
