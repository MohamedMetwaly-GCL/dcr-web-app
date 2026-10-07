"""blueprints/ai.py - DCR Native AI Assistant API (Gemini Flash).

Provides read-only AI Assistant capabilities using Google Gemini Flash API.
Strictly RBAC-aware: users can only query records and summaries from projects
they have explicit permission to view.
"""
import logging
import os
import re
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
    "You provide accurate, concise, factual answers about document statuses, "
    "overdue submittals, and cost approvals using ONLY the provided project register context. "
    "Never modify or fabricate data."
)


KEYWORD_MAP = {
    "محابس": "valve",
    "محبس": "valve",
    "فالف": "valve",
    "فالفات": "valve",
    "طلمبات": "pump",
    "طلمبة": "pump",
    "مضخات": "pump",
    "مضخة": "pump",
    "بمب": "pump",
    "مواسير": "pipe",
    "ماسورة": "pipe",
    "أنابيب": "pipe",
    "انابيب": "pipe",
    "عزل": "insulation",
    "عوازل": "insulation",
    "مخططات": "drawing",
    "مخطط": "drawing",
    "رسومات": "drawing",
    "رسم": "drawing",
    "شوب درونج": "shop drawing",
    "شوب دروينج": "shop drawing",
    "شوبدروينج": "shop drawing",
    "تكييف": "chiller",
    "تكييفات": "chiller",
    "مبردات": "chiller",
    "مبرد": "chiller",
    "تشيلر": "chiller",
    "تشيلرات": "chiller",
    "فلاتر": "filter",
    "فلتر": "filter",
    "تهوية": "fan",
    "مراوح": "fan",
    "مروحة": "fan",
    "لوحات": "panel",
    "لوحة": "panel",
    "كابلات": "cable",
    "كيبلات": "cable",
    "كابل": "cable",
    "محولات": "transformer",
    "محول": "transformer",
    "مولدات": "generator",
    "مولد": "generator",
    "إنذار": "alarm",
    "انذار": "alarm",
    "حريق": "fire",
    "إطفاء": "fire",
    "اطفاء": "fire",
    "دكت": "duct",
    "صاج": "duct",
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


def _build_ai_context(target_pids, user_prompt="", active_tab=None):
    """Generates concise, factual register context for the active project scope."""
    if not target_pids:
        return "No accessible projects found for the current user."

    context_lines = []

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

    # 2. Project Dashboard KPIs (Statuses, Totals, Overdues)
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

    # 3. Notice of Change (NOC) Cost & Approval Totals
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
        # Provide sample of approved and pending NOCs
        for item in (approved_nocs[:15] + pending_nocs[:10]):
            context_lines.append(f"  * {item}")
    else:
        context_lines.append("- No NOC (Notice of Change) records found in this scope.")
    context_lines.append("")

    # 4. Overdue Submittals
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

    # 5. Targeted Specific Search & Bilingual Semantic Mapping
    scoped_search_pids = list(target_pids)
    p_lower = user_prompt.lower()

    # Detect Project Mention in Prompt (e.g. CFC, Assiut, Suez, PEM-058, etc.)
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

    if "شوب درونج" in p_lower or "شوب دروينج" in p_lower or "شوبدروينج" in p_lower:
        if "SD" not in detected_doc_types:
            detected_doc_types.append("SD")
    if "ماتريال" in p_lower or "اعتماد مواد" in p_lower:
        if "MS" not in detected_doc_types:
            detected_doc_types.append("MS")

    if active_tab and active_tab.upper() in KNOWN_DOC_TYPES and not detected_doc_types:
        detected_doc_types.append(active_tab.upper())

    # Bilingual Keyword Extraction (Arabic to English mapping + raw keywords)
    search_terms = []
    for ar_kw, en_kw in KEYWORD_MAP.items():
        if ar_kw in p_lower:
            if en_kw not in search_terms:
                search_terms.append(en_kw)
            if ar_kw not in search_terms:
                search_terms.append(ar_kw)

    # General English/Alphanumeric tokens
    stop_words = {
        "and", "the", "for", "with", "all", "what", "show", "list", "give", "from",
        "find", "submittal", "submittals", "document", "documents", "project", "projects",
        "status", "approved", "noted", "date", "please", "cfc", "pem", "any", "our",
        "help", "query", "record", "records"
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

    # Approval Status Detection (e.g., معتمد, Approved, Status A/B)
    approval_indicators = [
        "معتمد", "معتمدة", "معتمدين", "موافقة", "موافق", "مقبول",
        "approved", "approval", "status a", "status b", "code a", "code b"
    ]
    is_approved_query = any(ind in p_lower for ind in approval_indicators)

    if search_terms or detected_doc_types:
        where_clauses = ["r.project_id = ANY(%s)"]
        params = [scoped_search_pids]

        if detected_doc_types:
            where_clauses.append("UPPER(r.dt_id) = ANY(%s)")
            params.append([dt.upper() for dt in detected_doc_types])

        if search_terms:
            kw_clauses = []
            for term in search_terms[:3]:
                kw_clauses.append("(r.data->>'title' ILIKE %s OR r.data->>'docNo' ILIKE %s)")
                params.extend([f"%{term}%", f"%{term}%"])
            where_clauses.append("(" + " OR ".join(kw_clauses) + ")")

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
            LIMIT 20;
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


def _call_gemini_api(prompt, context_text):
    """Calls Gemini Flash API with strict 15-second timeout and fast model execution."""
    global _working_gemini_model

    api_key = os.environ.get("GEMINI_API_KEY", "").strip()
    if not api_key:
        return None, "NO_API_KEY"

    full_contents = (
        f"### CONTEXT REGISTER DATA (READ-ONLY):\n{context_text}\n\n"
        f"### USER QUESTION:\n{prompt}\n\n"
        "Provide a structured, helpful, professional engineering response in Markdown format. "
        "If the user asks in Arabic, answer in clear, professional Arabic (باللغة العربية الهندسية). "
        "If the user asks in English, answer in English. "
        "Highlight document numbers, statuses, titles, and approval codes clearly. "
        "Use bullet points, bold key figures, and tables where appropriate."
    )

    models_to_try = []
    if _working_gemini_model:
        models_to_try.append(_working_gemini_model)
    for m in ["gemini-2.5-flash", "gemini-2.0-flash"]:
        if m not in models_to_try:
            models_to_try.append(m)

    last_err = None
    try:
        client = genai.Client(api_key=api_key, http_options={"timeout": 15.0})
        for model_name in models_to_try:
            try:
                resp = client.models.generate_content(
                    model=model_name,
                    contents=full_contents,
                    config=types.GenerateContentConfig(
                        system_instruction=SYSTEM_INSTRUCTION,
                        temperature=0.2,
                    ),
                )
                if resp and resp.text:
                    _working_gemini_model = model_name
                    return resp.text, None
            except Exception as e_model:
                last_err = e_model
                logger.warning("Gemini model %s failed or timed out: %s", model_name, e_model)
                continue
    except Exception as e_client:
        last_err = e_client
        logger.error("Gemini client initialization error: %s", e_client)

    return None, f"AI service response timed out or failed: {str(last_err)}"


@ai_bp.route("/query", methods=["POST"])
@ai_bp.route("/api/ai/query", methods=["POST"])
def api_ai_query():
    """Main AI Assistant Query Endpoint."""
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

    try:
        context_text = _build_ai_context(target_pids, user_prompt=prompt, active_tab=tab)
        reply, err = _call_gemini_api(prompt, context_text)

        if err == "NO_API_KEY":
            return jsonify(error="AI Assistant is not configured on this instance."), 503

        if err or not reply:
            if "time" in str(err).lower():
                return jsonify(error=f"AI service response timed out: {str(err)}"), 504
            return jsonify(
                reply=(
                    "⚠️ **Unable to complete AI analysis.**\n\n"
                    f"The AI service encountered an error: `{err or 'Empty response from model'}`. "
                    "Please verify network connection and API quotas."
                )
            ), 200

        return jsonify(reply=reply)
    except Exception as e:
        logger.exception("AI assistant query exception: %s", e)
        return jsonify(error=f"Internal error processing AI query: {str(e)}"), 500
