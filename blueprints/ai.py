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

    # 5. Targeted Specific Search (if user asks about a specific doc number or keyword)
    tokens = [t for t in re.findall(r"[A-Za-z0-9_-]{3,}", user_prompt) if not t.isdigit()]
    if tokens:
        # Check if user mentioned a document code pattern or specific word
        matches = []
        for token in tokens[:3]:
            try:
                found = db.q("""
                    SELECT r.project_id, r.dt_id, r.data, p.code as proj_code
                    FROM records r
                    JOIN projects p ON p.id = r.project_id
                    WHERE r.project_id = ANY(%s)
                      AND (r.data::text ILIKE %s)
                    LIMIT 6
                """, (target_pids, f"%{token}%"))
                for f in found:
                    if f not in matches:
                        matches.append(f)
            except Exception as e:
                logger.warning("Error searching specific records: %s", e)

        if matches:
            context_lines.append("### 5. MATCHING RECORDS IN REGISTER (RELEVANT TO QUERY):")
            for m in matches[:10]:
                d = m.get("data") or {}
                doc_no = d.get("docNo") or d.get("nocNo") or d.get("letterRef") or "—"
                title = d.get("title") or d.get("subject") or ""
                st = d.get("status") or d.get("partDStatus") or ""
                dt = m.get("dt_id")
                p_code = m.get("proj_code")
                issued = d.get("issuedDate") or d.get("partAIssueDate") or ""
                reply = d.get("actualReplyDate") or d.get("partDReturnDate") or ""
                context_lines.append(
                    f"  * [{p_code}] {doc_no} ({dt}): '{title}' | Status: {st} | Issued: {issued} | Reply: {reply}"
                )
            context_lines.append("")

    return "\n".join(context_lines)


_working_gemini_model = None


def _call_gemini_api(prompt, context_text):
    """Calls Gemini Flash API with dynamic model discovery and fallback."""
    global _working_gemini_model

    api_key = os.environ.get("GEMINI_API_KEY", "").strip()
    if not api_key:
        return None, "NO_API_KEY"

    full_contents = (
        f"### CONTEXT REGISTER DATA (READ-ONLY):\n{context_text}\n\n"
        f"### USER QUESTION:\n{prompt}\n\n"
        "Provide a structured, helpful, professional engineering response in Markdown format. "
        "Use bullet points, bold key figures, and tables where appropriate."
    )

    last_err = None

    # 1. Preferred modern models order (latest first)
    preferred_models = [
        "gemini-2.5-flash",
        "gemini-2.0-flash",
        "gemini-2.0-flash-exp",
        "gemini-flash-latest",
        "gemini-2.5-pro",
        "gemini-2.0-pro-exp-02-05",
    ]

    candidates = []
    if _working_gemini_model:
        candidates.append(_working_gemini_model)
    for m in preferred_models:
        if m not in candidates:
            candidates.append(m)

    # Try preferred candidates via official google-genai SDK
    try:
        client = genai.Client(api_key=api_key)
        for model_candidate in candidates:
            try:
                resp = client.models.generate_content(
                    model=model_candidate,
                    contents=full_contents,
                    config=types.GenerateContentConfig(
                        system_instruction=SYSTEM_INSTRUCTION,
                        temperature=0.2,
                    ),
                )
                if resp and resp.text:
                    _working_gemini_model = model_candidate
                    return resp.text, None
            except Exception as e_cand:
                last_err = e_cand
                logger.warning("Candidate model %s failed: %s", model_candidate, e_cand)
                continue

        # 2. Dynamic Discovery: Query API for available models supporting generateContent
        try:
            dynamic_models = []
            for m in client.models.list():
                m_name = getattr(m, "name", "") or ""
                clean_name = m_name.replace("models/", "").strip()
                if not clean_name:
                    continue
                supported = getattr(m, "supported_generation_methods", None) or getattr(m, "supported_actions", None)
                if supported and "generateContent" not in supported:
                    continue
                if "gemini" in clean_name.lower():
                    if "flash" in clean_name.lower():
                        dynamic_models.insert(0, clean_name)
                    else:
                        dynamic_models.append(clean_name)

            for model_dyn in dynamic_models:
                if model_dyn in candidates:
                    continue
                try:
                    resp = client.models.generate_content(
                        model=model_dyn,
                        contents=full_contents,
                        config=types.GenerateContentConfig(
                            system_instruction=SYSTEM_INSTRUCTION,
                            temperature=0.2,
                        ),
                    )
                    if resp and resp.text:
                        _working_gemini_model = model_dyn
                        return resp.text, None
                except Exception as e_dyn:
                    last_err = e_dyn
                    logger.warning("Dynamic model %s failed: %s", model_dyn, e_dyn)
                    continue
        except Exception as list_err:
            logger.warning("Dynamic model listing failed: %s", list_err)

    except Exception as e_client:
        last_err = e_client
        logger.warning("google-genai client error: %s", e_client)

    logger.error("All Gemini API attempts failed: %s", last_err)
    return None, str(last_err)


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
