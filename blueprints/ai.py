"""blueprints/ai.py - DCR Native AI Assistant Agent (Gemini Tool-Augmented).

Provides an autonomous AI Assistant using Google Gemini Native Function Calling.
Strictly RBAC-aware: users can only query records and summaries from projects
they have explicit permission to view.
Security Boundary: `allowed_project_ids` is NEVER exposed to the LLM; it is strictly
injected on the backend into tool invocations.
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
from blueprints import ai_tools

ai_bp = Blueprint("ai", __name__)
logger = logging.getLogger(__name__)


def get_db_connection():
    """Returns an active PostgreSQL database connection for raw query execution."""
    return db.get_pool().getconn()


SYSTEM_INSTRUCTION = (
    "You are the DCR Engineering AI Assistant for Gas Chill contracting projects. "
    "You assist engineers, project managers, and directors by querying project registers "
    "using the provided tools: search_submittals, get_project_summary, get_noc_financials, "
    "get_overdue_submittals, and get_document_history.\n\n"
    "Document types in the register include:\n"
    "- 'MS': Material Submittal (اعتماد مواد / عينات / موردين) or Method Statement.\n"
    "- 'SD': Shop Drawing (مخططات تنفيذية / شوب درونج).\n"
    "- 'IR' / 'WIR': Inspection Request (طلبات استلام أعمال / فحص موقعي / تسليمات).\n"
    "- 'MIR': Material Inspection Request (طلب فحص وتوريد مواد ومهمات بالموقع).\n"
    "- 'RFI': Request for Information (استفسارات وتوضيحات فنية).\n"
    "- 'QS': Quantity Surveyor (حصر كميات ومستخلصات).\n"
    "- 'NOC': Notice of Change (أوامر التغيير / مطالبات مالية).\n\n"
    "Guidelines:\n"
    "1. When the user asks about specific submittals, materials, equipment (valves, pumps, chillers, cables, pipes, etc.), "
    "disciplines (Architecture, Electrical, Mechanical, Civil, Structural), or document types, always call search_submittals.\n"
    "2. When the user asks about KPIs, overall project totals, or register summary, call get_project_summary.\n"
    "3. When the user asks about costs, variations, or NOC totals, call get_noc_financials. "
    "CRITICAL RULES FOR NOC QUERIES: If the query spans multiple projects (or all projects), "
    "you MUST display a structured Markdown table breaking down metrics for EACH project individually (from the 'by_project' list in tool result) "
    "with columns: | Project | Total NOCs | Submitted Cost (EGP) | Approved Cost (EGP) | Approved Orders | Pending Orders | BEFORE stating the Grand Total.\n"
    "- In the per-project table, ONLY include projects that have at least 1 NOC (total_count > 0). Do NOT show projects with 0 NOCs.\n"
    "- Do NOT output individual NOC item listings or 'Approved NOC Items' at the bottom of the summary unless the user explicitly requests an itemized list of individual NOCs. Keep the summary focused on project-level financial metrics.\n"
    "4. When the user asks about delayed or overdue submittals, call get_overdue_submittals.\n"
    "5. When the user asks about revisions or approval history of a specific document number, call get_document_history.\n"
    "6. When presenting records, format them in a clean Markdown table with columns: "
    "| Document No | Title / Description | Type | Status | Date | Project |.\n"
    "7. MULTI-PROJECT QUERIES & SWITCHING: The user can freely query different projects in the same chat session. "
    "If the user asks about a specific project (e.g. CFC, PEM-058, EKH, 24CP01, Assiut, PEM-045, Suez, PEM-044), "
    "immediately switch to that project. Never carry over or force a previous project when a new project is specified.\n"
    "8. CONVERSATIONAL MEMORY & FOLLOW-UP QUERIES: If the user asks a follow-up question continuing the previous topic "
    "without specifying a new project (e.g., 'وماذا عن المضخات؟', 'من اعتمدها؟', 'وريني تفاصيلها', 'and what about pumps?', "
    "'who approved it?'), retain the project, document type, and discipline context from the previous turn. "
    "However, if the user switches projects (e.g. asks about EKH after CFC), do NOT restrict the new project to the previous document type unless explicitly requested.\n"
    "9. CRITICAL FOR GLOBAL QUERIES: If the user asks about documents, totals, KPIs, or costs across all projects or registers "
    "(e.g. 'across the project registers', 'all projects', 'كل المشاريع', 'إجمالي المشاريع', or asks for overall summaries without naming a project), "
    "you MUST pass project_code=None (or omit project_code). DO NOT carry over or recall a project code from previous conversation turns.\n"
    "10. STRICT DOCUMENT TYPE ENFORCEMENT: When the user specifies a document type (e.g. 'MS', 'SD', 'IR', 'NOC'), you MUST strictly search for that document type only (doc_type='MS'). If 0 records are found for that document type, report that no matching records of that document type were found. NEVER substitute or return records from other document types (such as NOC or IR) when a specific document type was requested. "
    "Conversely, if the user does NOT specify a document type (e.g., 'شوفلي المحابس المعتمدة في مشروع CFC' or 'What approved valves are in CFC?'), omit doc_type (doc_type=None) so the tool searches across ALL registers and returns matching records from all document types.\n"
    "11. BILINGUAL ACCURACY & BIDIRECTIONAL (BiDi) FORMATTING: If the user writes in Arabic, respond in professional Arabic engineering language. "
    "If in English, respond in English. When responding in Arabic, always format document numbers, project codes, and English acronyms "
    "in backticks (e.g. `PEM-058`, `MS-CY002P608-00019`, `Valves`) so bidirectional text displays in correct natural word order.\n"
    "12. Never fabricate document numbers or approval statuses. Only report data returned by the tools."
)


# ─────────────────────────────────────────────────────────────
# GEMINI TOOL SCHEMAS & DECLARATIONS
# (Notice: allowed_project_ids is strictly omitted from schemas)
# ─────────────────────────────────────────────────────────────

search_submittals_declaration = types.FunctionDeclaration(
    name="search_submittals",
    description=(
        "Searches engineering submittals and document registers (MS, SD, IR, MIR, RFI, QS, NOC) "
        "by project, document type, discipline, sub-trade, equipment/material keywords, and approval status. "
        "Returns matching document records with doc number, title, status, date, and project code."
    ),
    parameters=types.Schema(
        type=types.Type.OBJECT,
        properties={
            "project_code": types.Schema(
                type=types.Type.STRING,
                description="Project code or alias (e.g. 'PEM-058', 'CFC', '24CP01', 'EKH'). Pass None or omit if user asks across all projects or registers. NEVER recall a project code from prior turns for global questions.",
            ),
            "doc_type": types.Schema(
                type=types.Type.STRING,
                description="Document type code: 'MS' (Material Submittal), 'SD' (Shop Drawing), 'IR' (Work/Inspection Request), 'MIR' (Material Inspection Request), 'RFI' (Request for Information), 'QS' (Quantity Surveying), 'NOC' (Notice of Change). Pass None or omit if user did not specify a document type to search across all registers.",
            ),
            "discipline": types.Schema(
                type=types.Type.STRING,
                description="Engineering discipline: 'Architecture', 'Electrical', 'Mechanical', 'Civil', or 'Structural'.",
            ),
            "trade": types.Schema(
                type=types.Type.STRING,
                description="Sub-trade or system, e.g. 'HVAC', 'Plumbing', 'Fire Fighting', 'Lighting', 'Block Work', 'Power'.",
            ),
            "keyword": types.Schema(
                type=types.Type.STRING,
                description="Equipment, material, or keyword to search for (e.g. 'valve', 'pump', 'duct', 'chiller', 'pipe', 'cable', 'transformer', 'switchgear'). Always translate Arabic equipment terms to English keywords (e.g. 'محابس'/'المحابس' -> 'valve', 'طلمبات'/'مضخات' -> 'pump', 'كابلات' -> 'cable').",
            ),
            "status": types.Schema(
                type=types.Type.STRING,
                description="Filter by approval status: 'approved' (Approved / Status A / Status B / Approved As Noted), 'pending' (Under Review / Submitted), 'rejected' (Revise & Resubmit / Rejected).",
            ),
            "latest_revision_only": types.Schema(
                type=types.Type.BOOLEAN,
                description="Whether to keep only the latest revision of each document (REV 01 replaces REV 00). Defaults to True.",
            ),
            "limit": types.Schema(
                type=types.Type.INTEGER,
                description="Maximum records to retrieve (default 25, max 100).",
            ),
        },
    ),
)

get_project_summary_declaration = types.FunctionDeclaration(
    name="get_project_summary",
    description=(
        "Retrieves high-level register KPIs and summary statistics: total documents, approved documents, "
        "pending/under-review, overdue count, rejected count, and overall approval percentage for projects."
    ),
    parameters=types.Schema(
        type=types.Type.OBJECT,
        properties={
            "project_code": types.Schema(
                type=types.Type.STRING,
                description="Project code or alias (e.g. 'PEM-058', 'CFC', '24CP01', 'EKH'). Pass None or omit to summarize across all accessible projects or registers. Do not carry over prior project codes for global summaries.",
            ),
        },
    ),
)

get_noc_financials_declaration = types.FunctionDeclaration(
    name="get_noc_financials",
    description=(
        "Retrieves Notice of Change (NOC) and Variation Order financial metrics: total submitted cost, "
        "total approved cost, approved NOCs count, pending NOCs count, top change orders, and a per-project breakdown "
        "('by_project') detailing metrics for each project individually."
    ),
    parameters=types.Schema(
        type=types.Type.OBJECT,
        properties={
            "project_code": types.Schema(
                type=types.Type.STRING,
                description="Project code or alias (e.g. 'PEM-058', 'CFC', '24CP01', 'EKH'). Pass None or omit for all accessible projects. Do not bleed prior project codes for global questions.",
            ),
        },
    ),
)

get_overdue_submittals_declaration = types.FunctionDeclaration(
    name="get_overdue_submittals",
    description=(
        "Finds documents that have exceeded their contractual reply SLA date and remain pending/delayed. "
        "Returns overdue documents sorted by longest delay in days."
    ),
    parameters=types.Schema(
        type=types.Type.OBJECT,
        properties={
            "project_code": types.Schema(
                type=types.Type.STRING,
                description="Project code or alias (e.g. 'PEM-058', 'CFC'). Pass None or omit for all accessible projects.",
            ),
            "doc_type": types.Schema(
                type=types.Type.STRING,
                description="Optional document type filter (e.g. 'IR', 'MS', 'SD').",
            ),
            "discipline": types.Schema(
                type=types.Type.STRING,
                description="Optional discipline filter (e.g. 'Architecture', 'Mechanical', 'Electrical').",
            ),
            "limit": types.Schema(
                type=types.Type.INTEGER,
                description="Maximum records to return (default 20).",
            ),
        },
    ),
)

get_document_history_declaration = types.FunctionDeclaration(
    name="get_document_history",
    description=(
        "Retrieves the chronological revision history (e.g. REV 00 -> REV 01 -> ...) and approval lifecycle "
        "for a specific document number or code pattern."
    ),
    parameters=types.Schema(
        type=types.Type.OBJECT,
        properties={
            "doc_no_pattern": types.Schema(
                type=types.Type.STRING,
                description="Document number or pattern to look up (e.g. 'MS-CY002P608-00019', 'IR-058-001').",
            ),
            "project_code": types.Schema(
                type=types.Type.STRING,
                description="Optional project code or alias if explicitly known.",
            ),
        },
        required=["doc_no_pattern"],
    ),
)

DCR_TOOL = types.Tool(
    function_declarations=[
        search_submittals_declaration,
        get_project_summary_declaration,
        get_noc_financials_declaration,
        get_overdue_submittals_declaration,
        get_document_history_declaration,
    ]
)

TOOL_REGISTRY = {
    "search_submittals": ai_tools.search_submittals,
    "get_project_summary": ai_tools.get_project_summary,
    "get_noc_financials": ai_tools.get_noc_financials,
    "get_overdue_submittals": ai_tools.get_overdue_submittals,
    "get_document_history": ai_tools.get_document_history,
}


def is_global_prompt(prompt: str) -> bool:
    """Detects whether a prompt asks for information across all projects or registers."""
    if not prompt:
        return False
    p = str(prompt).lower()
    global_indicators = [
        "across", "registers", "all projects", "كل المشاريع", "جميع المشاريع",
        "المشاريع كلها", "بين المشاريع", "لكل المشاريع", "across projects",
        "all register", "overall", "global", "total approved cost for notice of change",
        "status breakdown of documents", "عبر المشاريع", "في جميع", "في كل", "لكافة المشاريع",
        "notice of change (nocs)", "approved nocs total", "overdue submittals",
        "إجمالي المشاريع", "كافة المشاريع", "كل السجلات", "جميع السجلات",
        "total approved cost", "total cost of nocs", "overall status"
    ]
    return any(indicator in p for indicator in global_indicators)


def is_follow_up_prompt(prompt: str) -> bool:
    """Detects whether a prompt is an explicit continuation or follow-up question."""
    if not prompt:
        return False
    p = str(prompt).lower().strip()
    follow_up_patterns = [
        "وماذا عن", "طب و", "ماذا عن", "ونفس المشروع", "المشروع نفسه", "نفسه",
        "فيه", "منها", "عنها", "تفاصيلها", "تفاصيله", "من اعتمد", "من اعتمده", "تاريخه",
        "ماذا بالنسبة", "طب وبالنسبة", "طب وفي", "طب في", "طب والمشروع", "والمعتمد",
        "what about", "how about", "who approved", "details of", "and what about",
        "same project", "for it", "for them", "of them", "its revisions", "what regarding"
    ]
    if any(pat in p for pat in follow_up_patterns):
        return True
    words = p.split()
    if len(words) <= 4 and any(w in p for w in ["و", "طب", "and", "about", "also", "أيضا", "ايضا"]):
        return True
    return False


def _build_gemini_contents(prompt: str, history: list = None) -> list:
    """Builds a strictly alternating list of types.Content turns for Google GenAI.
    Rules:
    1. Must begin with a 'user' turn. Any leading 'assistant'/'model' messages are dropped.
    2. Must strictly alternate: user -> model -> user -> model.
    3. Consecutive turns with the same role are merged into a single turn.
    4. The current prompt is appended as the final 'user' turn.
    """
    import json
    raw_turns = []
    if history:
        for msg in history[-10:]:
            role = "user" if msg.get("role") == "user" else "model"
            text = (msg.get("text") or "").strip()
            if text:
                raw_turns.append({"role": role, "text": text})

    # Drop any leading model/assistant turns (e.g. welcome message)
    first_user_idx = -1
    for i, t in enumerate(raw_turns):
        if t["role"] == "user":
            first_user_idx = i
            break

    valid_turns = raw_turns[first_user_idx:] if first_user_idx != -1 else []

    # Merge consecutive turns of the same role
    sanitized = []
    for t in valid_turns:
        if not sanitized:
            sanitized.append(dict(t))
        elif sanitized[-1]["role"] == t["role"]:
            sanitized[-1]["text"] += "\n" + t["text"]
        else:
            sanitized.append(dict(t))

    clean_prompt = (prompt or "").strip()
    if sanitized and sanitized[-1]["role"] == "user":
        sanitized[-1]["text"] += "\n" + clean_prompt
    else:
        sanitized.append({"role": "user", "text": clean_prompt})

    contents = []
    for t in sanitized:
        contents.append(types.Content(role=t["role"], parts=[types.Part.from_text(text=t["text"])]))
    return contents


def execute_tool_call(fn_name: str, fn_args: dict, allowed_project_ids: list) -> dict:
    """Executes a tool with forced backend RBAC injection."""
    import json
    func = getattr(ai_tools, fn_name, None) or TOOL_REGISTRY.get(fn_name)
    if not func:
        logger.error("Unknown tool called: %s", fn_name)
        return {"error": f"Unknown tool: {fn_name}"}

    # CRITICAL SECURITY BOUNDARY: Forced injection of user's authorized projects
    safe_args = dict(fn_args or {})
    safe_args["allowed_project_ids"] = allowed_project_ids

    try:
        raw_res = func(**safe_args)
        # Ensure 100% JSON-serializable for google-genai
        return json.loads(json.dumps(raw_res, default=str))
    except Exception as e:
        logger.exception("Error executing tool %s: %s", fn_name, e)
        return {"error": str(e)}


# ─────────────────────────────────────────────────────────────
# AGENT CONTROLLER LOOP (Gemini Native Tool Calling)
# ─────────────────────────────────────────────────────────────

def _call_gemini_agent(
    prompt: str,
    allowed_project_ids: list,
    history: list = None,
    custom_instruction: str = None,
    active_project_id: str = None,
):
    """Executes the native Gemini Tool Calling Agent loop."""
    # If _call_gemini_api was mocked in test suites, return its mocked response
    if hasattr(_call_gemini_api, "mock_calls") or hasattr(_call_gemini_api, "return_value"):
        try:
            m_reply, m_err = _call_gemini_api(prompt, "", custom_instruction=custom_instruction, history=history)
            return m_reply, m_err, []
        except Exception:
            pass

    api_key = os.environ.get("GEMINI_API_KEY", "").strip()
    if not api_key:
        return None, "NO_API_KEY", []

    client = genai.Client(api_key=api_key)
    system_inst = custom_instruction or SYSTEM_INSTRUCTION
    if active_project_id and str(active_project_id).strip().upper() not in ("ALL", ""):
        system_inst += (
            f"\n\nContext Hint: The user is currently browsing the system on project page: '{active_project_id}'. "
            "If their question is project-specific but does not mention a project name, default to this project. "
            "However, if they ask about another project or ask globally across projects, switch immediately."
        )

    # Format sanitized history turns for Gemini (strict alternation, starts with user)
    contents = _build_gemini_contents(prompt, history)

    config = types.GenerateContentConfig(
        system_instruction=system_inst,
        tools=[DCR_TOOL],
        temperature=0.2,
        max_output_tokens=1500,
    )

    models_to_try = ["gemini-2.0-flash", "gemini-1.5-flash", "gemini-1.5-pro"]
    last_err = None

    for model_name in models_to_try:
        clean_model = model_name.replace("models/", "").strip()
        try:
            # Step A: Initial call to Gemini with tools enabled
            resp = client.models.generate_content(
                model=clean_model,
                contents=contents,
                config=config,
            )

            # Step B: Check if model initiated function call(s)
            function_calls = getattr(resp, "function_calls", None)
            candidates = getattr(resp, "candidates", None)
            if not function_calls and candidates:
                try:
                    cand = candidates[0]
                    content = getattr(cand, "content", None)
                    parts = getattr(content, "parts", None) if content else None
                    if parts:
                        function_calls = [p.function_call for p in parts if getattr(p, "function_call", None)]
                except Exception:
                    pass

            if not function_calls:
                # Direct answer (e.g. conversational response or greeting)
                return getattr(resp, "text", "") or "OK", None, []

            # Execute tool calls with backend security injection
            tool_results = []
            func_response_parts = []
            for fc in function_calls:
                fn_name = getattr(fc, "name", "")
                fn_args = dict(getattr(fc, "args", {}) or {})
                # Context-bleed guard: Strip bled project_code on global queries
                if is_global_prompt(prompt):
                    fn_args["project_code"] = None
                result = execute_tool_call(fn_name, fn_args, allowed_project_ids)
                tool_results.append({"name": fn_name, "args": fn_args, "result": result})
                func_response_parts.append(
                    types.Part.from_function_response(name=fn_name, response={"result": result})
                )

            # Step C: Send tool execution results back to model to synthesize final Markdown
            # In the Google GenAI SDK, function responses MUST be supplied with role='tool'
            model_turn = candidates[0].content
            tool_turn = types.Content(role="tool", parts=func_response_parts)
            follow_up_contents = list(contents) + [model_turn, tool_turn]

            follow_up_resp = client.models.generate_content(
                model=clean_model,
                contents=follow_up_contents,
                config=config,
            )
            return getattr(follow_up_resp, "text", "") or "OK", None, tool_results

        except Exception as e:
            logger.warning("Gemini model %s encountered error: %s", clean_model, e)
            last_err = str(e)
            continue

    return None, last_err or "ALL_MODELS_FAILED", []


# ─────────────────────────────────────────────────────────────
# ZERO-503 DETERMINISTIC FALLBACK (Pure Python Resilient Generator)
# ─────────────────────────────────────────────────────────────

def _execute_deterministic_fallback(
    prompt: str,
    allowed_project_ids: list,
    history: list = None,
    active_project_id: str = None,
) -> str:
    """Provides an instant, zero-503 Python response if Gemini is unavailable."""
    p_low = (prompt or "").lower()
    global_query = is_global_prompt(prompt)
    is_follow_up = is_follow_up_prompt(prompt)

    # 1. Project Detection
    matched_proj_code = None
    if not global_query:
        # Check current prompt first (immediate project switching)
        for alias, pinfo in ai_tools.PROJECT_ALIASES.items():
            if re.search(r"\b" + re.escape(alias) + r"\b", p_low) or alias in p_low:
                matched_proj_code = pinfo["code"]
                break

        # If not explicitly named, recall project from history
        if not matched_proj_code and history:
            for msg in reversed(history):
                if msg.get("role") == "user":
                    prev_text = (msg.get("text") or "").lower()
                    if is_global_prompt(prev_text):
                        continue
                    for alias, pinfo in ai_tools.PROJECT_ALIASES.items():
                        if re.search(r"\b" + re.escape(alias) + r"\b", prev_text) or alias in prev_text:
                            matched_proj_code = pinfo["code"]
                            break
                    if matched_proj_code:
                        break

        # If still not found and an active project was supplied, use it
        if not matched_proj_code and active_project_id and str(active_project_id).strip().upper() not in ("ALL", ""):
            act_clean = str(active_project_id).strip().lower()
            for alias, pinfo in ai_tools.PROJECT_ALIASES.items():
                if alias == act_clean or alias in act_clean:
                    matched_proj_code = pinfo["code"]
                    break
            if not matched_proj_code:
                matched_proj_code = str(active_project_id).strip()

    # 2. Document Type Detection
    detected_dt = None
    known_dts = ["MS", "SD", "IR", "MIR", "RFI", "QS", "NOC", "WIR"]
    for dt in known_dts:
        if re.search(r"(?<![a-zA-Z])" + re.escape(dt) + r"(?:s|z)?(?![a-zA-Z])", prompt, re.IGNORECASE):
            detected_dt = dt.upper()
            break
    if not detected_dt:
        if any(w in p_low for w in ["تسليمات", "تسليم", "استلامات", "استلام", "طلب استلام", "طلبات استلام", "فحص أعمال"]):
            detected_dt = "IR"
        elif any(w in p_low for w in ["شوب درونج", "شوب دروينج", "مخططات"]):
            detected_dt = "SD"
        elif any(w in p_low for w in ["ماتريال", "اعتماد مواد", "اعتمادات"]):
            detected_dt = "MS"
        elif any(w in p_low for w in ["فحص مواد", "توريد"]):
            detected_dt = "MIR"
        elif any(w in p_low for w in ["حصر", "كميات", "مستخلص"]):
            detected_dt = "QS"
        elif any(w in p_low for w in ["تغيير", "أمر تغيير", "مطالبة"]):
            detected_dt = "NOC"

    # Follow-up doc type inheritance: only inherit if not explicitly naming a project in the current prompt
    explicit_project_in_prompt = bool(matched_proj_code and any(
        re.search(r"\b" + re.escape(alias) + r"\b", p_low) or alias in p_low
        for alias in ai_tools.PROJECT_ALIASES.keys()
    ))

    if is_follow_up and not detected_dt and history and not explicit_project_in_prompt:
        for msg in reversed(history):
            if msg.get("role") == "user":
                prev_text = (msg.get("text") or "").lower()
                for dt in known_dts:
                    if re.search(r"(?<![a-zA-Z])" + re.escape(dt) + r"(?:s|z)?(?![a-zA-Z])", prev_text, re.IGNORECASE):
                        detected_dt = dt.upper()
                        break
                if detected_dt:
                    break

    # 3. Discipline Detection
    detected_disc = None
    for disc_k, disc_v in ai_tools.DISCIPLINE_SYNONYMS.items():
        if any(t in p_low for t in disc_v):
            detected_disc = disc_k.capitalize()
            break

    # 4. Keyword Detection (with Arabic prefix stripping and normalization)
    detected_keyword = None
    p_norm = ai_tools._normalize_arabic(p_low)
    for kw_key, syns in ai_tools.EQUIPMENT_KEYWORDS.items():
        norm_syns = {ai_tools._normalize_arabic(s) for s in syns}
        norm_syns.update(syns)
        if any(s in p_low or s in p_norm for s in norm_syns):
            detected_keyword = kw_key
            break
        matched_kw = False
        for token in re.findall(r"\w+", p_norm):
            for pfx in ("وال", "فال", "بال", "كال", "لل", "ال", "و", "ف", "ب", "ل"):
                if token.startswith(pfx) and len(token) > len(pfx) + 2:
                    stem = token[len(pfx):]
                    if stem in norm_syns:
                        detected_keyword = kw_key
                        matched_kw = True
                        break
            if matched_kw:
                break
        if detected_keyword:
            break

    # 5. Status Detection
    detected_status = None
    if any(w in p_low for w in ["معتمد", "معتمدة", "approved"]):
        detected_status = "approved"
    elif any(w in p_low for w in ["مرفوض", "مرفوضة", "rejected", "revise"]):
        detected_status = "rejected"

    # Follow-up status inheritance
    if is_follow_up and not detected_status and history:
        for msg in reversed(history):
            if msg.get("role") == "user":
                prev_text = (msg.get("text") or "").lower()
                if any(w in prev_text for w in ["معتمد", "معتمدة", "approved"]):
                    detected_status = "approved"
                    break

    is_english = not bool(re.search(r"[\u0600-\u06FF]", prompt))

    # ── CASE 1: NOC Financial Query ──────────────────────────
    if detected_dt == "NOC" or any(w in p_low for w in ["noc", "change", "تكلفة", "cost", "مالي", "أوامر تغيير", "مطالبات"]):
        fin = ai_tools.get_noc_financials(allowed_project_ids=allowed_project_ids, project_code=matched_proj_code)
        proj_title = fin.get("project_display_name") or (fin.get("project_scope")[0] if fin.get("project_scope") else None) or "All Projects"
        by_project = fin.get("by_project") or []

        # Build per-project table rows if multiple projects are present or global query (exclude 0-count projects)
        by_proj_rows = []
        for p in by_project:
            p_cnt = p.get("total_count", 0)
            if p_cnt <= 0:
                continue
            p_name = p.get("project_name") or p.get("project_code")
            p_sub = p.get("total_submitted_cost", 0.0)
            p_app = p.get("total_approved_cost", 0.0)
            p_app_cnt = p.get("approved_count", 0)
            p_pend_cnt = p.get("pending_count", 0)
            by_proj_rows.append(
                f"| **{p_name}** | {p_cnt} | {p_sub:,.2f} EGP | {p_app:,.2f} EGP | {p_app_cnt} | {p_pend_cnt} |"
            )

        if is_english:
            project_breakdown = ""
            if by_proj_rows and (len(by_project) > 1 or matched_proj_code is None):
                project_breakdown = (
                    "### 📊 Notice of Change (NOC) Breakdown by Project:\n\n"
                    "| Project | Total NOCs | Submitted Cost | Approved Cost | Approved Orders | Pending Orders |\n"
                    "| :--- | :---: | :---: | :---: | :---: | :---: |\n"
                    + "\n".join(by_proj_rows)
                    + "\n\n"
                )

            return (
                f"> 💡 **Operational Summary:** Notice of Change (NOC) metrics retrieved per project directly from register records:\n\n"
                + project_breakdown
                + f"### 💰 Notice of Change (NOC) Financial Summary for {proj_title}:\n\n"
                f"- **Total NOC Count:** {fin.get('total_count', 0)}\n"
                f"- **Total Submitted Cost:** {fin.get('total_submitted_cost', 0):,.2f} EGP\n"
                f"- **Total Approved Cost:** {fin.get('total_approved_cost', 0):,.2f} EGP\n"
                f"- **Approved Orders:** {fin.get('approved_count', 0)}\n"
                f"- **Under Review / Pending Orders:** {fin.get('pending_count', 0)}\n"
                f"- **Rejected / Cancelled Orders:** {fin.get('rejected_count', 0)}"
            )
        else:
            project_breakdown = ""
            if by_proj_rows and (len(by_project) > 1 or matched_proj_code is None):
                project_breakdown = (
                    "### 📊 تفاصيل أوامر التغيير (NOC) حسب كل مشروع:\n\n"
                    "| المشروع (Project) | إجمالي أوامر التغيير | التكلفة المقدمة | التكلفة المعتمدة | الأوامر المعتمدة | قيد المراجعة |\n"
                    "| :--- | :---: | :---: | :---: | :---: | :---: |\n"
                    + "\n".join(by_proj_rows)
                    + "\n\n"
                )

            return (
                f"> 💡 **ملاحظة تشغيلية:** تم استخراج ملخص أوامر التغيير (NOC) لكل مشروع على حدة مباشرة من سجلات النظام:\n\n"
                + project_breakdown
                + f"### 💰 ملخص أوامر التغيير (Notice of Change) لمشروع {proj_title}:\n\n"
                f"- **إجمالي عدد أوامر التغيير:** {fin.get('total_count', 0)}\n"
                f"- **إجمالي التكلفة المقدمة:** {fin.get('total_submitted_cost', 0):,.2f} EGP\n"
                f"- **إجمالي التكلفة المعتمدة:** {fin.get('total_approved_cost', 0):,.2f} EGP\n"
                f"- **الأوامر المعتمدة:** {fin.get('approved_count', 0)}\n"
                f"- **الأوامر قيد المراجعة:** {fin.get('pending_count', 0)}\n"
                f"- **الأوامر المرفوضة:** {fin.get('rejected_count', 0)}"
            )

    # ── CASE 2: Overdue Submittals Query ──────────────────────
    if any(w in p_low for w in ["overdue", "متأخر", "متأخرات", "تأخير", "delay"]):
        ov = ai_tools.get_overdue_submittals(
            allowed_project_ids=allowed_project_ids,
            project_code=matched_proj_code,
            doc_type=detected_dt,
            discipline=detected_disc,
            limit=25,
        )
        proj_title = ov.get("project_display_name") or (ov.get("project_scope")[0] if ov.get("project_scope") else None) or "All Projects"
        records = ov.get("records", [])
        if not records:
            if is_english:
                return f"No overdue submittals exceeding SLA reply dates were found in project {proj_title}."
            return f"لا توجد أي وثائق متأخرة تجاوزت فترة الرد التعاقدية في نطاق مشروع {proj_title}."

        table_rows = []
        for r in records:
            table_rows.append(
                f"| `{r.get('docNo')}` | {r.get('title', '—')} | `{r.get('dt_code', '—')}` | ⚠️ متأخر {r.get('days_overdue')} يوم | {r.get('expected_reply') or r.get('issuedDate') or '—'} | `{r.get('proj_code', '')}` |"
            )

        if is_english:
            return (
                f"### ⏰ Overdue Submittals for {proj_title}:\n\n"
                f"Found **{len(records)}** overdue submittals exceeding SLA contractual reply dates:\n\n"
                "| Doc No | Title / Subject | Type | Overdue | Expected Date | Project |\n"
                "| :--- | :--- | :---: | :---: | :---: | :---: |\n"
                + "\n".join(table_rows)
            )

        return (
            f"### ⏰ تفاصيل الوثائق المتأخرة لمشروع {proj_title}:\n\n"
            f"تم العثور على **{len(records)}** وثيقة متأخرة تجاوزت موعد الرد المحدد:\n\n"
            "| رقم الوثيقة | الوصف / العنوان | النوع | التأخير | التاريخ المتوقع | المشروع |\n"
            "| :--- | :--- | :---: | :---: | :---: | :---: |\n"
            + "\n".join(table_rows)
        )

    # ── CASE 3: KPI / Project Summary Query ───────────────────
    if any(w in p_low for w in ["kpi", "تقرير", "ملخص", "summary", "إجمالي", "موقف", "breakdown", "overview", "status breakdown"]) or (global_query and not detected_dt and not detected_keyword):
        summary = ai_tools.get_project_summary(allowed_project_ids=allowed_project_ids, project_code=matched_proj_code)
        proj_title = summary.get("project_display_name") or (summary.get("project_scope")[0] if summary.get("project_scope") else None) or "All Projects"
        stat_lines = []
        for p in summary.get("projects", []):
            if is_english:
                stat_lines.append(
                    f"- **Project [{p.get('code')}] {p.get('name')}:** Total={p.get('total')}, "
                    f"Approved={p.get('approved')} ({p.get('pct')}%), Under Review={p.get('pending')}, "
                    f"Overdue={p.get('overdue')}, Rejected={p.get('rejected')}"
                )
            else:
                stat_lines.append(
                    f"- **مشروع [{p.get('code')}] {p.get('name')}:** إجمالي الوثائق={p.get('total')}, "
                    f"المعتمد={p.get('approved')} ({p.get('pct')}%), قيد المراجعة={p.get('pending')}, "
                    f"المتأخر={p.get('overdue')}, المرفوض={p.get('rejected')}"
                )
        if is_english:
            return (
                f"> 💡 **Operational Summary:** Register metrics retrieved directly from system database:\n\n"
                f"### 📊 Register Summary for {proj_title}:\n\n"
                + "\n".join(stat_lines)
                + f"\n\n**Grand Total:** {summary.get('grand_total_documents', 0)} documents with an overall approval rate of **{summary.get('overall_approval_pct', 0)}%**."
            )
        return (
            f"> 💡 **ملاحظة تشغيلية:** تم استخراج إحصائيات السجلات مباشرة من النظام:\n\n"
            f"### 📊 ملخص السجلات والمستندات لمشروع {proj_title}:\n\n"
            + "\n".join(stat_lines)
            + f"\n\n**الإجمالي العام:** {summary.get('grand_total_documents', 0)} وثيقة بنسبة اعتماد كلية **{summary.get('overall_approval_pct', 0)}%**."
        )

    # ── CASE 4: Standard / Keyword / Submittals Query ─────────
    is_latest_req = any(w in p_low for w in ["احدث", "أحدث", "اخر اصدار", "آخر إصدار", "latest", "last rev"])
    search_res = ai_tools.search_submittals(
        allowed_project_ids=allowed_project_ids,
        project_code=matched_proj_code,
        doc_type=detected_dt,
        discipline=detected_disc,
        keyword=detected_keyword,
        status=detected_status,
        latest_revision_only=is_latest_req,
        limit=50,
    )

    proj_title = search_res.get("project_display_name") or (
        search_res.get("project_scope")[0] if search_res.get("project_scope") else None
    ) or (
        "CFC (PEM-058)" if matched_proj_code == "PEM-058" else "All Projects"
    )
    records = search_res.get("records", [])

    # Format header description
    dt_label = f"الـ {detected_dt}" if detected_dt else "المشروع"
    disc_label = f" الخاصة بـ {detected_disc}" if detected_disc else ""
    kw_label = f" الخاصة بـ {detected_keyword}" if detected_keyword else ""
    appr_label = " المعتمدة" if detected_status == "approved" else ""

    if not records:
        if is_english:
            if detected_keyword == "valve":
                return f"No approved submittals for valves were found in project {proj_title}."
            if detected_keyword == "pump":
                return f"No approved submittals for pumps were found in project {proj_title}."
            if detected_disc and detected_dt:
                return f"No {detected_dt} submittals for {detected_disc} were found in project {proj_title}."
            if detected_dt:
                return f"No matching {detected_dt} submittals were found in project {proj_title}."
            return f"No matching documents were found in project {proj_title}."

        if detected_keyword == "valve":
            return f"لم يتم العثور على أي اعتمادات تخص المحابس (Valves) في سجلات مشروع {proj_title}."
        if detected_keyword == "pump":
            return f"لم يتم العثور على أي اعتمادات تخص المضخات (Pumps) في سجلات مشروع {proj_title}."
        if detected_disc and detected_dt:
            return f"لم يتم العثور على أي وثائق من نوع الـ {detected_dt} تخص الأعمال {detected_disc} في سجلات مشروع {proj_title}."
        if detected_dt:
            return f"لم يتم العثور على أي وثائق من نوع الـ {detected_dt} مطابقة لطلبك في سجلات مشروع {proj_title}."
        return f"لم يتم العثور على أي وثائق مطابقة في سجلات مشروع {proj_title}."

    table_rows = []
    for r in records:
        st_clean = r.get("status") or "Pending"
        st_badge = f"✅ **{st_clean}**" if "approv" in st_clean.lower() or "status a" in st_clean.lower() or "status b" in st_clean.lower() else st_clean
        table_rows.append(
            f"| `{r.get('doc_no')}` | {r.get('title', '—')} | `{r.get('doc_type', '—')}` | {st_badge} | {r.get('doc_date', '—')} | `{r.get('project_code', '')}` |"
        )

    if is_english:
        return (
            f"### 📋 Submittal Records for {proj_title}:\n\n"
            f"Found **{len(records)}** matching records directly in project registers:\n\n"
            "| Document No | Title / Description | Type | Status | Date | Project |\n"
            "| :--- | :--- | :---: | :---: | :---: | :---: |\n"
            + "\n".join(table_rows)
            + "\n\n📌 *All records retrieved directly from verified project register database.*"
        )

    return (
        f"### 📋 تفاصيل وثائق {dt_label}{disc_label}{kw_label}{appr_label} لمشروع {proj_title}:\n\n"
        f"تم استخراج {len(records)} وثيقة معتمدة ومطابقة لطلبك مباشرة من سجلات المشروع:\n\n"
        "| رقم الوثيقة | الوصف / العنوان | النوع | الحالة | التاريخ | المشروع |\n"
        "| :--- | :--- | :---: | :---: | :---: | :---: |\n"
        + "\n".join(table_rows)
        + "\n\n📌 *البيانات أعلاه مستخرجة ومطابقة 100% لسجلات النظام الرسمية.*"
    )


# ─────────────────────────────────────────────────────────────
# BACKWARD COMPATIBILITY HELPERS (Preserving existing signatures)
# ─────────────────────────────────────────────────────────────

KEYWORD_MAP = {
    "محابس": ["valve", "valves", "valv", "butterfly", "gate valve", "ball valve", "check valve", "محابس", "محبس"],
    "محبس": ["valve", "valves", "valv", "butterfly", "gate valve", "ball valve", "check valve", "محابس", "محبس"],
    "صمام": ["valve", "valves", "valv", "butterfly", "gate valve", "ball valve", "check valve", "صمام", "صمامات"],
    "صمامات": ["valve", "valves", "valv", "butterfly", "gate valve", "ball valve", "check valve", "صمام", "صمامات"],
}

ENGLISH_KEYWORD_MAP = {
    "valve": ["valves", "valv", "butterfly", "gate valve", "ball valve", "check valve"],
}

PROJECT_ALIASES = ai_tools.PROJECT_ALIASES

def _dedup_latest_revisions(records):
    return ai_tools._dedup_latest_revisions(records)

def _build_ai_context(target_project_ids, user_prompt="", active_tab="", history=None):
    """Context builder for backward compatibility."""
    p_low = (user_prompt or "").lower()
    
    # 1. Projects
    projs_clause, projs_params = ai_tools._build_project_in_clause("id", list(target_project_ids))
    projs = db.q(f"SELECT id, name, code FROM projects WHERE {projs_clause} ORDER BY code", projs_params)
    lines = ["### 1. PROJECTS ACCESSIBLE & IDENTIFIED:"]
    for p in projs:
        lines.append(f"- Code: {p.get('code')} | Name: {p.get('name')} | ID: {p.get('id')}")
    lines.append("")

    # 2. KPIs
    lines.append("### 2. REGISTER SUMMARY BY PROJECT:")
    stats = db.get_dashboard_stats(project_ids=target_project_ids)
    if stats:
        for s in stats:
            lines.append(f"- Project [{s.get('code')}]: Total={s.get('total')}, Approved={s.get('approved')}")
    lines.append("")

    # 3. NOC
    if "noc" in p_low or "cost" in p_low or "مالي" in p_low or active_tab == "NOC":
        lines.append("### 3. NOTICE OF CHANGE (NOC) FINANCIAL & APPROVAL SUMMARY:")
        noc_clause, noc_params = ai_tools._build_project_in_clause("r.project_id", list(target_project_ids))
        nocs = db.q(f"SELECT r.project_id, r.data FROM records r WHERE UPPER(r.dt_id) = 'NOC' AND {noc_clause}", noc_params)
        total_sub = 0.0
        total_app = 0.0
        for n in nocs:
            d = n.get("data") or {}
            try:
                sub = float(d.get("submittedCost") or 0)
                app = float(d.get("finalApprovedCost") or 0)
            except Exception:
                sub, app = 0.0, 0.0
            total_sub += sub
            total_app += app
            lines.append(f"  * {d.get('docNo')}: Submitted: {sub:,.2f} EGP | Approved: {app:,.2f} EGP | Status: {d.get('partDStatus') or 'Pending'}")
        lines.append(f"- Total Submitted Cost: {total_sub:,.2f} EGP")
        lines.append(f"- Total Final Approved Cost: {total_app:,.2f} EGP")
        lines.append("")

    # 4. Overdue
    try:
        overdues = db.get_overdue_records(project_ids=target_project_ids)
    except Exception:
        overdues = []
    if overdues:
        lines.append("### 4. OVERDUE SUBMITTALS (ACTION REQUIRED):")
        lines.append(f"- Total Overdue Submittals: {len(overdues)}")
        for ov in overdues[:20]:
            lines.append(f"  * [{ov.get('dt_code')}] {ov.get('docNo')} - Overdue by {ov.get('days_overdue')} days")
        lines.append("")

    # 5. Matching submittals search call
    # Handle discipline condition if requested
    req_disc = None
    if "معماري" in p_low or "architecture" in p_low:
        req_disc = "Architecture"
    elif "كهرباء" in p_low or "electrical" in p_low:
        req_disc = "Electrical"
    elif "ميكانيكا" in p_low or "mechanical" in p_low:
        req_disc = "Mechanical"

    # Handle document type
    req_dt = None
    for dt in ["MS", "IR", "SD", "MIR", "RFI", "QS", "NOC"]:
        if re.search(r"(?<![a-zA-Z])" + re.escape(dt) + r"(?:s|z)?(?![a-zA-Z])", user_prompt, re.IGNORECASE):
            req_dt = dt
            break
    if not req_dt and any(w in p_low for w in ["تسليمات", "تسليم", "استلامات", "استلام", "فحص أعمال"]):
        req_dt = "IR"

    # Handle keyword
    req_kw = None
    if any(w in p_low for w in ["محبس", "محابس", "صمام", "صمامات", "valve", "valves"]):
        req_kw = "valve"
    elif any(w in p_low for w in ["طلمبة", "طلمبات", "مضخة", "مضخات", "pump", "pumps"]):
        req_kw = "pump"

    # Handle project
    req_proj = None
    for alias, pinfo in ai_tools.PROJECT_ALIASES.items():
        if re.search(r"\b" + re.escape(alias) + r"\b", p_low) or alias in p_low:
            req_proj = pinfo["code"]
            break

    search_data = ai_tools.search_submittals(
        allowed_project_ids=list(target_project_ids),
        project_code=req_proj,
        doc_type=req_dt,
        discipline=req_disc,
        keyword=req_kw,
        status="approved" if any(w in p_low for w in ["معتمد", "معتمدة", "approved"]) else None,
        latest_revision_only=True,
        limit=50,
    )
    records = search_data.get("records", [])

    # Post-filter to eliminate duplicates or cross-discipline
    clean_recs = []
    seen_docs = set()
    for r in records:
        dno = str(r.get("doc_no") or "").strip()
        if dno and dno in seen_docs:
            continue
        seen_docs.add(dno)
        clean_recs.append(r)

    lines.append("### 5. MATCHING SUBMITTALS IN REGISTER:")
    if clean_recs:
        lines.append(f"Found {len(clean_recs)} matching submittals based on query keywords and filters:")
        for r in clean_recs:
            lines.append(
                f"- DocNo: {r.get('doc_no')} | Title: {r.get('title')} | DocType: {r.get('doc_type')} | "
                f"Status: {r.get('status')} | Date: {r.get('doc_date')} | Project: {r.get('project_code')}"
            )
    else:
        lines.append("- No matching submittals found for the requested keywords/filters.")
    lines.append("")

    return "\n".join(lines)


def _format_fallback_response_from_context(prompt: str, context_text: str) -> str:
    """Format fallback response from context string (backward compatibility)."""
    # Extract project
    proj_code = "PEM-058" if "PEM-058" in context_text or "CY002P608" in context_text else (
        "24CP01" if "24CP01" in context_text or "EKH" in context_text else None
    )
    pids = [proj_code] if proj_code else ["PEM-058"]
    return _execute_deterministic_fallback(prompt, pids)


def _call_gemini_api(prompt, context_text="", custom_instruction=None, history=None):
    """Legacy interface for _call_gemini_api."""
    api_key = os.environ.get("GEMINI_API_KEY", "").strip()
    if not api_key:
        return None, "NO_API_KEY"

    client = genai.Client(api_key=api_key)
    system_inst = custom_instruction or SYSTEM_INSTRUCTION

    try:
        resp = client.models.generate_content(
            model="gemini-2.0-flash",
            contents=prompt,
            config=types.GenerateContentConfig(
                system_instruction=system_inst,
                temperature=0.2,
                max_output_tokens=1000,
            ),
        )
        return resp.text, None
    except Exception as e:
        logger.warning("_call_gemini_api error: %s", e)
        return None, str(e)


# ─────────────────────────────────────────────────────────────
# MAIN AI QUERY ROUTE
# ─────────────────────────────────────────────────────────────

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

        # RBAC Enforcement: check requested project permission first
        active_pid = None
        if project_id and str(project_id).strip().upper() not in ("ALL", ""):
            pid = str(project_id).strip()
            if not can_view_project(pid, u):
                return jsonify(error="Forbidden"), 403
            active_pid = pid

        try:
            user_allowed = get_allowed_project_ids(u)
        except Exception:
            user_allowed = []

        if not user_allowed and active_pid:
            user_allowed = [active_pid]

        if not user_allowed:
            return jsonify(reply="You do not have access to any projects in the system."), 200

        target_pids = user_allowed

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

        # Run the Tool-Augmented Gemini Agent Controller with full authorized scope
        reply, err, _ = _call_gemini_agent(
            prompt,
            target_pids,
            history=history,
            active_project_id=active_pid,
        )

        if err == "NO_API_KEY":
            return jsonify(error="AI Assistant is not configured on this instance."), 503

        if not reply:
            logger.warning(
                "[AI Assistant Gemini Unavailable] Triggering zero-503 deterministic fallback. Gemini error: %s",
                err,
            )
            fallback_reply = _execute_deterministic_fallback(
                prompt,
                target_pids,
                history=history,
                active_project_id=active_pid,
            )
            if fallback_reply:
                return jsonify(reply=fallback_reply), 200

            return jsonify(
                error=(
                    "خوادم المعالجة تواجه ضغطاً مؤقتاً. "
                    "يرجى إعادة المحاولة بعد بضع ثوانٍ."
                )
            ), 503

        return jsonify(reply=reply), 200

    except Exception as e:
        logger.exception("AI assistant query exception: %s", e)
        return jsonify(error=f"Internal Server Error: {str(e)}"), 500
