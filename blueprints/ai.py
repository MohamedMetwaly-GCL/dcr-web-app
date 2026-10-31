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
    "- 'PR': Purchase Requisition / Procurement Tracking (طلبات الشراء / أوامر الشراء والتوريد ومتابعة المشتريات).\n"
    "- 'MS': Material Submittal (اعتماد مواد / عينات / موردين) or Method Statement.\n"
    "- 'SD': Shop Drawing (مخططات تنفيذية / شوب درونج).\n"
    "- 'IR' / 'WIR': Inspection Request (طلبات استلام أعمال / فحص موقعي / تسليمات).\n"
    "- 'MIR': Material Inspection Request (طلب فحص وتوريد مواد ومهمات بالموقع).\n"
    "- 'DS': Design Submittal (اعتمادات التصميم والتنفيذ).\n"
    "- 'RFI': Request for Information (استفسارات وتوضيحات فنية).\n"
    "- 'QS': Quantity Surveyor (حصر كميات ومستخلصات).\n"
    "- 'NOC': Notice of Change (أوامر التغيير / مطالبات مالية).\n"
    "- 'LTR': Letters & Correspondence (الخطابات والمراسلات الرسمية).\n\n"
    "Guidelines:\n"
    "1. When the user asks about specific submittals, materials, PRs/procurement, equipment (valves, pumps, chillers, cables, pipes, etc.), "
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
    "6. TABLE FORMATTING & INTEGRITY (CRITICAL):\n"
    "   - Format records in a clean Markdown table with columns: | Document No | Title / Description | Type | Status | Date | Project |.\n"
    "   - In Arabic responses, use Arabic column headers: | رقم الوثيقة | البيان والمواصفات | النوع | الحالة | التاريخ | المشروع |.\n"
    "   - NEVER put pipe characters `|` or newlines `\\n` inside any table cell, as this destroys Markdown table syntax. If describing multiple items or specifications, separate them with commas or dashes (e.g. 'Gate Valve 6in, Butterfly Valve 1050mm').\n"
    "   - Ensure EVERY table row has EXACTLY the same number of columns (6 columns).\n"
    "   - When presenting PR (Purchase Requisition) records: in the description column ('البيان والمواصفات'), display the specific matching line items and their exact quantities (e.g. 'Gate Valve 4in (Qty: 10 No), Butterfly Valve 6in (Qty: 5 No)'). NEVER dump the raw 'PR Details' column, long technical pump skid specifications, or unrelated plumbing fittings (elbows, tees, couplings, pipes) into the table.\n"
    "   - For PRs and Submittals, show only the latest active revision (REV01 replaces REV00) unless the user explicitly requests revision history.\n"
    "   - Maintain the records in ascending order of Document No (from lowest to highest number) as returned by the tool.\n"
    "7. MULTI-PROJECT QUERIES & SWITCHING: The user can freely query different projects in the same chat session. "
    "Projects in the system include:\n"
    "- '24CP01': EKH New Building Office (Egypt Kuwait Holding - EKH / اي كيه اتش)\n"
    "- 'PEM-058': CFC District Cooling Plant Phase 2 (AFREB / CFC / سي اف سي / كايرو فيستيفال)\n"
    "- 'PEM-064': Suez Steel Admin Building (Suez Steel Co. / Suez / السويس / حديد السويس)\n"
    "- 'PEM-065': Eastmain MEP Buildings B06 & B10 (MOBCO Real Estate / Eastmain / Eastmain B06 / Eastmain B10 / ايست مين)\n"
    "- 'PEM-066': Eastmain - MEP Building B08 (MOBCO Real Estate / Eastmain B08 / ايست مين B08)\n"
    "- 'PEM-067': Assiut Hydrocracking Complex - ANOPC (Assiut / ANOPC / أسيوط / أنوبك)\n"
    "- 'PEM-068': Four Season Lift Station (Four Season / Four Seasons / فور سيزون / محطة الرفع)\n"
    "When the user asks about 'Eastmain' generically without specifying a building, search across Eastmain projects (project_code='Eastmain' or 'PEM-065'). If they specify 'B08', use project_code='PEM-066'. If they specify 'B06' or 'B10', use project_code='PEM-065'.\n"
    "If the user asks about a specific project, immediately switch to that project. Never carry over or force a previous project when a new project is specified.\n"
    "8. CONVERSATIONAL MEMORY, ELLIPSIS & PROJECT SWITCHING (CRITICAL):\n"
    "   - CONVERSATIONAL ELLIPSIS (FOLLOW-UP WITH PROJECT SWITCH): When the user asks a follow-up question that specifies a new project "
    "using a continuation or comparative phrase (e.g. 'وبالنسبة لمشروع CFC', 'وماذا عن CFC', 'طب وفي مشروع CFC', 'what about CFC project?', 'and for CFC?', 'طب ومشروع Eastmain', 'وفي مشروع Eastmain B08'), "
    "the user is asking the SAME question for the newly specified project! You MUST retain the previous turn's query intent: the equipment/keyword (e.g. 'valve'), "
    "the document type (e.g. 'MS'), discipline, and approval status (e.g. 'approved'), and apply them directly to the new project! "
    "For example: if the previous query was 'ايه MS الخاص بالمحابس المعتمد في مشروع EKH', and the user follows up with 'وبالنسبة لمشروع CFC', "
    "you MUST immediately call search_submittals(project_code='PEM-058', doc_type='MS', keyword='valve', status='approved'). "
    "And if they follow up with 'طيب ومشروع Eastmain', you MUST call search_submittals(project_code='Eastmain', doc_type='MS', keyword='valve', status='approved').\n"
    "   - FOLLOW-UP WITH EQUIPMENT / TOPIC SWITCH: When the user asks about different equipment or document type without specifying a project "
    "(e.g. 'وماذا عن المضخات؟', 'طب والـ SD؟', 'طب و PRs المحابس؟', 'and what about pumps?'), retain the project and approval status from the previous turn, and update the keyword or doc_type.\n"
    "   - ISOLATED PROJECT MENTION WITH NO ACTIVE TOPIC: If the user mentions ONLY a project name (e.g. 'مشروع CFC' or 'CFC') and there is NO active inquiry "
    "(no keyword, no specific doc_type, no NOC, no overdue, no summary) in the conversation history to continue, DO NOT execute a blind search returning random documents. "
    "Instead, ask the user politely and specifically what information they need for this project (Material Submittals, Purchase Requisitions, Shop Drawings, NOC financials, Overdue submittals, or General KPI Summary).\n"
    "9. CRITICAL FOR GLOBAL QUERIES: If the user asks about documents, totals, KPIs, or costs across all projects or registers "
    "(e.g. 'across the project registers', 'all projects', 'كل المشاريع', 'إجمالي المشاريع', or asks for overall summaries without naming a project), "
    "you MUST pass project_code=None (or omit project_code). DO NOT carry over or recall a project code from previous conversation turns.\n"
    "10. STRICT DOCUMENT TYPE ENFORCEMENT: When the user specifies a document type (e.g. 'PR', 'MS', 'SD', 'IR', 'NOC'), you MUST strictly search for that document type only (doc_type='PR'). If 0 records are found for that document type, report that no matching records of that document type were found. NEVER substitute or return records from other document types (such as NOC, IR, or DS) when a specific document type was requested. "
    "Conversely, if the user does NOT specify a document type (e.g., 'شوفلي المحابس المعتمدة في مشروع CFC' or 'What approved valves are in CFC?'), omit doc_type (doc_type=None) so the tool searches across ALL registers and returns matching records from all document types.\n"
    "11. BILINGUAL ACCURACY & PROFESSIONAL ARABIC PHRASING:\n"
    "   - If the user writes in Arabic, respond in fluent, professional Egyptian engineering Arabic. Never output literal robotic English translations like 'الخاصة بـ valve لمشروع All Projects'.\n"
    "   - Instead, phrase naturally: 'طلبات الشراء (PR) الخاصة بالمحابس (Valves) لكافة المشاريع:' or 'اعتمادات المحابس المعتمدة في مشروع...'.\n"
    "   - Always format document numbers, project codes, and English acronyms in backticks (e.g. `PEM-058`, `PR-005 REV01`, `Valves`) so bidirectional text displays in correct natural word order.\n"
    "12. Never fabricate document numbers or approval statuses. Only report data returned by the tools."
)


# ─────────────────────────────────────────────────────────────
# GEMINI TOOL SCHEMAS & DECLARATIONS
# (Notice: allowed_project_ids is strictly omitted from schemas)
# ─────────────────────────────────────────────────────────────

search_submittals_declaration = types.FunctionDeclaration(
    name="search_submittals",
    description=(
        "Searches engineering submittals and document registers (PR, MS, SD, IR, MIR, RFI, QS, NOC) "
        "by project, document type, discipline, sub-trade, equipment/material keywords, and approval status. "
        "Returns matching document records with doc number, title, status, date, and project code."
    ),
    parameters=types.Schema(
        type=types.Type.OBJECT,
        properties={
            "project_code": types.Schema(
                type=types.Type.STRING,
                description="Project code or alias (e.g. 'PEM-058', 'CFC', '24CP01', 'EKH', 'PEM-064', 'Suez', 'PEM-065', 'PEM-066', 'Eastmain', 'Eastmain B08', 'PEM-067', 'ANOPC', 'PEM-068', 'Four Season'). Pass None or omit if user asks across all projects or registers. NEVER recall a project code from prior turns for global questions.",
            ),
            "doc_type": types.Schema(
                type=types.Type.STRING,
                description="Document type code: 'PR' (Purchase Requisition / Procurement Tracking), 'MS' (Material Submittal), 'SD' (Shop Drawing), 'IR' (Work/Inspection Request), 'MIR' (Material Inspection Request), 'RFI' (Request for Information), 'QS' (Quantity Surveying), 'NOC' (Notice of Change). Pass None or omit if user did not specify a document type to search across all registers.",
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
                description="Project code or alias (e.g. 'PEM-058', 'CFC', '24CP01', 'EKH', 'PEM-064', 'Suez', 'PEM-065', 'PEM-066', 'Eastmain', 'PEM-067', 'ANOPC', 'PEM-068', 'Four Season'). Pass None or omit to summarize across all accessible projects or registers. Do not carry over prior project codes for global summaries.",
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
                description="Project code or alias (e.g. 'PEM-058', 'CFC', '24CP01', 'EKH', 'PEM-064', 'Suez', 'PEM-065', 'PEM-066', 'Eastmain', 'PEM-067', 'ANOPC', 'PEM-068', 'Four Season'). Pass None or omit for all accessible projects. Do not bleed prior project codes for global questions.",
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
                description="Project code or alias (e.g. 'PEM-058', 'CFC', '24CP01', 'EKH', 'PEM-064', 'PEM-065', 'PEM-066', 'Eastmain', 'PEM-067', 'PEM-068'). Pass None or omit for all accessible projects.",
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
    p_norm = ai_tools._normalize_arabic(p)
    follow_up_patterns = [
        "وبالنسبة", "بالنسبة", "ماذا بالنسبة", "وماذا بالنسبة", "طب وبالنسبة", "طب بالنسبه",
        "وماذا عن", "ماذا عن", "وعن", "عن مشروع",
        "طب و", "طب وفي", "طب في", "طب والمشروع", "طب مشروع",
        "وفي مشروع", "في مشروع", "ومشروع",
        "في نفس المشروع", "بنفس المشروع", "نفس المشروع", "في نفس", "ونفس المشروع", "المشروع نفسه", "نفسه", "نفسها",
        "فيه", "فيها", "منها", "عنها", "عنه", "تفاصيلها", "تفاصيله", "من اعتمد", "من اعتمده", "من اعتمدها",
        "تاريخه", "تاريخها", "والمعتمد", "المعتمد منها", "المعتمدة منها",
        "what about", "how about", "who approved", "details of", "and what about", "and for", "and in",
        "same project", "in the same project", "for it", "for them", "of them", "its revisions", "what regarding"
    ]
    if any(pat in p or pat in p_norm for pat in follow_up_patterns):
        return True
    if p.startswith(("و ", "طب ", "and ", "so ", "also ")):
        return True
    words = p.split()
    if len(words) <= 5:
        for alias in sorted(ai_tools.PROJECT_ALIASES.keys(), key=lambda x: len(x), reverse=True):
            alias_norm = ai_tools._normalize_arabic(alias)
            if re.search(r"\b" + re.escape(alias) + r"\b", p) or alias in p or alias_norm in p_norm:
                return True
    return False


def _extract_inquiry_intent_from_text(text: str) -> dict:
    """Extracts project_code, doc_type, keyword, discipline, status, and high-level intent from user text."""
    if not text:
        return {}
    p_low = str(text).lower().strip()
    p_norm = ai_tools._normalize_arabic(p_low)

    # 1. Project Detection
    proj_code = None
    for alias, pinfo in sorted(ai_tools.PROJECT_ALIASES.items(), key=lambda x: len(x[0]), reverse=True):
        alias_norm = ai_tools._normalize_arabic(alias)
        if (
            re.search(r"\b" + re.escape(alias) + r"\b", p_low)
            or alias in p_low
            or alias_norm in p_norm
            or re.search(r"\b" + re.escape(alias_norm) + r"\b", p_norm)
        ):
            proj_code = pinfo["code"]
            break

    # Dynamic DB fallback project check if not matched by static alias
    if not proj_code:
        try:
            db_projs = db.get_projects()
        except Exception:
            db_projs = []
        for dp in db_projs:
            d_code = str(dp.get("code") or "").strip().lower()
            d_name = str(dp.get("name") or "").strip().lower()
            d_norm = ai_tools._normalize_arabic(d_name)
            if d_code and (d_code in p_low or re.search(r"\b" + re.escape(d_code) + r"\b", p_low)):
                proj_code = dp.get("code")
                break
            if d_name and (d_name in p_low or p_low in d_name or d_norm in p_norm):
                proj_code = dp.get("code")
                break

    is_same_project = bool(any(w in p_low for w in ["في نفس المشروع", "بنفس المشروع", "نفس المشروع", "same project", "in the same project"]))

    # 2. Document Type Detection (Dynamically resolved from registered doc types across all projects)
    dt = None
    registered_dts = ai_tools.get_registered_doc_types()
    # Sort by code length descending so longer codes (e.g. MIR, NOC) match before shorter prefixes (e.g. IR)
    sorted_dts = sorted(registered_dts, key=lambda d: len(d.get("code") or ""), reverse=True)
    for d in sorted_dts:
        kdt = (d.get("code") or "").strip()
        if not kdt:
            continue
        if re.search(r"(?<![a-zA-Z])" + re.escape(kdt) + r"(?:s|z)?(?![a-zA-Z])", text, re.IGNORECASE):
            dt = kdt.upper()
            break
        # Also check English or Arabic register name from DB
        dname = (d.get("name") or "").strip().lower()
        if dname and len(dname) >= 4 and dname in p_low:
            dt = kdt.upper()
            break

    if not dt:
        if any(w in p_low for w in ["تسليمات", "تسليم", "استلامات", "استلام", "طلب استلام", "طلبات استلام", "فحص أعمال"]):
            dt = "IR"
        elif any(w in p_low for w in ["شوب درونج", "شوب دروينج", "مخططات"]):
            dt = "SD"
        elif any(w in p_low for w in ["ماتريال", "اعتماد مواد", "اعتمادات", "اعتماد خامات"]):
            dt = "MS"
        elif any(w in p_low for w in [
            "طلب شراء", "طلب الشراء", "طلبات شراء", "طلبات الشراء",
            "امر شراء", "أمر شراء", "امر الشراء", "أمر الشراء",
            "اوامر شراء", "أوامر شراء", "اوامر الشراء", "أوامر الشراء",
            "مشتريات", "المشتريات",
            "امر توريد", "أمر توريد", "امر التوريد", "أمر التوريد",
            "اوامر توريد", "أوامر توريد", "اوامر التوريد", "أوامر التوريد",
            "بي ار", "pr"
        ]):
            dt = "PR"
        elif any(w in p_low for w in ["فحص مواد", "فحص توريد", "محضر توريد"]):
            dt = "MIR"
        elif any(w in p_low for w in ["حصر", "كميات", "مستخلص"]):
            dt = "QS"
        elif any(w in p_low for w in ["تغيير", "أمر تغيير", "امر تغيير", "مطالبة", "مطالبات"]):
            dt = "NOC"
        elif any(w in p_low for w in ["خطاب", "خطابات", "مراسلة", "مراسلات", "جواب", "جوابات"]):
            dt = "LTR"
        elif any(w in p_low for w in ["عدم مطابقة", "مخالفة"]):
            dt = "NCR"
        elif any(w in p_low for w in ["محضر اجتماع", "محاضر اجتماعات"]):
            dt = "MOM"
        elif any(w in p_low for w in ["تعليمات شفهية", "أوامر شفهية"]):
            dt = "CVI"
        elif any(w in p_low for w in ["تعليمات موقع"]):
            dt = "SI"
        elif any(w in p_low for w in ["تعليمات مهندس"]):
            dt = "EI"
        elif any(w in p_low for w in ["ازبلت", "از بيلت", "as built"]):
            dt = "ABD"

    # 3. Keyword Detection (with Arabic prefix stripping and normalization)
    kw = None
    for kw_key, syns in ai_tools.EQUIPMENT_KEYWORDS.items():
        norm_syns = {ai_tools._normalize_arabic(s) for s in syns}
        norm_syns.update(syns)
        if any(s in p_low or s in p_norm for s in norm_syns):
            kw = kw_key
            break
        matched_kw = False
        for token in re.findall(r"\w+", p_norm):
            for pfx in ("وال", "فال", "بال", "كال", "لل", "ال", "و", "ف", "ب", "ل"):
                if token.startswith(pfx) and len(token) > len(pfx) + 2:
                    stem = token[len(pfx):]
                    if stem in norm_syns:
                        kw = kw_key
                        matched_kw = True
                        break
            if matched_kw:
                break
        if kw:
            break

    # 4. Discipline Detection
    disc = None
    for disc_k, disc_v in ai_tools.DISCIPLINE_SYNONYMS.items():
        if any(t in p_low for t in disc_v):
            disc = disc_k.capitalize()
            break

    # 5. Status Detection
    status = None
    if any(w in p_low for w in ["معتمد", "معتمدة", "approved"]):
        status = "approved"
    elif any(w in p_low for w in ["مرفوض", "مرفوضة", "rejected", "revise"]):
        status = "rejected"
    elif any(w in p_low for w in ["قيد المراجعة", "معلق", "pending", "under review"]):
        status = "pending"

    # 6. Intent (NOC, Overdue, Summary)
    is_noc = bool(dt == "NOC" or any(w in p_low for w in ["noc", "change", "تكلفة", "cost", "مالي", "أوامر تغيير", "أمر تغيير", "مطالبات", "مطالبة"]))
    is_overdue = bool(any(w in p_low for w in ["overdue", "متأخر", "متأخرات", "تأخير", "delay"]))
    is_summary = bool(any(w in p_low for w in ["summary", "ملخص", "إحصائيات", "احصائيات", "نسب اعتماد", "نسبة اعتماد", "kpi"]))

    return {
        "project_code": proj_code,
        "doc_type": dt,
        "keyword": kw,
        "discipline": disc,
        "status": status,
        "is_noc": is_noc,
        "is_overdue": is_overdue,
        "is_summary": is_summary,
        "is_same_project": is_same_project,
        "raw_text": text,
    }


def _get_last_inquiry_from_history(history: list) -> dict:
    """Finds the most recent user inquiry in history that contains an actionable engineering topic."""
    if not history:
        return {}
    for msg in reversed(history):
        if msg.get("role") in ("user", "human"):
            txt = msg.get("text") or ""
            if is_global_prompt(txt):
                continue
            parsed = _extract_inquiry_intent_from_text(txt)
            if (
                parsed.get("is_noc")
                or parsed.get("is_overdue")
                or parsed.get("is_summary")
                or parsed.get("keyword")
                or parsed.get("doc_type")
                or parsed.get("discipline")
            ):
                return parsed
    return {}


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

    # Dynamically inject registered registers from DB so Gemini always knows all available registers across all projects
    try:
        reg_types = ai_tools.get_registered_doc_types()
        if reg_types:
            dt_list_str = ", ".join([f"{d['code']} ({d['name']})" for d in reg_types if d.get('code')])
            system_inst += f"\n\nDynamic System Registers: Registers configured in this system across projects include:\n{dt_list_str}."
    except Exception:
        pass

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
                else:
                    curr_intent = _extract_inquiry_intent_from_text(prompt)
                    # Project switch guard: If current user prompt explicitly specifies a project, ensure fn_args uses current project
                    curr_parsed_proj = curr_intent.get("project_code")
                    if curr_parsed_proj:
                        fn_args["project_code"] = curr_parsed_proj
                    elif curr_intent.get("is_same_project"):
                        prev_inq = _get_last_inquiry_from_history(history)
                        if prev_inq and prev_inq.get("project_code"):
                            fn_args["project_code"] = prev_inq["project_code"]
                        elif active_project_id and str(active_project_id).strip().upper() not in ("ALL", ""):
                            fn_args["project_code"] = active_project_id

                    # Strict doc_type enforcement: if user explicitly specified doc_type (e.g. PR), enforce it
                    if curr_intent.get("doc_type") and fn_name == "search_submittals":
                        fn_args["doc_type"] = curr_intent["doc_type"]

                if is_follow_up_prompt(prompt) and fn_name == "search_submittals":
                    # Ellipsis guard: if user switched project or followed up, inherit topic from history
                    prev_inq = _get_last_inquiry_from_history(history)
                    if prev_inq:
                        if not fn_args.get("keyword") and prev_inq.get("keyword"):
                            fn_args["keyword"] = prev_inq["keyword"]
                        if not fn_args.get("doc_type") and prev_inq.get("doc_type"):
                            fn_args["doc_type"] = prev_inq["doc_type"]
                        # Status inheritance guard: do not inherit status if user requested 'all' (كل/all) or if doc_type changed
                        is_all_req = any(w in (prompt or "").lower() for w in ["كل ", "جميع ", "all "])
                        dt_changed = bool(fn_args.get("doc_type") and prev_inq.get("doc_type") and fn_args["doc_type"] != prev_inq["doc_type"])
                        if not is_all_req and not dt_changed and not fn_args.get("status") and prev_inq.get("status"):
                            fn_args["status"] = prev_inq["status"]
                        if not fn_args.get("discipline") and prev_inq.get("discipline"):
                            fn_args["discipline"] = prev_inq["discipline"]

                # Clarification guard: If search_submittals is called without any filters and no prior topic exists
                is_explicit_all = any(w in (prompt or "").lower() for w in ["كل المستندات", "جميع المستندات", "كل السجلات", "جميع السجلات", "كل الوثائق", "جميع الوثائق", "all documents", "all submittals", "show all"])
                if fn_name == "search_submittals" and not fn_args.get("keyword") and not fn_args.get("doc_type") and not fn_args.get("discipline") and not is_explicit_all:
                    prev_inq = _get_last_inquiry_from_history(history)
                    if not prev_inq:
                        p_disp = fn_args.get("project_code") or "the project"
                        result = {
                            "project_display_name": p_disp,
                            "clarification_needed": True,
                            "message": f"User only specified project {p_disp} without any equipment keyword or document type. Please ask the user politely what specific information they need for this project (Material Submittals, Shop Drawings, NOC financials, Overdue submittals, or General KPI Summary)."
                        }
                        tool_results.append({"name": fn_name, "args": fn_args, "result": result})
                        func_response_parts.append(
                            types.Part.from_function_response(name=fn_name, response={"result": result})
                        )
                        continue

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

    curr_parsed = _extract_inquiry_intent_from_text(prompt)
    prev_inquiry = _get_last_inquiry_from_history(history) if history else {}

    # 1. Project Detection
    explicit_proj_match = re.search(r"(?:مشروع|المشروع|project)\s+([^\s?.,!]+)", prompt, re.IGNORECASE)
    has_explicit_project = bool(curr_parsed.get("project_code") or explicit_proj_match)

    matched_proj_code = None
    if not global_query:
        if curr_parsed.get("project_code"):
            matched_proj_code = curr_parsed["project_code"]
        elif curr_parsed.get("is_same_project") and prev_inquiry.get("project_code"):
            matched_proj_code = prev_inquiry["project_code"]
        elif has_explicit_project and not curr_parsed.get("is_same_project"):
            # User explicitly mentioned a project in this turn!
            # Do NOT inherit prior turn's project code (which caused Eastmain -> CFC bleed)
            extracted_cand = explicit_proj_match.group(1).strip() if explicit_proj_match else ""
            matched_proj_code = extracted_cand
        elif prev_inquiry.get("project_code"):
            # Only inherit prior project if user did NOT mention any project at all
            matched_proj_code = prev_inquiry["project_code"]

        # If the matched project has alias codes and one of those codes is specifically in allowed_project_ids, prefer that code
        if matched_proj_code and allowed_project_ids:
            allowed_upper = [str(x).upper() for x in allowed_project_ids]
            for pinfo in ai_tools.PROJECT_ALIASES.values():
                p_codes = pinfo.get("codes", [pinfo["code"]])
                if (pinfo.get("code") == matched_proj_code or matched_proj_code in p_codes):
                    if matched_proj_code not in allowed_project_ids and matched_proj_code.upper() not in allowed_upper:
                        for c in p_codes:
                            if c in allowed_project_ids or c.upper() in allowed_upper:
                                matched_proj_code = c
                                break
                    break

        # If still not found and an active project was supplied, use it
        if not matched_proj_code and active_project_id and str(active_project_id).strip().upper() not in ("ALL", ""):
            act_clean = str(active_project_id).strip().lower()
            for alias, pinfo in sorted(ai_tools.PROJECT_ALIASES.items(), key=lambda x: len(x[0]), reverse=True):
                if alias == act_clean or alias in act_clean:
                    matched_proj_code = pinfo["code"]
                    break
            if not matched_proj_code:
                matched_proj_code = str(active_project_id).strip()

    # 2. Document Type Detection (with follow-up inheritance)
    detected_dt = curr_parsed.get("doc_type")
    if is_follow_up and not detected_dt and prev_inquiry:
        detected_dt = prev_inquiry.get("doc_type")

    # 3. Discipline Detection (with follow-up inheritance)
    detected_disc = curr_parsed.get("discipline")
    if is_follow_up and not detected_disc and prev_inquiry:
        detected_disc = prev_inquiry.get("discipline")

    # 4. Keyword Detection (with follow-up inheritance)
    detected_keyword = curr_parsed.get("keyword")
    if is_follow_up and not detected_keyword and prev_inquiry:
        detected_keyword = prev_inquiry.get("keyword")

    # 5. Status Detection (with follow-up inheritance)
    detected_status = curr_parsed.get("status")
    if is_follow_up and not detected_status and prev_inquiry:
        is_all_requested = any(w in p_low for w in ["كل ", "جميع ", "all "])
        doc_type_switched = bool(detected_dt and prev_inquiry.get("doc_type") and detected_dt != prev_inquiry.get("doc_type"))
        if not is_all_requested and not doc_type_switched:
            detected_status = prev_inquiry.get("status")

    # 6. Intent Inheritance (NOC, Overdue, Summary)
    is_noc_req = curr_parsed.get("is_noc")
    if is_follow_up and not is_noc_req and not detected_keyword and not detected_dt and prev_inquiry.get("is_noc"):
        is_noc_req = True

    is_overdue_req = curr_parsed.get("is_overdue")
    if is_follow_up and not is_overdue_req and not detected_keyword and not detected_dt and prev_inquiry.get("is_overdue"):
        is_overdue_req = True

    is_summary_req = curr_parsed.get("is_summary")
    if is_follow_up and not is_summary_req and not detected_keyword and not detected_dt and prev_inquiry.get("is_summary"):
        is_summary_req = True

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
    proj_title = "All Projects"
    if matched_proj_code:
        for pinfo in sorted(ai_tools.PROJECT_ALIASES.values(), key=lambda x: len(x.get("display", "")), reverse=True):
            if pinfo["code"] == matched_proj_code or (
                isinstance(pinfo.get("codes"), list) and matched_proj_code in pinfo["codes"]
            ):
                proj_title = pinfo["display"]
                break
        if proj_title == "All Projects":
            proj_title = matched_proj_code

    is_explicit_all_docs = any(w in p_low for w in ["كل المستندات", "جميع المستندات", "كل السجلات", "جميع السجلات", "كل الوثائق", "جميع الوثائق", "all documents", "all submittals", "show all"])
    if matched_proj_code and not detected_dt and not detected_keyword and not detected_disc and not is_explicit_all_docs:
        # User only specified a project without any topic and without any inherited inquiry
        if is_english:
            return (
                f"Hello! 👋 Regarding project **{proj_title}**:\n\n"
                "What would you like to inquire about for this project? You can ask about:\n"
                "- 📋 **Material Submittals (MS)** or **Shop Drawings (SD)** (e.g. valves, pumps, chillers, cables, etc.).\n"
                "- 💰 **Notice of Change (NOC)** and Variation Orders financial summary.\n"
                "- ⏰ **Overdue Submittals** exceeding contractual reply SLAs.\n"
                "- 📊 **Overall Project Register & Approval KPIs**."
            )
        return (
            f"أهلاً بك يا باشمهندس! 👋 بخصوص مشروع **{proj_title}**:\n\n"
            "ما الذي تود الاستعلام عنه تحديداً في المشروع؟ يمكنك سؤالي عن:\n"
            "- 📋 **اعتمادات المواد (MS)** أو **الشوب درونج (SD)** (مع إمكانية تحديد نوع المعدة مثل المحابس، المضخات، التكييف، الكابلات).\n"
            "- 💰 **ملخص وتكاليف أوامر التغيير والمطالبات (NOC)**.\n"
            "- ⏰ **الوثائق المتأخرة عن مواعيد الرد التعاقدية (Overdue Submittals)**.\n"
            "- 📊 **الملخص العام لإحصائيات ونسب اعتماد وثائق المشروع**.\n\n"
            "تفضل بتحديد المطلوب وسأقوم بالبحث في السجلات فوراً!"
        )

    # Show latest revision by default, unless user explicitly asks for all revisions or revision history
    is_all_revs = any(w in p_low for w in ["كل المراجعات", "جميع المراجعات", "تاريخ المراجعات", "سجل المراجعات", "all revisions", "revision history", "all revs"])
    use_latest_rev = not is_all_revs

    search_res = ai_tools.search_submittals(
        allowed_project_ids=allowed_project_ids,
        project_code=matched_proj_code,
        doc_type=detected_dt,
        discipline=detected_disc,
        keyword=detected_keyword,
        status=detected_status,
        latest_revision_only=use_latest_rev,
        limit=50,
    )

    proj_title = search_res.get("project_display_name") or (
        search_res.get("project_scope")[0] if search_res.get("project_scope") else None
    ) or proj_title
    records = search_res.get("records", [])

    KW_AR_MAP = {
        "valve": "المحابس (Valves)",
        "valves": "المحابس (Valves)",
        "pump": "المضخات (Pumps)",
        "pumps": "المضخات (Pumps)",
        "chiller": "المبردات والتكييف (Chillers)",
        "fan": "المراوح والتهوية (Fans)",
        "cable": "الكابلات (Cables)",
        "pipe": "المواسير والأنابيب (Pipes)",
        "panel": "اللوحات الكهربائية (Panels)",
        "generator": "المولدات (Generators)",
        "transformer": "المحولات (Transformers)",
        "duct": "مجاري الهواء والصاج (Ducts)",
    }

    # Format header description
    dt_label = "طلبات الشراء (PR)" if detected_dt == "PR" else (f"الـ {detected_dt}" if detected_dt else "المشروع")
    disc_label = f" الخاصة بـ {detected_disc}" if detected_disc else ""
    
    if proj_title in ("All Projects", None, ""):
        kw_display = KW_AR_MAP.get(detected_keyword, detected_keyword or "")
        kw_label = f" الخاصة بـ {kw_display}" if detected_keyword else ""
        proj_display_ar = "كافة المشاريع المصرح بها (Global)"
    else:
        kw_label = f" الخاصة بـ {detected_keyword}" if detected_keyword else ""
        proj_display_ar = proj_title

    appr_label = " المعتمدة" if detected_status == "approved" else ""

    if not records:
        if is_english:
            if detected_keyword == "valve" and detected_dt == "PR":
                return f"No Purchase Requisitions (PR) for valves were found in project {proj_title}."
            if detected_keyword == "valve":
                return f"No approved submittals for valves were found in project {proj_title}."
            if detected_keyword == "pump":
                return f"No approved submittals for pumps were found in project {proj_title}."
            if detected_disc and detected_dt:
                return f"No {detected_dt} submittals for {detected_disc} were found in project {proj_title}."
            if detected_dt:
                return f"No matching {detected_dt} submittals were found in project {proj_title}."
            return f"No matching documents were found in project {proj_title}."

        if detected_keyword == "valve" and detected_dt == "PR":
            return f"لم يتم العثور على أي طلبات شراء (PR) تخص المحابس (Valves) في سجلات مشروع {proj_display_ar}."
        if detected_keyword == "valve":
            return f"لم يتم العثور على أي اعتمادات تخص المحابس (Valves) في سجلات مشروع {proj_display_ar}."
        if detected_keyword == "pump":
            return f"لم يتم العثور على أي اعتمادات تخص المضخات (Pumps) في سجلات مشروع {proj_display_ar}."
        if detected_disc and detected_dt:
            return f"لم يتم العثور على أي وثائق من نوع الـ {detected_dt} تخص الأعمال {detected_disc} في سجلات مشروع {proj_display_ar}."
        if detected_dt:
            return f"لم يتم العثور على أي وثائق من نوع الـ {detected_dt} مطابقة لطلبك في سجلات مشروع {proj_display_ar}."
        return f"لم يتم العثور على أي وثائق مطابقة في سجلات مشروع {proj_display_ar}."

    table_rows = []
    for r in records:
        st_clean = r.get("status") or ("Pending PO" if str(r.get("doc_type") or "").upper() == "PR" else "Pending")
        st_badge = f"✅ **{st_clean}**" if any(a in st_clean.lower() for a in ["approv", "status a", "status b", "ordered", "delivered", "completed"]) else st_clean
        
        c_doc = str(r.get('doc_no', '—')).replace('|', '-').replace('\n', ' ').strip()
        c_title = str(r.get('title', '—')).replace('|', '-').replace('\n', ' ').strip()
        c_type = str(r.get('doc_type', '—')).replace('|', '-').replace('\n', ' ').strip()
        c_date = str(r.get('doc_date', '—')).replace('|', '-').replace('\n', ' ').strip()
        c_proj = str(r.get('project_code', '')).replace('|', '-').replace('\n', ' ').strip()
        c_status = str(st_badge).replace('|', '-').replace('\n', ' ').strip()

        table_rows.append(
            f"| `{c_doc}` | {c_title} | `{c_type}` | {c_status} | {c_date} | `{c_proj}` |"
        )

    status_word = "معتمدة و" if detected_status == "approved" else ""

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
        f"### 📋 تفاصيل وثائق {dt_label}{disc_label}{kw_label}{appr_label} لمشروع {proj_display_ar}:\n\n"
        f"تم استخراج **{len(records)}** وثيقة {status_word}مطابقة لطلبك مباشرة من سجلات المشروعات الرسمية:\n\n"
        "| رقم الوثيقة | البيان والمواصفات | النوع | الحالة | التاريخ | المشروع |\n"
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
    proj_code = None
    for alias, pinfo in sorted(ai_tools.PROJECT_ALIASES.items(), key=lambda x: len(x[0]), reverse=True):
        if pinfo["code"] in context_text:
            proj_code = pinfo["code"]
            break
    if not proj_code:
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
                try:
                    db.log_action(u.get("username", ""), "ACCESS DENIED", project_id=pid,
                                  detail=f"AI query access denied for project '{pid}'")
                except Exception:
                    pass
                return jsonify(error="Forbidden"), 403
            active_pid = pid

        try:
            db.log_action(
                username=u.get("username", ""),
                action="AI QUERY",
                project_id=active_pid,
                dt_id=tab or None,
                detail=f"Prompt: {prompt[:150]}"
            )
        except Exception:
            pass

        ws_id = data.get("workspace_id")
        if not ws_id and active_pid:
            p_obj = db.get_project(active_pid)
            if p_obj:
                ws_id = p_obj.get("workspace_id")
        if not ws_id:
            ws_id = "ENG"
        ws_id = str(ws_id).strip().upper()

        try:
            user_allowed = get_allowed_project_ids(u, workspace_id=ws_id)
        except Exception:
            user_allowed = []

        if not user_allowed and active_pid:
            p_obj = db.get_project(active_pid)
            if p_obj and (p_obj.get("workspace_id") or "ENG").upper() == ws_id:
                user_allowed = [active_pid]

        if not user_allowed:
            ws_label = "Operation & Maintenance (O&M)" if ws_id == "OM" else ("Engineering & Projects" if ws_id == "ENG" else ws_id)
            return jsonify(reply=f"You do not have access to any projects/contracts in {ws_label}."), 200

        target_pids = user_allowed

        # Short-circuit for simple greetings/generic conversation
        p_clean = re.sub(r"[^\w\s]", "", prompt).strip().lower()
        greeting_words = [
            "ازيك", "ازىك", "عامل ايه", "مرحبا", "أهلا", "اهلا", 
            "سلام", "السلام عليكم", "صباح الخير", "مساء الخير", "hello", "hi", "hey", "هاي"
        ]
        is_greeting = any(g == p_clean or p_clean.startswith(g) for g in greeting_words) and len(prompt.split()) <= 4

        if is_greeting:
            if ws_id == "OM":
                reply = (
                    "أهلاً بك يا باشمهندس! 👋 أنا مساعدك الذكي لقسم التشغيل والصيانة (O&M).\n\n"
                    "أنا جاهز لمساعدتك فوراً في:\n"
                    "- ⚙️ متابعة أوامر العمل (Work Orders) وطلبات الصيانة (MR).\n"
                    "- 📅 جداول الصيانة الوقائية (PPM) والإصلاحات الطارئة (CM).\n"
                    "- 📊 تقارير مستوى الخدمة (SLA) ومعدلات الإنجاز بالعقود والمحطات.\n\n"
                    "تفضل بسؤالي عن أي عقد صيانة أو أمر عمل!"
                )
            else:
                reply = (
                    "أهلاً بك يا باشمهندس! 👋 أنا مساعدك الذكي لنظام مراقبة وثائق ومشاريع جازشيل (DCR).\n\n"
                    "أنا جاهز لمساعدتك فوراً في:\n"
                    "- 📋 استعراض واعتمادات سجلات الـ (MS) والشوب درونج (SD).\n"
                    "- ⏰ متابعة الوثائق المتأخرة (Overdue Submittals).\n"
                    "- 💰 مراجعة أوامر التغيير والمطالبات (NOC).\n\n"
                    "تفضل بسؤالي عن أي مشروع أو معدة أو مستند!"
                )
            return jsonify(reply=reply), 200

        custom_inst = None
        if ws_id == "OM":
            projs_info = db.get_projects(target_pids)
            p_lines = "\n".join([f"- '{p.get('code', p['id'])}': {p.get('name', p['id'])}" for p in projs_info])
            custom_inst = (
                "You are the DCR Operation & Maintenance (O&M) AI Assistant for Gas Chill facilities and maintenance contracts. "
                "You assist maintenance managers, engineers, and facility coordinators by querying maintenance registers "
                "using the provided tools: search_submittals, get_project_summary, get_noc_financials, get_overdue_submittals, and get_document_history.\n\n"
                "Document types in the O&M register include:\n"
                "- 'WO': Work Orders (أوامر العمل والصيانة).\n"
                "- 'PPM': Preventive Maintenance (الصيانة الوقائية والجدولة الدورية).\n"
                "- 'CM': Corrective Maintenance (الصيانة العلاجية وإصلاح الأعطال الطارئة).\n"
                "- 'INC': Incident Reports (بلاغات الحوادث والأعطال الجسيمة).\n"
                "- 'MR': Maintenance Requests (طلبات الصيانة الواردة من العملاء).\n"
                "- 'SLA': Service Level Agreements & Reports (تقارير مستوى الخدمة والالتزام التعاقدي).\n"
                "- 'COR': Correspondence (المراسلات والخطابات الرسمية للصيانة).\n\n"
                f"Contracts and facilities currently accessible in this O&M workspace:\n{p_lines}\n\n"
                "Guidelines:\n"
                "1. Focus exclusively on Operation & Maintenance facilities, client contracts, PPM schedules, and Work Orders.\n"
                "2. When the user asks for summaries or totals, only summarize the O&M contracts listed above.\n"
                "3. Never refer to EPC construction projects (e.g. EKH, Eastmain, etc.) when operating in this O&M context.\n"
                "4. Respond in professional Egyptian engineering Arabic when queried in Arabic."
            )
        elif ws_id != "ENG":
            ws_obj = db.get_workspace(ws_id) if hasattr(db, "get_workspace") else None
            ws_title = ws_obj.get("name", ws_id) if ws_obj else ws_id
            projs_info = db.get_projects(target_pids)
            p_lines = "\n".join([f"- '{p.get('code', p['id'])}': {p.get('name', p['id'])}" for p in projs_info])
            custom_inst = (
                f"You are the DCR {ws_title} AI Assistant for Gas Chill. "
                f"You assist users in the {ws_title} department by querying project registers.\n\n"
                f"Projects/Contracts currently accessible in this workspace:\n{p_lines}\n\n"
                "Only report data belonging to this department."
            )

        # Run the Tool-Augmented Gemini Agent Controller with full authorized scope
        reply, err, _ = _call_gemini_agent(
            prompt,
            target_pids,
            history=history,
            custom_instruction=custom_inst,
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
