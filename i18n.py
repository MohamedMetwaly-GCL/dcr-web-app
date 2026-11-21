"""i18n.py - DCR Bilingual (English LTR / Arabic RTL) Internationalization Engine.

Zero-migration localization dictionary, request language resolution, BiDi/RTL helpers,
and client-side dictionary serialization for the DCR platform.

Invariants:
- Database values (statuses, roles, column keys, LTR directions) remain canonical English.
- Western Arabic numerals (0-9) are preserved in both English and Arabic modes.
"""
import json
from flask import request as _flask_request, has_request_context

SUPPORTED_LANGS = ("en", "ar")
DEFAULT_LANG = "en"
LANG_COOKIE_NAME = "dcr_lang"
LANG_COOKIE_MAX_AGE = 31536000  # 1 year


def normalize_lang(raw) -> str:
    """Return 'ar' or 'en' strictly; fall back to DEFAULT_LANG ('en')."""
    if not raw or not isinstance(raw, str):
        return DEFAULT_LANG
    val = raw.strip().lower()
    if val in SUPPORTED_LANGS:
        return val
    return DEFAULT_LANG


def get_lang(req=None) -> str:
    """Resolve active UI language with priority: ?lang= query param > dcr_lang cookie > 'en'."""
    target_req = req
    if target_req is None and has_request_context():
        target_req = _flask_request
    if target_req is None:
        return DEFAULT_LANG

    try:
        args = getattr(target_req, "args", None)
        if args:
            q_lang = args.get("lang")
            if q_lang and isinstance(q_lang, str) and q_lang.strip().lower() in SUPPORTED_LANGS:
                return q_lang.strip().lower()
    except Exception:
        pass

    try:
        cookies = getattr(target_req, "cookies", None)
        if cookies:
            c_lang = cookies.get(LANG_COOKIE_NAME)
            if c_lang and isinstance(c_lang, str) and c_lang.strip().lower() in SUPPORTED_LANGS:
                return c_lang.strip().lower()
    except Exception:
        pass

    return DEFAULT_LANG


def is_rtl(lang: str) -> bool:
    """Return True when language uses Right-to-Left layout ('ar')."""
    return normalize_lang(lang) == "ar"


def html_dir(lang: str) -> str:
    """Return 'rtl' for Arabic and 'ltr' for English."""
    return "rtl" if is_rtl(lang) else "ltr"


TRANSLATIONS = {
    "en": {
        # ── Language Switcher ──
        "lang.current": "en",
        "lang.dir": "ltr",
        "lang.switch_label": "عربي",
        "lang.switch_target": "ar",
        "lang.switch_title": "التبديل إلى اللغة العربية (RTL)",

        # ── Global Topbar & Shared Chrome ──
        "app.title_full": "Document Control Register",
        "app.title_short": "DCR",
        "topbar.workspaces": "Workspaces",
        "topbar.workspaces_title": "Return to Company Workspaces Launchpad",
        "topbar.active_workspace": "Active Workspace: {name}",
        "topbar.enterprise_workspaces": "Enterprise Workspaces",
        "topbar.global_search_placeholder": "Search all projects...",
        "topbar.global_search_title": "Global Search (Ctrl+K)",
        "topbar.ai_btn": "AI",
        "topbar.ai_title": "DCR AI Assistant (Gemini Flash)",
        "topbar.about_btn": "About",
        "topbar.about_title": "About System",
        "topbar.dark_toggle_title": "Toggle dark mode",
        "topbar.admin_btn": "Admin",
        "topbar.change_pw_title": "Change Password",
        "topbar.logout_title": "Logout",
        "topbar.login_btn": "Login",
        "topbar.dashboard_btn": "Dashboard",

        # ── Roles ──
        "role.superadmin": "SUPER ADMIN",
        "role.admin": "ADMIN",
        "role.dept_admin": "DEPARTMENT ADMIN",
        "role.editor": "EDITOR",
        "role.dc_editor": "DOCUMENT CONTROL • EDITOR",
        "role.viewer": "VIEWER",
        "role.dc_viewer": "DOCUMENT CONTROL • VIEWER",
        "role.guest": "GUEST",

        # ── Global Footer ──
        "footer.designed_by": "Designed & Developed by",
        "footer.developer_name": "Mohamed Metwaly",
        "footer.department": "Document Control Department",
        "footer.inhouse_by": "Designed & Developed in-house by Mohamed Metwaly",

        # ── About System Modal ──
        "about.header_prefix": "About",
        "about.header_suffix": "System",
        "about.system_title": "Document Control Register",
        "about.release_badge": "v1.0.0 — Stable Release",
        "about.purpose_title": "🎯 PURPOSE & BUSINESS IMPACT",
        "about.purpose_body": (
            "An enterprise-grade Document Management System developed in-house to digitize engineering workflows, "
            "ensure zero data loss, and accelerate the decision-making process. Built to eliminate manual tracking "
            "and centralize document control for Project Management and Operations teams."
        ),
        "about.tech_title": "⚙️ CORE TECHNOLOGIES",
        "about.tech_db_label": "Database:",
        "about.tech_db_val": "PostgreSQL with dynamic JSONB architecture for high scalability.",
        "about.tech_backend_label": "Backend Engine:",
        "about.tech_backend_val": "Python / Flask.",
        "about.tech_frontend_label": "Frontend:",
        "about.tech_frontend_val": "Vanilla JavaScript & CSS (Optimized for speed and zero latency).",
        "about.credits_title": "👨‍💻 CREDITS",
        "about.credits_role": "Architect & Full-Stack Developer:",
        "about.close_btn": "Close",

        # ── Login Page (/login) ──
        "login.page_title": "Gas Chill DCR — Sign In",
        "login.tagline": "Energy. Redefined.",
        "login.platform_desc": (
            "Enterprise Document Control System — Unified hub managing company workspaces, "
            "registers, and submittals with precision."
        ),
        "login.stat_platform": "Platform",
        "login.stat_version": "Version",
        "login.stat_access": "Access",
        "login.left_footer": "Gas Chill • A Valmore Holding Company",
        "login.welcome_title": "Welcome Back",
        "login.welcome_sub": "Sign in to your workspace to continue",
        "login.username_label": "Username",
        "login.username_placeholder": "Enter your username",
        "login.password_label": "Password",
        "login.password_placeholder": "Enter your password",
        "login.signin_btn": "Sign In",
        "login.form_footer": "Gas Chill Document Control System — Secure Access",
        "login.err_missing": "Please enter username and password",
        "login.err_format": "Invalid username format: letters, numbers, dots, hyphens, and underscores only (3-40 chars).",
        "login.err_invalid": "Invalid credentials",
        "login.err_connection": "Connection error. Please try again.",

        # ── Launchpad Hub (/hub) ──
        "hub.page_title": "Gas Chill — Company Workspaces Launchpad",
        "hub.subtitle_pill": "🏢 Enterprise Document Control Hub",
        "hub.admin_btn": "Admin",
        "hub.admin_title": "Open Enterprise Admin Panel",
        "hub.theme_dark": "Dark Mode",
        "hub.theme_light": "Light Mode",
        "hub.theme_title": "Toggle Light / Dark Mode",
        "hub.theme_switched_dark": "🌙 Switched to Dark Mode",
        "hub.theme_switched_light": "☀️ Switched to Light Mode",
        "hub.signout_btn": "Sign Out",
        "hub.headline_title": "Company Workspaces Launchpad",
        "hub.headline_sub": "Select your authorized workspace to access live engineering registers, contracts, and submittals.",
        "hub.legend_active": "Authorized Workspace (Active)",
        "hub.legend_restricted": "🔒 Restricted (Contact Administrator)",
        "hub.badge_primary": "★ Primary",
        "hub.badge_active": "● Active",
        "hub.badge_locked": "🔒 Locked",
        "hub.btn_enter": "Enter Workspace",
        "hub.btn_restricted": "🔒 Restricted Access",
        "hub.restricted_tooltip": "🔒 Access Restricted: Contact System Administrator to request permissions",
        "hub.footer_left": "Gas Chill — Enterprise Document Control System (DCR)",
        "hub.kpi_projects": "Projects",
        "hub.kpi_contracts": "Contracts",
        "hub.kpi_registers": "Registers",
        "hub.kpi_records": "Records",
        "hub.kpi_documents": "Documents",

        # ── Workspace Names & Descriptions ──
        "ws.ENG.name": "Engineering & Projects",
        "ws.ENG.sub": "EPC & Site Submittals",
        "ws.ENG.desc": "EPC Submittals, Drawings, Material Approvals, IR, MIR, RFI & NOC Variation Orders.",
        "ws.OM.name": "Operation & Maintenance (O&M)",
        "ws.OM.sub": "Facilities & Maintenance Contracts",
        "ws.OM.desc": "Facilities maintenance contracts, Work Orders, PPM schedules and SLA compliance metrics.",
        "ws.HR.name": "Human Resources",
        "ws.HR.sub": "Personnel & Corporate Policies",
        "ws.HR.desc": "Personnel archives, employee onboarding, credential files, attendance and corporate policies.",
        "ws.FIN.name": "Finance & Accounting",
        "ws.FIN.sub": "IPCs, Invoices & Cost Audits",
        "ws.FIN.desc": "Interim Payment Certificates (IPC), subcontractor invoices, project budgets and cost audits.",
        "ws.PROC.name": "Procurement",
        "ws.PROC.sub": "POs, Vendors & Requisitions",
        "ws.PROC.desc": "Purchase Orders (PO), vendor quotations, material requisition packages and approved suppliers.",
        "ws.WH.name": "Warehousing",
        "ws.WH.sub": "Inventory, MRNs & Dispatch",
        "ws.WH.desc": "Central store inventory, Material Receipt Notes (MRN), site dispatch and stock balances.",
        "ws.HSE.name": "Health, Safety & Environment (HSE)",
        "ws.HSE.sub": "Permits, Observations, Incidents & FEI",
        "ws.HSE.desc": "Occupational safety audits, site inspection checklists, incident registers and ISO certifications.",
        "ws.TND.name": "Tendering",
        "ws.TND.sub": "BOQs, Bids & Award Tracking",
        "ws.TND.desc": "Tender submissions, BOQ bill of quantities, tender addendums, commercial bids and award tracking.",

        # ── Common Client UI Actions & Prompts ──
        "common.close": "Close",
        "common.cancel": "Cancel",
        "common.save": "Save",
        "common.done": "Done",
        "common.create": "Create",
        "common.delete": "Delete",
        "common.loading": "Loading...",
        "common.forbidden": "Forbidden",
        "common.error": "Error",
        "common.new_password_prompt": "New password (min 4 chars):",
        "common.password_changed": "✔ Password changed",
        "common.project_info": "Project Info",
        "common.hide_info": "Hide Info",

        # ── Global Spotlight Search (SPOTLIGHT_MODAL_COMPONENT) ──
        "spotlight.input_placeholder": "Search across all accessible projects (min. 2 chars)...",
        "spotlight.clear_title": "Clear search",
        "spotlight.close_title": "Close (Esc)",
        "spotlight.filter_project_title": "Filter by Project",
        "spotlight.all_projects": "All Projects",
        "spotlight.all_contracts_facilities": "All Contracts / Facilities",
        "spotlight.filter_status_title": "Filter by Status",
        "spotlight.all_statuses": "All Statuses",
        "spotlight.status_approved": "Approved",
        "spotlight.status_under_review": "Under Review",
        "spotlight.status_revise": "Revise & Resubmit",
        "spotlight.status_rejected": "Rejected",
        "spotlight.date_range_title": "Date Range",
        "spotlight.date_from_title": "From Date",
        "spotlight.date_to_title": "To Date",
        "spotlight.reset_btn": "↺ Reset",
        "spotlight.reset_title": "Reset all filters",
        "spotlight.scope_rbac": "🔒 RBAC Scope",
        "spotlight.scope_checking": "Checking permissions...",
        "spotlight.scope_admin": "👑 Admin Scope",
        "spotlight.scope_admin_text": "Searching across all projects (Admin)",
        "spotlight.scope_assigned": "🔒 Assigned Scope",
        "spotlight.scope_assigned_one": "Searching in your 1 assigned project",
        "spotlight.scope_assigned_many": "Searching in your {count} assigned projects",
        "spotlight.scope_none": "No projects assigned (Results restricted)",
        "spotlight.count_zero": "0 found",
        "spotlight.count_records_found": "{count} records found",
        "spotlight.count_of_total_found": "{shown} of {total} records found",
        "spotlight.count_dt_found": "{count} {dt} records found",
        "spotlight.count_dt_records": "{count} {dt} records",
        "spotlight.chip_all": "All ({count})",
        "spotlight.empty_title": "Instant Global Search",
        "spotlight.empty_sub": "Type at least 2 characters to search across all accessible projects, document numbers, subjects, and statuses...",
        "spotlight.empty_initial_sub": "Type at least 2 characters to search across document numbers, titles, subjects, descriptions, and statuses across your authorized projects.",
        "spotlight.kbd_navigate": "Navigate",
        "spotlight.kbd_open_tab": "Open in new tab",
        "spotlight.kbd_close": "Close",
        "spotlight.footer_rbac": "🔒 RBAC-Filtered Results",
        "spotlight.err_title": "Search Error",
        "spotlight.err_default": "Error executing search",
        "spotlight.no_match_title": "No Matching Records Found",
        "spotlight.no_match_sub": "No records matched \"{query}\" in your accessible projects.",
        "spotlight.untitled_doc": "Untitled Document",
        "spotlight.open_drive_title": "Open attached Drive link",
        "spotlight.submission_date_title": "Submission Date",

        # ── AI Assistant Drawer (AI_DRAWER_COMPONENT) ──
        "ai.fab_title": "DCR AI Assistant (Gemini Flash) — Drag to reposition",
        "ai.fab_aria": "Open AI Assistant",
        "ai.fab_text": "AI",
        "ai.drawer_aria": "DCR AI Assistant",
        "ai.resizer_title": "Drag horizontally to resize width",
        "ai.badge_gemini": "✨ Gemini Flash",
        "ai.badge_readonly": "🔒 Read-Only",
        "ai.title_eng": "DCR Engineering Assistant",
        "ai.title_om": "DCR O&M Assistant",
        "ai.title_hse": "DCR HSE Safety Assistant",
        "ai.title_ws": "DCR {ws} Assistant",
        "ai.scope_detecting": "Scope: Detecting...",
        "ai.scope_active_project": "Active Project: {name}",
        "ai.scope_om": "Scope: O&M Contracts & Facilities",
        "ai.scope_global": "Scope: All Authorized Projects (Global)",
        "ai.btn_expand_title": "Maximize / Restore width",
        "ai.btn_expand_max": "Maximize width (920px)",
        "ai.btn_expand_restore": "Restore normal width",
        "ai.btn_clear_title": "Clear Chat History",
        "ai.btn_close_title": "Close (Esc)",
        "ai.quick_label": "Quick Queries:",
        "ai.chip_eng_overdue": "📌 Overdue Submittals",
        "ai.chip_eng_noc": "💵 Approved NOCs Total",
        "ai.chip_eng_status": "📊 Status Summary",
        "ai.chip_om_summary": "⚡ Maintenance Summary",
        "ai.chip_om_wo": "📌 Open Work Orders",
        "ai.chip_om_ppm": "📊 PPM & SLA Status",
        "ai.chip_hse_kpis": "🛡️ HSE KPIs & Rates",
        "ai.chip_hse_expiring": "⚠️ Expiring PTW / FEI",
        "ai.chip_hse_ncr": "📝 Draft Safety NCR",
        "ai.loading_text": "Querying registers & analyzing with Gemini Flash...",
        "ai.placeholder_eng": "Ask about submittals, NOCs, costs, or document statuses...",
        "ai.placeholder_om": "Ask about work orders, PPM schedules, maintenance requests, SLA...",
        "ai.placeholder_hse": "Ask about HSE KPIs, permits (PTW), fire equipment (FEI), or safety compliance...",
        "ai.send_title": "Send (Enter)",
        "ai.hint_send": "Send",
        "ai.hint_newline": "New line",
        "ai.hint_factual": "Factual responses from register data only",
        "ai.sender_you": "You",
        "ai.sender_gemini": "Gemini",
        "ai.welcome_eng": (
            "Hello! I am your **DCR Engineering AI Assistant** powered by Gemini Flash.\n\n"
            "I have real-time, read-only access to your authorized project registers and can help you with:\n"
            "- 📌 **Overdue submittals** and delay days\n"
            "- 💵 **Notice of Change (NOC)** cost totals and approvals\n"
            "- 📊 **Register status breakdowns** (Approved, Under Review, Rejected)\n"
            "- 🔍 **Specific document numbers** or disciplines\n\n"
            "*How can I assist you today?*"
        ),
        "ai.welcome_om": (
            "Hello! I am your **DCR Operation & Maintenance (O&M) AI Assistant** powered by Gemini Flash.\n\n"
            "I have real-time, read-only access to your maintenance contracts, facilities, and registers:\n"
            "- ⚙️ **Work Orders (WO)** and Maintenance Requests (MR)\n"
            "- 📅 **Preventive Maintenance (PPM)** schedules & Corrective Maintenance (CM)\n"
            "- 📊 **SLA Compliance & Service Reports**\n"
            "- 🔍 **Contract submittals & incident logs**\n\n"
            "*How can I assist you with your maintenance contracts today?*"
        ),
        "ai.welcome_hse": (
            "Hello! I am your **DCR HSE Safety & Compliance Assistant** powered by Gemini Flash.\n\n"
            "I have real-time, RBAC-scoped access to your authorized HSE sites and can help with:\n"
            "- 🛡️ **Safety KPIs & Streaks:** LTI-Free Days, LTIF, TRIR, Severity Rate & Man-Hours\n"
            "- ⚠️ **Permits & Equipment Alerts:** Overdue/expiring PTWs, FEI extinguishers & stale observations\n"
            "- ⚖️ **Compliance & NCR Drafting:** OSHA 29 CFR 1910/1926, NFPA 10/51B & Egyptian Labor Law Decision 211/2003\n\n"
            "*How can I help keep your sites safe today?*"
        ),
        "ai.err_503": "⚠️ **AI Assistant is not configured on this instance.**\n\nPlease configure the `GEMINI_API_KEY` environment variable on Railway to enable Gemini Flash.",
        "ai.err_403": "🔒 **Authentication Required.**\n\nYour session may have expired. Please refresh the page and log in.",
        "ai.err_service": "⚠️ **AI Service Error:** {error}",
        "ai.err_http": "⚠️ **Error ({status}):** {error}",
        "ai.err_server": "Server error",
        "ai.err_empty": "⚠️ Received empty response from AI service.",
        "ai.err_timeout": "⏱️ **Request Timed Out (60s):** Analysis took too long. Please retry with a simpler query.",
        "ai.err_network": "⚠️ **Network Error:** {error}",
        "ai.err_unreachable": "Could not reach AI Assistant endpoint.",

        # ── Admin, Department & Permissions Modals (SHARED_JS) ──
        "admin.modal_title": "User Management & Access Control",
        "admin.rbac_active": "Scoped Multi-Workspace RBAC Active",
        "admin.tab_users": "👥 Users & Access Control",
        "admin.tab_depts": "🏢 Departments & Workspaces",
        "admin.all_depts_opt": "All Departments",
        "admin.all_roles_opt": "All Roles",
        "admin.role_superadmin": "Super Admin",
        "admin.role_admin": "Admin",
        "admin.role_editor": "Editor",
        "admin.role_viewer": "Viewer",
        "admin.search_placeholder": "🔍 Search by user or email...",
        "admin.btn_add_user": "Add User",
        "admin.btn_export_matrix": "Export Matrix",
        "admin.btn_export_matrix_title": "Export Users Matrix to Excel/CSV",
        "admin.btn_gap_audit": "Gap Audit",
        "admin.btn_gap_audit_title": "Detect Sequence Gaps",
        "admin.add_user_title": "➕ Add New User",
        "admin.close_x": "Close ✕",
        "admin.lbl_username": "Username",
        "admin.ph_username": "username",
        "admin.lbl_email": "Email",
        "admin.ph_email": "user@company.com",
        "admin.lbl_department": "Department",
        "admin.lbl_role": "Role",
        "admin.lbl_password": "Password",
        "admin.lbl_downloads": "Downloads",
        "admin.lbl_allow_dl": "Allow 📥",
        "admin.audit_title": "Global Sequence Gap Audit",
        "admin.audit_desc": "Detect missing sequence numbers (gaps) in document logs for any project.",
        "admin.audit_select_proj": "Select Project",
        "admin.audit_select_dt": "Select Doc Type",
        "admin.audit_run_btn": "Run Audit",
        "admin.audit_select_both": "Please select both Project and Document Type.",
        "admin.audit_running": "Auditing document sequences...",
        "admin.audit_missing_title": "⚠️ Missing Sequences Found!",
        "admin.audit_expected_range": "Expected range: {min} to {max}",
        "admin.audit_missing_nums": "Missing numbers: {list}",
        "admin.audit_no_docs": "No numbered documents found to audit.",
        "admin.audit_intact_title": "✅ 100% Intact.",
        "admin.audit_intact_desc": "No missing sequence numbers. Range: {min} to {max}.",
        "admin.audit_failed": "Audit failed: {error}",
        "admin.no_users_match": "No users match the selected filters.",
        "admin.th_user": "User",
        "admin.th_email": "Email",
        "admin.th_department": "Department",
        "admin.th_role": "Role",
        "admin.th_actions": "Actions",
        "admin.badge_all_depts": "🌐 All Departments",
        "admin.btn_edit": "Edit",
        "admin.btn_open": "Open",
        "admin.btn_assign": "➕ Assign",
        "admin.no_depts": "No Departments",
        "admin.manage_depts_title": "Manage department assignments",
        "admin.add_remove_depts_title": "Add or remove departments for {user}",
        "admin.role_all_viewer": "All as Viewer",
        "admin.role_all_editor": "All as Editor",
        "admin.role_all_admin": "All as Admin",
        "admin.role_mixed": "Mixed Roles",
        "admin.badge_multi_dept": "🌐 Multi-Dept",
        "admin.badge_multi_dept_title": "Shared across multiple departments — global account controls are managed by Super Admin",
        "admin.dept_count_one": "1 Dept",
        "admin.dept_count_many": "{count} Depts",
        "admin.ph_add_email": "Add email...",
        "admin.dl_full": "📥 Full",
        "admin.dl_full_title": "Full file download rights included with {role} role",
        "admin.dl_on": "Downloads: ON",
        "admin.dl_on_badge": "📥 Downloads: ON",
        "admin.dl_view_only": "View Only",
        "admin.dl_view_only_badge": "🔒 View Only",
        "admin.dl_allowed_title": "File Downloads: Allowed (Click to Restrict)",
        "admin.dl_restricted_title": "File Downloads: Restricted (Click to Allow)",
        "admin.dl_global_title": "Global download setting managed by Super Admin",
        "admin.lock_protected": "🔒 Protected",
        "admin.lock_you": "👤 You",
        "admin.lock_admin": "🔒 Admin",
        "admin.btn_permissions": "Permissions",
        "admin.btn_permissions_title": "Manage project & contract permissions",
        "admin.btn_reset_pw_title": "Reset Password",
        "admin.btn_delete_user_title": "Delete User",
        "admin.err_open": "Error opening admin: {error}",

        # ── Departments Management Tab & Modal ──
        "dept.enterprise_title": "🏢 Enterprise Departments & Workspaces",
        "dept.enterprise_sub": "Configure company departments, codes, descriptions, and icons across the platform",
        "dept.btn_add": "Add Department",
        "dept.form_add_title": "➕ Add New Department (إضافة قسم جديد)",
        "dept.form_edit_title": "✏️ Edit Department: {name}",
        "dept.btn_save": "Save Department (حفظ القسم)",
        "dept.btn_update": "Update Department (حفظ التعديلات)",
        "dept.btn_cancel": "Cancel (إلغاء)",
        "dept.lbl_code": "Code (كود القسم)",
        "dept.ph_code": "e.g. O&M, HSE",
        "dept.lbl_name_en": "Name in English (اسم القسم بالإنجليزي)",
        "dept.ph_name_en": "e.g. Quality Assurance & Control",
        "dept.lbl_name_ar": "Name in Arabic (اسم القسم بالعربي)",
        "dept.ph_name_ar": "e.g. توكيد ومراقبة الجودة",
        "dept.lbl_desc": "Subtitle / Description (التعريف البسيط أسفل اسم القسم في القوائم والشاشات)",
        "dept.ph_desc": "e.g. Facilities & Maintenance Contracts",
        "dept.lbl_icon": "Icon (أيقونة القسم):",
        "dept.icon_hint": "اضغط على أي أيقونة جاهزة أو اكتب أي رمز تريده:",
        "dept.th_workspace": "Department / Workspace",
        "dept.th_code": "Code",
        "dept.th_desc": "Subtitle / Description",
        "dept.th_projects": "Projects / Contracts",
        "dept.th_status": "Status",
        "dept.th_actions": "Actions",
        "dept.count_contracts": "{count} Contracts",
        "dept.count_projects": "{count} Projects",
        "dept.status_active": "Active",
        "dept.move_up_title": "Move Up",
        "dept.move_down_title": "Move Down",
        "dept.edit_title": "Edit Department Details",
        "dept.open_title": "Open Department Dashboard",
        "dept.err_code_name_req": "Please enter Department Code and English Name.",
        "dept.toast_saved": "✔ Department saved successfully!",
        "dept.err_save": "Failed to save department",
        "dept.toast_order_updated": "✔ Department order updated",
        "dept.err_reorder": "Error reordering: {error}",
        "dept.user_modal_title": "Department Access & Roles — {user}",
        "dept.user_modal_sub": "Assign company departments and define access level for this user",
        "dept.company_depts_count": "Company Departments ({count})",
        "dept.lbl_role_colon": "Role:",
        "dept.btn_done": "Done (إتمام)",

        # ── Scoped Permissions Modal ──
        "perm.modal_title_default": "⚙ User Permissions",
        "perm.modal_sub_default": "Department Projects & Contracts",
        "perm.modal_title_user": "⚙️ Project / Contract Permissions —",
        "perm.search_placeholder": "Search projects/contracts...",
        "perm.select_all": "Select All",
        "perm.clear_all": "Clear All",
        "perm.status_immediate": "Permissions are applied immediately.",
        "perm.status_saving": "Saving...",
        "perm.status_saved": "✔ Changes saved.",
        "perm.status_failed": "Failed to save.",
        "perm.dl_full_role": "📥 Full File Downloads ({role})",
        "perm.dl_allow_chk": "📥 Allow File Downloads",
        "perm.dl_shared_suffix": "(Shared Account)",
        "perm.scoped_strictly": "Scoped strictly to user's assigned department(s)",
        "perm.no_projects": "No projects or contracts found for this department.",
        "perm.th_project": "Project / Contract",
        "perm.th_access": "Access",
        "perm.th_doc_types": "Document Types",
        "perm.client_prefix": "Client:",
        "perm.dt_all": "All",
        "perm.dt_none": "0 types (None)",
        "perm.dt_count": "{count} types",
        "perm.allowed_dts_for": "Allowed Document Types for {name}:",
        "perm.loading_dts": "Loading document types...",
        "perm.no_dts_configured": "No document types configured for this project.",

        # ── Admin Action Toasts & Prompts ──
        "admin.toast_user_not_found": "User not found",
        "admin.toast_all_roles_updated": "✔ All department roles updated to {role}",
        "admin.toast_dl_enabled": "📥 File downloads enabled for {user}",
        "admin.toast_dl_restricted": "🔒 File downloads restricted for {user}",
        "admin.err_dl_update": "Failed to update download permission",
        "admin.toast_role_updated_ws": "✔ Role updated for {ws}",
        "admin.toast_role_updated": "✔ Role updated",
        "admin.err_role_update": "Failed to update role",
        "admin.toast_assigned_all_depts": "✔ Assigned to All Departments",
        "admin.toast_reset_default_dept": "✔ Reset to default department",
        "admin.toast_min_one_dept": "User must belong to at least one department",
        "admin.toast_dept_updated": "✔ Department updated",
        "admin.toast_dept_updated_to": "✔ Department updated to {dept}",
        "admin.err_dept_update": "Failed to update department",
        "admin.err_dt_perm_save": "Error saving document type permissions: {error}",
        "admin.err_perm_update": "Failed to update permission: {error}",
        "admin.err_user_pw_req": "Username and password required",
        "admin.toast_user_added": "✔ User added successfully",
        "admin.err_user_add": "Failed to add user",
        "admin.toast_email_updated": "✔ Email updated",
        "admin.err_email_update": "Failed to update email",
        "admin.confirm_delete_user": "Are you sure you want to delete user \"{user}\"?\nThis cannot be undone.",
        "admin.toast_user_deleted": "✔ User deleted",
        "admin.err_user_delete": "Failed to delete user",
        "admin.prompt_new_pw_user": "Set new password for \"{user}\" (min 4 characters):",
        "admin.err_pw_change": "Error changing password",
        "admin.err_no_export_data": "No user data to export",
        "admin.toast_csv_exported": "✔ Users Matrix exported to CSV",
        "admin.err_export": "Export failed: {error}",
    },
    "ar": {
        # ── Language Switcher ──
        "lang.current": "ar",
        "lang.dir": "rtl",
        "lang.switch_label": "EN",
        "lang.switch_target": "en",
        "lang.switch_title": "Switch to English (LTR)",

        # ── Global Topbar & Shared Chrome ──
        "app.title_full": "سجل مراقبة الوثائق والمستندات",
        "app.title_short": "DCR",
        "topbar.workspaces": "الأقسام",
        "topbar.workspaces_title": "العودة إلى منصة أقسام الشركة",
        "topbar.active_workspace": "القسم النشط: {name}",
        "topbar.enterprise_workspaces": "أقسام وإدارات الشركة",
        "topbar.global_search_placeholder": "البحث في جميع المشروعات...",
        "topbar.global_search_title": "البحث الشامل (Ctrl+K)",
        "topbar.ai_btn": "الذكاء الاصطناعي",
        "topbar.ai_title": "مساعد DCR الذكي (Gemini Flash)",
        "topbar.about_btn": "عن النظام",
        "topbar.about_title": "معلومات عن النظام",
        "topbar.dark_toggle_title": "تبديل الوضع الليلي / النهاري",
        "topbar.admin_btn": "الإدارة",
        "topbar.change_pw_title": "تغيير كلمة المرور",
        "topbar.logout_title": "تسجيل الخروج",
        "topbar.login_btn": "تسجيل الدخول",
        "topbar.dashboard_btn": "لوحة المؤشرات",

        # ── Roles ──
        "role.superadmin": "مدير النظام العام",
        "role.admin": "مدير قسم",
        "role.dept_admin": "مدير الإدارة / القسم",
        "role.editor": "محرر",
        "role.dc_editor": "مراقبة الوثائق • محرر",
        "role.viewer": "مُطّلع",
        "role.dc_viewer": "مراقبة الوثائق • مُطّلع",
        "role.guest": "زائر",

        # ── Global Footer ──
        "footer.designed_by": "تصميم وتطوير",
        "footer.developer_name": "محمد متولي (Mohamed Metwaly)",
        "footer.department": "إدارة مراقبة الوثائق والمستندات (Document Control)",
        "footer.inhouse_by": "تصميم وتطوير داخلي بواسطة Mohamed Metwaly",

        # ── About System Modal ──
        "about.header_prefix": "عن",
        "about.header_suffix": "النظام",
        "about.system_title": "نظام سجل مراقبة الوثائق (DCR)",
        "about.release_badge": "v1.0.0 — الإصدار المستقر",
        "about.purpose_title": "🎯 الهدف والأثر التشغيلي",
        "about.purpose_body": (
            "نظام مؤسسي متكامل لإدارة ومراقبة الوثائق تم تطويره داخلياً لرقمنة مسارات العمل الهندسية والتشغيلية، "
            "وضمان عدم فقدان البيانات، وتسريع اتخاذ القرار. صُمم النظام لإلغاء التتبع اليدوي وتوحيد سجلات مراقبة "
            "المستندات لفرق إدارة المشروعات والتشغيل والصيانة والسلامة."
        ),
        "about.tech_title": "⚙️ التقنيات الأساسية",
        "about.tech_db_label": "قاعدة البيانات:",
        "about.tech_db_val": "PostgreSQL مع بنية JSONB ديناميكية عالية التوسع.",
        "about.tech_backend_label": "محرك الخادم:",
        "about.tech_backend_val": "Python / Flask.",
        "about.tech_frontend_label": "واجهة المستخدم:",
        "about.tech_frontend_val": "Vanilla JavaScript & CSS (مُحسّنة للسرعة الفائقة وزمن استجابة صفري).",
        "about.credits_title": "👨‍💻 تصميم وتطوير",
        "about.credits_role": "معماري النظام ومطور الواجهات المتكاملة:",
        "about.close_btn": "إغلاق",

        # ── Login Page (/login) ──
        "login.page_title": "Gas Chill DCR — تسجيل الدخول",
        "login.tagline": "طاقة بمعايير متجددة.",
        "login.platform_desc": (
            "النظام المؤسسي لمراقبة الوثائق — منصة موحدة لإدارة أقسام الشركة والسجلات الهندسية "
            "والاعتمادات والمراسلات بدقة متناهية."
        ),
        "login.stat_platform": "المنصة",
        "login.stat_version": "الإصدار",
        "login.stat_access": "الإتاحة",
        "login.left_footer": "Gas Chill • إحدى شركات مجموعة Valmore Holding",
        "login.welcome_title": "مرحباً بعودتك",
        "login.welcome_sub": "سجّل الدخول إلى مساحة العمل الخاصة بك للمتابعة",
        "login.username_label": "اسم المستخدم",
        "login.username_placeholder": "أدخل اسم المستخدم",
        "login.password_label": "كلمة المرور",
        "login.password_placeholder": "أدخل كلمة المرور",
        "login.signin_btn": "تسجيل الدخول",
        "login.form_footer": "نظام مراقبة الوثائق Gas Chill — وصول آمن ومُشفّر",
        "login.err_missing": "يرجى إدخال اسم المستخدم وكلمة المرور",
        "login.err_format": "صيغة اسم المستخدم غير صحيحة: يُسمح بالحروف والأرقام والنقاط والشرطات فقط (3-40 حرفاً).",
        "login.err_invalid": "بيانات الدخول غير صحيحة",
        "login.err_connection": "خطأ في الاتصال بالخادم، يرجى المحاولة مرة أخرى.",

        # ── Launchpad Hub (/hub) ──
        "hub.page_title": "Gas Chill — منصة أقسام وإدارات الشركة",
        "hub.subtitle_pill": "🏢 المنصة المركزية لمراقبة الوثائق",
        "hub.admin_btn": "الإدارة",
        "hub.admin_title": "فتح لوحة إدارة المستخدمين والصلاحيات",
        "hub.theme_dark": "الوضع الليلي",
        "hub.theme_light": "الوضع النهاري",
        "hub.theme_title": "تبديل الوضع الليلي / النهاري",
        "hub.theme_switched_dark": "🌙 تم التبديل إلى الوضع الليلي",
        "hub.theme_switched_light": "☀️ تم التبديل إلى الوضع النهاري",
        "hub.signout_btn": "خروج",
        "hub.headline_title": "منصة أقسام وإدارات الشركة",
        "hub.headline_sub": "اختر القسم المصرّح لك بالدخول إليه لاستعراض السجلات الهندسية والعقود والاعتمادات المباشرة.",
        "hub.legend_active": "قسم مصرّح به (نشط)",
        "hub.legend_restricted": "🔒 مقيّد (تواصل مع مدير النظام)",
        "hub.badge_primary": "★ رئيسي",
        "hub.badge_active": "● نشط",
        "hub.badge_locked": "🔒 مغلق",
        "hub.btn_enter": "دخول القسم",
        "hub.btn_restricted": "🔒 وصول مقيّد",
        "hub.restricted_tooltip": "🔒 الوصول مقيّد: يرجى التواصل مع مدير النظام لطلب صلاحية الدخول",
        "hub.footer_left": "Gas Chill — النظام المؤسسي لمراقبة الوثائق والمستندات (DCR)",
        "hub.kpi_projects": "المشروعات",
        "hub.kpi_contracts": "العقود",
        "hub.kpi_registers": "السجلات",
        "hub.kpi_records": "السجلات",
        "hub.kpi_documents": "الوثائق",

        # ── Workspace Names & Descriptions ──
        "ws.ENG.name": "المشروعات الهندسية",
        "ws.ENG.sub": "اعتمادات المشروعات والموقع (EPC)",
        "ws.ENG.desc": "تقديمات المشروعات، المخططات التنفيذية، اعتمادات المواد، طلبات الفحص والاستلام، الاستفسارات الفنية وأوامر التغيير (NOC).",
        "ws.OM.name": "التشغيل والصيانة (O&M)",
        "ws.OM.sub": "عقود المرافق والصيانة",
        "ws.OM.desc": "عقود صيانة المرافق، أوامر العمل، جداول الصيانة الوقائية الدورية (PPM)، وتقارير مستوى الخدمة (SLA).",
        "ws.HR.name": "الموارد البشرية",
        "ws.HR.sub": "ملفات العاملين والسياسات",
        "ws.HR.desc": "أرشيف شؤون العاملين، إجراءات التعيين، ملفات المؤهلات، الحضور والانصراف واللوائح المؤسسية.",
        "ws.FIN.name": "المالية والحسابات",
        "ws.FIN.sub": "المستخلصات والفواتير والموازنات",
        "ws.FIN.desc": "المستخلصات الجارية (IPC)، فواتير مقاولي الباطن، موازنات المشروعات والمراجعات المالية.",
        "ws.PROC.name": "المشتريات والتعاقدات",
        "ws.PROC.sub": "أوامر الشراء والموردين",
        "ws.PROC.desc": "أوامر الشراء (PO)، عروض أسعار الموردين، حزم طلبات الشراء وقوائم الموردين المعتمدين.",
        "ws.WH.name": "المخازن والمستودعات",
        "ws.WH.sub": "المخزون وأذون الاستلام والصرف",
        "ws.WH.desc": "جرد المخازن المركزية، أذون استلام المواد (MRN)، حركة الصرف للمواقع وأرصدة المخزون.",
        "ws.HSE.name": "السلامة والصحة المهنية والبيئة (HSE)",
        "ws.HSE.sub": "تصاريح العمل، الملاحظات، الحوادث وفحص المعدات",
        "ws.HSE.desc": "تصاريح العمل، ملاحظات السلامة، سجلات الحوادث والإسعافات، ساعات العمل، وفحص طفايات ومعدات الحريق.",
        "ws.TND.name": "العطاءات والمناقصات",
        "ws.TND.sub": "مقايسات الأعمال والعروض الفنية والمالية",
        "ws.TND.desc": "مظاريف العطاءات، جداول الكميات (BOQ)، ملاحق المناقصات، العروض المالية ومتابعة الترسية.",

        # ── Common Client UI Actions & Prompts ──
        "common.close": "إغلاق",
        "common.cancel": "إلغاء",
        "common.save": "حفظ",
        "common.done": "تم",
        "common.create": "إنشاء",
        "common.delete": "حذف",
        "common.loading": "جاري التحميل...",
        "common.forbidden": "غير مصرّح",
        "common.error": "خطأ",
        "common.new_password_prompt": "كلمة المرور الجديدة (4 أحرف على الأقل):",
        "common.password_changed": "✔ تم تغيير كلمة المرور بنجاح",
        "common.project_info": "بيانات المشروع",
        "common.hide_info": "إخفاء البيانات",

        # ── Global Spotlight Search (SPOTLIGHT_MODAL_COMPONENT) ──
        "spotlight.input_placeholder": "ابحث في جميع المشروعات المتاحة (حرفان على الأقل)...",
        "spotlight.clear_title": "مسح البحث",
        "spotlight.close_title": "إغلاق (Esc)",
        "spotlight.filter_project_title": "تصفية حسب المشروع",
        "spotlight.all_projects": "جميع المشروعات",
        "spotlight.all_contracts_facilities": "جميع العقود والمرافق",
        "spotlight.filter_status_title": "تصفية حسب الحالة",
        "spotlight.all_statuses": "جميع الحالات",
        "spotlight.status_approved": "معتمد (Approved)",
        "spotlight.status_under_review": "قيد المراجعة (Under Review)",
        "spotlight.status_revise": "تعديل وإعادة تقديم (Revise & Resubmit)",
        "spotlight.status_rejected": "مرفوض (Rejected)",
        "spotlight.date_range_title": "النطاق الزمني",
        "spotlight.date_from_title": "من تاريخ",
        "spotlight.date_to_title": "إلى تاريخ",
        "spotlight.reset_btn": "↺ إعادة ضبط",
        "spotlight.reset_title": "إعادة ضبط جميع الفلاتر",
        "spotlight.scope_rbac": "🔒 نطاق الصلاحيات",
        "spotlight.scope_checking": "جاري التحقق من الصلاحيات...",
        "spotlight.scope_admin": "👑 صلاحية مدير",
        "spotlight.scope_admin_text": "البحث في جميع المشروعات (صلاحية إدارة)",
        "spotlight.scope_assigned": "🔒 النطاق المخصص",
        "spotlight.scope_assigned_one": "البحث في مشروع واحد مخصص لك",
        "spotlight.scope_assigned_many": "البحث في {count} مشروعات مخصصة لك",
        "spotlight.scope_none": "لا توجد مشروعات مخصصة (النتائج مقيّدة)",
        "spotlight.count_zero": "0 نتيجة",
        "spotlight.count_records_found": "تم العثور على {count} سجل",
        "spotlight.count_of_total_found": "عرض {shown} من أصل {total} سجل",
        "spotlight.count_dt_found": "تم العثور على {count} سجل ({dt})",
        "spotlight.count_dt_records": "{count} سجل ({dt})",
        "spotlight.chip_all": "الكل ({count})",
        "spotlight.empty_title": "البحث الشامل الفوري",
        "spotlight.empty_sub": "اكتب حرفين على الأقل للبحث في جميع المشروعات المتاحة وأرقام الوثائق والموضوعات والحالات...",
        "spotlight.empty_initial_sub": "اكتب حرفين على الأقل للبحث في أرقام الوثائق والعناوين والموضوعات والأوصاف والحالات عبر المشروعات المصرّح لك بها.",
        "spotlight.kbd_navigate": "للتنقل",
        "spotlight.kbd_open_tab": "فتح في علامة تبويب جديدة",
        "spotlight.kbd_close": "إغلاق",
        "spotlight.footer_rbac": "🔒 نتائج مفلترة حسب صلاحيات الوصول (RBAC)",
        "spotlight.err_title": "خطأ في البحث",
        "spotlight.err_default": "حدث خطأ أثناء تنفيذ البحث",
        "spotlight.no_match_title": "لا توجد سجلات مطابقة",
        "spotlight.no_match_sub": "لم يتم العثور على سجلات تطابق \"{query}\" في المشروعات المتاحة لك.",
        "spotlight.untitled_doc": "وثيقة بدون عنوان",
        "spotlight.open_drive_title": "فتح رابط الملف المرفق على Drive",
        "spotlight.submission_date_title": "تاريخ التقديم / الإصدار",

        # ── AI Assistant Drawer (AI_DRAWER_COMPONENT) ──
        "ai.fab_title": "مساعد DCR الذكي (Gemini Flash) — اسحب لتغيير الموضع",
        "ai.fab_aria": "فتح المساعد الذكي",
        "ai.fab_text": "AI",
        "ai.drawer_aria": "مساعد DCR الذكي",
        "ai.resizer_title": "اسحب أفقياً لتغيير عرض النافذة",
        "ai.badge_gemini": "✨ Gemini Flash",
        "ai.badge_readonly": "🔒 قراءة فقط",
        "ai.title_eng": "مساعد DCR الهندسي الذكي",
        "ai.title_om": "مساعد DCR للتشغيل والصيانة",
        "ai.title_hse": "مساعد DCR للسلامة والصحة المهنية",
        "ai.title_ws": "مساعد DCR الذكي ({ws})",
        "ai.scope_detecting": "النطاق: جاري التحديد...",
        "ai.scope_active_project": "المشروع النشط: {name}",
        "ai.scope_om": "النطاق: عقود ومرافق التشغيل والصيانة",
        "ai.scope_global": "النطاق: جميع المشروعات المصرّح بها (شامل)",
        "ai.btn_expand_title": "تكبير / استعادة العرض",
        "ai.btn_expand_max": "تكبير عرض النافذة (920px)",
        "ai.btn_expand_restore": "استعادة العرض الافتراضي",
        "ai.btn_clear_title": "مسح سجل المحادثة",
        "ai.btn_close_title": "إغلاق (Esc)",
        "ai.quick_label": "استفسارات سريعة:",
        "ai.chip_eng_overdue": "📌 التقديمات المتأخرة",
        "ai.chip_eng_noc": "💵 إجمالي أوامر التغيير المعتمدة (NOC)",
        "ai.chip_eng_status": "📊 ملخص حالات الوثائق",
        "ai.chip_om_summary": "⚡ ملخص عقود الصيانة",
        "ai.chip_om_wo": "📌 أوامر العمل المفتوحة",
        "ai.chip_om_ppm": "📊 حالة الصيانة الوقائية و SLA",
        "ai.chip_hse_kpis": "🛡️ مؤشرات ومعدلات السلامة (HSE)",
        "ai.chip_hse_expiring": "⚠️ تصاريح وطفايات قاربت على الانتهاء",
        "ai.chip_hse_ncr": "📝 صياغة عدم مطابقة سلامة (NCR)",
        "ai.loading_text": "جاري استعلام السجلات والتحليل عبر Gemini Flash...",
        "ai.placeholder_eng": "اسأل عن التقديمات أو أوامر التغيير (NOC) أو التكاليف أو حالات الوثائق...",
        "ai.placeholder_om": "اسأل عن أوامر العمل أو جداول الصيانة الوقائية (PPM) أو طلبات الصيانة...",
        "ai.placeholder_hse": "اسأل عن مؤشرات السلامة أو تصاريح العمل (PTW) أو فحص الطفايات (FEI)...",
        "ai.send_title": "إرسال (Enter)",
        "ai.hint_send": "إرسال",
        "ai.hint_newline": "سطر جديد",
        "ai.hint_factual": "إجابات موثقة من بيانات السجلات فقط",
        "ai.sender_you": "أنت",
        "ai.sender_gemini": "Gemini",
        "ai.welcome_eng": (
            "مرحباً بك! أنا **مساعد DCR الهندسي الذكي** المدعوم بمحرك Gemini Flash.\n\n"
            "لديّ وصول فوري (للقراءة فقط) إلى سجلات المشروعات المصرّح لك بها، ويمكنني مساعدتك في:\n"
            "- 📌 **التقديمات المتأخرة** وعدد أيام التأخير\n"
            "- 💵 **أوامر التغيير (NOC)** وإجمالي التكاليف المعتمدة\n"
            "- 📊 **تحليل حالات السجلات** (معتمد، قيد المراجعة، مرفوض)\n"
            "- 🔍 **البحث عن أرقام وثائق محددة** أو تخصصات هندسية\n\n"
            "*كيف يمكنني مساعدتك اليوم؟*"
        ),
        "ai.welcome_om": (
            "مرحباً بك! أنا **مساعد DCR الذكي للتشغيل والصيانة (O&M)** المدعوم بمحرك Gemini Flash.\n\n"
            "لديّ وصول فوري (للقراءة فقط) إلى عقود الصيانة والمرافق والسجلات الخاصة بك:\n"
            "- ⚙️ **أوامر العمل (WO)** وطلبات الصيانة (MR)\n"
            "- 📅 **جداول الصيانة الوقائية (PPM)** والصيانة التصحيحية (CM)\n"
            "- 📊 **الالتزام بمستوى الخدمة (SLA) وتقارير الأداء**\n"
            "- 🔍 **تقديمات العقود وسجلات الحوادث والأعطال**\n\n"
            "*كيف يمكنني مساعدتك في عقود الصيانة اليوم؟*"
        ),
        "ai.welcome_hse": (
            "مرحباً بك! أنا **مساعد DCR للسلامة والصحة المهنية والامتثال (HSE)** المدعوم بمحرك Gemini Flash.\n\n"
            "لديّ وصول فوري ومُحكم حسب الصلاحيات إلى مواقع السلامة المصرّح لك بها، ويمكنني المساعدة في:\n"
            "- 🛡️ **مؤشرات السلامة:** أيام العمل بدون إصابات (LTI-Free)، ومعدلات LTIF و TRIR وساعات العمل\n"
            "- ⚠️ **تنبيهات التصاريح والمعدات:** تصاريح العمل (PTW) وفحص الطفايات (FEI) والملاحظات المتأخرة\n"
            "- ⚖️ **الامتثال وصياغة تقارير NCR:** وفق معايير OSHA و NFPA وقرار وزارة القوى العاملة المصرية رقم 211/2003\n\n"
            "*كيف يمكنني المساهمة في سلامة مواقعك اليوم؟*"
        ),
        "ai.err_503": "⚠️ **المساعد الذكي غير مُفعّل على هذا الخادم.**\n\nيرجى ضبط متغير البيئة `GEMINI_API_KEY` لتفعيل Gemini Flash.",
        "ai.err_403": "🔒 **يلزم تسجيل الدخول.**\n\nانتهت صلاحية الجلسة، يرجى تحديث الصفحة وتسجيل الدخول مجدداً.",
        "ai.err_service": "⚠️ **خطأ في خدمة الذكاء الاصطناعي:** {error}",
        "ai.err_http": "⚠️ **خطأ ({status}):** {error}",
        "ai.err_server": "خطأ في الخادم",
        "ai.err_empty": "⚠️ تم استلام استجابة فارغة من خدمة الذكاء الاصطناعي.",
        "ai.err_timeout": "⏱️ **انتهت مهلة الطلب (60 ثانية):** استغرق التحليل وقتاً طويلاً، يرجى المحاولة باستفسار أبسط.",
        "ai.err_network": "⚠️ **خطأ في الاتصال بالشبكة:** {error}",
        "ai.err_unreachable": "تعذّر الوصول إلى نقطة اتصال المساعد الذكي.",

        # ── Admin, Department & Permissions Modals (SHARED_JS) ──
        "admin.modal_title": "إدارة المستخدمين وصلاحيات الوصول",
        "admin.rbac_active": "نظام التحكم في الصلاحيات متعدد الأقسام (RBAC) نشط",
        "admin.tab_users": "👥 المستخدمون وصلاحيات الوصول",
        "admin.tab_depts": "🏢 الأقسام ومساحات العمل",
        "admin.all_depts_opt": "جميع الأقسام",
        "admin.all_roles_opt": "جميع الأدوار",
        "admin.role_superadmin": "مدير النظام العام (Super Admin)",
        "admin.role_admin": "مدير قسم (Admin)",
        "admin.role_editor": "محرر (Editor)",
        "admin.role_viewer": "مُطّلع (Viewer)",
        "admin.search_placeholder": "🔍 ابحث باسم المستخدم أو البريد الإلكتروني...",
        "admin.btn_add_user": "إضافة مستخدم",
        "admin.btn_export_matrix": "تصدير المصفوفة",
        "admin.btn_export_matrix_title": "تصدير مصفوفة صلاحيات المستخدمين إلى ملف CSV/Excel",
        "admin.btn_gap_audit": "فحص التسلسل",
        "admin.btn_gap_audit_title": "اكتشاف الفجوات في التسلسل الرقمي للوثائق",
        "admin.add_user_title": "➕ إضافة مستخدم جديد",
        "admin.close_x": "إغلاق ✕",
        "admin.lbl_username": "اسم المستخدم",
        "admin.ph_username": "username",
        "admin.lbl_email": "البريد الإلكتروني",
        "admin.ph_email": "user@company.com",
        "admin.lbl_department": "القسم",
        "admin.lbl_role": "الدور الوظيفي",
        "admin.lbl_password": "كلمة المرور",
        "admin.lbl_downloads": "تحميل الملفات",
        "admin.lbl_allow_dl": "سماح 📥",
        "admin.audit_title": "فحص فجوات التسلسل الرقمي للوثائق",
        "admin.audit_desc": "اكتشاف الأرقام التسلسلية المفقودة في سجلات الوثائق لأي مشروع.",
        "admin.audit_select_proj": "اختر المشروع",
        "admin.audit_select_dt": "اختر نوع الوثيقة",
        "admin.audit_run_btn": "بدء الفحص",
        "admin.audit_select_both": "يرجى اختيار كل من المشروع ونوع الوثيقة.",
        "admin.audit_running": "جاري فحص التسلسل الرقمي للوثائق...",
        "admin.audit_missing_title": "⚠️ تم العثور على أرقام تسلسلية مفقودة!",
        "admin.audit_expected_range": "النطاق المتوقع: من {min} إلى {max}",
        "admin.audit_missing_nums": "الأرقام المفقودة: {list}",
        "admin.audit_no_docs": "لم يتم العثور على وثائق مرقّمة لفحصها.",
        "admin.audit_intact_title": "✅ التسلسل مكتمل بنسبة 100%.",
        "admin.audit_intact_desc": "لا توجد أرقام تسلسلية مفقودة. النطاق: من {min} إلى {max}.",
        "admin.audit_failed": "فشل الفحص: {error}",
        "admin.no_users_match": "لا يوجد مستخدمون يطابقون معايير التصفية المحددة.",
        "admin.th_user": "المستخدم",
        "admin.th_email": "البريد الإلكتروني",
        "admin.th_department": "القسم",
        "admin.th_role": "الدور",
        "admin.th_actions": "الإجراءات",
        "admin.badge_all_depts": "🌐 جميع الأقسام",
        "admin.btn_edit": "تعديل",
        "admin.btn_open": "فتح",
        "admin.btn_assign": "➕ تعيين",
        "admin.no_depts": "بدون أقسام",
        "admin.manage_depts_title": "إدارة تعيينات الأقسام",
        "admin.add_remove_depts_title": "إضافة أو إزالة الأقسام للمستخدم {user}",
        "admin.role_all_viewer": "الكل كـ مُطّلع (Viewer)",
        "admin.role_all_editor": "الكل كـ محرر (Editor)",
        "admin.role_all_admin": "الكل كـ مدير قسم (Admin)",
        "admin.role_mixed": "أدوار متعددة",
        "admin.badge_multi_dept": "🌐 متعدد الأقسام",
        "admin.badge_multi_dept_title": "مشترك عبر عدة أقسام — تتم إدارة إعدادات الحساب العامة بواسطة مدير النظام العام",
        "admin.dept_count_one": "قسم واحد",
        "admin.dept_count_many": "{count} أقسام",
        "admin.ph_add_email": "إضافة بريد إلكتروني...",
        "admin.dl_full": "📥 كامل",
        "admin.dl_full_title": "صلاحية تحميل الملفات الكاملة مشمولة مع دور {role}",
        "admin.dl_on": "التحميل: مفعّل",
        "admin.dl_on_badge": "📥 التحميل: مفعّل",
        "admin.dl_view_only": "عرض فقط",
        "admin.dl_view_only_badge": "🔒 عرض فقط",
        "admin.dl_allowed_title": "تحميل الملفات: مسموح (اضغط للتقييد)",
        "admin.dl_restricted_title": "تحميل الملفات: مقيّد (اضغط للسماح)",
        "admin.dl_global_title": "إعداد التحميل العام يُدار بواسطة مدير النظام العام",
        "admin.lock_protected": "🔒 محمي",
        "admin.lock_you": "👤 حسابك",
        "admin.lock_admin": "🔒 مدير",
        "admin.btn_permissions": "الصلاحيات",
        "admin.btn_permissions_title": "إدارة صلاحيات المشروعات والعقود",
        "admin.btn_reset_pw_title": "إعادة تعيين كلمة المرور",
        "admin.btn_delete_user_title": "حذف المستخدم",
        "admin.err_open": "خطأ أثناء فتح لوحة الإدارة: {error}",

        # ── Departments Management Tab & Modal ──
        "dept.enterprise_title": "🏢 أقسام وإدارات الشركة ومساحات العمل",
        "dept.enterprise_sub": "إعداد أقسام الشركة والأكواد والأوصاف والأيقونات عبر المنصة",
        "dept.btn_add": "إضافة قسم",
        "dept.form_add_title": "➕ إضافة قسم جديد (Add New Department)",
        "dept.form_edit_title": "✏️ تعديل القسم: {name}",
        "dept.btn_save": "حفظ القسم (Save Department)",
        "dept.btn_update": "حفظ التعديلات (Update Department)",
        "dept.btn_cancel": "إلغاء (Cancel)",
        "dept.lbl_code": "كود القسم (Code)",
        "dept.ph_code": "مثال: O&M, HSE",
        "dept.lbl_name_en": "اسم القسم بالإنجليزي (Name in English)",
        "dept.ph_name_en": "مثال: Quality Assurance & Control",
        "dept.lbl_name_ar": "اسم القسم بالعربي (Name in Arabic)",
        "dept.ph_name_ar": "مثال: توكيد ومراقبة الجودة",
        "dept.lbl_desc": "التعريف البسيط / الوصف (يظهر أسفل اسم القسم في القوائم والشاشات)",
        "dept.ph_desc": "مثال: عقود المرافق والتشغيل والصيانة",
        "dept.lbl_icon": "أيقونة القسم (Icon):",
        "dept.icon_hint": "اضغط على أي أيقونة جاهزة أو اكتب أي رمز تريده:",
        "dept.th_workspace": "القسم / مساحة العمل",
        "dept.th_code": "الكود",
        "dept.th_desc": "الوصف المختصر",
        "dept.th_projects": "المشروعات / العقود",
        "dept.th_status": "الحالة",
        "dept.th_actions": "الإجراءات",
        "dept.count_contracts": "{count} عقود",
        "dept.count_projects": "{count} مشروعات",
        "dept.status_active": "نشط",
        "dept.move_up_title": "تحريك لأعلى",
        "dept.move_down_title": "تحريك لأسفل",
        "dept.edit_title": "تعديل بيانات القسم",
        "dept.open_title": "فتح لوحة القسم",
        "dept.err_code_name_req": "يرجى إدخال كود القسم والاسم باللغة الإنجليزية.",
        "dept.toast_saved": "✔ تم حفظ القسم بنجاح!",
        "dept.err_save": "تعذّر حفظ القسم",
        "dept.toast_order_updated": "✔ تم تحديث ترتيب الأقسام",
        "dept.err_reorder": "خطأ في إعادة الترتيب: {error}",
        "dept.user_modal_title": "صلاحيات وأدوار الأقسام — {user}",
        "dept.user_modal_sub": "تعيين أقسام الشركة وتحديد مستوى الصلاحية لهذا المستخدم",
        "dept.company_depts_count": "أقسام الشركة ({count})",
        "dept.lbl_role_colon": "الدور:",
        "dept.btn_done": "إتمام (Done)",

        # ── Scoped Permissions Modal ──
        "perm.modal_title_default": "⚙ صلاحيات المستخدم",
        "perm.modal_sub_default": "مشروعات وعقود القسم",
        "perm.modal_title_user": "⚙️ صلاحيات المشروعات والعقود —",
        "perm.search_placeholder": "ابحث في المشروعات أو العقود...",
        "perm.select_all": "تحديد الكل",
        "perm.clear_all": "إلغاء تحديد الكل",
        "perm.status_immediate": "يتم تطبيق الصلاحيات فوراً.",
        "perm.status_saving": "جاري الحفظ...",
        "perm.status_saved": "✔ تم حفظ التغييرات.",
        "perm.status_failed": "فشل الحفظ.",
        "perm.dl_full_role": "📥 تحميل كامل للملفات ({role})",
        "perm.dl_allow_chk": "📥 السماح بتحميل الملفات",
        "perm.dl_shared_suffix": "(حساب مشترك)",
        "perm.scoped_strictly": "مُقيّد حصرياً بالأقسام المعيّنة للمستخدم",
        "perm.no_projects": "لا توجد مشروعات أو عقود مسجلة لهذا القسم.",
        "perm.th_project": "المشروع / العقد",
        "perm.th_access": "الوصول",
        "perm.th_doc_types": "أنواع الوثائق",
        "perm.client_prefix": "العميل:",
        "perm.dt_all": "الكل",
        "perm.dt_none": "0 أنواع (لا يوجد)",
        "perm.dt_count": "{count} أنواع",
        "perm.allowed_dts_for": "أنواع الوثائق المسموح بها في {name}:",
        "perm.loading_dts": "جاري تحميل أنواع الوثائق...",
        "perm.no_dts_configured": "لا توجد أنواع وثائق مهيأة لهذا المشروع.",

        # ── Admin Action Toasts & Prompts ──
        "admin.toast_user_not_found": "لم يتم العثور على المستخدم",
        "admin.toast_all_roles_updated": "✔ تم تحديث دور المستخدم في جميع الأقسام إلى {role}",
        "admin.toast_dl_enabled": "📥 تم تفعيل تحميل الملفات للمستخدم {user}",
        "admin.toast_dl_restricted": "🔒 تم تقييد تحميل الملفات للمستخدم {user}",
        "admin.err_dl_update": "تعذّر تحديث صلاحية التحميل",
        "admin.toast_role_updated_ws": "✔ تم تحديث الدور في قسم {ws}",
        "admin.toast_role_updated": "✔ تم تحديث الدور الوظيفي",
        "admin.err_role_update": "تعذّر تحديث الدور الوظيفي",
        "admin.toast_assigned_all_depts": "✔ تم التعيين في جميع الأقسام",
        "admin.toast_reset_default_dept": "✔ تمت إعادة الضبط إلى القسم الافتراضي",
        "admin.toast_min_one_dept": "يجب أن ينتمي المستخدم إلى قسم واحد على الأقل",
        "admin.toast_dept_updated": "✔ تم تحديث القسم",
        "admin.toast_dept_updated_to": "✔ تم تحديث القسم إلى {dept}",
        "admin.err_dept_update": "تعذّر تحديث القسم",
        "admin.err_dt_perm_save": "خطأ في حفظ صلاحيات أنواع الوثائق: {error}",
        "admin.err_perm_update": "تعذّر تحديث الصلاحية: {error}",
        "admin.err_user_pw_req": "اسم المستخدم وكلمة المرور مطلوبان",
        "admin.toast_user_added": "✔ تمت إضافة المستخدم بنجاح",
        "admin.err_user_add": "تعذّرت إضافة المستخدم",
        "admin.toast_email_updated": "✔ تم تحديث البريد الإلكتروني",
        "admin.err_email_update": "تعذّر تحديث البريد الإلكتروني",
        "admin.confirm_delete_user": "هل أنت متأكد من رغبتك في حذف المستخدم \"{user}\"؟\nلا يمكن التراجع عن هذا الإجراء.",
        "admin.toast_user_deleted": "✔ تم حذف المستخدم",
        "admin.err_user_delete": "تعذّر حذف المستخدم",
        "admin.prompt_new_pw_user": "أدخل كلمة المرور الجديدة للمستخدم \"{user}\" (4 أحرف على الأقل):",
        "admin.err_pw_change": "خطأ أثناء تغيير كلمة المرور",
        "admin.err_no_export_data": "لا توجد بيانات مستخدمين لتصديرها",
        "admin.toast_csv_exported": "✔ تم تصدير مصفوفة المستخدمين إلى ملف CSV",
        "admin.err_export": "فشل التصدير: {error}",
    },
}


def t(key: str, lang: str = DEFAULT_LANG, **kwargs) -> str:
    """Translate a key into the target language with fallback to English, then key."""
    norm = normalize_lang(lang)
    lang_dict = TRANSLATIONS.get(norm) or TRANSLATIONS[DEFAULT_LANG]
    val = lang_dict.get(key)
    if not val:
        val = TRANSLATIONS[DEFAULT_LANG].get(key, key)
    if kwargs and isinstance(val, str):
        try:
            return val.format(**kwargs)
        except Exception:
            return val
    return val


def get_client_dict(lang: str = DEFAULT_LANG) -> dict:
    """Return a JSON-serializable translation dictionary for client-side JS."""
    norm = normalize_lang(lang)
    base = dict(TRANSLATIONS[DEFAULT_LANG])
    if norm != DEFAULT_LANG and norm in TRANSLATIONS:
        base.update(TRANSLATIONS[norm])
    return base


def get_client_i18n_json(lang: str = DEFAULT_LANG) -> str:
    """Return safe inline JSON string for embedding in <script> tags."""
    payload = get_client_dict(lang)
    return json.dumps(payload, ensure_ascii=False).replace("</", "<\\/")


def translate_workspace_meta(ws: dict, lang: str = DEFAULT_LANG) -> dict:
    """Return localized display name and subtitle/description for a workspace dict without mutating canonical IDs."""
    norm = normalize_lang(lang)
    ws_id = str((ws or {}).get("id") or (ws or {}).get("code") or "ENG").strip().upper()
    raw_name = str((ws or {}).get("name") or ws_id).strip()
    raw_name_ar = str((ws or {}).get("name_ar") or "").strip()
    raw_desc = (ws or {}).get("description")
    raw_desc_str = str(raw_desc).strip() if raw_desc is not None else ""

    known_en_defaults = {
        "EPC & Site Submittals",
        "Facilities & Maintenance Contracts",
        "Permits, Safety Observations, Incidents, Man-Hours & Fire Equipment",
        "EPC Submittals, Drawings, Material Approvals, IR, MIR, RFI & NOC Variation Orders.",
        "Facilities maintenance contracts, Work Orders, PPM schedules and SLA compliance metrics.",
        "Personnel archives, employee onboarding, credential files, attendance and corporate policies.",
        "Interim Payment Certificates (IPC), subcontractor invoices, project budgets and cost audits.",
        "Purchase Orders (PO), vendor quotations, material requisition packages and approved suppliers.",
        "Central store inventory, Material Receipt Notes (MRN), site dispatch and stock balances.",
        "Occupational safety audits, site inspection checklists, incident registers and ISO certifications.",
        "Tender submissions, BOQ bill of quantities, tender addendums, commercial bids and award tracking.",
    }

    if norm == "ar":
        disp_name = raw_name_ar or t(f"ws.{ws_id}.name", "ar")
        if disp_name == f"ws.{ws_id}.name":
            disp_name = raw_name
        if raw_desc_str and raw_desc_str not in known_en_defaults:
            disp_desc = raw_desc_str
            disp_sub = raw_desc_str
        else:
            disp_desc = t(f"ws.{ws_id}.desc", "ar")
            if disp_desc == f"ws.{ws_id}.desc":
                disp_desc = raw_desc_str or raw_name_ar or raw_name
            disp_sub = t(f"ws.{ws_id}.sub", "ar")
            if disp_sub == f"ws.{ws_id}.sub":
                disp_sub = raw_name_ar or raw_desc_str or disp_name
    else:
        disp_name = raw_name or t(f"ws.{ws_id}.name", "en")
        if raw_desc_str:
            disp_desc = raw_desc_str
            disp_sub = raw_desc_str
        else:
            disp_desc = t(f"ws.{ws_id}.desc", "en")
            if disp_desc == f"ws.{ws_id}.desc":
                disp_desc = ""
            disp_sub = t(f"ws.{ws_id}.sub", "en")
            if disp_sub == f"ws.{ws_id}.sub":
                disp_sub = raw_name_ar or ""

    return {
        "name": disp_name,
        "sub": disp_sub,
        "desc": disp_desc,
    }
