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
