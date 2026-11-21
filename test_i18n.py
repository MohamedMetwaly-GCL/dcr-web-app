"""test_i18n.py - Tier 1 isolated unit tests for Phase 1A i18n engine, RTL CSS, language switcher, /login, /hub, and shared chrome."""
import os
os.environ["DATABASE_URL"] = ""
os.environ["SECRET_KEY"] = "synthetic-test-secret-key-123456789012345678901234"
os.environ["MAGIC_SECRET"] = "synthetic-test-magic-secret-only"
os.environ["GEMINI_API_KEY"] = "synthetic-test-gemini-key-only"

import json
import unittest
from unittest.mock import patch

import i18n
import html_render
from app import app


class TestI18nCoreModule(unittest.TestCase):
    """Verify i18n.py constants, language resolution, RTL helpers, and EN/AR dictionary parity."""

    def test_constants_and_direction_helpers(self):
        self.assertEqual(i18n.SUPPORTED_LANGS, ("en", "ar"))
        self.assertEqual(i18n.DEFAULT_LANG, "en")
        self.assertEqual(i18n.LANG_COOKIE_NAME, "dcr_lang")
        self.assertFalse(i18n.is_rtl("en"))
        self.assertTrue(i18n.is_rtl("ar"))
        self.assertFalse(i18n.is_rtl("invalid"))
        self.assertEqual(i18n.html_dir("en"), "ltr")
        self.assertEqual(i18n.html_dir("ar"), "rtl")
        self.assertEqual(i18n.html_dir(None), "ltr")

    def test_normalize_lang_strict_whitelist(self):
        self.assertEqual(i18n.normalize_lang("en"), "en")
        self.assertEqual(i18n.normalize_lang("ar"), "ar")
        self.assertEqual(i18n.normalize_lang("AR"), "ar")
        self.assertEqual(i18n.normalize_lang(" ar "), "ar")
        self.assertEqual(i18n.normalize_lang("ar-EG"), "en")
        self.assertEqual(i18n.normalize_lang("en-US"), "en")
        self.assertEqual(i18n.normalize_lang("fr"), "en")
        self.assertEqual(i18n.normalize_lang("<script>alert(1)</script>"), "en")
        self.assertEqual(i18n.normalize_lang(""), "en")
        self.assertEqual(i18n.normalize_lang(None), "en")

    def test_get_lang_priority_and_invalid_input_rejection(self):
        # 1. Outside request context falls back to 'en'
        self.assertEqual(i18n.get_lang(), "en")

        # 2. Query parameter ?lang=ar takes highest priority over cookie dcr_lang=en
        with app.test_request_context("/login?lang=ar", headers={"Cookie": "dcr_lang=en"}):
            self.assertEqual(i18n.get_lang(), "ar")

        # 3. Query parameter ?lang=en takes priority over cookie dcr_lang=ar
        with app.test_request_context("/hub?lang=en", headers={"Cookie": "dcr_lang=ar"}):
            self.assertEqual(i18n.get_lang(), "en")

        # 4. Valid cookie dcr_lang=ar used when query param is absent
        with app.test_request_context("/hub", headers={"Cookie": "dcr_lang=ar"}):
            self.assertEqual(i18n.get_lang(), "ar")

        # 5. Invalid query param falls back to valid cookie
        with app.test_request_context("/hub?lang=fr", headers={"Cookie": "dcr_lang=ar"}):
            self.assertEqual(i18n.get_lang(), "ar")

        # 6. Invalid query param and invalid cookie fall back to 'en'
        with app.test_request_context("/login?lang=de", headers={"Cookie": "dcr_lang=es"}):
            self.assertEqual(i18n.get_lang(), "en")

    def test_dictionary_parity_between_en_and_ar(self):
        en_dict = i18n.TRANSLATIONS.get("en", {})
        ar_dict = i18n.TRANSLATIONS.get("ar", {})
        self.assertGreater(len(en_dict), 50)
        en_keys = set(en_dict.keys())
        ar_keys = set(ar_dict.keys())

        missing_in_ar = en_keys - ar_keys
        missing_in_en = ar_keys - en_keys
        self.assertEqual(missing_in_ar, set(), f"Keys missing in Arabic dictionary: {sorted(missing_in_ar)}")
        self.assertEqual(missing_in_en, set(), f"Keys missing in English dictionary: {sorted(missing_in_en)}")

        for key, val in en_dict.items():
            self.assertIsInstance(val, str, f"EN key '{key}' is not a string")
            self.assertTrue(len(val.strip()) > 0, f"EN key '{key}' is empty")

        for key, val in ar_dict.items():
            self.assertIsInstance(val, str, f"AR key '{key}' is not a string")
            self.assertTrue(len(val.strip()) > 0, f"AR key '{key}' is empty")
            # Ensure standard Western Arabic digits (0-9) are used, no Eastern Arabic-Indic digits (٠-٩)
            for eastern_digit in "٠١٢٣٤٥٦٧٨٩":
                self.assertNotIn(eastern_digit, val, f"AR key '{key}' contains Eastern numeral '{eastern_digit}'")

    def test_t_translation_lookup_formatting_and_fallback(self):
        self.assertEqual(i18n.t("login.signin_btn", "en"), "Sign In")
        self.assertEqual(i18n.t("login.signin_btn", "ar"), "تسجيل الدخول")
        # Placeholder interpolation with Western numerals
        msg_ar = i18n.t("topbar.active_workspace", "ar", name="ENG-5")
        self.assertIn("ENG-5", msg_ar)
        self.assertIn("القسم النشط", msg_ar)
        # Missing key falls back gracefully to key name
        self.assertEqual(i18n.t("nonexistent.custom.key", "ar"), "nonexistent.custom.key")

    def test_client_dict_and_json_serialization(self):
        client_ar = i18n.get_client_dict("ar")
        self.assertIsInstance(client_ar, dict)
        self.assertIn("login.signin_btn", client_ar)
        self.assertEqual(client_ar["login.signin_btn"], "تسجيل الدخول")
        raw_json = i18n.get_client_i18n_json("ar")
        parsed = json.loads(raw_json)
        self.assertEqual(parsed["login.signin_btn"], "تسجيل الدخول")


class TestPhase1ARenderersAndRoutes(unittest.TestCase):
    """Verify render_login, render_launchpad, shared chrome, and Flask route wiring for EN and AR."""

    def setUp(self):
        self.app = app
        self.app.config["TESTING"] = True
        self.client = self.app.test_client()

    def test_render_login_en_and_ar(self):
        html_en = html_render.render_login(lang="en")
        self.assertIn('<html lang="en" dir="ltr">', html_en)
        self.assertIn("Enterprise Document Control System", html_en)
        self.assertIn("Sign In", html_en)
        self.assertIn("setLanguage('ar')", html_en)
        self.assertIn("عربي", html_en)

        html_ar = html_render.render_login(lang="ar")
        self.assertIn('<html lang="ar" dir="rtl">', html_ar)
        self.assertIn("Cairo", html_ar)
        self.assertIn("تسجيل الدخول", html_ar)
        self.assertIn("النظام المؤسسي لمراقبة الوثائق", html_ar)
        self.assertIn("setLanguage('en')", html_ar)
        self.assertIn("EN", html_ar)
        self.assertIn('html[dir="rtl"]', html_ar)

    def test_render_launchpad_en_and_ar(self):
        user = {"username": "m.metwaly", "role": "superadmin", "primary_workspace": "ENG"}
        mock_workspaces = [
            {
                "id": "ENG",
                "code": "ENG",
                "name": "Engineering & Projects",
                "name_ar": "المشروعات الهندسية",
                "icon": "🏗️",
                "status": "active",
            },
            {
                "id": "OM",
                "code": "OM",
                "name": "Operation & Maintenance (O&M)",
                "name_ar": "الصيانة والتشغيل",
                "icon": "⚙️",
                "status": "active",
            },
        ]
        with patch("db.get_workspaces", return_value=mock_workspaces), \
             patch("db.get_user_workspaces", return_value=mock_workspaces), \
             patch("db.get_projects_for_user", return_value=[{"id": "24CP01", "name": "Alpha", "code": "CP01"}]), \
             patch("db.q", return_value=[{"project_id": "24CP01", "c": 15}]):
            html_en = html_render.render_launchpad(user, lang="en")
            self.assertIn('<html lang="en" dir="ltr">', html_en)
            self.assertIn("Company Workspaces Launchpad", html_en)
            self.assertIn("Engineering & Projects", html_en)
            self.assertIn("15", html_en)
            self.assertIn("Records", html_en)
            self.assertIn("setLanguage('ar')", html_en)

            html_ar = html_render.render_launchpad(user, lang="ar")
            self.assertIn('<html lang="ar" dir="rtl">', html_ar)
            self.assertIn("منصة أقسام وإدارات الشركة", html_ar)
            self.assertIn("المشروعات الهندسية", html_ar)
            self.assertIn("الصيانة والتشغيل", html_ar)
            self.assertIn("15", html_ar)
            self.assertIn("السجلات", html_ar)
            self.assertIn("دخول القسم", html_ar)
            self.assertIn("setLanguage('en')", html_ar)
            self.assertIn("bidi-ltr", html_ar)

    def test_shared_chrome_and_dashboard_topbar_rtl(self):
        user = {"username": "admin_user", "role": "superadmin"}
        mock_workspaces = [
            {"id": "ENG", "code": "ENG", "name": "Engineering & Projects", "name_ar": "المشروعات الهندسية", "icon": "🏗️"},
            {"id": "OM", "code": "OM", "name": "Operation & Maintenance", "name_ar": "الصيانة والتشغيل", "icon": "⚙️"},
        ]
        with patch("db.get_user_workspaces", return_value=mock_workspaces), \
             patch("db.get_projects_for_user", return_value=[]):
            dash_ar = html_render.render_dashboard(user, active_ws_id="ENG", lang="ar")
            self.assertIn('<html lang="ar" dir="rtl">', dash_ar)
            self.assertIn("سجل مراقبة الوثائق والمستندات", dash_ar)
            self.assertIn("المشروعات الهندسية", dash_ar)
            self.assertIn("أقسام وإدارات الشركة", dash_ar)
            self.assertIn("مدير النظام العام", dash_ar)
            self.assertIn("window.DCR_LANG = \"ar\"", dash_ar)
            self.assertIn("window.setLanguage", dash_ar)
            self.assertIn(".bidi-ltr", dash_ar)

    def test_routes_respect_lang_query_and_persist_cookie(self):
        # 1. GET /login?lang=ar returns Arabic RTL HTML and sets dcr_lang=ar cookie
        resp = self.client.get("/login?lang=ar")
        self.assertEqual(resp.status_code, 200)
        body = resp.get_data(as_text=True)
        self.assertIn('<html lang="ar" dir="rtl">', body)
        self.assertIn("تسجيل الدخول", body)
        set_cookie_hdr = "\n".join(resp.headers.getlist("Set-Cookie"))
        self.assertIn("dcr_lang=ar", set_cookie_hdr)

        # 2. Subsequent GET /login without ?lang= uses persisted dcr_lang=ar cookie
        resp2 = self.client.get("/login")
        self.assertEqual(resp2.status_code, 200)
        self.assertIn('<html lang="ar" dir="rtl">', resp2.get_data(as_text=True))

        # 3. Switching back via ?lang=en updates cookie to en and renders LTR
        resp3 = self.client.get("/login?lang=en")
        self.assertEqual(resp3.status_code, 200)
        self.assertIn('<html lang="en" dir="ltr">', resp3.get_data(as_text=True))
        set_cookie_hdr3 = "\n".join(resp3.headers.getlist("Set-Cookie"))
        self.assertIn("dcr_lang=en", set_cookie_hdr3)


class TestPhase1BSharedOverlays(unittest.TestCase):
    """Verify Spotlight Global Search, AI Assistant Drawer, and Admin/Permissions Modals EN & AR localization and RTL hooks."""

    def test_spotlight_modal_en_and_ar_and_rtl_css(self):
        sl_en = html_render._spotlight_modal_html("en")
        self.assertIn('id="spotlight-modal"', sl_en)
        self.assertIn('placeholder="Search across all accessible projects (min. 2 chars)..."', sl_en)
        self.assertIn("<option value=\"ALL\">All Projects</option>", sl_en)
        self.assertIn("<option value=\"ALL\">All Statuses</option>", sl_en)
        self.assertIn("Instant Global Search", sl_en)
        self.assertIn("🔒 RBAC-Filtered Results", sl_en)

        sl_ar = html_render._spotlight_modal_html("ar")
        self.assertIn('id="spotlight-modal"', sl_ar)
        self.assertIn('placeholder="ابحث في جميع المشروعات المتاحة (حرفان على الأقل)..."', sl_ar)
        self.assertIn("<option value=\"ALL\">جميع المشروعات</option>", sl_ar)
        self.assertIn("<option value=\"ALL\">جميع الحالات</option>", sl_ar)
        self.assertIn("<option value=\"Approved\">معتمد (Approved)</option>", sl_ar)
        self.assertIn("البحث الشامل الفوري", sl_ar)
        self.assertIn("🔒 نتائج مفلترة حسب صلاحيات الوصول (RBAC)", sl_ar)
        self.assertIn("↺ إعادة ضبط", sl_ar)

        # Verify RTL CSS hooks and .bidi-ltr wrappers for doc_no, project_code, and dates
        self.assertIn('html[dir="rtl"] #spotlight-modal', sl_ar)
        self.assertIn('html[dir="rtl"] .spotlight-item-meta', sl_ar)
        self.assertIn('spotlight-item-docno bidi-ltr', sl_ar)
        self.assertIn('spotlight-item-proj-code bidi-ltr', sl_ar)
        self.assertIn('_slT(', sl_ar)

    def test_ai_drawer_en_and_ar_and_rtl_docking(self):
        ai_en = html_render._ai_drawer_html("en")
        self.assertIn('id="ai-assistant-drawer"', ai_en)
        self.assertIn('id="ai-fab-btn"', ai_en)
        self.assertIn('id="ai-drawer-resizer"', ai_en)
        self.assertIn("DCR Engineering Assistant", ai_en)
        self.assertIn("Quick Queries:", ai_en)
        self.assertIn("📌 Overdue Submittals", ai_en)
        self.assertIn("💵 Approved NOCs Total", ai_en)
        self.assertIn("📊 Status Summary", ai_en)

        ai_ar = html_render._ai_drawer_html("ar")
        self.assertIn('id="ai-assistant-drawer"', ai_ar)
        self.assertIn("مساعد DCR الهندسي الذكي", ai_ar)
        self.assertIn("🔒 قراءة فقط", ai_ar)
        self.assertIn("استفسارات سريعة:", ai_ar)
        self.assertIn("📌 التقديمات المتأخرة", ai_ar)
        self.assertIn("💵 إجمالي أوامر التغيير المعتمدة (NOC)", ai_ar)
        self.assertIn("📊 ملخص حالات الوثائق", ai_ar)
        self.assertIn("جاري استعلام السجلات والتحليل عبر Gemini Flash...", ai_ar)
        self.assertIn("إجابات موثقة من بيانات السجلات فقط", ai_ar)

        # Verify RTL drawer docking, resizer math, and FAB positioning rules
        self.assertIn('html[dir="rtl"] #ai-assistant-drawer', ai_ar)
        self.assertIn('html[dir="rtl"] .ai-drawer-resizer', ai_ar)
        self.assertIn('html[dir="rtl"] #ai-fab-btn', ai_ar)
        self.assertIn('transform: translateX(-100%)', ai_ar)
        self.assertIn("const diff = isRtl ? (e.clientX - startX) : (startX - e.clientX);", ai_ar)
        self.assertIn('_aiT(', ai_ar)

    def test_admin_and_permissions_modals_en_and_ar(self):
        shared_en = html_render._shared_js_html("en")
        self.assertIn('id="admin-modal"', shared_en)
        self.assertIn('id="permissions-modal"', shared_en)
        self.assertIn("User Management & Access Control", shared_en)
        self.assertIn("Scoped Multi-Workspace RBAC Active", shared_en)
        self.assertIn("⚙ User Permissions", shared_en)
        self.assertIn("Department Projects & Contracts", shared_en)
        self.assertIn("Permissions are applied immediately.", shared_en)

        shared_ar = html_render._shared_js_html("ar")
        self.assertIn('window.DCR_LANG = "ar"', shared_ar)
        self.assertIn('window.DCR_DIR = "rtl"', shared_ar)
        self.assertIn("إدارة المستخدمين وصلاحيات الوصول", shared_ar)
        self.assertIn("نظام التحكم في الصلاحيات متعدد الأقسام (RBAC) نشط", shared_ar)
        self.assertIn("⚙ صلاحيات المستخدم", shared_ar)
        self.assertIn("مشروعات وعقود القسم", shared_ar)
        self.assertIn("ابحث في المشروعات أو العقود...", shared_ar)
        self.assertIn("تحديد الكل", shared_ar)
        self.assertIn("إلغاء تحديد الكل", shared_ar)
        self.assertIn("يتم تطبيق الصلاحيات فوراً.", shared_ar)

        # Verify client dictionary includes keys for Spotlight, AI Drawer, and Admin/Permissions
        client_ar = i18n.get_client_dict("ar")
        for required_key in (
            "spotlight.scope_admin",
            "spotlight.no_match_title",
            "ai.title_eng",
            "ai.title_om",
            "ai.title_hse",
            "ai.welcome_eng",
            "admin.tab_users",
            "admin.tab_depts",
            "dept.enterprise_title",
            "perm.modal_title_user",
            "perm.dt_none",
        ):
            self.assertIn(required_key, client_ar)

        # Verify RTL CSS rules in ADMIN_MODAL_CSS and BASE_CSS
        self.assertIn('html[dir="rtl"] .user-tbl-th-actions', html_render.ADMIN_MODAL_CSS)
        self.assertIn('html[dir="rtl"] .user-tbl-th-actions', html_render.BASE_CSS)

    def test_dashboard_and_register_embed_localized_overlays_in_arabic(self):
        user = {"username": "super_admin", "role": "superadmin"}
        proj = {"id": "24CP01", "code": "24CP01", "name": "Alpha Plant", "workspace_id": "ENG"}
        mock_workspaces = [
            {"id": "ENG", "code": "ENG", "name": "Engineering & Projects", "name_ar": "المشروعات الهندسية", "icon": "🏗️"},
        ]
        with patch("db.get_user_workspaces", return_value=mock_workspaces), \
             patch("db.get_projects_for_user", return_value=[proj]), \
             patch("db.get_doc_types", return_value=[]), \
             patch("db.get_logo", return_value=None), \
             patch("html_render.can_edit", return_value=True):
            dash_ar = html_render.render_dashboard(user, active_ws_id="ENG", lang="ar")
            self.assertIn("ابحث في جميع المشروعات المتاحة (حرفان على الأقل)...", dash_ar)
            self.assertIn("مساعد DCR الهندسي الذكي", dash_ar)
            self.assertIn("إدارة المستخدمين وصلاحيات الوصول", dash_ar)

            reg_ar = html_render.render_register(user, proj, active_ws_id="ENG", lang="ar")
            self.assertIn('<html lang="ar" dir="rtl">', reg_ar)
            self.assertIn("ابحث في جميع المشروعات المتاحة (حرفان على الأقل)...", reg_ar)
            self.assertIn("مساعد DCR الهندسي الذكي", reg_ar)
            self.assertIn("إدارة المستخدمين وصلاحيات الوصول", reg_ar)


class TestPhase2AWorkspaceDashboards(unittest.TestCase):
    """Verify I18N-02A localization, BiDi/RTL CSS/JS hooks, and Chart.js RTL config for ENG, O&M/Dynamic, and HSE dashboards."""

    def test_display_helpers_preserve_abbreviations_and_canonical_en(self):
        self.assertEqual(i18n.t_status("Approved", "en"), "Approved")
        self.assertEqual(i18n.t_status("Approved", "ar"), "معتمد")
        self.assertEqual(i18n.t_status("Under Review", "ar"), "قيد المراجعة")
        self.assertEqual(i18n.t_status("CustomStatus", "ar"), "CustomStatus")

        self.assertEqual(i18n.t_dt("NOC", "No Objection Certificates", "en"), "No Objection Certificates")
        self.assertEqual(i18n.t_dt("NOC", "No Objection Certificates", "ar"), "طلبات عدم الممانعة (NOC)")
        self.assertEqual(i18n.t_dt("RFI", "Requests for Information", "ar"), "طلبات المعلومات (RFI)")
        self.assertEqual(i18n.t_dt("VO", "Variation Orders", "ar"), "أوامر التغيير (VO)")
        self.assertEqual(i18n.t_dt("PR", "Purchase Requisitions", "ar"), "طلبات الشراء (PR)")
        self.assertEqual(i18n.t_dt("PTW", "Permit to Work", "ar"), "تصاريح العمل (PTW)")

        client_ar = i18n.get_client_dict("ar")
        for key in ("dash.kpi_total_submittals", "om.kpi_contracts", "hse.kpi_ltif", "dt_name.NOC", "status.approved"):
            self.assertIn(key, client_ar)

    def test_eng_dashboard_en_and_ar_rtl_hooks(self):
        user = {"username": "super_admin", "role": "superadmin"}
        proj = {"id": "24CP01", "code": "CP01", "name": "Alpha Plant", "workspace_id": "ENG"}
        mock_ws = [{"id": "ENG", "code": "ENG", "name": "Engineering & Projects", "name_ar": "المشروعات الهندسية", "icon": "🏗️"}]
        with patch("db.get_user_workspaces", return_value=mock_ws), \
             patch("db.get_projects_for_user", return_value=[proj]):
            html_en = html_render.render_dashboard(user, active_ws_id="ENG", lang="en")
            self.assertIn('<html lang="en" dir="ltr">', html_en)
            self.assertIn("Submittals KPI Summary", html_en)
            self.assertIn("PR Analytics", html_en)

            html_ar = html_render.render_dashboard(user, active_ws_id="ENG", lang="ar")
            self.assertIn('<html lang="ar" dir="rtl">', html_ar)
            self.assertIn("مؤشرات أداء الاعتمادات (KPIs)", html_ar)
            self.assertIn("تحليلات طلبات الشراء (PR)", html_ar)
            self.assertIn("نظرة عامة على الخطابات (LTR)", html_ar)
            self.assertIn("سجل التدقيق", html_ar)
            # RTL sticky first column, accent borders, and audit column resizer hooks
            self.assertIn('html[dir="rtl"] .dt-tbl th:first-child', html_ar)
            self.assertIn("right:0", html_ar)
            self.assertIn('html[dir="rtl"] #aud-tbl th .aud-rz', html_ar)
            self.assertIn("document.documentElement.dir==='rtl'", html_ar)
            self.assertIn('html[dir="rtl"] .pr-total', html_ar)
            self.assertIn('html[dir="rtl"] .ltr-summary-card', html_ar)
            # Chart.js RTL config & prValueLabels RTL support
            self.assertIn("rtl:isRtl,textDirection:isRtl?'rtl':'ltr'", html_ar)
            self.assertIn("prValueLabels", html_ar)
            # BiDi isolation on KPIs and inputs
            self.assertIn('class="kval bidi-ltr" id="kpi-total"', html_ar)
            self.assertIn('id="digest-date" class="bidi-ltr"', html_ar)

    def test_om_and_dynamic_dashboard_en_and_ar(self):
        user = {"username": "om_admin", "role": "admin"}
        proj = {"id": "OM-01", "code": "OM01", "name": "CFC Facility", "workspace_id": "OM", "data": {"client": "Al-Futtaim"}}
        ws_om = {"id": "OM", "code": "OM", "name": "Operation & Maintenance (O&M)", "name_ar": "الصيانة والتشغيل", "icon": "⚙️"}
        with patch("db.get_user_workspaces", return_value=[ws_om]), \
             patch("db.get_workspace", return_value=ws_om), \
             patch("db.get_projects_for_user", return_value=[proj]), \
             patch("db.q", return_value=[]), \
             patch("db.get_overdue_records", return_value=[]):
            html_en = html_render.render_dashboard(user, active_ws_id="OM", lang="en")
            self.assertIn('<html lang="en" dir="ltr">', html_en)
            self.assertIn("Total Facilities & Contracts", html_en)

            html_ar = html_render.render_dashboard(user, active_ws_id="OM", lang="ar")
            self.assertIn('<html lang="ar" dir="rtl">', html_ar)
            self.assertIn("الصيانة والتشغيل", html_ar)
            self.assertIn("إجمالي المنشآت والعقود", html_ar)
            self.assertIn("فتح السجل", html_ar)
            self.assertIn('class="om-code-badge bidi-ltr"', html_ar)
            self.assertIn('<span class="rtl-flip">➔</span>', html_ar)
            self.assertIn('html[dir="rtl"] .om-header-bar', html_ar)

    def test_hse_dashboard_en_and_ar_rtl_and_charts(self):
        user = {"username": "hse_admin", "role": "admin"}
        proj = {"id": "HSE-01", "code": "HSE01", "name": "Main Plant HSE", "workspace_id": "HSE", "data": {"client": "GasChill"}}
        ws_hse = {"id": "HSE", "code": "HSE", "name": "Health, Safety & Environment", "name_ar": "الصحة والسلامة والبيئة", "icon": "🛡️"}
        with patch("db.get_user_workspaces", return_value=[ws_hse]), \
             patch("db.get_workspace", return_value=ws_hse), \
             patch("db.get_projects_for_user", return_value=[proj]), \
             patch("hse_analytics.get_scoped_hse_analytics", side_effect=RuntimeError("mocked")):
            html_en = html_render.render_dashboard(user, active_ws_id="HSE", lang="en")
            self.assertIn('<html lang="en" dir="ltr">', html_en)
            self.assertIn("Executive Command Center", html_en)
            self.assertIn("LTIF (per 1M Man-Hours)", html_en)

            html_ar = html_render.render_dashboard(user, active_ws_id="HSE", lang="ar")
            self.assertIn('<html lang="ar" dir="rtl">', html_ar)
            self.assertIn("الصحة والسلامة والبيئة", html_ar)
            self.assertIn("مركز القيادة التنفيذي (HSE)", html_ar)
            self.assertIn("LTIF", html_ar)
            self.assertIn("TRIR", html_ar)
            self.assertIn("PTW", html_ar)
            self.assertIn('html[dir="rtl"] .hse-header-bar', html_ar)
            self.assertIn('html[dir="rtl"] .hse-alert-item', html_ar)
            self.assertIn("rtl: isRtl, textDirection: isRtl ? 'rtl' : 'ltr'", html_ar)
            self.assertIn('class="om-code-badge bidi-ltr"', html_ar)
            self.assertIn('<span class="rtl-flip">➔</span>', html_ar)


class TestPhase2B1RegisterGridAndSettings(unittest.TestCase):
    """Verify I18N-02B-1: t_col(), reg.*/col.* parity, render_register() EN/AR, RTL sticky cols/resizer, and BiDi isolation."""

    def test_t_col_built_in_and_custom_fallback(self):
        # Built-in system columns translate in Arabic, preserve default_label in English
        self.assertEqual(i18n.t_col("docNo", "Document No.", "en"), "Document No.")
        self.assertEqual(i18n.t_col("docNo", "Document No.", "ar"), "رقم المستند")
        self.assertEqual(i18n.t_col("partBStatus", "Part B Status", "ar"), "حالة Part B")
        self.assertEqual(i18n.t_col("parentLetterRef", "Response Ref", "ar"), "مرجع الرد")
        self.assertEqual(i18n.t_col("facility", "Facility / Location", "ar"), "المنشأة / الموقع")
        self.assertEqual(i18n.t_col("extinguisherStatus", "Extinguisher Status", "ar"), "حالة الطفاية")
        self.assertEqual(i18n.t_col("hours", "Worked Hours", "ar"), "ساعات العمل")

        # Custom user-created columns fall back to default_label
        self.assertEqual(i18n.t_col("col_custom_999", "Site Engineer Note", "en"), "Site Engineer Note")
        self.assertEqual(i18n.t_col("col_custom_999", "Site Engineer Note", "ar"), "Site Engineer Note")

        # Client dictionary exposes reg.*, col.*, status.*, dt_name.* and preserves excludes
        client_ar = i18n.get_client_dict("ar")
        for key in ("reg.btn_add", "reg.th_sr", "reg.proj_modal_title", "col.docNo", "col.partBStatus", "status.a_approved", "dt_name.NOC"):
            self.assertIn(key, client_ar)
        for excl in i18n._SERVER_ONLY_CLIENT_EXCLUDES:
            self.assertNotIn(excl, client_ar)

    def test_register_en_and_ar_localization_and_rtl_hooks(self):
        user = {"username": "super_admin", "role": "superadmin"}
        proj = {
            "id": "24CP01",
            "code": "CP01",
            "name": "Alpha Plant",
            "workspace_id": "ENG",
            "client": "GasChill",
            "status_colors": {},
        }
        dts = [
            {"id": "SD", "code": "SD", "name": "Shop Drawing Submittal"},
            {"id": "NOC", "code": "NOC", "name": "Notice of Change"},
            {"id": "LTR", "code": "LTR", "name": "Letters"},
        ]
        with patch("html_render.can_edit", return_value=True), \
             patch("html_render.is_project_admin", return_value=True), \
             patch("html_render.get_allowed_doc_type_filter", return_value=None), \
             patch("db.can_user_download", return_value=True), \
             patch("db.get_projects_for_user", return_value=[proj]), \
             patch("db.get_doc_types", return_value=dts), \
             patch("db.get_logo", return_value=None), \
             patch("db.get_user_workspaces", return_value=[{"id": "ENG", "code": "ENG", "name": "Engineering & Projects", "name_ar": "المشروعات الهندسية", "icon": "🏗️"}]):
            html_en = html_render.render_register(user, proj, active_ws_id="ENG", lang="en")
            self.assertIn('<html lang="en" dir="ltr">', html_en)
            self.assertIn("➕ Add", html_en)
            self.assertIn("Latest Revisions Only", html_en)
            self.assertIn("⚙ Manage Columns", html_en)
            self.assertIn("📋 Dropdown Lists", html_en)
            self.assertIn("🎨 Status Colors", html_en)

            html_ar = html_render.render_register(user, proj, active_ws_id="ENG", lang="ar")
            self.assertIn('<html lang="ar" dir="rtl">', html_ar)
            # Toolbar, search, quickbar, bulk bar, status bar
            self.assertIn("➕ إضافة", html_ar)
            self.assertIn("⚙ الأعمدة", html_ar)
            self.assertIn("🗓 الإجازات", html_ar)
            self.assertIn("📋 القوائم", html_ar)
            self.assertIn("🎨 ألوان الحالات", html_ar)
            self.assertIn("🏗 المشروع", html_ar)
            self.assertIn("📥 تصدير ▾", html_ar)
            self.assertIn("🖨 طباعة", html_ar)
            self.assertIn("📤 استيراد", html_ar)
            self.assertIn("أحدث المراجعات فقط", html_ar)
            self.assertIn('placeholder="بحث..."', html_ar)
            self.assertIn("عرض الخطابات", html_ar)
            self.assertIn("0 محدد", html_ar)
            self.assertIn("الإجمالي: 0", html_ar)
            self.assertIn("المعروض: 0", html_ar)
            self.assertIn("المتأخر: 0", html_ar)

            # Register settings modals
            self.assertIn("🏗 بيانات المشروع", html_ar)
            self.assertIn("📇 مصفوفة التوزيع", html_ar)
            self.assertIn("📊 مركز التصدير", html_ar)
            self.assertIn("📋 القوائم المنسدلة", html_ar)
            self.assertIn("إضافة نوع مستند", html_ar)
            self.assertIn("⚙ إدارة الأعمدة", html_ar)
            self.assertIn("إضافة عمود", html_ar)
            self.assertIn("📤 استيراد Excel / CSV", html_ar)
            self.assertIn("⚙ إعدادات المشروع", html_ar)

            # Mandatory RTL CSS & JS fixes
            self.assertIn('html[dir="rtl"] #regtbl th:nth-child(1)', html_ar)
            self.assertIn('right: 0 !important;', html_ar)
            self.assertIn('right: 32px !important;', html_ar)
            self.assertIn('right: 66px !important;', html_ar)
            self.assertIn('box-shadow: -2px 0 5px -2px rgba(0,0,0,0.1) !important;', html_ar)
            self.assertIn('html[dir="rtl"] #regtbl tbody tr.row-selected td:first-child', html_ar)
            self.assertIn('html[dir="rtl"] .rz{right:auto;left:0}', html_ar)
            self.assertIn("const isRtl=document.documentElement.dir==='rtl';", html_ar)
            self.assertIn('html[dir="rtl"] #srchbox{padding:4px 28px 4px 10px;background-position:calc(100% - 7px) center}', html_ar)

            # Mandatory BiDi isolation & client i18n helpers
            self.assertIn("td.classList.add('docno-cell');td.classList.add('bidi-ltr')", html_ar)
            self.assertIn("rBadge.className='bidi-ltr'", html_ar)
            self.assertIn('<span id="s-clock" class="bidi-ltr"></span>', html_ar)
            self.assertIn("tsr.className='sr bidi-ltr'", html_ar)
            self.assertIn("window.tCol", html_ar)
            self.assertIn("window.tStatus", html_ar)
            self.assertIn("window.tDt", html_ar)


class TestPhase2B2RegisterModalsAndViewers(unittest.TestCase):
    def test_phase2b2_dictionary_namespaces_and_client_exposure(self):
        en = i18n.TRANSLATIONS["en"]
        ar = i18n.TRANSLATIONS["ar"]
        for prefix in ("form.", "noc.", "ltr.", "pr.", "hse_form.", "drive."):
            en_keys = {k for k in en if k.startswith(prefix)}
            ar_keys = {k for k in ar if k.startswith(prefix)}
            self.assertTrue(len(en_keys) > 0, f"Missing {prefix} keys")
            self.assertEqual(en_keys, ar_keys, f"Parity mismatch for {prefix}")
        client_ar = i18n.get_client_dict("ar")
        self.assertIn("form.add_title", client_ar)
        self.assertIn("noc.part_a_title", client_ar)
        self.assertIn("ltr.dir_sent", client_ar)
        self.assertIn("pr.badge_pending_po", client_ar)
        self.assertIn("hse_form.region_head_face", client_ar)
        self.assertIn("drive.dv_fullscreen", client_ar)

    def test_register_modals_and_viewers_en_ar_rtl_and_canonical_values(self):
        import html_render
        user = {"id": "u1", "username": "testuser", "name": "Test User", "role": "admin", "workspace": "ENG"}
        proj = {"id": "p1", "code": "PRJ-01", "name": "Test Project", "workspace_id": "ENG", "drive_folder_id": "fld1", "status_colors": {}}
        dts = [{"id": "SD", "code": "SD", "name": "Shop Drawings", "visible": True}]
        with patch("html_render.can_edit", return_value=True), \
             patch("html_render.is_project_admin", return_value=True), \
             patch("html_render.get_allowed_doc_type_filter", return_value=None), \
             patch("db.can_user_download", return_value=True), \
             patch("db.get_projects_for_user", return_value=[proj]), \
             patch("db.get_doc_types", return_value=dts), \
             patch("db.get_logo", return_value=None), \
             patch("db.get_user_workspaces", return_value=[{"id": "ENG", "code": "ENG", "name": "Engineering & Projects", "name_ar": "المشروعات الهندسية", "icon": "🏗️"}]):
            html_en = html_render.render_register(user, proj, active_ws_id="ENG", lang="en")
            self.assertIn('<span id="rec-title">Add Document</span>', html_en)
            self.assertIn("Batch Upload Files to Google Drive", html_en)
            self.assertIn("Batch Extinguisher / Equipment Inspection", html_en)

            html_ar = html_render.render_register(user, proj, active_ws_id="ENG", lang="ar")
            # Record Add/Edit, Thread & Timeline modals
            self.assertIn('<span id="rec-title">إضافة مستند</span>', html_ar)
            self.assertIn('<span id="thread-title">سلسلة المراسلات</span>', html_ar)
            self.assertIn('<span id="timeline-title">الجدول الزمني للمراسلات</span>', html_ar)

            # Batch Drive Upload & HSE FEI Batch Inspect modals
            self.assertIn("☁ الرفع الجماعي للملفات إلى Google Drive", html_ar)
            self.assertIn("مجلد Drive المستهدف:", html_ar)
            self.assertIn("Part A: إشعار التغيير", html_ar)
            self.assertIn("🧯 الفحص الجماعي للطفايات / المعدات (FEI)", html_ar)
            # Canonical English <option value="..."> preserved with Arabic display text
            self.assertIn('<option value="Valid">ساري</option>', html_ar)
            self.assertIn('<option value="Needs Maintenance">يحتاج صيانة</option>', html_ar)
            self.assertIn('<option value="Under Refill">قيد إعادة التعبئة</option>', html_ar)

            # In-App Drive Viewer
            self.assertIn("⛶ ملء الشاشة", html_ar)
            self.assertIn("📁 الملفات", html_ar)
            self.assertIn("👁️ معاينة", html_ar)
            self.assertIn('placeholder="🔍 ابحث عن ملفات في المجلد..."', html_ar)

            # Mandatory RTL CSS & JS fixes for .letter-flow and #doc-viewer-modal
            self.assertIn('html[dir="rtl"] .letter-flow{padding-left:0;padding-right:10px}', html_ar)
            self.assertIn('html[dir="rtl"] .letter-flow-line{left:auto;right:13px}', html_ar)
            self.assertIn('html[dir="rtl"] .letter-node{padding-left:0;padding-right:30px}', html_ar)
            self.assertIn('html[dir="rtl"] .letter-dot{left:auto;right:8px}', html_ar)
            self.assertIn('html[dir="rtl"] .letter-card{border-left:1px solid #d7e0e6;border-right:4px solid #d7e0e6}', html_ar)
            self.assertIn("isRtl?`margin-right:${margin}px`:`margin-left:${margin}px`", html_ar)
            self.assertIn('html[dir="rtl"] #dv-side { border-right: none !important; border-left: 1px solid var(--bd) !important; }', html_ar)

            # BiDi isolation & canonical body-map preservation
            self.assertIn('id="batch-drv-target-path-txt" class="bidi-ltr"', html_ar)
            self.assertIn('class="letter-card-ref bidi-ltr"', html_ar)
            self.assertIn('class="letter-date-badge bidi-ltr"', html_ar)
            self.assertIn('class="pri-qty bidi-ltr"', html_ar)
            self.assertIn("setDynamicFieldValue('bodyPart', meta.label);", html_ar)
            self.assertIn("if(regEl) regEl.value = meta.id;", html_ar)


class TestPhase3StandaloneViews(unittest.TestCase):
    """Verify Phase 3 standalone public views (/magic/<pid> and /qr/<qr_token>) in EN/LTR and AR/RTL."""

    def setUp(self):
        self.app = app
        self.app.config["TESTING"] = True
        self.client = self.app.test_client()

    def test_phase3_dictionary_parity_magic_and_qr(self):
        en = i18n.TRANSLATIONS["en"]
        ar = i18n.TRANSLATIONS["ar"]
        magic_keys = [k for k in en if k.startswith("magic.")]
        qr_keys = [k for k in en if k.startswith("qr.")]
        self.assertGreaterEqual(len(magic_keys), 15)
        self.assertGreaterEqual(len(qr_keys), 12)
        for k in magic_keys + qr_keys:
            self.assertIn(k, ar, f"Missing Arabic key: {k}")
            self.assertTrue(str(en[k]).strip(), f"Empty EN key: {k}")
            self.assertTrue(str(ar[k]).strip(), f"Empty AR key: {k}")
        self.assertIn("HSE", ar["qr.badge_header"])
        self.assertIn("FEI", ar["qr.reg_type_fei"])
        self.assertIn("PTW", ar["qr.reg_type_ptw"])

    def test_magic_master_view_en_and_ar_rtl_and_bidi(self):
        proj = {"id": "p1", "code": "PRJ-01", "name": "Tower A"}
        with patch("app._generate_master_magic_token", return_value="tok123"), \
             patch("db.get_projects", return_value=[proj]), \
             patch("app._head_icon_tags", return_value=""):
            # Forbidden without valid token
            bad = self.client.get("/magic/p1?token=wrong")
            self.assertEqual(bad.status_code, 403)

            # English (default / explicit)
            res_en = self.client.get("/magic/p1?token=tok123&lang=en")
            self.assertEqual(res_en.status_code, 200)
            html_en = res_en.get_data(as_text=True)
            self.assertIn('<html lang="en" dir="ltr">', html_en)
            self.assertIn("Smart Project Digest", html_en)
            self.assertIn("Documents Received", html_en)
            self.assertIn("switchDcrLang('ar')", html_en)
            self.assertIn("عربي", html_en)

            # Arabic RTL
            res_ar = self.client.get("/magic/p1?token=tok123&lang=ar")
            self.assertEqual(res_ar.status_code, 200)
            html_ar = res_ar.get_data(as_text=True)
            self.assertIn('<html lang="ar" dir="rtl">', html_ar)
            self.assertIn("/static/fonts/Cairo-Regular.ttf", html_ar)
            self.assertIn('html[dir="rtl"]', html_ar)
            self.assertIn("switchDcrLang('en')", html_ar)
            self.assertIn(">EN</button>", html_ar)
            self.assertIn("الملخص الذكي للمشروع", html_ar)
            self.assertIn("المستندات المستلمة", html_ar)
            self.assertIn("المستندات المرسلة", html_ar)
            self.assertIn("المستندات المجاب عنها", html_ar)
            self.assertIn("الخطابات الواردة", html_ar)
            self.assertIn("الخطابات الصادرة", html_ar)
            self.assertIn("عرض الكل", html_ar)
            self.assertIn("ملخص مشروع ذكي للعرض فقط • منصة DCR", html_ar)
            self.assertIn('<span class="bidi-ltr">PRJ-01</span>', html_ar)
            self.assertIn('class="bidi-ltr"', html_ar)
            self.assertIn("dcr_lang=ar", res_ar.headers.get("Set-Cookie", ""))

    def test_public_qr_verify_en_and_ar_valid_and_invalid(self):
        fei_row = {
            "id": "rec-fei-1",
            "project_id": "p-hse",
            "dt_id": "dt-fei",
            "proj_name": "Safety Plant",
            "proj_code": "HSE-01",
            "workspace_id": "HSE",
            "dt_code": "FEI",
            "data": {
                "docNo": "FEI-2026-009",
                "legacySubtype": "CO2",
                "location": "Gate B",
                "extinguisherStatus": "Valid",
                "expiryDate": "2027-06-01",
                "lastInspection": "2026-10-15",
            },
        }
        ptw_row = {
            "id": "rec-ptw-1",
            "project_id": "p-hse",
            "dt_id": "dt-ptw",
            "proj_name": "Safety Plant",
            "proj_code": "HSE-01",
            "workspace_id": "HSE",
            "dt_code": "PTW",
            "data": {
                "docNo": "PTW-2026-042",
                "legacySubtype": "Hot Work",
                "location": "Zone 3",
                "status": "Open",
                "validUntil": "2027-12-31",
                "issuedDate": "2026-11-01",
            },
        }
        with patch("html_render._head_icon_tags", return_value=""):
            # Valid FEI in Arabic
            with patch("db.q", return_value=fei_row):
                res_ar = self.client.get("/qr/validtoken123?lang=ar")
                self.assertEqual(res_ar.status_code, 200)
                html_ar = res_ar.get_data(as_text=True)
                self.assertIn('<html lang="ar" dir="rtl">', html_ar)
                self.assertIn("/static/fonts/Cairo-Regular.ttf", html_ar)
                self.assertIn('html[dir="rtl"]', html_ar)
                self.assertIn("switchDcrLang('en')", html_ar)
                self.assertIn("التحقق الميداني للسلامة (HSE)", html_ar)
                self.assertIn("فحص معدات الحريق والسلامة (FEI)", html_ar)
                self.assertIn("المشروع", html_ar)
                self.assertIn("نوع السجل", html_ar)
                self.assertIn("رقم السجل", html_ar)
                self.assertIn("تاريخ الصلاحية / الانتهاء", html_ar)
                self.assertIn("تاريخ آخر فحص", html_ar)
                self.assertIn("ساري", html_ar)
                self.assertIn("شارة تحقق سلامة (HSE) للعرض فقط • منصة DCR للسلامة", html_ar)
                self.assertIn('<span class="bidi-ltr">FEI-2026-009</span>', html_ar)
                self.assertIn('<span class="bidi-ltr">HSE-01</span>', html_ar)
                self.assertIn('<span class="bidi-ltr">2027-06-01</span>', html_ar)
                self.assertIn('<span class="bidi-ltr">2026-10-15</span>', html_ar)

            # Valid PTW in English
            with patch("db.q", return_value=ptw_row):
                res_en = self.client.get("/qr/validtoken456?lang=en")
                self.assertEqual(res_en.status_code, 200)
                html_en = res_en.get_data(as_text=True)
                self.assertIn('<html lang="en" dir="ltr">', html_en)
                self.assertIn("HSE Field Verification", html_en)
                self.assertIn("Permit to Work (PTW)", html_en)
                self.assertIn(">Open<", html_en)
                self.assertIn("switchDcrLang('ar')", html_en)
                self.assertIn("عربي", html_en)

            # Invalid / Not Found in Arabic
            with patch("db.q", return_value=None):
                res_nf = self.client.get("/qr/missingtoken1?lang=ar")
                self.assertEqual(res_nf.status_code, 404)
                html_nf = res_nf.get_data(as_text=True)
                self.assertIn('<html lang="ar" dir="rtl">', html_nf)
                self.assertIn("/static/fonts/Cairo-Regular.ttf", html_nf)
                self.assertIn("سجل رمز الاستجابة (QR) غير موجود", html_nf)
                self.assertIn("switchDcrLang('en')", html_nf)


class TestI18nVisualQAFixes(unittest.TestCase):
    """Regression coverage for I18N-FIX-01 visual QA defects (Defects 1-7)."""

    def test_bidi_ltr_table_cell_and_docno_nowrap_css(self):
        self.assertIn("td.bidi-ltr, th.bidi-ltr {", html_render.BASE_CSS)
        self.assertIn("display: table-cell;", html_render.BASE_CSS)
        self.assertIn("#regtbl td.docno-cell, #regtbl td:nth-child(3)", html_render.BASE_CSS)
        self.assertIn("white-space: nowrap !important;", html_render.BASE_CSS)
        self.assertIn("min-width: 130px;", html_render.BASE_CSS)

    def test_dt_name_coverage_and_no_raw_key_leakage(self):
        import db
        import hse_defaults

        required_codes = ("DS", "WN", "QS", "CEV", "EVAC", "WKR", "SD", "MS", "IR", "MIR",
                          "RFI", "CVI", "PCQ", "SI", "EI", "SVP", "ABD", "NCR", "NOC", "PR", "LTR", "OM", "IOM")
        for code in required_codes:
            self.assertIn(f"dt_name.{code}", i18n.TRANSLATIONS["en"])
            self.assertIn(f"dt_name.{code}", i18n.TRANSLATIONS["ar"])
            ar_val = i18n.t_dt(code, code, "ar")
            self.assertFalse(ar_val.startswith("dt_name."), f"Raw dt_name key returned for {code}: {ar_val}")
            self.assertIn(code, ar_val)

        for code, name, _ in (db.DEFAULT_DOC_TYPES + db.DEFAULT_OM_DOC_TYPES + list(hse_defaults.HSE_REGISTERS)):
            ar_val = i18n.t_dt(code, name, "ar")
            self.assertFalse(ar_val.startswith("dt_name."), f"Raw key for {code}")
            self.assertNotEqual(ar_val, name, f"Expected Arabic translation for built-in {code}")

        # Unknown doc type code never leaks raw 'dt_name.*' key
        self.assertEqual(i18n.t_dt("CUSTOMX", "Custom Register", "ar"), "Custom Register")
        self.assertEqual(i18n.t_dt("CUSTOMX", "", "ar"), "CUSTOMX")

        # Client JS window.tDt and HSE register cards never leak raw 'dt_name.*'
        self.assertIn("window.tDt(r.code, r.name)", html_render.HSE_DASHBOARD_JS)
        self.assertNotIn("_hseT('dt_name.'", html_render.HSE_DASHBOARD_JS)

    def test_hse_11_register_submetrics_localized(self):
        for k in ("hse.reg_open_closed", "hse.reg_inc", "hse.reg_fei", "hse.reg_mhr",
                  "hse.reg_wkr", "hse.reg_jsa", "hse.reg_tbt", "hse.reg_cev", "hse.reg_evac"):
            self.assertIn(k, i18n.TRANSLATIONS["en"])
            self.assertIn(k, i18n.TRANSLATIONS["ar"])
        self.assertIn("LTI:", i18n.TRANSLATIONS["ar"]["hse.reg_inc"])
        self.assertIn("MTC:", i18n.TRANSLATIONS["ar"]["hse.reg_inc"])
        self.assertIn("FAC:", i18n.TRANSLATIONS["ar"]["hse.reg_inc"])
        self.assertIn("_formatHseRegSubtitle", html_render.HSE_DASHBOARD_JS)

    def test_bc_status_translation_and_no_comma_split(self):
        ar_nospace = i18n.t_status("B,C - Approved & Resubmit", "ar")
        ar_space = i18n.t_status("B, C - Approved & Resubmit", "ar")
        self.assertEqual(ar_nospace, "B,C - معتمد ويعاد التقديم")
        self.assertEqual(ar_space, "B,C - معتمد ويعاد التقديم")
        self.assertEqual(i18n.t_status("B, C - Approved & Resubmit", "en"), "B, C - Approved & Resubmit")

        user = {"username": "admin_u", "role": "superadmin"}
        proj = {"id": "p1", "code": "CP01", "name": "Alpha", "workspace_id": "ENG", "status_colors": {}}
        with patch("html_render.can_edit", return_value=True), \
             patch("html_render.is_project_admin", return_value=True), \
             patch("html_render.get_allowed_doc_type_filter", return_value=None), \
             patch("db.can_user_download", return_value=True), \
             patch("db.get_projects_for_user", return_value=[proj]), \
             patch("db.get_doc_types", return_value=[]), \
             patch("db.get_logo", return_value=None), \
             patch("db.get_user_workspaces", return_value=[]):
            reg_ar = html_render.render_register(user, proj, active_ws_id="ENG", lang="ar")
            self.assertIn("splitStatusValues(val)", reg_ar)
            self.assertNotIn("String(val).split(',').map", reg_ar)
            # Inline Edit button sits inside topbar-proj-info; no stray empty #projbar row
            self.assertIn("topbar-proj-edit-btn", reg_ar)
            self.assertNotIn('<div id="projbar">', reg_ar)

    def test_topbar_rtl_nowrap_and_ai_drawer_fixes(self):
        self.assertIn('html[dir="rtl"] #topbar .tb-btn', html_render.BASE_CSS)
        self.assertIn("white-space: nowrap;", html_render.BASE_CSS)
        # AI Drawer active project reads PROJ_NAME/PROJ_CODE/.pf-val rather than .pf-lbl
        self.assertIn(".pf-val[data-key=\"code\"]", html_render.AI_DRAWER_COMPONENT)
        self.assertNotIn("topbarInfo.querySelector('span')", html_render.AI_DRAWER_COMPONENT)
        # AI Drawer markdown list strips '- ' cleanly (no raw '\\s*' double-backslash bug producing '• -')
        self.assertIn(r"line.replace(/^[-*•]\s*/, '')", html_render.AI_DRAWER_COMPONENT)
        self.assertNotIn(r"line.replace(/^[-*•]\\s*/, '')", html_render.AI_DRAWER_COMPONENT)

    def test_app_renderers_do_not_swallow_typeerror(self):
        import inspect
        import app as app_mod
        for fn in (app_mod.login, app_mod.launchpad_hub, app_mod.home, app_mod.register):
            src = inspect.getsource(fn)
            self.assertNotIn("except TypeError", src)


if __name__ == "__main__":
    unittest.main()
