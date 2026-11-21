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


if __name__ == "__main__":
    unittest.main()
