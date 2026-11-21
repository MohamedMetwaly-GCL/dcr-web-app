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


if __name__ == "__main__":
    unittest.main()
