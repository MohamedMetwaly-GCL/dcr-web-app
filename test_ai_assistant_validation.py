"""test_ai_assistant_validation.py - Comprehensive test suite for DCR AI Assistant.

Validates:
1. Blueprint and endpoint mapping (/api/ai/query).
2. Authentication (rejects unauthenticated requests with 403 LOGIN_REQUIRED).
3. Payload validation (rejects empty prompt with 400).
4. Missing API key handling (graceful 503 with exact error string).
5. RBAC isolation (forbidden projects return 403; admin allowed across all).
6. Context generator (_build_ai_context builds KPIs, NOC financials, overdues).
7. Successful Gemini API invocation returns { "reply": markdown_text }.
8. Frontend verification (drawer component & topbar buttons in both dashboard and register).
"""
import os
import unittest
from unittest.mock import patch

os.environ["SECRET_KEY"] = "test-secret-key-1234567890"

from app import app
from blueprints.ai import _build_ai_context


class TestAiAssistantValidation(unittest.TestCase):

    def setUp(self):
        self.client = app.test_client()

    def test_01_unauthenticated_request_rejected(self):
        """Unauthenticated requests must be rejected with 403 LOGIN_REQUIRED."""
        with patch("app.current_user", return_value=None), \
             patch("blueprints.ai.current_user", return_value=None):
            res = self.client.post("/api/ai/query", json={"prompt": "Hello"})
            self.assertEqual(res.status_code, 403)
            data = res.get_json()
            self.assertEqual(data.get("error"), "LOGIN_REQUIRED")
        print("PASS: Unauthenticated requests rejected with 403 LOGIN_REQUIRED.")

    def test_02_missing_prompt_returns_400(self):
        """Empty or missing prompt must return 400."""
        user = {"username": "eng_viewer", "role": "viewer"}
        with patch("app.current_user", return_value=user), \
             patch("blueprints.ai.current_user", return_value=user), \
             patch.dict(os.environ, {"GEMINI_API_KEY": "AIzaSyTestFakeKey"}):
            res = self.client.post("/api/ai/query", json={"prompt": "   "})
            self.assertEqual(res.status_code, 400)
            data = res.get_json()
            self.assertIn("required", data.get("error", "").lower())
        print("PASS: Missing prompt rejected with 400.")

    def test_03_missing_gemini_api_key_returns_503(self):
        """If GEMINI_API_KEY is not set, must return graceful 503."""
        user = {"username": "admin_user", "role": "admin"}
        with patch("app.current_user", return_value=user), \
             patch("blueprints.ai.current_user", return_value=user), \
             patch.dict(os.environ, {}, clear=True):
            os.environ["SECRET_KEY"] = "test-secret-key-1234567890"
            if "GEMINI_API_KEY" in os.environ:
                del os.environ["GEMINI_API_KEY"]
            res = self.client.post("/api/ai/query", json={"prompt": "List overdue submittals"})
            self.assertEqual(res.status_code, 503)
            data = res.get_json()
            self.assertEqual(data.get("error"), "AI Assistant is not configured on this instance.")
        print("PASS: Missing API key returns graceful 503.")

    def test_04_rbac_forbidden_project_returns_403(self):
        """User requesting a project they do not have access to receives 403 Forbidden."""
        user = {"username": "viewer_user", "role": "viewer"}
        with patch("app.current_user", return_value=user), \
             patch("blueprints.ai.current_user", return_value=user), \
             patch("blueprints.ai.can_view_project", return_value=False), \
             patch("app.can_view_project", return_value=False), \
             patch.dict(os.environ, {"GEMINI_API_KEY": "AIzaSyTestFakeKey"}):
            res = self.client.post("/api/ai/query", json={
                "prompt": "Show status",
                "project_id": "UNAUTHORIZED_PROJECT"
            })
            self.assertEqual(res.status_code, 403)
            data = res.get_json()
            self.assertEqual(data.get("error"), "Forbidden")
        print("PASS: RBAC blocks queries to unauthorized projects with 403 Forbidden.")

    def test_05_rbac_user_with_no_projects(self):
        """User with no assigned projects receives graceful response without error."""
        user = {"username": "empty_user", "role": "viewer"}
        with patch("app.current_user", return_value=user), \
             patch("blueprints.ai.current_user", return_value=user), \
             patch("blueprints.ai.get_allowed_project_ids", return_value=[]), \
             patch.dict(os.environ, {"GEMINI_API_KEY": "AIzaSyTestFakeKey"}):
            res = self.client.post("/api/ai/query", json={"prompt": "Any updates?"})
            self.assertEqual(res.status_code, 200)
            data = res.get_json()
            self.assertIn("do not have access to any projects", data.get("reply", ""))
        print("PASS: User with zero accessible projects handled gracefully.")

    def test_06_context_generation_with_noc_and_overdue(self):
        """Verify _build_ai_context gathers project details, KPIs, NOC totals, and overdues."""
        mock_projects = [{"id": "p1", "name": "CFC Ph2", "code": "PEM-058"}]
        mock_stats = [{
            "id": "p1", "code": "PEM-058", "name": "CFC Ph2",
            "total": 650, "approved": 450, "pending": 120,
            "rejected": 80, "overdue": 15, "pct": 69
        }]
        mock_nocs = [
            {
                "project_id": "p1", "proj_code": "PEM-058", "proj_name": "CFC Ph2",
                "data": {
                    "docNo": "NOC-058-001",
                    "title": "Chiller Additional Power Supply",
                    "submittedCost": "125000",
                    "finalApprovedCost": "110000",
                    "partDStatus": "Approved"
                }
            },
            {
                "project_id": "p1", "proj_code": "PEM-058", "proj_name": "CFC Ph2",
                "data": {
                    "docNo": "NOC-058-002",
                    "title": "Pump Relocation",
                    "submittedCost": "45000",
                    "finalApprovedCost": "0",
                    "partDStatus": "Under Review"
                }
            }
        ]
        mock_overdue = [
            {
                "project_id": "p1", "dt_code": "MS", "dt_name": "Material Submittal",
                "docNo": "MS-058-042", "title": "BMS Sensors Specification",
                "days_overdue": 18, "status": "Pending", "issuedDate": "2026-08-10"
            }
        ]

        with patch("db.q") as mock_q, \
             patch("db.get_dashboard_stats", return_value=mock_stats), \
             patch("db.get_overdue_records", return_value=mock_overdue):
            def q_side_effect(sql, params=()):
                if "FROM projects" in sql:
                    return mock_projects
                if "UPPER(r.dt_id) = 'NOC'" in sql:
                    return mock_nocs
                return []
            mock_q.side_effect = q_side_effect

            ctx = _build_ai_context(["p1"], user_prompt="What is the NOC cost?", active_tab="NOC")
            self.assertIn("PEM-058", ctx)
            self.assertIn("110,000.00 EGP", ctx)
            self.assertIn("125,000.00 EGP", ctx)
            self.assertIn("MS-058-042", ctx)
            self.assertIn("18 days", ctx)
        print("PASS: Context generation compiles KPIs, NOC financials, and overdue submittals accurately.")

    def test_07_successful_query_returns_reply(self):
        """End-to-end endpoint call with mocked Gemini client returns markdown reply."""
        user = {"username": "admin_user", "role": "admin"}
        mock_gemini_reply = "### Summary\n- Total approved NOC value: **110,000.00 EGP**\n- 1 overdue submittal pending."

        with patch("app.current_user", return_value=user), \
             patch("blueprints.ai.current_user", return_value=user), \
             patch("blueprints.ai.can_view_project", return_value=True), \
             patch("blueprints.ai._build_ai_context", return_value="Context data"), \
             patch("blueprints.ai._call_gemini_api", return_value=(mock_gemini_reply, None)), \
             patch.dict(os.environ, {"GEMINI_API_KEY": "AIzaSyTestFakeKey"}):
            res = self.client.post("/api/ai/query", json={
                "prompt": "Give me an overview",
                "project_id": "p1"
            })
            self.assertEqual(res.status_code, 200)
            data = res.get_json()
            self.assertIn("reply", data)
            self.assertIn("110,000.00 EGP", data["reply"])
        print("PASS: Successful AI query returns structured reply.")

    def test_08_frontend_drawer_and_buttons_rendered(self):
        """Verify AI drawer markup, quick action chips, and topbar button exist in templates."""
        from html_render import AI_DRAWER_COMPONENT, render_dashboard, render_register

        self.assertIn('id="ai-assistant-drawer"', AI_DRAWER_COMPONENT)
        self.assertIn('ai-chip', AI_DRAWER_COMPONENT)
        self.assertIn('Overdue Submittals', AI_DRAWER_COMPONENT)
        self.assertIn('Approved NOCs Total', AI_DRAWER_COMPONENT)
        self.assertIn('Status Summary', AI_DRAWER_COMPONENT)

        # Dashboard rendering check
        user = {"username": "admin", "role": "admin"}
        with app.test_request_context("/"):
            dash_html = render_dashboard(user)
            self.assertIn("toggleAiDrawer()", dash_html)
            self.assertIn('id="ai-assistant-drawer"', dash_html)
            self.assertIn('ai-btn', dash_html)

        # Register rendering check
        proj = {"id": "test_proj", "name": "Test Project", "code": "TST-01"}
        with app.test_request_context("/app?p=test_proj"), \
             patch("db.get_doc_types", return_value=[]), \
             patch("db.get_columns", return_value=[]), \
             patch("db.get_records", return_value=[]), \
             patch("db.get_projects", return_value=[proj]), \
             patch("db.get_logo", return_value=None), \
             patch("auth.can_edit", return_value=True):
            reg_html = render_register(user, proj)
            self.assertIn("toggleAiDrawer()", reg_html)
            self.assertIn('id="ai-assistant-drawer"', reg_html)
            self.assertIn('ai-btn', reg_html)
        print("PASS: Frontend drawer and navbar buttons verified in Dashboard and Register views.")

    def test_09_bilingual_keyword_and_doc_type_search(self):
        """Verify Arabic keyword 'محابس' maps to 'valve', detects 'MS', and injects matching submittals."""
        mock_projects = [{"id": "p_cfc", "name": "CFC Ph2", "code": "PEM-058"}]
        mock_stats = []
        mock_overdue = []
        mock_submittals = [
            {
                "id": "rec_120",
                "project_id": "p_cfc",
                "proj_code": "PEM-058",
                "doc_type": "MS",
                "doc_no": "MS-CY002P608-00120",
                "title": "Double Regulating Valves",
                "status": "B - Approved As Noted",
                "issued_date": "12-05-2022",
                "actual_reply": "20-05-2022"
            }
        ]

        with patch("db.q") as mock_q, \
             patch("db.get_dashboard_stats", return_value=mock_stats), \
             patch("db.get_overdue_records", return_value=mock_overdue):
            def q_side_effect(sql, params=()):
                if "FROM projects" in sql:
                    return mock_projects
                if "UPPER(r.dt_id) = 'NOC'" in sql:
                    return []
                if "UPPER(r.dt_id) = ANY" in sql or "ORDER BY" in sql:
                    # Targeted submittal search query
                    return mock_submittals
                return []
            mock_q.side_effect = q_side_effect

            user_query = "شوفلي معتمد MS محابس ايه في CFC"
            ctx = _build_ai_context(["p_cfc"], user_prompt=user_query)

            self.assertIn("MATCHING SUBMITTALS IN REGISTER", ctx)
            self.assertIn("MS-CY002P608-00120", ctx)
            self.assertIn("Double Regulating Valves", ctx)
            self.assertIn("B - Approved As Noted", ctx)
        print("PASS: Bilingual keyword mapping, doc type detection, and matching submittal injection verified.")

    def test_10_simple_greeting_fast_response(self):
        """Verify simple greetings like 'ازيك' return immediate response without DB context."""
        user = {"username": "admin_user", "role": "admin"}
        with patch("app.current_user", return_value=user), \
             patch("blueprints.ai.current_user", return_value=user), \
             patch("blueprints.ai.get_allowed_project_ids", return_value=["p1"]), \
             patch("blueprints.ai._build_ai_context") as mock_build_ctx, \
             patch("blueprints.ai._call_gemini_api", return_value=("أهلاً بك يا باشمهندس!", None)), \
             patch.dict(os.environ, {"GEMINI_API_KEY": "AIzaSyTestFakeKey"}):
            res = self.client.post("/api/ai/query", json={"prompt": "ازيك"})
            self.assertEqual(res.status_code, 200)
            data = res.get_json()
            self.assertIn("reply", data)
            self.assertIn("أهلاً بك", data["reply"])
            # Ensure heavy _build_ai_context was NOT called for simple greeting
            mock_build_ctx.assert_not_called()
        print("PASS: Simple greeting short-circuit verified (no heavy DB queries called).")

    def test_11_call_gemini_api_custom_instruction(self):
        """Verify _call_gemini_api accepts custom_instruction parameter without TypeError."""
        from blueprints.ai import _call_gemini_api
        with patch("google.genai.Client") as mock_client_cls, \
             patch.dict(os.environ, {"GEMINI_API_KEY": "AIzaSyTestFakeKey"}):
            mock_client = mock_client_cls.return_value
            mock_resp = unittest.mock.MagicMock()
            mock_resp.text = "مرحبا يا هندسة"
            mock_client.models.generate_content.return_value = mock_resp

            reply, err = _call_gemini_api("ازيك", "context", custom_instruction="Custom greeting")
            self.assertEqual(reply, "مرحبا يا هندسة")
            self.assertIsNone(err)
        print("PASS: _call_gemini_api with custom_instruction parameter verified.")

    def test_12_format_fallback_response_direct_records(self):
        """Verify _format_fallback_response_from_context builds markdown table from context."""
        from blueprints.ai import _format_fallback_response_from_context
        sample_context = (
            "### 1. PROJECTS ACCESSIBLE & IDENTIFIED:\n"
            "- Code: PEM-058 | Name: CFC Ph2 | ID: p1\n\n"
            "### 5. MATCHING SUBMITTALS IN REGISTER:\n"
            "Found 2 matching submittals based on query keywords and filters:\n"
            "- DocNo: MS-058-012 | Title: Butterfly Valves | DocType: MS | Status: Approved | Date: 2026-04-12 | Project: PEM-058\n"
            "- DocNo: MS-058-014 | Title: Check Valves | DocType: MS | Status: Approved with Comments | Date: 2026-05-02 | Project: PEM-058\n"
        )
        res = _format_fallback_response_from_context("CFC محابس ايه في MS شوفلي معتمد", sample_context)
        self.assertIsNotNone(res)
        self.assertIn("MS-058-012", res)
        self.assertIn("Butterfly Valves", res)
        self.assertIn("✅ **Approved**", res)
        self.assertIn("🟡 **Approved with Comments**", res)
        self.assertIn("503 High Demand", res)
        print("PASS: Direct context formatting produces structured Markdown table on 503.")

    def test_13_gemini_503_fallback_to_direct_context_endpoint(self):
        """When Gemini throws 503 UNAVAILABLE, query endpoint falls back to direct DB formatting."""
        user = {"username": "admin_user", "role": "admin"}
        sample_context = (
            "### 5. MATCHING SUBMITTALS IN REGISTER:\n"
            "- DocNo: MS-058-012 | Title: Butterfly Valves | DocType: MS | Status: Approved | Date: 2026-04-12 | Project: PEM-058\n"
        )
        with patch("app.current_user", return_value=user), \
             patch("blueprints.ai.current_user", return_value=user), \
             patch("blueprints.ai.can_view_project", return_value=True), \
             patch("blueprints.ai._build_ai_context", return_value=sample_context), \
             patch("blueprints.ai._call_gemini_api", return_value=(None, "Gemini API Error: 503 UNAVAILABLE")), \
             patch.dict(os.environ, {"GEMINI_API_KEY": "AIzaSyTestFakeKey"}):
            res = self.client.post("/api/ai/query", json={
                "prompt": "CFC محابس ايه في MS شوفلي معتمد",
                "project_id": "p1"
            })
            self.assertEqual(res.status_code, 200)
            data = res.get_json()
            self.assertIn("reply", data)
            self.assertIn("MS-058-012", data["reply"])
            self.assertIn("Butterfly Valves", data["reply"])
        print("PASS: Endpoint gracefully recovers from Gemini 503 using direct DB context fallback.")


if __name__ == "__main__":
    unittest.main()
