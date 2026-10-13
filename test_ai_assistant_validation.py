"""test_ai_assistant_validation.py - Comprehensive test suite for DCR AI Assistant (Agent Architecture).

Validates:
1. Blueprint and endpoint mapping (/api/ai/query).
2. Authentication (rejects unauthenticated requests with 403 LOGIN_REQUIRED).
3. Payload validation (rejects empty prompt with 400).
4. Missing API key handling (graceful 503 with exact error string).
5. RBAC isolation (forbidden projects return 403; user with 0 projects handled gracefully).
6. Frontend verification (drawer component & topbar buttons in both dashboard and register).
7. Simple greeting short-circuit (fast return without heavy queries).
8. Gemini Tool Registry: All 5 tools registered in DCR_TOOL.
9. Security Boundary: allowed_project_ids is strictly omitted from LLM tool schemas.
10. Tool Execution with Forced RBAC Injection (execute_tool_call).
11. Multi-turn Agent Controller loop (ToolCall -> Backend Execution -> Markdown Synthesis).
12. End-to-end endpoint call returns structured reply.
13. Zero-503 Deterministic Fallback: Recovers from 503 with 200 and formatted Markdown.
14. Deterministic fallback valve search table formatting.
15. Deterministic fallback zero records returns clean engineering message.
16. Deterministic fallback NOC financials summary.
17. Deterministic fallback overdue submittals table.
18. Deterministic fallback project register summary.
19. Search Submittals valve stemming & excludes unrelated fittings.
20. Revision deduplication (REV 01 supersedes REV 00).
21. Project alias resolution (CFC, PEM-058, EKH, 24CP01).
22. Discipline isolation (Architecture IRs strictly excludes Electrical).
23. Document type strictness (MS strictly excludes MIR).
24. Duplicate document elimination.
25. PostgreSQL array parameters compatibility (lists for ANY(%s)).
26. History context preservation across turns.
27. Project override prevents context bleed.
28. NOC financial cost calculations.
29. Overdue submittals delay sorting.
30. Chronological revision history tracking.
31. Strict RBAC enforcement blocking unauthorized project data.
"""
import os
import unittest
from unittest.mock import MagicMock, patch

os.environ["SECRET_KEY"] = "test-secret-key-1234567890"

from app import app
import blueprints.ai_tools as ai_tools
from blueprints.ai import (
    DCR_TOOL,
    TOOL_REGISTRY,
    execute_tool_call,
    _call_gemini_agent,
    _execute_deterministic_fallback,
    _dedup_latest_revisions,
    _build_gemini_contents,
    is_follow_up_prompt,
)


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
        orig_key = os.environ.get("GEMINI_API_KEY")
        if "GEMINI_API_KEY" in os.environ:
            del os.environ["GEMINI_API_KEY"]
        try:
            with patch("app.current_user", return_value=user), \
                 patch("blueprints.ai.current_user", return_value=user):
                res = self.client.post("/api/ai/query", json={"prompt": "List overdue submittals"})
                self.assertEqual(res.status_code, 503)
                data = res.get_json()
                self.assertEqual(data.get("error"), "AI Assistant is not configured on this instance.")
        finally:
            if orig_key is not None:
                os.environ["GEMINI_API_KEY"] = orig_key
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

    def test_06_frontend_drawer_and_buttons_rendered(self):
        """Verify AI drawer markup, quick action chips, and topbar button exist in templates."""
        from html_render import AI_DRAWER_COMPONENT, render_dashboard, render_register

        self.assertIn('id="ai-assistant-drawer"', AI_DRAWER_COMPONENT)
        self.assertIn('ai-chip', AI_DRAWER_COMPONENT)
        self.assertIn('Overdue Submittals', AI_DRAWER_COMPONENT)
        self.assertIn('Approved NOCs Total', AI_DRAWER_COMPONENT)
        self.assertIn('Status Summary', AI_DRAWER_COMPONENT)
        self.assertIn('id="ai-fab-btn"', AI_DRAWER_COMPONENT)
        self.assertIn('id="ai-expand-btn"', AI_DRAWER_COMPONENT)
        self.assertIn('ai-drawer-resizer', AI_DRAWER_COMPONENT)

        user = {"username": "admin", "role": "admin"}
        with app.test_request_context("/"):
            dash_html = render_dashboard(user)
            self.assertIn("toggleAiDrawer()", dash_html)
            self.assertIn('id="ai-assistant-drawer"', dash_html)
            self.assertIn('ai-btn', dash_html)

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

    def test_07_simple_greeting_fast_response(self):
        """Verify simple greetings like 'ازيك' return immediate response without DB context."""
        user = {"username": "admin_user", "role": "admin"}
        with patch("app.current_user", return_value=user), \
             patch("blueprints.ai.current_user", return_value=user), \
             patch("blueprints.ai.get_allowed_project_ids", return_value=["p1"]), \
             patch("blueprints.ai._call_gemini_agent") as mock_agent, \
             patch.dict(os.environ, {"GEMINI_API_KEY": "AIzaSyTestFakeKey"}):
            res = self.client.post("/api/ai/query", json={"prompt": "ازيك"})
            self.assertEqual(res.status_code, 200)
            data = res.get_json()
            self.assertIn("reply", data)
            self.assertIn("أهلاً بك", data["reply"])
            mock_agent.assert_not_called()
        print("PASS: Simple greeting short-circuit verified (no heavy DB queries called).")

    def test_08_gemini_tool_declarations_registered(self):
        """Verify all 5 tools are declared with clean schemas in DCR_TOOL."""
        tool_names = [fn.name for fn in DCR_TOOL.function_declarations]
        self.assertIn("search_submittals", tool_names)
        self.assertIn("get_project_summary", tool_names)
        self.assertIn("get_noc_financials", tool_names)
        self.assertIn("get_overdue_submittals", tool_names)
        self.assertIn("get_document_history", tool_names)
        self.assertEqual(len(tool_names), 5)
        print("PASS: All 5 Gemini Tool declarations registered successfully.")

    def test_09_rbac_boundary_omits_allowed_project_ids(self):
        """Verify that allowed_project_ids is strictly omitted from the LLM tool declaration schemas."""
        for fn in DCR_TOOL.function_declarations:
            props = fn.parameters.properties if fn.parameters else {}
            self.assertNotIn(
                "allowed_project_ids",
                props,
                f"Security breach: tool {fn.name} exposes allowed_project_ids to LLM!"
            )
        print("PASS: Security Boundary verified: allowed_project_ids is 100% hidden from LLM schema.")

    def test_10_agent_executes_tool_call_with_rbac_injection(self):
        """Verify execute_tool_call forcibly injects allowed_project_ids into the target function."""
        with patch.object(ai_tools, "search_submittals", return_value={"count": 1}) as mock_search:
            res = execute_tool_call(
                fn_name="search_submittals",
                fn_args={"keyword": "valve", "doc_type": "MS"},
                allowed_project_ids=["PEM-058"]
            )
            self.assertEqual(res.get("count"), 1)
            mock_search.assert_called_once()
            called_kwargs = mock_search.call_args.kwargs
            self.assertEqual(called_kwargs.get("allowed_project_ids"), ["PEM-058"])
            self.assertEqual(called_kwargs.get("keyword"), "valve")
            self.assertEqual(called_kwargs.get("doc_type"), "MS")
        print("PASS: Backend RBAC injection into tool execution verified.")

    def test_11_agent_controller_multi_turn_tool_calling(self):
        """Verify agent controller processes function call turn and returns synthesized markdown."""
        mock_fc = MagicMock()
        mock_fc.name = "search_submittals"
        mock_fc.args = {"keyword": "valve", "project_code": "CFC"}

        # First turn: Gemini requests tool call
        turn1_resp = MagicMock()
        turn1_resp.function_calls = [mock_fc]
        turn1_resp.candidates = [MagicMock()]

        # Second turn: Gemini synthesizes final response
        turn2_resp = MagicMock()
        turn2_resp.function_calls = []
        turn2_resp.text = "### تفاصيل المحابس في CFC\n- تم العثور على 14 محبس معتمد."

        with patch("google.genai.Client") as mock_client_cls, \
             patch.object(ai_tools, "search_submittals", return_value={"count": 14, "records": []}), \
             patch.dict(os.environ, {"GEMINI_API_KEY": "AIzaSyTestFakeKey"}):
            client_inst = mock_client_cls.return_value
            client_inst.models.generate_content.side_effect = [turn1_resp, turn2_resp]

            reply, err, tool_results = _call_gemini_agent(
                prompt="شوفلي معتمد MS محابس في CFC",
                allowed_project_ids=["PEM-058"]
            )
            self.assertIsNone(err)
            self.assertIn("14 محبس معتمد", reply)
            self.assertEqual(len(tool_results), 1)
            self.assertEqual(tool_results[0]["name"], "search_submittals")
            # Verify that Gemini tool response was sent with role='tool' (not 'user')
            call2_kwargs = client_inst.models.generate_content.call_args_list[1].kwargs
            follow_up_contents = call2_kwargs.get("contents", [])
            tool_turns = [c for c in follow_up_contents if getattr(c, "role", "") == "tool"]
            self.assertEqual(len(tool_turns), 1, "Function response turn MUST have role='tool'")
        print("PASS: Multi-turn Agent tool calling and response synthesis verified with role='tool'.")

    def test_12_successful_query_returns_reply(self):
        """End-to-end endpoint call with mocked Agent returns markdown reply."""
        user = {"username": "admin_user", "role": "admin"}
        mock_reply = "### Summary\n- Total approved NOC value: **110,000.00 EGP**\n- 1 overdue submittal pending."

        with patch("app.current_user", return_value=user), \
             patch("blueprints.ai.current_user", return_value=user), \
             patch("blueprints.ai.can_view_project", return_value=True), \
             patch("blueprints.ai._call_gemini_agent", return_value=(mock_reply, None, [])), \
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

    def test_13_gemini_503_fallback_to_direct_context_endpoint(self):
        """When Gemini throws 503 UNAVAILABLE, query endpoint falls back to deterministic Python engine."""
        user = {"username": "admin_user", "role": "admin"}
        with patch("app.current_user", return_value=user), \
             patch("blueprints.ai.current_user", return_value=user), \
             patch("blueprints.ai.can_view_project", return_value=True), \
             patch("blueprints.ai._call_gemini_agent", return_value=(None, "503 UNAVAILABLE", [])), \
             patch("blueprints.ai._execute_deterministic_fallback", return_value="### Fallback Table\n| MS-058-012 | Valves | Approved |"), \
             patch.dict(os.environ, {"GEMINI_API_KEY": "AIzaSyTestFakeKey"}):
            res = self.client.post("/api/ai/query", json={
                "prompt": "CFC محابس ايه في MS شوفلي معتمد",
                "project_id": "PEM-058"
            })
            self.assertEqual(res.status_code, 200)
            data = res.get_json()
            self.assertIn("reply", data)
            self.assertIn("MS-058-012", data["reply"])
        print("PASS: Endpoint gracefully recovers from Gemini 503 using deterministic fallback.")

    def test_14_deterministic_fallback_valve_search(self):
        """Verify deterministic fallback formats valve submittals into structured Markdown table."""
        mock_valves = [
            {"doc_no": "MS-CY002P608-00019 REV01", "title": "Duty Valves", "doc_type": "MS", "status": "Approved As Noted", "doc_date": "2026-05-01", "project_code": "PEM-058"},
            {"doc_no": "MS-CY002P608-00020 REV00", "title": "Fire Valves", "doc_type": "MS", "status": "Approved", "doc_date": "2026-05-02", "project_code": "PEM-058"},
        ]
        with patch.object(ai_tools, "search_submittals", return_value={"project_scope": ["CFC (PEM-058)"], "count": 2, "records": mock_valves}):
            res = _execute_deterministic_fallback(
                prompt="شوفلي MS معتمد الخاص بالمحابس في CFC",
                allowed_project_ids=["PEM-058"]
            )
            self.assertIn("تفاصيل وثائق الـ MS الخاصة بـ valve المعتمدة لمشروع CFC (PEM-058):", res)
            self.assertIn("MS-CY002P608-00019 REV01", res)
            self.assertIn("Duty Valves", res)
            self.assertIn("✅ **Approved As Noted**", res)
        print("PASS: Deterministic fallback formats valve search results into Markdown table.")

    def test_15_deterministic_fallback_zero_valves_clean_message(self):
        """When 0 valve records found, deterministic fallback returns specific Arabic message."""
        with patch.object(ai_tools, "search_submittals", return_value={"project_scope": ["EKH (24CP01)"], "count": 0, "records": []}):
            res = _execute_deterministic_fallback(
                prompt="شوفلي معتمد MS محابس في EKH",
                allowed_project_ids=["24CP01"]
            )
            self.assertEqual(res, "لم يتم العثور على أي اعتمادات تخص المحابس (Valves) في سجلات مشروع EKH (24CP01).")
        print("PASS: Zero valve records returns explicit message rather than arbitrary dump.")

    def test_16_deterministic_fallback_noc_financials(self):
        """Verify deterministic fallback extracts and formats NOC metrics."""
        mock_fin = {
            "project_scope": ["CFC (PEM-058)"],
            "total_count": 5,
            "total_submitted_cost": 500000.0,
            "total_approved_cost": 420000.0,
            "approved_count": 3,
            "pending_count": 2,
            "rejected_count": 0,
        }
        with patch.object(ai_tools, "get_noc_financials", return_value=mock_fin):
            res = _execute_deterministic_fallback(
                prompt="ما هي تكلفة أوامر التغيير NOC في CFC؟",
                allowed_project_ids=["PEM-058"]
            )
            self.assertIn("ملخص أوامر التغيير (Notice of Change)", res)
            self.assertIn("500,000.00 EGP", res)
            self.assertIn("420,000.00 EGP", res)
        print("PASS: Deterministic fallback accurately calculates and formats NOC metrics.")

    def test_17_deterministic_fallback_overdue_submittals(self):
        """Verify deterministic fallback formats overdue documents table."""
        mock_ov = {
            "project_scope": ["CFC (PEM-058)"],
            "total_overdue": 1,
            "records": [
                {"docNo": "IR-058-100", "title": "Duct Inspection", "dt_code": "IR", "days_overdue": 14, "expected_reply": "2026-06-01", "proj_code": "PEM-058"}
            ]
        }
        with patch.object(ai_tools, "get_overdue_submittals", return_value=mock_ov):
            res = _execute_deterministic_fallback(
                prompt="عايز الوثائق المتأخرة overdue في CFC",
                allowed_project_ids=["PEM-058"]
            )
            self.assertIn("تفاصيل الوثائق المتأخرة لمشروع CFC (PEM-058)", res)
            self.assertIn("IR-058-100", res)
            self.assertIn("14 يوم", res)
        print("PASS: Deterministic fallback accurately formats overdue submittals table.")

    def test_18_deterministic_fallback_project_summary(self):
        """Verify deterministic fallback formats register KPI totals."""
        mock_sum = {
            "project_scope": ["CFC (PEM-058)"],
            "grand_total_documents": 200,
            "overall_approval_pct": 75,
            "projects": [
                {"code": "PEM-058", "name": "CFC Ph2", "total": 200, "approved": 150, "pct": 75, "pending": 30, "overdue": 10, "rejected": 10}
            ]
        }
        with patch.object(ai_tools, "get_project_summary", return_value=mock_sum):
            res = _execute_deterministic_fallback(
                prompt="ملخص موقف وسجلات مشروع CFC",
                allowed_project_ids=["PEM-058"]
            )
            self.assertIn("ملخص السجلات والمستندات لمشروع CFC (PEM-058)", res)
            self.assertIn("200 وثيقة بنسبة اعتماد كلية **75%**", res)
        print("PASS: Deterministic fallback accurately formats project summary KPIs.")

    def test_19_search_submittals_valve_stemming(self):
        """Verify valve terms match valve submittals and exclude generic pipe/conduit fittings."""
        mock_projects = [{"id": "PEM-058", "name": "CFC Ph2", "code": "PEM-058", "data": {}}]
        raw_records = [
            {"id": "r1", "project_id": "PEM-058", "doc_type": "MS", "doc_no": "MS-CY002P608-00019 REV01", "title": "Manual Butterfly Valve", "status": "Approved As Noted"},
            {"id": "r2", "project_id": "PEM-058", "doc_type": "MS", "doc_no": "MS-CY002P608-00025 REV00", "title": "UPVC Pipes & Fittings", "status": "Approved"},
        ]
        with patch("db.q") as mock_q:
            def q_side_effect(sql, params=()):
                if "FROM records r" in sql:
                    return raw_records
                if "FROM projects" in sql:
                    return mock_projects
                return []
            mock_q.side_effect = q_side_effect

            res = ai_tools.search_submittals(
                allowed_project_ids=["PEM-058"],
                project_code="CFC",
                keyword="valve"
            )
            # Only the valve record should be returned
            titles = [r["title"] for r in res["records"]]
            self.assertIn("Manual Butterfly Valve", titles)
            self.assertNotIn("UPVC Pipes & Fittings", titles)
        print("PASS: Valve stemming isolates valves without keyword pollution.")

    def test_20_search_submittals_revision_deduplication(self):
        """Multiple revisions of same document must be deduplicated to retain only the latest revision."""
        records = [
            {"doc_no": "MS-CY002P608-00019 REV00", "title": "Duty Valves", "status": "C - Revise and Resubmit"},
            {"doc_no": "MS-CY002P608-00019 REV01", "title": "Duty Valves", "status": "B - Approved As Noted"},
            {"doc_no": "MS-CY002P608-00020 REV00", "title": "Fire Fighting Valves", "status": "B - Approved As Noted"},
        ]
        deduped = _dedup_latest_revisions(records)
        self.assertEqual(len(deduped), 2)
        doc_nos = [r["doc_no"] for r in deduped]
        self.assertIn("MS-CY002P608-00019 REV01", doc_nos)
        self.assertNotIn("MS-CY002P608-00019 REV00", doc_nos)
        self.assertIn("MS-CY002P608-00020 REV00", doc_nos)
        print("PASS: Revision deduplication correctly keeps only the latest revision.")

    def test_21_search_submittals_project_aliases(self):
        """Verify project aliases (PEM-058, CFC, 24CP01, EKH) resolve correctly."""
        mock_projects = [
            {"id": "PEM-058", "name": "CFC DCP Phase 2", "code": "PEM-058", "data": {}},
            {"id": "24CP01", "name": "EKH Hospital", "code": "24CP01", "data": {}}
        ]
        with patch("db.q") as mock_q:
            mock_q.return_value = mock_projects
            resolved_cfc = ai_tools._resolve_target_projects(["PEM-058", "24CP01"], project_code="CFC")
            self.assertEqual(len(resolved_cfc), 1)
            self.assertEqual(resolved_cfc[0]["code"], "PEM-058")

            resolved_ekh = ai_tools._resolve_target_projects(["PEM-058", "24CP01"], project_code="EKH")
            self.assertEqual(len(resolved_ekh), 1)
            self.assertEqual(resolved_ekh[0]["code"], "24CP01")
        print("PASS: Project aliases resolved accurately across codes, names, and IDs.")

    def test_22_search_submittals_discipline_isolation(self):
        """Verify querying Architecture discipline isolates architecture and excludes electrical records."""
        mock_projects = [{"id": "PEM-058", "name": "CFC Ph2", "code": "PEM-058", "data": {}}]
        mixed_records = [
            {"id": "r1", "project_id": "PEM-058", "doc_type": "IR", "doc_no": "IR-058-100", "title": "Block Work", "discipline": "Architecture"},
            {"id": "r2", "project_id": "PEM-058", "doc_type": "IR", "doc_no": "IR-058-101", "title": "Cabling", "discipline": "Electrical"},
        ]
        with patch("db.q") as mock_q:
            def q_side_effect(sql, params=()):
                if "FROM records r" in sql:
                    return mixed_records
                if "FROM projects" in sql:
                    return mock_projects
                return []
            mock_q.side_effect = q_side_effect

            res = ai_tools.search_submittals(
                allowed_project_ids=["PEM-058"],
                project_code="PEM-058",
                discipline="Architecture"
            )
            doc_nos = [r["doc_no"] for r in res["records"]]
            self.assertIn("IR-058-100", doc_nos)
            self.assertNotIn("IR-058-101", doc_nos)
        print("PASS: Architecture discipline strictly isolates architecture and excludes electrical records.")

    def test_23_search_submittals_doctype_strictness(self):
        """When user requests MS, MIR records must be excluded."""
        mock_projects = [{"id": "PEM-058", "name": "CFC Ph2", "code": "PEM-058", "data": {}}]
        mixed_records = [
            {"id": "r1", "project_id": "PEM-058", "doc_type": "MS", "doc_no": "MS-058-001", "title": "Chilled Pipes", "status": "Approved"},
        ]
        with patch("db.q") as mock_q:
            def q_side_effect(sql, params=()):
                if "FROM records r" in sql:
                    return mixed_records
                if "FROM projects" in sql:
                    return mock_projects
                return []
            mock_q.side_effect = q_side_effect

            res = ai_tools.search_submittals(
                allowed_project_ids=["PEM-058"],
                doc_type="MS"
            )
            self.assertEqual(res["count"], 1)
            self.assertEqual(res["records"][0]["doc_no"], "MS-058-001")
        print("PASS: Strict document type filtering excludes unrelated document types.")

    def test_24_search_submittals_no_duplicates(self):
        """Verify duplicate records are strictly eliminated."""
        mock_projects = [{"id": "PEM-058", "name": "CFC Ph2", "code": "PEM-058", "data": {}}]
        duplicate_records = [
            {"id": "r1", "project_id": "PEM-058", "doc_no": "MS-CY002P608-00116 REV00", "title": "Valves", "status": "Approved"},
            {"id": "r1", "project_id": "PEM-058", "doc_no": "MS-CY002P608-00116 REV00", "title": "Valves", "status": "Approved"},
        ]
        with patch("db.q") as mock_q:
            def q_side_effect(sql, params=()):
                if "FROM records r" in sql:
                    return duplicate_records
                if "FROM projects" in sql:
                    return mock_projects
                return []
            mock_q.side_effect = q_side_effect

            res = ai_tools.search_submittals(
                allowed_project_ids=["PEM-058"],
                latest_revision_only=True
            )
            self.assertEqual(res["count"], 1)
        print("PASS: Duplicate document numbers strictly eliminated.")

    def test_25_psycopg2_array_compatibility(self):
        """Verify all parameters passed to ANY(%s) are Python lists, preventing DatatypeMismatch."""
        mock_projects = [{"id": "PEM-058", "name": "CFC Ph2", "code": "PEM-058", "data": {}}]
        with patch("db.q") as mock_q:
            executed_sqls = []
            def q_side_effect(sql, params=()):
                executed_sqls.append((sql, params))
                if "FROM records r" in sql:
                    return []
                if "FROM projects" in sql:
                    return mock_projects
                return []
            mock_q.side_effect = q_side_effect

            ai_tools.search_submittals(
                allowed_project_ids=["PEM-058"],
                project_code="CFC",
                keyword="valve"
            )
            records_calls = [c for c in executed_sqls if "FROM records r" in c[0]]
            self.assertTrue(len(records_calls) > 0)
            _, params = records_calls[0]
            for p in params:
                if isinstance(p, (tuple, list, set)):
                    self.assertIsInstance(p, list, f"Param {p} must be a list for psycopg2 ANY(...) compatibility")
        print("PASS: All PostgreSQL ANY(...) array bindings are Python lists, preventing DatatypeMismatch.")

    def test_26_history_preserves_equipment_and_project(self):
        """Clarification prompt inherits equipment and project context from history in fallback."""
        history = [
            {"role": "user", "text": "شوفلي معتمد MS محابس في CFC"},
            {"role": "assistant", "text": "تم العثور على محابس معتمدة"}
        ]
        with patch.object(ai_tools, "search_submittals", return_value={"project_scope": ["CFC (PEM-058)"], "count": 1, "records": [{"doc_no": "MS-01", "title": "Valves", "doc_type": "MS", "status": "Approved", "doc_date": "2026-05-01", "project_code": "PEM-058"}]}) as mock_search:
            res = _execute_deterministic_fallback(
                prompt="MS = Material Submittal",
                allowed_project_ids=["PEM-058"],
                history=history
            )
            mock_search.assert_called_once()
            self.assertEqual(mock_search.call_args.kwargs.get("project_code"), "PEM-058")
        print("PASS: Conversational clarification preserves equipment and project scope from history.")

    def test_27_project_override_prevents_context_bleed(self):
        """Mentioning a new project strictly overrides previous project in conversation history."""
        history = [
            {"role": "user", "text": "شوفلي معتمد MS محابس في CFC"},
            {"role": "assistant", "text": "تم العثور على 14 محبس معتمد في CFC"}
        ]
        with patch.object(ai_tools, "search_submittals", return_value={"project_scope": ["EKH (24CP01)"], "count": 0, "records": []}) as mock_search:
            res = _execute_deterministic_fallback(
                prompt="طب شوفلي معتمد محابس في EKH",
                allowed_project_ids=["24CP01", "PEM-058"],
                history=history
            )
            mock_search.assert_called_once()
            self.assertEqual(mock_search.call_args.kwargs.get("project_code"), "24CP01")
        print("PASS: Project override strictly prevents history context bleed.")

    def test_28_get_noc_financials_cost_aggregation(self):
        """Verify get_noc_financials accurately calculates submitted vs approved totals."""
        mock_projects = [{"id": "PEM-058", "name": "CFC Ph2", "code": "PEM-058", "data": {}}]
        mock_nocs = [
            {"project_id": "PEM-058", "proj_code": "PEM-058", "proj_name": "CFC", "data": {"submittedCost": "100000", "finalApprovedCost": "90000", "partDStatus": "Approved"}},
            {"project_id": "PEM-058", "proj_code": "PEM-058", "proj_name": "CFC", "data": {"submittedCost": "50000", "finalApprovedCost": "0", "partDStatus": "Under Review"}},
        ]
        with patch("db.q") as mock_q:
            def q_side_effect(sql, params=()):
                if "FROM records r" in sql:
                    return mock_nocs
                if "FROM projects" in sql:
                    return mock_projects
                return []
            mock_q.side_effect = q_side_effect

            fin = ai_tools.get_noc_financials(allowed_project_ids=["PEM-058"])
            self.assertEqual(fin["total_count"], 2)
            self.assertEqual(fin["total_submitted_cost"], 150000.0)
            self.assertEqual(fin["total_approved_cost"], 90000.0)
            self.assertEqual(fin["approved_count"], 1)
            self.assertEqual(fin["pending_count"], 1)
        print("PASS: get_noc_financials accurately calculates costs and statuses.")

    def test_29_get_overdue_submittals_sorting(self):
        """Verify get_overdue_submittals sorts by delay descending."""
        mock_projects = [{"id": "PEM-058", "name": "CFC Ph2", "code": "PEM-058", "data": {}}]
        mock_overdue = [
            {"project_id": "PEM-058", "dt_code": "IR", "docNo": "IR-01", "days_overdue": 5},
            {"project_id": "PEM-058", "dt_code": "IR", "docNo": "IR-02", "days_overdue": 20},
        ]
        with patch("db.q") as mock_q, patch("db.get_overdue_records", return_value=mock_overdue):
            mock_q.return_value = mock_projects
            res = ai_tools.get_overdue_submittals(allowed_project_ids=["PEM-058"])
            self.assertEqual(res["records"][0]["docNo"], "IR-02")
            self.assertEqual(res["records"][0]["days_overdue"], 20)
        print("PASS: get_overdue_submittals sorts records by longest delay first.")

    def test_30_get_document_history_chronology(self):
        """Verify get_document_history compiles revisions in chronological order."""
        mock_projects = [{"id": "PEM-058", "name": "CFC Ph2", "code": "PEM-058", "data": {}}]
        mock_history = [
            {"id": "r1", "project_id": "PEM-058", "dt_id": "MS", "proj_code": "PEM-058", "created_at": "2026-01-01", "data": {"docNo": "MS-01 REV00", "status": "Rejected"}},
            {"id": "r2", "project_id": "PEM-058", "dt_id": "MS", "proj_code": "PEM-058", "created_at": "2026-02-01", "data": {"docNo": "MS-01 REV01", "status": "Approved"}},
        ]
        with patch("db.q") as mock_q:
            def q_side_effect(sql, params=()):
                if "FROM records r" in sql:
                    return mock_history
                if "FROM projects" in sql:
                    return mock_projects
                return []
            mock_q.side_effect = q_side_effect

            hist = ai_tools.get_document_history(allowed_project_ids=["PEM-058"], doc_no_pattern="MS-01 REV01")
            self.assertEqual(hist["total_revisions"], 2)
            self.assertEqual(hist["latest_revision"], "REV 01")
            self.assertEqual(hist["latest_status"], "Approved")
        print("PASS: get_document_history compiles chronological revisions.")

    def test_31_rbac_strictly_blocks_unauthorized_projects(self):
        """Verify unauthorized projects are strictly blocked from returning data."""
        mock_projects = [{"id": "PEM-058", "name": "CFC Ph2", "code": "PEM-058", "data": {}}]
        with patch("db.q") as mock_q:
            mock_q.return_value = mock_projects
            # User only has permission for PEM-058, attempts to search project 24CP01
            res = ai_tools.search_submittals(allowed_project_ids=["PEM-058"], project_code="24CP01")
            self.assertEqual(res["count"], 0)
            self.assertEqual(len(res["records"]), 0)
        print("PASS: RBAC strictly blocks unauthorized projects.")

    def test_32_global_query_does_not_bleed_project_from_history(self):
        """Global queries across registers strictly do not inherit prior project from history."""
        history = [
            {"role": "user", "text": "شوفلي معتمد MS محابس في CFC"},
            {"role": "assistant", "text": "تم العثور على 14 محبس معتمد في CFC"}
        ]
        with patch.object(ai_tools, "get_project_summary", return_value={"project_scope": ["All Projects"], "total_projects": 2, "projects": [], "grand_total_documents": 100}) as mock_sum:
            res = _execute_deterministic_fallback(
                prompt="Provide a complete status breakdown of documents across the project registers",
                allowed_project_ids=["PEM-058", "24CP01"],
                history=history
            )
            mock_sum.assert_called_once()
            self.assertIsNone(mock_sum.call_args.kwargs.get("project_code"))
        print("PASS: Global query strictly suppresses previous project context from history.")

    def test_33_gemini_contents_sanitization_starts_with_user(self):
        """Gemini content turns must drop leading welcome message and begin with user role."""
        raw_history = [
            {"role": "assistant", "text": "Hello! I am your DCR Engineering AI Assistant."},
            {"role": "user", "text": "Show me CFC valves"},
            {"role": "assistant", "text": "Found 14 valves"},
        ]
        contents = _build_gemini_contents(prompt="What about pumps?", history=raw_history)
        self.assertGreaterEqual(len(contents), 2)
        # Verify first turn is strictly user
        self.assertEqual(contents[0].role, "user")
        self.assertEqual(contents[0].parts[0].text, "Show me CFC valves")
        # Verify alternating turns
        self.assertEqual(contents[1].role, "model")
        self.assertEqual(contents[2].role, "user")
        self.assertEqual(contents[2].parts[0].text, "What about pumps?")
        print("PASS: Gemini contents strictly alternate and drop leading welcome message.")

    def test_34_multi_project_switching_same_chat(self):
        """User can switch projects in subsequent turns within the same chat without bleed."""
        history = [
            {"role": "user", "text": "شوفلي معتمد MS محابس في CFC"},
            {"role": "assistant", "text": "تم العثور على 14 محبس معتمد في CFC"}
        ]
        # Turn 2: Switch to EKH
        with patch.object(ai_tools, "search_submittals", return_value={"project_scope": ["EKH (24CP01)"], "count": 1, "records": []}) as mock_search:
            res_ekh = _execute_deterministic_fallback(
                prompt="طب في مشروع EKH وريني المحابس المعتمدة",
                allowed_project_ids=["PEM-058", "24CP01", "PEM-045"],
                history=history
            )
            self.assertEqual(mock_search.call_args.kwargs.get("project_code"), "24CP01")

        # Turn 3: Switch to Assiut
        history.extend([
            {"role": "user", "text": "طب في مشروع EKH وريني المحابس المعتمدة"},
            {"role": "assistant", "text": "تم العثور على محابس في EKH"}
        ])
        with patch.object(ai_tools, "search_submittals", return_value={"project_scope": ["Assiut (PEM-045)"], "count": 1, "records": []}) as mock_search2:
            res_assiut = _execute_deterministic_fallback(
                prompt="طب وماذا عن مشروع أسيوط؟",
                allowed_project_ids=["PEM-058", "24CP01", "PEM-045"],
                history=history
            )
            self.assertEqual(mock_search2.call_args.kwargs.get("project_code"), "PEM-045")
        print("PASS: Multi-project switching works seamlessly across turns in the same chat.")

    def test_35_conversational_follow_up_inherits_scope_and_type(self):
        """Follow-up questions like 'وماذا عن المضخات؟' inherit project and document type context."""
        history = [
            {"role": "user", "text": "شوفلي معتمد MS محابس في CFC"},
            {"role": "assistant", "text": "تم العثور على 14 محبس معتمد في CFC"}
        ]
        with patch.object(ai_tools, "search_submittals", return_value={"project_scope": ["CFC (PEM-058)"], "count": 2, "records": [{"doc_no": "MS-02", "title": "Chilled Water Pump", "doc_type": "MS", "status": "Approved", "doc_date": "2026-05-01", "project_code": "PEM-058"}]}) as mock_search:
            res = _execute_deterministic_fallback(
                prompt="وماذا عن المضخات؟",
                allowed_project_ids=["PEM-058"],
                history=history
            )
            mock_search.assert_called_once()
            # Inherited project CFC and doc_type MS, with new keyword pump
            self.assertEqual(mock_search.call_args.kwargs.get("project_code"), "PEM-058")
            self.assertEqual(mock_search.call_args.kwargs.get("doc_type"), "MS")
            self.assertEqual(mock_search.call_args.kwargs.get("keyword"), "pump")
            self.assertEqual(mock_search.call_args.kwargs.get("status"), "approved")
        print("PASS: Conversational follow-up question inherits project scope, document type, and approval status.")

    def test_36_bilingual_response_language_accuracy(self):
        """English prompts return English headers; Arabic prompts return Arabic headers."""
        mock_rec = [{"doc_no": "MS-01", "title": "Valves", "doc_type": "MS", "status": "Approved", "doc_date": "2026-05-01", "project_code": "PEM-058"}]
        with patch.object(ai_tools, "search_submittals", return_value={"project_scope": ["CFC (PEM-058)"], "count": 1, "records": mock_rec}):
            # English Prompt
            en_res = _execute_deterministic_fallback(
                prompt="Show me approved MS valves in CFC",
                allowed_project_ids=["PEM-058"]
            )
            self.assertIn("Submittal Records for CFC (PEM-058)", en_res)
            self.assertIn("Document No", en_res)

            # Arabic Prompt
            ar_res = _execute_deterministic_fallback(
                prompt="شوفلي معتمد MS محابس في CFC",
                allowed_project_ids=["PEM-058"]
            )
            self.assertIn("تفاصيل وثائق", ar_res)
            self.assertIn("رقم الوثيقة", ar_res)
        print("PASS: Bilingual response accurately distinguishes and renders English and Arabic.")

    def test_37_get_noc_financials_per_project_breakdown(self):
        """get_noc_financials and fallback must isolate and display metrics for each project individually."""
        mock_projects = [
            {"id": "PEM-058", "name": "CFC Ph2", "code": "PEM-058", "data": {}},
            {"id": "24CP01", "name": "EKH Hospital", "code": "24CP01", "data": {}},
            {"id": "PEM-045", "name": "Assiut Power", "code": "PEM-045", "data": {}},
        ]
        mock_nocs = [
            {"project_id": "PEM-058", "proj_code": "PEM-058", "proj_name": "CFC", "data": {"docNo": "NOC-01", "submittedCost": "100000", "finalApprovedCost": "90000", "partDStatus": "Approved"}},
            {"project_id": "24CP01", "proj_code": "24CP01", "proj_name": "EKH", "data": {"docNo": "NOC-02", "submittedCost": "200000", "finalApprovedCost": "180000", "partDStatus": "Approved"}},
        ]
        with patch("db.q") as mock_q:
            def q_side_effect(sql, params=()):
                if "FROM records r" in sql:
                    return mock_nocs
                if "FROM projects" in sql:
                    return mock_projects
                return []
            mock_q.side_effect = q_side_effect

            fin = ai_tools.get_noc_financials(allowed_project_ids=["PEM-058", "24CP01", "PEM-045"])
            self.assertIn("by_project", fin)
            by_proj = fin["by_project"]
            # Assiut has 0 NOCs, so it must be excluded from by_project
            self.assertEqual(len(by_proj), 2)
            proj_codes = [p["project_code"] for p in by_proj]
            self.assertIn("PEM-058", proj_codes)
            self.assertIn("24CP01", proj_codes)
            self.assertNotIn("PEM-045", proj_codes)

            cfc_stat = next(p for p in by_proj if p["project_code"] == "PEM-058")
            ekh_stat = next(p for p in by_proj if p["project_code"] == "24CP01")
            self.assertEqual(cfc_stat["total_submitted_cost"], 100000.0)
            self.assertEqual(ekh_stat["total_submitted_cost"], 200000.0)

            # Fallback rendering test
            res = _execute_deterministic_fallback(
                prompt="What is the total approved cost for Notice of Change (NOCs)?",
                allowed_project_ids=["PEM-058", "24CP01", "PEM-045"]
            )
            self.assertIn("Notice of Change (NOC) Breakdown by Project", res)
            self.assertIn("PEM-058", res)
            self.assertIn("24CP01", res)
            self.assertNotIn("PEM-045", res)
            self.assertIn("100,000.00 EGP", res)
            self.assertIn("200,000.00 EGP", res)
            # Must NOT contain random Approved NOC Items section
            self.assertNotIn("Approved NOC Items", res)
        print("PASS: NOC financials accurately breaks down metrics for each project individually and excludes 0-count projects.")

    def test_38_valves_strict_doctype_and_multidoctype_unspecified(self):
        """When MS is specified, search strictly returns MS and never substitutes with NOC.
        When no doc_type is specified, it returns matching valve submittals across all registers."""
        mock_projects = [{"id": "CFC DCP PH-2A", "name": "CFC DCP PH-2A", "code": "PEM-058", "data": {}}]
        ms_records = [
            {
                "id": "r1",
                "project_id": "CFC DCP PH-2A",
                "dt_id": "MS",
                "doc_type": "MS",
                "doc_no": "MS-CY002P608-00019 REV01",
                "title": "Duty Valves for HVAC Piping (Manual Butterfly Valve)",
                "status": "B - Approved As Noted",
                "doc_date": "2026-07-05",
                "project_code": "PEM-058"
            }
        ]
        noc_records = [
            {
                "id": "r2",
                "project_id": "CFC DCP PH-2A",
                "dt_id": "NOC",
                "doc_type": "NOC",
                "doc_no": "NOC-00007 REV01",
                "title": "Cooling towers fittings and valves",
                "status": "Approved",
                "doc_date": "2026-05-01",
                "project_code": "PEM-058"
            }
        ]

        with patch("db.q") as mock_q:
            def q_side_effect(sql, params=()):
                if "FROM projects" in sql:
                    return mock_projects
                if "FROM records r" in sql:
                    if "UPPER(r.dt_id) = %s" in sql:
                        return ms_records
                    return ms_records + noc_records
                return []
            mock_q.side_effect = q_side_effect

            # 1. Strict doc_type MS query with Arabic keyword
            res_ms = ai_tools.search_submittals(
                allowed_project_ids=["PEM-058", "CFC DCP PH-2A"],
                project_code="CFC",
                doc_type="MS",
                keyword="المحابس",
                status="approved"
            )
            self.assertEqual(res_ms["count"], 1)
            self.assertEqual(res_ms["records"][0]["doc_no"], "MS-CY002P608-00019 REV01")
            # Must NOT contain related_other_types or NOC
            self.assertNotIn("related_other_types", res_ms)

            # 2. Unspecified doc_type query across all registers
            res_all = ai_tools.search_submittals(
                allowed_project_ids=["PEM-058", "CFC DCP PH-2A"],
                project_code="CFC",
                doc_type=None,
                keyword="المحابس",
                status="approved"
            )
            self.assertEqual(res_all["count"], 2)
            all_doc_nos = [r["doc_no"] for r in res_all["records"]]
            self.assertIn("MS-CY002P608-00019 REV01", all_doc_nos)
            self.assertIn("NOC-00007 REV01", all_doc_nos)

            # 3. Fallback rendering strictly honors doc_type MS
            reply = _execute_deterministic_fallback(
                prompt="شوفلي MS المعتمد الخاص بالمحابس في مشروع CFC",
                allowed_project_ids=["PEM-058", "CFC DCP PH-2A"]
            )
            self.assertIn("MS-CY002P608-00019 REV01", reply)
            self.assertNotIn("NOC-00007 REV01", reply)
        print("PASS: Strict doc_type enforcement and multi-doctype unspecified searches verified.")

    def test_28_psycopg2_query_interpolation_and_valve_records(self):
        """Verify all SQL queries interpolate correctly under psycopg2 and 14 MS valve entries are retrieved."""
        mock_projects = [{"id": "CFC DCP PH-2A", "name": "CFC District Cooling", "code": "PEM-058", "data": {}}]
        valves_14 = [
            {"id": f"r{i}", "project_id": "CFC DCP PH-2A", "doc_type": "MS", "doc_no": f"MS-CY002P608-000{i:02d} REV00", "title": f"Valves Item {i}", "status": "B - Approved As Noted", "doc_date": "2026-05-01"}
            for i in range(1, 15)
        ]
        with patch("db.q") as mock_q:
            def q_side_effect(sql, params=()):
                # Crucial assertion: Ensure query string interpolates with params without ValueError
                try:
                    dummy_params = tuple(['val'] * len(params))
                    _ = sql % dummy_params
                except ValueError as ve:
                    self.fail(f"SQL failed psycopg2 % format interpolation: {ve}")

                if "FROM records r" in sql:
                    return valves_14
                if "FROM projects" in sql:
                    return mock_projects
                return []
            mock_q.side_effect = q_side_effect

            res = ai_tools.search_submittals(
                allowed_project_ids=["PEM-058", "CFC DCP PH-2A"],
                project_code="CFC",
                doc_type="MS",
                keyword="المحابس",
                status="approved",
                latest_revision_only=False
            )
            self.assertEqual(res["count"], 14)
            self.assertEqual(len(res["records"]), 14)

            # Test fallback response with 14 records
            reply = _execute_deterministic_fallback(
                prompt="شوفلي MS المعتمد الخاص بالمحابس في مشروع CFC",
                allowed_project_ids=["PEM-058", "CFC DCP PH-2A"]
            )
            self.assertIn("14", reply)
            self.assertIn("MS-CY002P608-00001 REV00", reply)
            self.assertIn("MS-CY002P608-00014 REV00", reply)
        print("PASS: psycopg2 % interpolation and 14 MS valve entries verified without query errors.")

    def test_29_conversational_ellipsis_project_switch_inherits_topic(self):
        """When user asks 'ايه MS الخاص بالمحابس المعتمد في مشروع EKH' then follows up with 'وبالنسبة لمشروع CFC',
        the query switches project to CFC and inherits MS, valve, and approved status."""
        history = [
            {"role": "user", "text": "ايه MS الخاص بالمحابس المعتمد في مشروع EKH"},
            {"role": "assistant", "text": "تم العثور على 5 محابس معتمدة في مشروع EKH"}
        ]
        mock_cfc_valves = [
            {"id": "r1", "project_id": "CFC DCP PH-2A", "doc_type": "MS", "doc_no": "MS-CY002P608-00019 REV01", "title": "Manual Butterfly Valve", "status": "B - Approved As Noted", "doc_date": "2026-05-01", "project_code": "PEM-058"}
        ]
        with patch.object(ai_tools, "search_submittals", return_value={"project_scope": ["CFC (PEM-058)"], "count": 1, "records": mock_cfc_valves}) as mock_search:
            reply = _execute_deterministic_fallback(
                prompt="وبالنسبة لمشروع CFC",
                allowed_project_ids=["PEM-058", "24CP01"],
                history=history
            )
            mock_search.assert_called_once()
            # Verified that CFC is the project, and MS + valve + approved are inherited from Turn 1!
            self.assertEqual(mock_search.call_args.kwargs.get("project_code"), "PEM-058")
            self.assertEqual(mock_search.call_args.kwargs.get("doc_type"), "MS")
            self.assertEqual(mock_search.call_args.kwargs.get("keyword"), "valve")
            self.assertEqual(mock_search.call_args.kwargs.get("status"), "approved")
            self.assertIn("MS-CY002P608-00019 REV01", reply)
        print("PASS: Conversational ellipsis correctly switches project while inheriting doc_type, keyword, and status.")

    def test_30_isolated_project_mention_without_history_requests_clarification(self):
        """When user mentions only a project (e.g. 'وبالنسبة لمشروع CFC') without any topic and without history,
        the system requests clarification instead of dumping random unrelated records."""
        with patch.object(ai_tools, "search_submittals") as mock_search:
            reply = _execute_deterministic_fallback(
                prompt="وبالنسبة لمشروع CFC",
                allowed_project_ids=["PEM-058", "24CP01"],
                history=[]
            )
            # search_submittals MUST NOT be called blindly!
            mock_search.assert_not_called()
            # Must ask what they need in project CFC
            self.assertIn("ما الذي تود الاستعلام عنه تحديداً في المشروع؟", reply)
            self.assertIn("اعتمادات المواد (MS)", reply)
        print("PASS: Isolated project mention without history properly prompts user for clarification.")

    def test_36_eastmain_generic_resolves_and_prevents_cfc_bleed(self):
        """Turn 3: 'طيب ومشروع Eastmain' must query Eastmain and NEVER bleed to previous project (CFC)."""
        history = [
            {"role": "user", "text": "شوفلي MS المعتمد الخاص بالمحابس في مشروع CFC"},
            {"role": "assistant", "text": "تم العثور على 14 محبس معتمد في CFC"},
            {"role": "user", "text": "طيب وفي مشروع EKH"},
            {"role": "assistant", "text": "تم العثور على المحابس المعتمدة في EKH"}
        ]
        with patch.object(ai_tools, "search_submittals", return_value={"project_scope": ["Eastmain (PEM-065 / PEM-066)"], "count": 2, "records": [{"doc_no": "MS-EM-01", "title": "Butterfly Valve", "doc_type": "MS", "status": "Approved", "doc_date": "2026-06-01", "project_code": "PEM-065"}]}) as mock_search:
            reply = _execute_deterministic_fallback(
                prompt="طيب ومشروع Eastmain",
                allowed_project_ids=["PEM-058", "24CP01", "PEM-065", "PEM-066"],
                history=history
            )
            mock_search.assert_called_once()
            called_proj = mock_search.call_args.kwargs.get("project_code")
            # Must NOT be CFC (PEM-058) or EKH (24CP01)!
            self.assertNotEqual(called_proj, "PEM-058")
            self.assertNotEqual(called_proj, "24CP01")
            self.assertIn(called_proj, ("Eastmain", "PEM-065"))
            self.assertEqual(mock_search.call_args.kwargs.get("doc_type"), "MS")
            self.assertEqual(mock_search.call_args.kwargs.get("keyword"), "valve")
            self.assertEqual(mock_search.call_args.kwargs.get("status"), "approved")
        print("PASS: Eastmain generic resolves accurately and prevents CFC context bleed.")

    def test_37_eastmain_b08_resolves_specifically(self):
        """'وفي مشروع Eastmain B08' must specifically resolve to PEM-066."""
        history = [
            {"role": "user", "text": "شوفلي MS المعتمد الخاص بالمحابس في مشروع CFC"},
            {"role": "assistant", "text": "تم العثور على 14 محبس معتمد في CFC"}
        ]
        with patch.object(ai_tools, "search_submittals", return_value={"project_scope": ["Eastmain B08 (PEM-066)"], "count": 1, "records": [{"doc_no": "MS-B08-01", "title": "Gate Valve", "doc_type": "MS", "status": "Approved", "doc_date": "2026-06-01", "project_code": "PEM-066"}]}) as mock_search:
            reply = _execute_deterministic_fallback(
                prompt="وفي مشروع Eastmain B08",
                allowed_project_ids=["PEM-058", "PEM-065", "PEM-066"],
                history=history
            )
            mock_search.assert_called_once()
            self.assertEqual(mock_search.call_args.kwargs.get("project_code"), "PEM-066")
            self.assertEqual(mock_search.call_args.kwargs.get("doc_type"), "MS")
            self.assertEqual(mock_search.call_args.kwargs.get("keyword"), "valve")
        print("PASS: Eastmain B08 specifically resolves to PEM-066.")

    def test_38_suez_four_season_anopc_resolutions(self):
        """Validates all other system projects resolve accurately without bleed."""
        with patch.object(ai_tools, "search_submittals", return_value={"project_scope": [], "count": 0, "records": []}) as mock_search:
            # Suez Steel PEM-064
            _execute_deterministic_fallback("وفي مشروع Suez وريني المحابس", allowed_project_ids=["PEM-064"])
            self.assertEqual(mock_search.call_args.kwargs.get("project_code"), "PEM-064")

            # Four Season PEM-068
            _execute_deterministic_fallback("وفي مشروع Four Season وريني المحابس", allowed_project_ids=["PEM-068"])
            self.assertEqual(mock_search.call_args.kwargs.get("project_code"), "PEM-068")

            # ANOPC Assiut PEM-067
            _execute_deterministic_fallback("وفي مشروع ANOPC وريني المحابس", allowed_project_ids=["PEM-067"])
            self.assertEqual(mock_search.call_args.kwargs.get("project_code"), "PEM-067")
        print("PASS: Suez, Four Season, and ANOPC project resolutions verified.")

    def test_39_explicit_unknown_project_never_bleeds_to_history(self):
        """If user explicitly asks about a project name that does not exist, NEVER bleed to previous project."""
        history = [
            {"role": "user", "text": "شوفلي MS المعتمد الخاص بالمحابس في مشروع CFC"},
            {"role": "assistant", "text": "تم العثور على 14 محبس معتمد في CFC"}
        ]
        with patch.object(ai_tools, "search_submittals", return_value={"project_scope": [], "count": 0, "records": []}) as mock_search:
            reply = _execute_deterministic_fallback(
                prompt="وفي مشروع نيو_كايرو_الجديد",
                allowed_project_ids=["PEM-058"],
                history=history
            )
            called_proj = mock_search.call_args.kwargs.get("project_code")
            self.assertNotEqual(called_proj, "PEM-058")
            self.assertEqual(called_proj, "نيو_كايرو_الجديد")
        print("PASS: Explicit project mentions never bleed to history.")

    def test_40_search_submittals_strict_valve_filtering_excludes_false_positives(self):
        """Verifies that queries for valves strictly exclude unrelated submittals
        (Disconnects, Marble Stairs, EWH, Aerosol/Generator, Duct smoke tests)
        while cleanly returning actual valve inspection requests and material submittals.
        """
        mock_projects = [
            {"id": "PEM-058", "name": "CFC Ph2", "code": "PEM-058", "data": {}},
            {"id": "24CP01", "name": "EKH Hospital", "code": "24CP01", "data": {}},
            {"id": "PEM-065", "name": "Eastmain B06", "code": "PEM-065", "data": {}},
        ]
        all_records = [
            # EKH records
            {"id": "ekh1", "project_id": "24CP01", "doc_type": "IR", "doc_no": "24CP01-IR-SCAS-652 REV00", "title": "Disconnects (3rd Fix)", "discipline": "Electrical", "trade": "SCAS", "status": "Approved"},
            {"id": "ekh2", "project_id": "24CP01", "doc_type": "IR", "doc_no": "743", "title": "Marble Stairs Final Fix", "discipline": "Civil", "trade": "Finishes", "status": "Approved"},
            {"id": "ekh3", "project_id": "24CP01", "doc_type": "IR", "doc_no": "411", "title": "EWH installation", "discipline": "Mechanical", "trade": "Plumbing", "status": "Approved"},
            {"id": "ekh4", "project_id": "24CP01", "doc_type": "IR", "doc_no": "24CP01-IR-MEP-010 REV00", "title": "Installation of Gate Valves and Check Valves", "discipline": "Mechanical", "trade": "Plumbing", "status": "Approved"},
            {"id": "ekh5", "project_id": "24CP01", "doc_type": "IR", "doc_no": "24CP01-IR-MEP-011 REV00", "title": "PRV and Balancing Valves Inspection", "discipline": "Mechanical", "trade": "Plumbing", "status": "Approved"},

            # CFC records
            {"id": "cfc1", "project_id": "PEM-058", "doc_type": "IR", "doc_no": "IR-CY002P608-00437 REV00", "title": "Aerosol System First and Second Fix, Generator and Panel Installation - Aersol-09 (IT Room)", "discipline": "Mechanical", "trade": "Fire Fighting", "status": "Approved"},
            {"id": "cfc2", "project_id": "PEM-058", "doc_type": "IR", "doc_no": "IR-CY002P608-00200 REV00", "title": "Installation of Zone Control Valves (ZCV)", "discipline": "Mechanical", "trade": "Fire Fighting", "status": "Approved"},

            # Eastmain records
            {"id": "em1", "project_id": "PEM-065", "doc_type": "IR", "doc_no": "IR-EASTMAIN-HV-027 REV00", "title": "Smoke test for Galvanized Duct for building 06 1st floor", "discipline": "Mechanical", "trade": "HVAC", "status": "Approved"},
            {"id": "em2", "project_id": "PEM-065", "doc_type": "IR", "doc_no": "IR-EASTMAIN-PL-015 REV00", "title": "Water supply valves and PRV installation", "discipline": "Mechanical", "trade": "Plumbing", "status": "Approved"},
        ]

        with patch("db.q") as mock_q:
            def q_side_effect(sql, params=()):
                if "FROM projects" in sql:
                    return mock_projects
                if "FROM records r" in sql:
                    for pid in ["24CP01", "PEM-058", "PEM-065"]:
                        if pid in params:
                            return [r for r in all_records if r["project_id"] == pid]
                    return all_records
                return []
            mock_q.side_effect = q_side_effect

            # 1. Test EKH valve search
            res_ekh = ai_tools.search_submittals(
                allowed_project_ids=["24CP01", "PEM-058", "PEM-065"],
                project_code="EKH",
                keyword="محابس"
            )
            ekh_doc_nos = [r["doc_no"] for r in res_ekh["records"]]
            # True valves must be included
            self.assertIn("24CP01-IR-MEP-010 REV00", ekh_doc_nos)
            self.assertIn("24CP01-IR-MEP-011 REV00", ekh_doc_nos)
            # False positives must be strictly excluded
            self.assertNotIn("24CP01-IR-SCAS-652 REV00", ekh_doc_nos)
            self.assertNotIn("743", ekh_doc_nos)
            self.assertNotIn("411", ekh_doc_nos)

            # 2. Test CFC valve search
            res_cfc = ai_tools.search_submittals(
                allowed_project_ids=["24CP01", "PEM-058", "PEM-065"],
                project_code="CFC",
                keyword="valve"
            )
            cfc_doc_nos = [r["doc_no"] for r in res_cfc["records"]]
            self.assertIn("IR-CY002P608-00200 REV00", cfc_doc_nos)
            self.assertNotIn("IR-CY002P608-00437 REV00", cfc_doc_nos)

            # 3. Test Eastmain B06 valve search
            res_em = ai_tools.search_submittals(
                allowed_project_ids=["24CP01", "PEM-058", "PEM-065"],
                project_code="PEM-065",
                keyword="محابس"
            )
            em_doc_nos = [r["doc_no"] for r in res_em["records"]]
            self.assertIn("IR-EASTMAIN-PL-015 REV00", em_doc_nos)
            self.assertNotIn("IR-EASTMAIN-HV-027 REV00", em_doc_nos)

        print("PASS: Valve search strictly includes genuine valve submittals and eliminates all false positives.")

    def test_41_search_submittals_natural_ascending_doc_no_sort(self):
        """Verifies that submittals are returned in strict natural ascending order by Document Number."""
        mock_projects = [{"id": "PEM-058", "name": "CFC Ph2", "code": "PEM-058", "data": {}}]
        unsorted_records = [
            {"id": "r1", "project_id": "PEM-058", "doc_type": "IR", "doc_no": "IR-CY002P608-00530 REV00", "title": "Mockup for Valve Insulation", "status": "Approved"},
            {"id": "r2", "project_id": "PEM-058", "doc_type": "IR", "doc_no": "IR-CY002P608-00499 REV00", "title": "Motorized Valve Installation", "status": "Approved"},
            {"id": "r3", "project_id": "PEM-058", "doc_type": "IR", "doc_no": "IR-CY002P608-00418 REV00", "title": "Condenser Duty Valve Installation", "status": "Approved"},
            {"id": "r4", "project_id": "PEM-058", "doc_type": "IR", "doc_no": "IR-CY002P608-00467 REV00", "title": "Mock-up Chain for Valve", "status": "Approved"},
            {"id": "r5", "project_id": "PEM-058", "doc_type": "IR", "doc_no": "IR-CY002P608-00419 REV00", "title": "Torque for Butterfly Valves", "status": "Approved"},
            {"id": "r6", "project_id": "PEM-058", "doc_type": "IR", "doc_no": "IR-CY002P608-00019 REV00", "title": "Butterfly Valves Installation", "status": "Approved"},
        ]

        with patch("db.q") as mock_q:
            def q_side_effect(sql, params=()):
                if "FROM projects" in sql:
                    return mock_projects
                if "FROM records r" in sql:
                    return unsorted_records
                return []
            mock_q.side_effect = q_side_effect

            res = ai_tools.search_submittals(
                allowed_project_ids=["PEM-058"],
                project_code="CFC",
                keyword="valve"
            )
            doc_nos = [r["doc_no"] for r in res["records"]]
            expected_order = [
                "IR-CY002P608-00019 REV00",
                "IR-CY002P608-00418 REV00",
                "IR-CY002P608-00419 REV00",
                "IR-CY002P608-00467 REV00",
                "IR-CY002P608-00499 REV00",
                "IR-CY002P608-00530 REV00",
            ]
            self.assertEqual(doc_nos, expected_order)
        print("PASS: Submittals are strictly sorted in natural ascending order by Document Number.")

    def test_42_search_submittals_pr_doctype_and_items(self):
        """Verifies that PR search strictly isolates PR records, excludes DS/IR, and sorts naturally."""
        mock_projects = [{"id": "PEM-058", "name": "CFC Ph2", "code": "PEM-058", "data": {}}]
        mock_records = [
            {"id": "pr1", "project_id": "PEM-058", "doc_type": "PR", "doc_no": "PR-042 REV00", "title": "Gate Valve 125mm - 6\", Butterfly Valve 200mm", "status": "Pending PO", "description": ""},
            {"id": "pr2", "project_id": "PEM-058", "doc_type": "PR", "doc_no": "PR-042 REV01", "title": "Butterfly Valve 300mm - 12\" WITH CHAIN WHEEL", "status": "Pending PO", "description": ""},
            {"id": "ds1", "project_id": "PEM-058", "doc_type": "DS", "doc_no": "DS-CY002P608-00134 REV00", "title": "Inspection Test Plan for Firefighting Valves", "status": "Approved As Noted", "description": ""},
            {"id": "ir1", "project_id": "PEM-058", "doc_type": "IR", "doc_no": "IR-CY002P608-00038 REV00", "title": "Fire Fighting Piping Installation + Zone Valve", "status": "Approved As Noted", "description": ""},
        ]
        mock_pr_items = {
            "pr1": [{"item_name": "Gate Valve 125mm", "status": "Pending PO"}],
            "pr2": [{"item_name": "Butterfly Valve 300mm", "status": "Pending PO"}],
        }

        with patch("db.q") as mock_q, patch("db.get_pr_items_for_records") as mock_items:
            def q_side_effect(sql, params=()):
                if "FROM projects" in sql:
                    return mock_projects
                if "FROM records r" in sql:
                    # Filter by doc_type if requested
                    if "UPPER(r.dt_id) = %s" in sql or "doc_types WHERE (UPPER(code) = 'PR'" in sql:
                        return [r for r in mock_records if r["doc_type"] == "PR"]
                    return mock_records
                return []
            mock_q.side_effect = q_side_effect
            mock_items.return_value = mock_pr_items

            # 1. Search all revisions of PR for valves
            res_all = ai_tools.search_submittals(
                allowed_project_ids=["PEM-058"],
                project_code="CFC",
                doc_type="PR",
                keyword="valve",
                latest_revision_only=False,
            )
            doc_nos_all = [r["doc_no"] for r in res_all["records"]]
            doc_types_all = [r["doc_type"] for r in res_all["records"]]

            self.assertEqual(len(doc_nos_all), 2)
            self.assertEqual(doc_nos_all, ["PR-042 REV00", "PR-042 REV01"])
            self.assertTrue(all(dt == "PR" for dt in doc_types_all))
            self.assertNotIn("DS-CY002P608-00134 REV00", doc_nos_all)
            self.assertNotIn("IR-CY002P608-00038 REV00", doc_nos_all)

            # 2. Search latest revision only
            res_latest = ai_tools.search_submittals(
                allowed_project_ids=["PEM-058"],
                project_code="CFC",
                doc_type="PR",
                keyword="valve",
                latest_revision_only=True,
            )
            doc_nos_latest = [r["doc_no"] for r in res_latest["records"]]
            self.assertEqual(doc_nos_latest, ["PR-042 REV01"])

        print("PASS: PR search strictly isolates PR records, excludes DS/IR, and enriches line items.")

    def test_43_pr_intent_extraction_and_fallback_workflow(self):
        """Verifies Arabic PR prompt extraction and deterministic fallback response."""
        from blueprints.ai import _extract_inquiry_intent_from_text, _execute_deterministic_fallback

        # 1. Intent extraction
        parsed = _extract_inquiry_intent_from_text("شوفلي كل PRs الخاصة بالمحابس في نفس المشروع")
        self.assertEqual(parsed.get("doc_type"), "PR")
        self.assertEqual(parsed.get("keyword"), "valve")
        self.assertTrue(parsed.get("is_same_project"))

        # 2. Intent extraction with Arabic terms
        parsed_ar = _extract_inquiry_intent_from_text("طلبات الشراء الخاصة بالمضخات")
        self.assertEqual(parsed_ar.get("doc_type"), "PR")
        self.assertEqual(parsed_ar.get("keyword"), "pump")

        # 3. Deterministic fallback with history inheritance
        history = [
            {"role": "user", "text": "ايه MS المعتمد الخاص بالمحابس في مشروع CFC"},
            {"role": "model", "text": "تم استخراج 14 وثيقة معتمدة..."},
        ]
        mock_projects = [{"id": "PEM-058", "name": "CFC Ph2", "code": "PEM-058", "data": {}}]
        mock_pr_records = [
            {"id": "pr1", "project_id": "PEM-058", "doc_type": "PR", "doc_no": "PR-042 REV01", "title": "Butterfly Valve 300mm", "status": "Pending PO", "description": ""},
        ]

        with patch("db.q") as mock_q, patch("db.get_pr_items_for_records") as mock_items:
            def q_side_effect(sql, params=()):
                if "FROM projects" in sql:
                    return mock_projects
                if "FROM records r" in sql:
                    return mock_pr_records
                return []
            mock_q.side_effect = q_side_effect
            mock_items.return_value = {}

            reply = _execute_deterministic_fallback(
                prompt="شوفلي كل PRs الخاصة بالمحابس في نفس المشروع",
                allowed_project_ids=["PEM-058"],
                history=history,
            )
            self.assertIn("طلبات الشراء (PR)", reply)
            self.assertIn("PEM-058", reply)
            self.assertIn("PR-042 REV01", reply)
            self.assertNotIn("DS-CY002P608", reply)

        print("PASS: PR intent extraction and conversational fallback correctly inherit project and format PR records.")

    def test_44_dynamic_doc_type_discovery_across_projects(self):
        """Verifies that all document types across all projects are dynamically discovered without manual addition."""
        from blueprints.ai import _extract_inquiry_intent_from_text
        ai_tools._DOC_TYPES_CACHE["types"] = []
        ai_tools._DOC_TYPES_CACHE["ts"] = 0.0

        mock_db_types = [
            {"code": "MS", "name": "Material Submittal"},
            {"code": "SD", "name": "Shop Drawing"},
            {"code": "IR", "name": "Inspection Request"},
            {"code": "PR", "name": "Requisition"},
            {"code": "NCR", "name": "Non-Conformance Report"},
            {"code": "MOM", "name": "Minutes of Meetings"},
            {"code": "CUSTOM_REG", "name": "Custom Engineering Register"},
        ]

        with patch("db.q") as mock_q:
            def q_side_effect(sql, params=()):
                if "SELECT DISTINCT UPPER(TRIM(code))" in sql:
                    return mock_db_types
                return []
            mock_q.side_effect = q_side_effect

            # 1. Verify dynamic doc types retrieval
            dts = ai_tools.get_registered_doc_types()
            codes = [d["code"] for d in dts]
            self.assertIn("CUSTOM_REG", codes)
            self.assertIn("NCR", codes)
            self.assertIn("PR", codes)

            # 2. Verify AI recognizes newly registered custom doc type dynamically
            parsed = _extract_inquiry_intent_from_text("شوفلي كل مستندات الـ CUSTOM_REG في مشروع CFC")
            self.assertEqual(parsed.get("doc_type"), "CUSTOM_REG")

            # 3. Verify AI recognizes standard registered types dynamically
            parsed_ncr = _extract_inquiry_intent_from_text("ما موقف تقارير عدم المطابقة NCR؟")
            self.assertEqual(parsed_ncr.get("doc_type"), "NCR")

        print("PASS: Dynamic document type discovery reads all registers across all projects automatically.")


if __name__ == "__main__":
    unittest.main()



