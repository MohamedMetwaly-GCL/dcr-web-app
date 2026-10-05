#!/usr/bin/env python3
"""
test_global_search_validation.py

Comprehensive test suite verifying:
1. Endpoint registration and unauthenticated access rejection (403 LOGIN_REQUIRED).
2. RBAC isolation:
   - Admin sees all active projects (user_scope == 'all').
   - Restricted user sees only their assigned projects (user_scope == 'filtered').
   - User with 0 assigned projects receives immediate empty result list [] without DB search.
3. Query parameters validation:
   - Query < 2 chars returns empty result list immediately.
   - dt_id filter parameter handling.
4. Frontend UI integration in html_render.py:
   - Navbar search trigger in render_dashboard and render_register.
   - Spotlight modal palette component with dark theme aesthetics.
   - Record highlight and auto-scroll handling.
"""

import os
import unittest
from unittest.mock import patch, MagicMock
from flask import Flask, jsonify

from blueprints.records import records_bp
import html_render

class TestGlobalSearch(unittest.TestCase):

    def setUp(self):
        self.app = Flask(__name__)
        self.app.secret_key = "test-secret-key-global-search"
        self.app.register_blueprint(records_bp)
        self.client = self.app.test_client()

    # -------------------------------------------------------------------------
    # 1. Unauthenticated Security Test
    # -------------------------------------------------------------------------
    @patch("blueprints.records.current_user")
    def test_unauthenticated_request_rejected(self, mock_user):
        mock_user.return_value = None
        resp = self.client.get("/api/records/search/global?q=valves")
        self.assertEqual(resp.status_code, 403)
        data = resp.get_json()
        self.assertEqual(data.get("error"), "LOGIN_REQUIRED")
        print("PASS: Unauthenticated search requests are rejected with 403 LOGIN_REQUIRED.")

    # -------------------------------------------------------------------------
    # 2. User with 0 Assigned Projects
    # -------------------------------------------------------------------------
    @patch("blueprints.records.current_user")
    @patch("blueprints.records.db.get_user_projects")
    @patch("blueprints.records.db.q")
    def test_zero_assigned_projects_returns_empty_immediately(self, mock_db_q, mock_get_user_projs, mock_user):
        mock_user.return_value = {"username": "restricted_user", "role": "viewer"}
        mock_get_user_projs.return_value = [] # 0 assigned projects

        resp = self.client.get("/api/records/search/global?q=valves")
        self.assertEqual(resp.status_code, 200)
        data = resp.get_json()
        self.assertEqual(data.get("total_found"), 0)
        self.assertEqual(data.get("user_scope"), "filtered")
        self.assertEqual(data.get("results"), [])
        # Ensure heavy database search query was never executed
        mock_db_q.assert_not_called()
        print("PASS: User with 0 assigned projects receives [] immediately without executing DB query.")

    # -------------------------------------------------------------------------
    # 3. Minimum Character Validation
    # -------------------------------------------------------------------------
    @patch("blueprints.records.current_user")
    @patch("blueprints.records.db.get_projects")
    def test_min_query_length_validation(self, mock_get_projs, mock_user):
        mock_user.return_value = {"username": "admin", "role": "admin"}
        mock_get_projs.return_value = [{"id": "P1", "code": "PEM-058"}]

        # 1 character query
        resp = self.client.get("/api/records/search/global?q=v")
        self.assertEqual(resp.status_code, 200)
        data = resp.get_json()
        self.assertEqual(data.get("total_found"), 0)
        self.assertEqual(data.get("results"), [])
        print("PASS: Search query under 2 characters returns empty results immediately.")

    # -------------------------------------------------------------------------
    # 4. RBAC Isolation: Regular User Restricted to Assigned Project
    # -------------------------------------------------------------------------
    @patch("blueprints.records.current_user")
    @patch("blueprints.records.db.get_user_projects")
    @patch("blueprints.records.db.get_projects")
    @patch("blueprints.records.db.q")
    def test_restricted_user_scope_enforcement(self, mock_db_q, mock_get_projs, mock_get_user_projs, mock_user):
        mock_user.return_value = {"username": "site_engineer", "role": "viewer"}
        # User only assigned to PEM-058
        mock_get_user_projs.return_value = [{"project_id": "PEM-058", "is_dc": False}]
        mock_get_projs.return_value = [{"id": "CFC DCP PH-2A", "code": "PEM-058"}]

        # Mock database returning 1 record
        mock_db_q.return_value = [
            {
                "id": "rec_001",
                "project_id": "CFC DCP PH-2A",
                "project_code": "PEM-058",
                "project_name": "CFC District Cooling Plant Phase 2",
                "dt_id": "MIR",
                "doc_no": "MIR-CY002P608-00124 REV00",
                "title": "Material Submittal for Butterfly Valves",
                "status": "Approved",
                "drive_link": "https://drive.google.com/test",
                "total_count": 1
            }
        ]

        resp = self.client.get("/api/records/search/global?q=valves")
        self.assertEqual(resp.status_code, 200)
        data = resp.get_json()
        self.assertEqual(data.get("user_scope"), "filtered")
        self.assertEqual(data.get("total_found"), 1)
        self.assertEqual(len(data.get("results")), 1)
        self.assertEqual(data["results"][0]["doc_no"], "MIR-CY002P608-00124 REV00")

        # Verify SQL arguments passed to db.q
        sql_call_args = mock_db_q.call_args[0]
        sql_query = sql_call_args[0]
        sql_params = sql_call_args[1]
        
        self.assertIn("(p.id = ANY(%s) OR p.code = ANY(%s))", sql_query)
        # Check that allowed_list contains only PEM-058 / CFC DCP PH-2A
        allowed_list = sql_params[0]
        self.assertIn("PEM-058", allowed_list)
        self.assertIn("CFC DCP PH-2A", allowed_list)
        # Confirm other projects (e.g. SUEZ, PEM-042) are NOT in the allowed list
        self.assertNotIn("PEM-042", allowed_list)
        self.assertNotIn("SUEZ", allowed_list)
        print("PASS: RBAC isolation restricts query parameters strictly to user's assigned project IDs.")

    # -------------------------------------------------------------------------
    # 5. RBAC Admin Scope: Global Access
    # -------------------------------------------------------------------------
    @patch("blueprints.records.current_user")
    @patch("blueprints.records.db.get_projects")
    @patch("blueprints.records.db.q")
    def test_admin_scope_enforcement(self, mock_db_q, mock_get_projs, mock_user):
        mock_user.return_value = {"username": "admin", "role": "admin"}
        mock_get_projs.return_value = [
            {"id": "P1", "code": "PEM-058"},
            {"id": "P2", "code": "PEM-042"},
            {"id": "P3", "code": "SUEZ"}
        ]
        mock_db_q.return_value = [
            {"id": "r1", "project_id": "P1", "project_code": "PEM-058", "dt_id": "MIR", "doc_no": "MIR-001", "title": "Valves 1", "status": "Approved", "drive_link": "", "total_count": 2},
            {"id": "r2", "project_id": "P2", "project_code": "PEM-042", "dt_id": "MIR", "doc_no": "MIR-002", "title": "Valves 2", "status": "Pending", "drive_link": "", "total_count": 2}
        ]

        resp = self.client.get("/api/records/search/global?q=valves")
        self.assertEqual(resp.status_code, 200)
        data = resp.get_json()
        self.assertEqual(data.get("user_scope"), "all")
        self.assertEqual(data.get("total_found"), 2)
        self.assertEqual(len(data.get("results")), 2)
        print("PASS: Admin search correctly returns user_scope='all' across all projects.")

    # -------------------------------------------------------------------------
    # 6. Document Type (dt_id) Filtering
    # -------------------------------------------------------------------------
    @patch("blueprints.records.current_user")
    @patch("blueprints.records.db.get_projects")
    @patch("blueprints.records.db.q")
    def test_dt_id_filtering(self, mock_db_q, mock_get_projs, mock_user):
        mock_user.return_value = {"username": "admin", "role": "admin"}
        mock_get_projs.return_value = [{"id": "P1", "code": "PEM-058"}]
        mock_db_q.return_value = []

        resp = self.client.get("/api/records/search/global?q=valves&dt_id=NOC")
        self.assertEqual(resp.status_code, 200)

        sql_call_args = mock_db_q.call_args[0]
        sql_query = sql_call_args[0]
        sql_params = sql_call_args[1]

        self.assertIn("UPPER(r.dt_id) = %s", sql_query)
        self.assertIn("NOC", sql_params)
        print("PASS: Optional dt_id filter parameter is correctly applied in SQL.")

    # -------------------------------------------------------------------------
    # 7. Frontend Integration in html_render.py
    # -------------------------------------------------------------------------
    @patch("html_render.can_edit")
    @patch("html_render.db.get_doc_types")
    @patch("html_render.db.get_logo")
    def test_frontend_markup_integration(self, mock_logo, mock_dts, mock_can_edit):
        mock_can_edit.return_value = True
        mock_dts.return_value = [{"id": "MIR", "name": "Material Submittal", "code": "MIR"}]
        mock_logo.return_value = None

        with self.app.test_request_context():
            # 1. Dashboard rendering
            dash_html = html_render.render_dashboard({"username": "admin", "role": "admin"})
            self.assertIn("global-search-trigger", dash_html, "Dashboard topbar must contain global search trigger")
            self.assertIn("openGlobalSearch()", dash_html, "Dashboard trigger must call openGlobalSearch()")
            self.assertIn("id=\"spotlight-modal\"", dash_html, "Dashboard must render spotlight-modal overlay")
            self.assertIn("id=\"spotlight-input\"", dash_html, "Dashboard must contain spotlight-input")
            self.assertIn("Ctrl+K", dash_html, "Dashboard must mention Ctrl+K shortcut")

            # 2. Register rendering
            proj = {"id": "PEM-058", "name": "CFC DCP PH-2A", "code": "PEM-058"}
            reg_html = html_render.render_register({"username": "admin", "role": "admin"}, proj)
            self.assertIn("global-search-trigger", reg_html, "Register topbar must contain global search trigger")
            self.assertIn("id=\"spotlight-modal\"", reg_html, "Register must render spotlight-modal overlay")
            self.assertIn("pendingHighlightId", reg_html, "Register must include highlight handling logic")
            print("PASS: Frontend markup and scripts verified in both Dashboard and Register views.")

if __name__ == "__main__":
    unittest.main()
