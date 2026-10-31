import unittest
import json
import inspect
import os
from unittest.mock import MagicMock, patch

os.environ['SECRET_KEY'] = 'test_secret_key_12345678901234567890123456789012'

import html_render
import db
import auth
from app import app
import blueprints.records as records_mod
import blueprints.users as users_mod


class TestDynamicDepartmentsAndDashboard(unittest.TestCase):
    def setUp(self):
        self.app = app
        self.client = app.test_client()
        self.app.config["TESTING"] = True

    def test_01_om_dashboard_has_global_search_and_unified_toolbar(self):
        """Global Search (SPOTLIGHT) and Unified Navigation Toolbar are embedded in _render_om_dashboard."""
        mock_user = {"username": "a.raffat", "role": "admin", "primary_workspace": "OM"}
        accessible_projs = [{"id": "02", "name": "CFC District Cooling", "code": "02", "workspace_id": "OM"}]
        with patch.object(db, 'get_workspace', return_value={"id": "OM", "code": "OM", "name": "Operation & Maintenance (O&M)", "icon": "⚙️", "name_ar": "الصيانة والتشغيل"}):
            html = html_render._render_om_dashboard(
                mock_user, accessible_projs, ws_switcher_html="<div>WS</div>",
                btns="", uname="a.raffat", rlbl="Admin", rbg="#10b981", active_ws="OM"
            )
        # Global search modal and trigger must be present
        self.assertIn('id="spotlight-modal"', html)
        self.assertIn('openGlobalSearch()', html)
        
        # Unified toolbar tabs must be present
        self.assertIn('showOmTab', html)
        self.assertIn('📊 Overview', html)
        self.assertIn('📈 Analytics', html)
        self.assertIn('⚠️ Overdue', html)
        self.assertIn('📋 Executive', html)
        self.assertIn('🌟 Daily Digest', html)
        self.assertIn('📜 Audit Log', html)
        
        # In-page tab panes must be present (including Audit Log!)
        self.assertIn('id="tab-om-overview"', html)
        self.assertIn('id="tab-om-daily-digest"', html)
        self.assertIn('id="tab-om-executive"', html)
        self.assertIn('id="tab-om-overdue"', html)
        self.assertIn('id="tab-om-analytics"', html)
        self.assertIn('id="tab-om-audit"', html)
        
        # Old popup modal must NOT be present
        self.assertNotIn('id="om-audit-modal"', html)
        self.assertIn('showOmTab(&#39;audit&#39;, this)', html)
        self.assertIn('omAuditPage', html)

    def test_02_register_exposes_columns_and_lists_to_department_admin(self):
        """Department Admin (role=admin) can see Columns, Lists, and Holidays in Register toolbar."""
        admin_user = {"username": "a.raffat", "role": "admin", "primary_workspace": "OM"}
        proj = {"id": "02", "name": "CFC Facility", "workspace_id": "OM", "dts": []}
        
        with patch.object(auth, 'can_edit', return_value=True), \
             patch.object(db, 'get_columns', return_value=[]), \
             patch.object(db, 'get_doc_types', return_value=[{"id": "WO", "code": "WO", "name": "Work Order"}]), \
             patch.object(db, 'get_user_workspaces', return_value=[{"id": "OM"}]), \
             patch.object(db, 'get_logo', return_value=None), \
             patch.object(db, 'get_lists', return_value={}), \
             patch.object(db, 'get_records', return_value=[]):
            html = html_render.render_register(admin_user, proj, active_ws_id="OM")
                            
        self.assertIn("manageColumns()", html)
        self.assertIn("openLists()", html)
        self.assertIn("openSettings()", html)  # Holidays button

    def test_03_global_search_scopes_to_workspace(self):
        """api_search_global accepts workspace_id and scopes to workspace projects."""
        with app.test_request_context("/api/records/search/global?q=cooler&workspace_id=OM"):
            with patch("auth.current_user", return_value={"username": "a.raffat", "role": "admin"}), \
                 patch("blueprints.records.current_user", return_value={"username": "a.raffat", "role": "admin"}), \
                 patch("db.get_projects_for_user", return_value=[{"id": "02", "code": "02", "name": "CFC"}]) as mock_gp, \
                 patch("db.log_action"), \
                 patch("db.q", return_value=[]):
                resp = records_mod.api_search_global()
                mock_gp.assert_called_with("a.raffat", role="admin", workspace_id="OM")

    def test_04_workspaces_api_crud_and_superadmin_guard(self):
        """Superadmin can add and update workspaces; non-superadmins are rejected."""
        # Non-superadmin cannot create workspace
        with app.test_request_context("/api/workspaces", method="POST", json={"action": "add", "id": "QA", "name": "Quality"}):
            with patch("auth.current_user", return_value={"username": "a.raffat", "role": "admin"}), \
                 patch("blueprints.users.current_user", return_value={"username": "a.raffat", "role": "admin"}):
                resp = users_mod.api_workspaces()
                status = resp[1] if isinstance(resp, tuple) else resp.status_code
                self.assertEqual(status, 403)

        # Superadmin can add workspace
        with app.test_request_context("/api/workspaces", method="POST", json={
            "action": "add", "id": "QA", "code": "QA", "name": "Quality Assurance", "name_ar": "توكيد الجودة", "icon": "🛡️"
        }):
            with patch("auth.current_user", return_value={"username": "admin", "role": "superadmin"}), \
                 patch("blueprints.users.current_user", return_value={"username": "admin", "role": "superadmin"}), \
                 patch("db.create_workspace") as mock_cw, \
                 patch("db.log_action"):
                resp = users_mod.api_workspaces()
                raw_resp = resp[0] if isinstance(resp, tuple) else resp
                data = json.loads(raw_resp.get_data(as_text=True))
                self.assertTrue(data.get("ok"))
                mock_cw.assert_called_with("QA", "QA", "Quality Assurance", "توكيد الجودة", "🛡️", "general", 0)

        # Superadmin can update workspace
        with app.test_request_context("/api/workspaces", method="POST", json={
            "action": "update", "id": "OM", "name": "Operation & Maintenance (O&M)", "code": "O&M", "icon": "⚙️"
        }):
            with patch("auth.current_user", return_value={"username": "admin", "role": "superadmin"}), \
                 patch("blueprints.users.current_user", return_value={"username": "admin", "role": "superadmin"}), \
                 patch("db.update_workspace") as mock_uw, \
                 patch("db.log_action"):
                resp = users_mod.api_workspaces()
                raw_resp = resp[0] if isinstance(resp, tuple) else resp
                data = json.loads(raw_resp.get_data(as_text=True))
                self.assertTrue(data.get("ok"))
                mock_uw.assert_called_with("OM", code="O&M", name="Operation & Maintenance (O&M)", name_ar=None, icon="⚙️", sort_order=None, is_active=None)

    def test_05_admin_modal_tab_navigation(self):
        """Admin modal JS includes switchAdminTab, renderDepartmentsTable, and editDepartment."""
        js = html_render.SHARED_JS
        self.assertIn("switchAdminTab", js)
        self.assertIn("renderDepartmentsTable", js)
        self.assertIn("editDepartment", js)
        self.assertIn("submitDepartmentForm", js)
        self.assertIn("admin-view-depts", js)
        self.assertIn("admin-tab-nav", js)

    def test_06_dynamic_dashboard_routing_for_any_workspace(self):
        """render_dashboard routes any non-ENG workspace to dynamic workspace dashboard."""
        mock_qa_admin = {"username": "qa.manager", "role": "admin", "primary_workspace": "QA"}
        with patch.object(db, 'get_user_workspaces', return_value=[{"id": "QA", "name": "Quality Assurance"}]), \
             patch.object(db, 'get_projects_for_user', return_value=[]), \
             patch.object(db, 'get_workspace', return_value={
                 "id": "QA", "code": "QA", "name": "Quality Assurance", "icon": "🛡️", "name_ar": "توكيد الجودة"
             }):
            html = html_render.render_dashboard(mock_qa_admin, active_ws_id="QA")
            self.assertIn('data-workspace="QA"', html)
            self.assertIn('Quality Assurance', html)
            self.assertIn('🛡️', html)
            self.assertIn('showOmTab', html)
            self.assertIn('id="tab-om-audit"', html)

    def test_07_audit_api_scopes_to_workspace(self):
        """api_audit scopes to workspace when workspace_id is provided."""
        with app.test_request_context("/api/audit?workspace_id=OM"):
            with patch("auth.current_user", return_value={"username": "a.raffat", "role": "admin"}), \
                 patch("app.current_user", return_value={"username": "a.raffat", "role": "admin"}), \
                 patch("db.get_admin_workspaces_for_user", return_value=["OM"]), \
                 patch("db.get_projects_for_user", return_value=[{"id": "02", "name": "CFC", "code": "02"}]) as mock_gpu, \
                 patch("db.get_users_for_admin", return_value=[{"username": "a.raffat"}]), \
                 patch("db.get_audit_log", return_value=[]), \
                 patch("db.get_projects", return_value=[{"id": "02", "name": "CFC", "code": "02"}]), \
                 patch("db.get_audit_actions", return_value=[]):
                client = self.client
                # Invoke route
                resp = client.get("/api/audit?workspace_id=OM")
                self.assertEqual(resp.status_code, 200)
                mock_gpu.assert_called_with("a.raffat", "admin", workspace_id="OM")


if __name__ == "__main__":
    unittest.main()
