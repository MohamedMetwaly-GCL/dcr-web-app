# test_step4_user_management.py
import unittest
import json
import os

os.environ['SECRET_KEY'] = 'test_secret_key_12345678901234567890123456789012'

from html_render import render_dashboard, render_register, SHARED_JS, BASE_CSS
import blueprints.users as users_bp
import auth
import db
import app


class TestStep4UserManagement(unittest.TestCase):

    def test_01_modals_and_js_present(self):
        """Verify new admin modal, permissions modal, and JS functions exist in SHARED_JS."""
        self.assertIn('id="admin-modal"', SHARED_JS)
        self.assertIn('id="permissions-modal"', SHARED_JS)
        self.assertIn('async function openAdmin()', SHARED_JS)
        self.assertIn('function openPermissionsModal(username)', SHARED_JS)
        self.assertIn('function renderAdminTable()', SHARED_JS)
        self.assertIn('async function togglePermAccess', SHARED_JS)
        self.assertIn('async function togglePermDC', SHARED_JS)
        self.assertIn('async function permSelectAll', SHARED_JS)
        self.assertIn('async function permClearAll', SHARED_JS)
        self.assertIn('async function submitAddUser()', SHARED_JS)
        self.assertIn('async function updUsrDept', SHARED_JS)
        self.assertIn('async function exportUsersCSV()', SHARED_JS)

    def test_02_css_styling_present(self):
        """Verify modern responsive table & badge CSS exists in BASE_CSS."""
        self.assertIn('.user-tbl-wrap', BASE_CSS)
        self.assertIn('.user-tbl', BASE_CSS)
        self.assertIn('.dept-badge-eng', BASE_CSS)
        self.assertIn('.dept-badge-om', BASE_CSS)
        self.assertIn('.perm-table', BASE_CSS)
        self.assertIn('.uact-btn-perm', BASE_CSS)
        self.assertIn('.add-user-card', BASE_CSS)

    def test_03_rendered_pages_include_modals(self):
        """Verify that rendered dashboard and register pages contain the new modals."""
        orig_get_dts = db.get_doc_types
        orig_get_cols = getattr(db, 'get_columns', None)
        orig_get_lists = getattr(db, 'get_lists', None)
        orig_get_records = getattr(db, 'get_records', None)
        orig_get_logo = getattr(db, 'get_logo', None)
        orig_get_projects_for_user = getattr(db, 'get_projects_for_user', None)
        orig_get_workspaces = getattr(db, 'get_workspaces', None)
        orig_get_user_workspaces = getattr(db, 'get_user_workspaces', None)

        db.get_doc_types = lambda pid: [{'id': 1, 'code': 'DWG', 'name': 'Drawings'}]
        db.get_columns = lambda dt_id: []
        db.get_lists = lambda dt_id: []
        db.get_records = lambda dt_id: []
        db.get_logo = lambda pid, key: None
        db.get_projects_for_user = lambda uname, role, workspace_id=None: []
        db.get_workspaces = lambda active_only=True: [{'id': 'ENG', 'code': 'ENG', 'name': 'Engineering'}]
        db.get_user_workspaces = lambda uname: [{'id': 'ENG', 'code': 'ENG', 'name': 'Engineering', 'user_role': 'admin'}]

        try:
            dummy_admin = {'username': 'eng_admin', 'role': 'admin'}
            dash_html = render_dashboard(dummy_admin)
            self.assertIn('id="admin-modal"', dash_html)
            self.assertIn('id="permissions-modal"', dash_html)

            dummy_proj = {
                'id': 'TEST_P1',
                'name': 'Test Project',
                'code': 'TP1',
                'workspace_id': 'ENG',
                'data': {}
            }
            reg_html = render_register(dummy_admin, dummy_proj)
            self.assertIn('id="admin-modal"', reg_html)
            self.assertIn('id="permissions-modal"', reg_html)
        finally:
            db.get_doc_types = orig_get_dts
            if orig_get_cols: db.get_columns = orig_get_cols
            if orig_get_lists: db.get_lists = orig_get_lists
            if orig_get_records: db.get_records = orig_get_records
            if orig_get_logo: db.get_logo = orig_get_logo
            if orig_get_projects_for_user: db.get_projects_for_user = orig_get_projects_for_user
            if orig_get_workspaces: db.get_workspaces = orig_get_workspaces
            if orig_get_user_workspaces: db.get_user_workspaces = orig_get_user_workspaces

    def test_04_backend_update_department(self):
        """Verify that POST /api/users supports update_department / update_workspace."""
        mock_user = {'username': 'superadmin', 'role': 'superadmin'}
        target_username = 'test_engineer'

        assigned = {}
        def mock_set_primary(uname, ws_id):
            assigned['primary'] = (uname, ws_id)

        def mock_assign_ws(uname, ws_id, role):
            assigned['ws'] = (uname, ws_id, role)

        def mock_get_user(uname):
            if uname == target_username:
                return {'username': target_username, 'role': 'editor', 'primary_workspace': 'ENG'}
            return None

        orig_set_primary = getattr(db, 'set_user_primary_workspace', None)
        orig_assign_ws = getattr(db, 'assign_user_workspace', None)
        orig_get_user = db.get_user
        orig_log = db.log_action
        orig_current_user = auth.current_user
        orig_bp_current_user = users_bp.current_user

        db.set_user_primary_workspace = mock_set_primary
        db.assign_user_workspace = mock_assign_ws
        db.get_user = mock_get_user
        db.log_action = lambda *args, **kwargs: None
        auth.current_user = lambda: mock_user
        users_bp.current_user = lambda: mock_user

        try:
            with app.app.test_request_context(
                '/api/users',
                method='POST',
                data=json.dumps({
                    'action': 'update_department',
                    'username': target_username,
                    'department': 'OM'
                }),
                content_type='application/json'
            ):
                resp = users_bp.api_users_action()
                if isinstance(resp, tuple):
                    res_obj, status = resp
                else:
                    res_obj, status = resp, 200
                self.assertEqual(status, 200)
                data = res_obj.get_json() if hasattr(res_obj, 'get_json') else res_obj
                self.assertTrue(data.get('ok'))
                self.assertEqual(assigned.get('primary'), (target_username, 'OM'))
                self.assertEqual(assigned.get('ws'), (target_username, 'OM', 'editor'))
        finally:
            db.get_user = orig_get_user
            db.log_action = orig_log
            auth.current_user = orig_current_user
            users_bp.current_user = orig_bp_current_user
            if orig_set_primary:
                db.set_user_primary_workspace = orig_set_primary
            if orig_assign_ws:
                db.assign_user_workspace = orig_assign_ws

    def test_05_dept_admin_cannot_assign_other_dept_projects(self):
        """Verify that a Department Admin cannot assign projects outside their workspace."""
        mock_dept_admin = {'username': 'eng_admin', 'role': 'admin'}

        orig_current_user = auth.current_user
        orig_bp_current_user = users_bp.current_user
        orig_is_super = users_bp.is_superadmin
        orig_auth_is_super = auth.is_superadmin
        orig_is_ws_admin = getattr(db, 'is_user_workspace_admin', None)
        orig_get_admin_ws = db.get_admin_workspaces_for_user
        orig_get_proj = db.get_project
        orig_get_user = db.get_user

        auth.current_user = lambda: mock_dept_admin
        users_bp.current_user = lambda: mock_dept_admin
        users_bp.is_superadmin = lambda u: False
        auth.is_superadmin = lambda u: False
        db.is_user_workspace_admin = lambda uname: True
        db.get_admin_workspaces_for_user = lambda u: ['ENG']
        db.get_user = lambda u: mock_dept_admin
        # Project belongs to OM
        db.get_project = lambda pid: {'id': pid, 'workspace_id': 'OM'}

        try:
            with app.app.test_request_context(
                '/api/users',
                method='POST',
                data=json.dumps({
                    'action': 'assign',
                    'username': 'some_user',
                    'project_id': 'OM_PROJ_1',
                    'is_dc': False
                }),
                content_type='application/json'
            ):
                resp = users_bp.api_users_action()
                if isinstance(resp, tuple):
                    res_obj, status = resp
                else:
                    res_obj, status = resp, 200
                self.assertEqual(status, 403)
                data = res_obj.get_json() if hasattr(res_obj, 'get_json') else res_obj
                self.assertFalse(data.get('ok'))
                self.assertIn('Cannot assign project from another department', data.get('error'))
        finally:
            auth.current_user = orig_current_user
            users_bp.current_user = orig_bp_current_user
            users_bp.is_superadmin = orig_is_super
            auth.is_superadmin = orig_auth_is_super
            if orig_is_ws_admin: db.is_user_workspace_admin = orig_is_ws_admin
            db.get_admin_workspaces_for_user = orig_get_admin_ws
            db.get_project = orig_get_proj
            db.get_user = orig_get_user

    def test_06_refinements_and_deprovisioning(self):
        """Verify 8 refinements: All Departments badge, O&M label, Checkboxes, No DC column, OM Audit modal."""
        # 1. Verify DC column is removed from perm table header
        self.assertNotIn('Document Controller (DC)', SHARED_JS)
        self.assertIn('async function toggleUserDept', SHARED_JS)
        self.assertIn('dept-checkboxes', SHARED_JS)
        self.assertIn('🌐 All Departments', SHARED_JS)
        self.assertIn('Operation & Maintenance (O&M)', SHARED_JS)

        # 2. Verify backend update_workspaces action
        orig_current_user = auth.current_user
        orig_bp_current_user = users_bp.current_user
        orig_is_super = users_bp.is_superadmin
        orig_set_ws = getattr(db, 'set_user_workspaces', None)
        orig_get_user = db.get_user

        mock_super = {'username': 'admin', 'role': 'superadmin'}
        auth.current_user = lambda: mock_super
        users_bp.current_user = lambda: mock_super
        users_bp.is_superadmin = lambda u: True
        db.get_user = lambda u: {'username': 'sakr', 'role': 'editor', 'primary_workspace': 'ENG'}

        captured = {}
        def mock_set_workspaces(uname, ws_list, role):
            captured['set'] = (uname, ws_list, role)
            return True
        db.set_user_workspaces = mock_set_workspaces

        try:
            with app.app.test_request_context(
                '/api/users',
                method='POST',
                data=json.dumps({
                    'action': 'update_workspaces',
                    'username': 'sakr',
                    'workspaces': ['ENG']
                }),
                content_type='application/json'
            ):
                resp = users_bp.api_users_action()
                res_obj = resp[0] if isinstance(resp, tuple) else resp
                data = res_obj.get_json() if hasattr(res_obj, 'get_json') else res_obj
                self.assertTrue(data.get('ok'))
                self.assertEqual(captured.get('set'), ('sakr', ['ENG'], 'editor'))
        finally:
            auth.current_user = orig_current_user
            users_bp.current_user = orig_bp_current_user
            users_bp.is_superadmin = orig_is_super
            db.get_user = orig_get_user
            if orig_set_ws:
                db.set_user_workspaces = orig_set_ws

    def test_07_om_dashboard_defaults_and_audit(self):
        """Verify O&M admin defaults to OM dashboard and sees Audit Log modal."""
        orig_get_user_ws = getattr(db, 'get_user_workspaces', None)
        orig_get_projs_user = getattr(db, 'get_projects_for_user', None)

        mock_om_admin = {'username': 'om_manager', 'role': 'admin', 'primary_workspace': 'OM'}
        db.get_user_workspaces = lambda u: [{'id': 'OM', 'code': 'O&M', 'name': 'Operation & Maintenance (O&M)', 'user_role': 'admin'}]
        db.get_projects_for_user = lambda u, r, workspace_id=None: [{'id': 'OM_FAC_1', 'name': 'CFC Cooling', 'code': 'CFC01', 'workspace_id': 'OM'}]

        try:
            # When active_ws_id is None, it should default to OM because om_manager primary_workspace is OM
            html = render_dashboard(mock_om_admin, active_ws_id=None)
            self.assertIn('id="om-audit-modal"', html)
            self.assertIn('openOmAuditModal()', html)
            self.assertIn('Audit Log', html)
            self.assertIn('Operation & Maintenance (O&M)', html)
        finally:
            if orig_get_user_ws:
                db.get_user_workspaces = orig_get_user_ws
            if orig_get_projs_user:
                db.get_projects_for_user = orig_get_projs_user


if __name__ == '__main__':
    unittest.main()
