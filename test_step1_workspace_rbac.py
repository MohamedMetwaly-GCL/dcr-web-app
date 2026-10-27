"""test_step1_workspace_rbac.py - Comprehensive Unit Tests for Step 1 Multi-Workspace Schema & RBAC."""
import os
os.environ['SECRET_KEY'] = 'test_secret_key_12345678901234567890123456789012'

import unittest
from unittest.mock import MagicMock, patch
import app
import db
import auth


class TestWorkspaceRBAC(unittest.TestCase):

    def setUp(self):
        self.app = app.app.test_client()

    def test_auth_is_superadmin(self):
        self.assertTrue(auth.is_superadmin({'username': 'admin', 'role': 'superadmin'}))
        self.assertFalse(auth.is_superadmin({'username': 'om_mgr', 'role': 'admin'}))
        self.assertFalse(auth.is_superadmin({'username': 'eng_user', 'role': 'viewer'}))
        self.assertFalse(auth.is_superadmin(None))

    @patch('db.get_user')
    @patch('db.q')
    def test_is_user_workspace_admin(self, mock_q, mock_get_user):
        # 1. Superadmin is admin everywhere
        mock_get_user.return_value = {'username': 'super', 'role': 'superadmin'}
        self.assertTrue(db.is_user_workspace_admin('super', 'ENG'))
        self.assertTrue(db.is_user_workspace_admin('super', 'OM'))

        # 2. OM Admin is admin in OM but not in ENG
        mock_get_user.return_value = {'username': 'om_boss', 'role': 'admin', 'primary_workspace': 'OM'}
        mock_q.side_effect = lambda sql, params, **kw: {'role': 'admin'} if params == ('om_boss', 'OM') else None
        self.assertTrue(db.is_user_workspace_admin('om_boss', 'OM'))
        self.assertFalse(db.is_user_workspace_admin('om_boss', 'ENG'))

    @patch('db.get_user')
    @patch('db.get_admin_workspaces_for_user')
    @patch('db.q')
    def test_get_users_for_admin_isolation(self, mock_q, mock_admin_ws, mock_get_user):
        # Superadmin sees all users
        mock_get_user.return_value = {'username': 'admin', 'role': 'superadmin'}
        mock_q.return_value = [
            {'username': 'admin', 'role': 'superadmin', 'primary_workspace': 'ENG'},
            {'username': 'om_user', 'role': 'viewer', 'primary_workspace': 'OM'},
            {'username': 'eng_user', 'role': 'editor', 'primary_workspace': 'ENG'}
        ]
        users = db.get_users_for_admin('admin')
        self.assertEqual(len(users), 3)

        # OM Admin ONLY sees OM users, strictly excluding superadmin and ENG users
        mock_get_user.return_value = {'username': 'om_admin', 'role': 'admin', 'primary_workspace': 'OM'}
        mock_admin_ws.return_value = ['OM']
        mock_q.return_value = [
            {'username': 'om_user', 'role': 'viewer', 'primary_workspace': 'OM'}
        ]
        dept_users = db.get_users_for_admin('om_admin')
        self.assertEqual(len(dept_users), 1)
        self.assertEqual(dept_users[0]['username'], 'om_user')

    @patch('auth.current_user')
    @patch('blueprints.users.current_user')
    @patch('blueprints.users.is_superadmin')
    @patch('db.is_user_workspace_admin')
    @patch('db.get_admin_workspaces_for_user')
    @patch('db.add_user')
    def test_dept_admin_cannot_create_superadmin(self, mock_add_user, mock_admin_ws, mock_is_ws_admin, mock_is_super, mock_bp_user, mock_auth_user):
        user_dict = {'username': 'om_admin', 'role': 'admin'}
        mock_auth_user.return_value = user_dict
        mock_bp_user.return_value = user_dict
        mock_is_super.return_value = False
        mock_is_ws_admin.return_value = True
        mock_admin_ws.return_value = ['OM']

        with app.app.test_request_context('/api/users', method='POST', json={
            'action': 'add',
            'username': 'bad_actor',
            'password': 'secret_password_123',
            'role': 'superadmin',
            'workspace_id': 'OM'
        }):
            from blueprints.users import api_users_action
            resp = api_users_action()
            if isinstance(resp, tuple):
                data, code = resp[0].get_json(), resp[1]
            else:
                data, code = resp.get_json(), 200
            self.assertEqual(code, 403)
            self.assertFalse(data.get('ok'))
            self.assertIn('superadmin', data.get('error', '').lower())
            mock_add_user.assert_not_called()

    @patch('auth.current_user')
    @patch('blueprints.users.current_user')
    @patch('blueprints.users.is_superadmin')
    @patch('db.is_user_workspace_admin')
    @patch('db.get_admin_workspaces_for_user')
    @patch('db.get_user')
    @patch('db.get_user_workspaces')
    @patch('db.delete_user')
    def test_dept_admin_cannot_delete_other_dept_user(self, mock_delete_user, mock_get_ws, mock_get_user, mock_admin_ws, mock_is_ws_admin, mock_is_super, mock_bp_user, mock_auth_user):
        user_dict = {'username': 'om_admin', 'role': 'admin'}
        mock_auth_user.return_value = user_dict
        mock_bp_user.return_value = user_dict
        mock_is_super.return_value = False
        mock_is_ws_admin.return_value = True
        mock_admin_ws.return_value = ['OM']
        mock_get_user.return_value = {'username': 'eng_user', 'role': 'viewer', 'primary_workspace': 'ENG'}
        mock_get_ws.return_value = [{'id': 'ENG'}]

        with app.app.test_request_context('/api/users', method='POST', json={
            'action': 'delete',
            'username': 'eng_user'
        }):
            from blueprints.users import api_users_action
            resp = api_users_action()
            if isinstance(resp, tuple):
                data, code = resp[0].get_json(), resp[1]
            else:
                data, code = resp.get_json(), 200
            self.assertEqual(code, 403)
            self.assertFalse(data.get('ok'))
            mock_delete_user.assert_not_called()

    @patch('auth.current_user')
    @patch('blueprints.projects.current_user')
    @patch('blueprints.projects.is_superadmin')
    @patch('db.is_user_workspace_admin')
    @patch('db.get_admin_workspaces_for_user')
    @patch('db.create_project')
    def test_dept_admin_cannot_create_project_in_another_workspace(self, mock_create_proj, mock_admin_ws, mock_is_ws_admin, mock_is_super, mock_bp_user, mock_auth_user):
        user_dict = {'username': 'om_admin', 'role': 'admin'}
        mock_auth_user.return_value = user_dict
        mock_bp_user.return_value = user_dict
        mock_is_super.return_value = False
        mock_is_ws_admin.return_value = True
        mock_admin_ws.return_value = ['OM']

        with app.app.test_request_context('/api/projects/create', method='POST', json={
            'id': 'ENG_PROJ_99',
            'name': 'Civil Tower',
            'code': 'CT99',
            'workspace_id': 'ENG'
        }):
            from blueprints.projects import api_create_project
            resp = api_create_project()
            if isinstance(resp, tuple):
                data, code = resp[0].get_json(), resp[1]
            else:
                data, code = resp.get_json(), 200
            self.assertEqual(code, 403)
            self.assertFalse(data.get('ok'))
            mock_create_proj.assert_not_called()


if __name__ == '__main__':
    unittest.main()
