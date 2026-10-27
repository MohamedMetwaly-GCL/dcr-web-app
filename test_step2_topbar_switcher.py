"""test_step2_topbar_switcher.py - Unit tests for Topbar Workspace Switcher."""
import os
os.environ['SECRET_KEY'] = 'test_secret_key_12345678901234567890123456789012'

import unittest
from unittest.mock import patch
import html_render
import app


class TestTopbarWorkspaceSwitcher(unittest.TestCase):

    def test_single_workspace_user(self):
        # A user who only has 1 workspace (e.g. ENG)
        user = {'username': 'site_eng', 'role': 'viewer'}
        with patch('db.get_user_workspaces', return_value=[
            {'id': 'ENG', 'code': 'ENG', 'name': 'Engineering & Projects', 'name_ar': 'المشروعات الهندسية', 'icon': '🏗️'}
        ]):
            html = html_render._workspace_switcher_html(user, active_ws_id='ENG')
            self.assertIn('ws-switcher', html)
            self.assertIn('المشروعات الهندسية', html)
            self.assertIn('single-ws', html)
            self.assertNotIn('ws-dropdown-menu', html)  # No dropdown for single workspace

    def test_multi_workspace_user(self):
        # A superadmin or multi-dept user
        user = {'username': 'admin', 'role': 'superadmin'}
        workspaces = [
            {'id': 'ENG', 'code': 'ENG', 'name': 'Engineering & Projects', 'name_ar': 'المشروعات الهندسية', 'icon': '🏗️'},
            {'id': 'OM', 'code': 'OM', 'name': 'Operation & Maintenance', 'name_ar': 'الصيانة والتشغيل', 'icon': '⚙️'}
        ]
        with patch('db.get_user_workspaces', return_value=workspaces):
            # When ENG is active
            html_eng = html_render._workspace_switcher_html(user, active_ws_id='ENG')
            self.assertIn('ws-dropdown-menu', html_eng)
            self.assertIn('المشروعات الهندسية', html_eng)
            self.assertIn('الصيانة والتشغيل', html_eng)
            self.assertIn('href="/?ws=ENG"', html_eng)
            self.assertIn('href="/?ws=OM"', html_eng)
            self.assertIn('ws-arrow', html_eng)

            # When OM is active
            html_om = html_render._workspace_switcher_html(user, active_ws_id='OM')
            self.assertIn('الصيانة والتشغيل', html_om)

    def test_render_dashboard_includes_switcher(self):
        user = {'username': 'admin', 'role': 'superadmin'}
        with patch('db.get_user_workspaces', return_value=[
            {'id': 'ENG', 'code': 'ENG', 'name': 'Engineering & Projects', 'name_ar': 'المشروعات الهندسية', 'icon': '🏗️'},
            {'id': 'OM', 'code': 'OM', 'name': 'Operation & Maintenance', 'name_ar': 'الصيانة والتشغيل', 'icon': '⚙️'}
        ]):
            html = html_render.render_dashboard(user)
            self.assertIn('id="ws-switcher"', html)
            self.assertIn('toggleWsDropdown', html)
            self.assertIn('selectWorkspace', html)


if __name__ == '__main__':
    unittest.main()
