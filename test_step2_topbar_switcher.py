"""test_step2_topbar_switcher.py - Unit tests for Topbar Workspace Switcher & O&M Dashboard."""
import os
os.environ['SECRET_KEY'] = 'test_secret_key_12345678901234567890123456789012'

import unittest
from unittest.mock import patch
import html_render
import app
import db


class TestTopbarWorkspaceSwitcher(unittest.TestCase):

    def test_single_workspace_user_english(self):
        user = {'username': 'site_eng', 'role': 'viewer'}
        with patch('db.get_user_workspaces', return_value=[
            {'id': 'ENG', 'code': 'ENG', 'name': 'Engineering & Projects', 'name_ar': 'المشروعات الهندسية', 'icon': '🏗️'}
        ]):
            html = html_render._workspace_switcher_html(user, active_ws_id='ENG')
            self.assertIn('ws-switcher', html)
            self.assertIn('Engineering & Projects', html)
            self.assertIn('single-ws', html)
            self.assertNotIn('ws-dropdown-menu', html)

    def test_multi_workspace_user_english(self):
        user = {'username': 'admin', 'role': 'superadmin'}
        workspaces = [
            {'id': 'ENG', 'code': 'ENG', 'name': 'Engineering & Projects', 'name_ar': 'المشروعات الهندسية', 'icon': '🏗️'},
            {'id': 'OM', 'code': 'OM', 'name': 'Operation & Maintenance', 'name_ar': 'الصيانة والتشغيل', 'icon': '⚙️'}
        ]
        with patch('db.get_user_workspaces', return_value=workspaces):
            # When ENG is active
            html_eng = html_render._workspace_switcher_html(user, active_ws_id='ENG')
            self.assertIn('ws-dropdown-menu', html_eng)
            self.assertIn('Enterprise Workspaces', html_eng)
            self.assertIn('Engineering & Projects', html_eng)
            self.assertIn('Operation & Maintenance', html_eng)
            self.assertIn('href="/?ws=ENG"', html_eng)
            self.assertIn('href="/?ws=OM"', html_eng)
            self.assertIn('ws-arrow', html_eng)

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

    def test_render_register_does_not_contain_switcher(self):
        user = {'username': 'admin', 'role': 'superadmin'}
        proj = {'id': '24CP01', 'name': 'Alpha', 'code': 'A01', 'workspace_id': 'ENG'}
        with patch('db.get_user_workspaces', return_value=[
            {'id': 'ENG', 'code': 'ENG', 'name': 'Engineering & Projects', 'name_ar': 'المشروعات الهندسية', 'icon': '🏗️'},
            {'id': 'OM', 'code': 'OM', 'name': 'Operation & Maintenance', 'name_ar': 'الصيانة والتشغيل', 'icon': '⚙️'}
        ]), patch('db.get_doc_types', return_value=[]), patch('db.get_logo', return_value=None):
            html = html_render.render_register(user, proj)
            self.assertNotIn('id="ws-switcher"', html)

    def test_render_dashboard_om_workspace(self):
        user = {'username': 'om_admin', 'role': 'admin'}
        with patch('db.get_user_workspaces', return_value=[
            {'id': 'OM', 'code': 'OM', 'name': 'Operation & Maintenance', 'name_ar': 'الصيانة والتشغيل', 'icon': '⚙️'}
        ]), patch('db.get_projects_for_user', return_value=[
            {'id': 'OM-FAC-01', 'name': 'CFC District Cooling O&M', 'code': 'OM01', 'workspace_id': 'OM', 'data': {'client': 'Al-Futtaim'}}
        ]), patch('db.q', return_value={'total': 12, 'completed': 8, 'pending': 4}):
            html = html_render.render_dashboard(user, active_ws_id='OM')
            self.assertIn('Operation & Maintenance (O&M)', html)
            self.assertIn('CFC District Cooling O&M', html)
            self.assertIn('Al-Futtaim', html)
            self.assertNotIn('Submittals KPI Summary', html)  # Zero consultant submittal jargon in O&M!


if __name__ == '__main__':
    unittest.main()
