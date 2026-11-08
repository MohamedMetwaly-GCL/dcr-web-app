"""
test_workspace_assigned_projects_rbac.py - Scoped Multi-Workspace RBAC Unit Tests
Verifies that users with admin or mixed roles can access explicitly assigned projects
in departments where they hold non-admin roles (e.g. viewer/editor in ENG while admin in WH).
"""

import unittest
from unittest.mock import patch, MagicMock
import sys, os

import db
import html_render

class TestWorkspaceAssignedProjectsRBAC(unittest.TestCase):
    def test_admin_user_assigned_projects_in_non_admin_workspace(self):
        """
        Verify that a user with admin/mixed role who is admin in WH
        but viewer in ENG can access their explicitly assigned projects in ENG,
        and that get_projects_for_user returns those projects instead of [].
        """
        with patch("db.get_admin_workspaces_for_user", return_value=["WH"]), \
             patch("db.get_user_projects", return_value=[
                 {"project_id": "P01", "is_dc": False},
                 {"project_id": "P02", "is_dc": False},
                 {"project_id": "P03", "is_dc": False},
                 {"project_id": "P04", "is_dc": False},
                 {"project_id": "P05", "is_dc": False},
                 {"project_id": "P06", "is_dc": False},
                 {"project_id": "P07", "is_dc": False},
             ]), \
             patch("db.get_projects") as mock_get_projs:

            mock_projs = [
                {"id": f"P0{i}", "code": f"C0{i}", "name": f"Project {i}", "workspace_id": "ENG"}
                for i in range(1, 8)
            ]
            mock_get_projs.return_value = mock_projs

            # Call get_projects_for_user for ENG with role="admin"
            res = db.get_projects_for_user("noura", role="admin", workspace_id="ENG")
            self.assertEqual(len(res), 7)
            self.assertEqual([p["id"] for p in res], ["P01", "P02", "P03", "P04", "P05", "P06", "P07"])

            # Verify get_projects was called with pids
            mock_get_projs.assert_called_with(
                ["P01", "P02", "P03", "P04", "P05", "P06", "P07"],
                workspace_id="ENG"
            )

    def test_launchpad_counts_rendered_correctly_for_assigned_projects(self):
        """
        Verify that render_launchpad correctly computes project count and records
        for noura in Engineering & Projects instead of 0 PROJECTS, 0 RECORDS.
        """
        mock_user = {
            "username": "noura",
            "role": "admin",
            "primary_workspace": "ENG"
        }
        all_workspaces = [
            {"id": "ENG", "code": "ENG", "name": "Engineering & Projects", "icon": "🏗️", "description": "EPC"},
            {"id": "WH", "code": "WH", "name": "Warehousing", "icon": "🏬", "description": "Stores"}
        ]
        user_workspaces = [
            {"id": "ENG", "code": "ENG", "user_role": "viewer"},
            {"id": "WH", "code": "WH", "user_role": "admin"}
        ]
        eng_projs = [
            {"id": f"P0{i}", "code": f"C0{i}", "name": f"Project {i}", "workspace_id": "ENG"}
            for i in range(1, 8)
        ]

        def fake_get_projects_for_user(uname, role, workspace_id=None):
            if workspace_id == "ENG":
                return eng_projs
            elif workspace_id == "WH":
                return [{"id": "WH01", "name": "Main Warehouse", "workspace_id": "WH"}]
            return []

        def fake_q(sql, params=(), one=False):
            if "FROM records GROUP BY project_id" in sql:
                # 5,545 records total across 7 projects
                return [
                    {"project_id": "P01", "c": 1000},
                    {"project_id": "P02", "c": 1000},
                    {"project_id": "P03", "c": 1000},
                    {"project_id": "P04", "c": 1000},
                    {"project_id": "P05", "c": 1000},
                    {"project_id": "P06", "c": 500},
                    {"project_id": "P07", "c": 45},
                ]
            return []

        with patch("db.get_workspaces", return_value=all_workspaces), \
             patch("db.get_user_workspaces", return_value=user_workspaces), \
             patch("db.get_projects_for_user", side_effect=fake_get_projects_for_user), \
             patch("db.q", side_effect=fake_q):

            html = html_render.render_launchpad(mock_user)
            # Verify 7 Projects is rendered in the card
            self.assertIn("7", html)
            self.assertIn("5,545", html)
            self.assertNotIn(">0</span><span class=\"hc-stat-lbl\">PROJECTS", html)

if __name__ == "__main__":
    unittest.main()
