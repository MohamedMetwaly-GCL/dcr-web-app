"""test_daily_digest_rbac.py - Validation test suite for Daily Digest RBAC & Document Visibility.

Validates:
1. Unauthenticated requests are rejected with 403 LOGIN_REQUIRED.
2. Non-admin users (editors, viewers) assigned to a project can view all issued, received, and replied documents for that project.
3. Non-admin users cannot access daily digest for projects they are not assigned to (403 Forbidden).
4. Calling /api/daily_digest/all for a non-admin only aggregates authorized projects.
5. Admins can view daily digest across all projects and document types.
"""
import os
import unittest
from unittest.mock import MagicMock, patch

os.environ["SECRET_KEY"] = "test-secret-key-1234567890"

from app import app
import db


class TestDailyDigestRBAC(unittest.TestCase):

    def setUp(self):
        self.client = app.test_client()

    def test_unauthenticated_rejected(self):
        """Unauthenticated calls must return 403 LOGIN_REQUIRED."""
        with patch("app.current_user", return_value=None):
            resp = self.client.get("/api/daily_digest/P01")
            self.assertEqual(resp.status_code, 403)
            data = resp.get_json()
            self.assertEqual(data.get("error"), "LOGIN_REQUIRED")

    def test_editor_assigned_project_sees_all_documents(self):
        """A non-admin editor assigned to P01 must see today's submitted documents."""
        mock_user = {"username": "sakr", "role": "editor"}
        mock_projects = [
            {"id": "P01", "name": "Eastmain MEP Buildings B06 & B10", "code": "EASTMAIN"}
        ]
        mock_dts = [
            {"id": 101, "project_id": "P01", "code": "IR", "name": "Inspection Request"},
            {"id": 102, "project_id": "P01", "code": "RFI", "name": "Request For Information"}
        ]
        mock_digest_return = {
            "received": [],
            "issued": [
                {
                    "id": "rec-1",
                    "project_id": "P01",
                    "project_name": "Eastmain MEP Buildings B06 & B10",
                    "dt_id": "101",
                    "docNo": "IR-EASTMAIN-FF-045 REV00",
                    "title": "Fire Fighting Inspection",
                    "status": "Under Review"
                }
            ],
            "replied": []
        }

        with patch("app.current_user", return_value=mock_user), \
             patch("app.can_view_project", return_value=True), \
             patch("db.get_projects", return_value=mock_projects), \
             patch("db.get_doc_types", return_value=mock_dts) as mock_get_dts, \
             patch("db.get_daily_digest", return_value=mock_digest_return) as mock_get_digest:

            resp = self.client.get("/api/daily_digest/P01?date=2026-10-17")
            self.assertEqual(resp.status_code, 200)
            data = resp.get_json()

            # Ensure doc types for P01 were queried
            mock_get_dts.assert_called_with("P01")
            # Ensure get_daily_digest was called with assigned doc types [101, 102]
            mock_get_digest.assert_called_with("P01", [101, 102], "2026-10-17")

            # Validate that the submitted document is present in the issued list
            self.assertEqual(len(data["issued"]), 1)
            self.assertEqual(data["issued"][0]["docNo"], "IR-EASTMAIN-FF-045 REV00")
            self.assertEqual(data["issued"][0]["project_name"], "Eastmain MEP Buildings B06 & B10")

    def test_editor_forbidden_for_unassigned_project(self):
        """A non-admin editor cannot access a project they are not assigned to."""
        mock_user = {"username": "sakr", "role": "editor"}

        with patch("app.current_user", return_value=mock_user), \
             patch("app.can_view_project", return_value=False):

            resp = self.client.get("/api/daily_digest/P99")
            self.assertEqual(resp.status_code, 403)
            data = resp.get_json()
            self.assertEqual(data.get("error"), "Forbidden")

    def test_editor_daily_digest_all_only_aggregates_assigned_projects(self):
        """When calling /api/daily_digest/all, only projects assigned to the editor are queried."""
        mock_user = {"username": "sakr", "role": "editor"}
        # User is allowed only P01, not P02
        allowed_pids = ["P01"]
        mock_dts_p01 = [{"id": 101, "project_id": "P01", "code": "IR", "name": "Inspection Request"}]
        mock_digest_p01 = {
            "received": [],
            "issued": [{"docNo": "IR-P01-001"}],
            "replied": []
        }

        with patch("app.current_user", return_value=mock_user), \
             patch("auth.get_allowed_project_ids", return_value=allowed_pids), \
             patch("db.get_doc_types", return_value=mock_dts_p01) as mock_get_dts, \
             patch("db.get_daily_digest", return_value=mock_digest_p01) as mock_get_digest:

            resp = self.client.get("/api/daily_digest/all")
            self.assertEqual(resp.status_code, 200)
            data = resp.get_json()

            # P01 should be queried
            mock_get_dts.assert_called_once_with("P01")
            mock_get_digest.assert_called_once_with("P01", [101], None)
            self.assertEqual(len(data["issued"]), 1)
            self.assertEqual(data["issued"][0]["docNo"], "IR-P01-001")

    def test_admin_daily_digest_all_aggregates_all_projects(self):
        """Admin calling /api/daily_digest/all queries all projects."""
        mock_admin = {"username": "admin", "role": "admin"}
        all_pids = ["P01", "P02"]
        mock_dts = [{"id": 1, "project_id": "P01", "code": "IR", "name": "IR"}]
        mock_digest = {"received": [], "issued": [{"docNo": "DOC-1"}], "replied": []}

        with patch("app.current_user", return_value=mock_admin), \
             patch("auth.get_allowed_project_ids", return_value=all_pids), \
             patch("db.get_doc_types", return_value=mock_dts), \
             patch("db.get_daily_digest", return_value=mock_digest):

            resp = self.client.get("/api/daily_digest/all")
            self.assertEqual(resp.status_code, 200)
            data = resp.get_json()
            # 2 projects queried, each returning 1 document -> total 2
            self.assertEqual(len(data["issued"]), 2)


if __name__ == "__main__":
    unittest.main()
