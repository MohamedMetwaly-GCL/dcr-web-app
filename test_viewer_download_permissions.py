import os
import unittest
from unittest.mock import patch, MagicMock

os.environ["SECRET_KEY"] = "test-secret-key-1234567890abcdef"
os.environ["ALLOW_INSECURE_COOKIES_FOR_LOCAL_DEV"] = "1"

import db
import blueprints.drive_viewer as dv
from app import app

class TestViewerDownloadPermissions(unittest.TestCase):
    def test_db_can_user_download_logic(self):
        # 1. Non-existent or empty user
        self.assertFalse(db.can_user_download(""))
        self.assertFalse(db.can_user_download(None))

        # 2. Superadmin / Admin / Editor roles always True
        with patch("db.get_user", return_value={"username": "super1", "role": "superadmin"}):
            self.assertTrue(db.can_user_download("super1"))

        with patch("db.get_user", return_value={"username": "admin1", "role": "admin"}):
            self.assertTrue(db.can_user_download("admin1"))

        with patch("db.get_user", return_value={"username": "editor1", "role": "editor"}):
            self.assertTrue(db.can_user_download("editor1"))

        # 3. Viewer user with can_download=False (default)
        with patch("db.get_user", return_value={"username": "viewer1", "role": "viewer", "can_download": False}):
            self.assertFalse(db.can_user_download("viewer1"))

        # 4. Viewer user with can_download=True
        with patch("db.get_user", return_value={"username": "viewer2", "role": "viewer", "can_download": True}):
            self.assertTrue(db.can_user_download("viewer2"))

    def test_drive_viewer_context_honors_download_flag(self):
        # Test that _context sets can_download=True for a viewer when can_user_download is True
        mock_rec = {
            "id": "REC-001",
            "_project_id": "PEM-058",
            "_dt_id": "DT-01",
            "docNo": "D-001",
            "attachments": {
                "submission": {"file_id": "drive_file_123", "name": "drawing.dwg", "url": "https://drive.google.com/file/d/drive_file_123/view"}
            }
        }
        
        with patch("blueprints.drive_viewer.current_user", return_value={"username": "eng_noura", "role": "viewer"}), \
             patch("db.get_record_by_id", return_value=mock_rec), \
             patch("blueprints.drive_viewer.can_view_project", return_value=True), \
             patch("blueprints.drive_viewer.is_superadmin", return_value=False), \
             patch("db.is_user_workspace_admin", return_value=False), \
             patch("db.get_user_project_doc_types", return_value=None), \
             patch("blueprints.drive_viewer.can_edit", return_value=False), \
             patch("blueprints.drive_viewer._is_workspace_admin", return_value=False), \
             patch("db.can_user_download", return_value=True):
            
            with app.test_request_context("/"):
                ctx, err = dv._context("REC-001", "submission")
                self.assertIsNone(err)
                self.assertIsNotNone(ctx)
                self.assertTrue(ctx["can_download"])
                self.assertFalse(ctx["can_open_drive"])

        # Test that _context sets can_download=False when can_user_download is False
        with patch("blueprints.drive_viewer.current_user", return_value={"username": "eng_noura", "role": "viewer"}), \
             patch("db.get_record_by_id", return_value=mock_rec), \
             patch("blueprints.drive_viewer.can_view_project", return_value=True), \
             patch("blueprints.drive_viewer.is_superadmin", return_value=False), \
             patch("db.is_user_workspace_admin", return_value=False), \
             patch("db.get_user_project_doc_types", return_value=None), \
             patch("blueprints.drive_viewer.can_edit", return_value=False), \
             patch("blueprints.drive_viewer._is_workspace_admin", return_value=False), \
             patch("db.can_user_download", return_value=False):
            
            with app.test_request_context("/"):
                ctx, err = dv._context("REC-001", "submission")
                self.assertIsNone(err)
                self.assertIsNotNone(ctx)
                self.assertFalse(ctx["can_download"])

    def test_api_toggle_download_permission(self):
        client = app.test_client()

        # 1. Unauthenticated request -> 401, 403 or redirect
        res = client.post("/api/users", json={"action": "toggle_download_permission", "username": "noura", "can_download": True})
        self.assertIn(res.status_code, (401, 403, 302))

        # 2. Authenticated as Superadmin via session cookie
        client.set_cookie("dcr_token", "test_admin_token")
        admin_user = {"username": "admin", "role": "superadmin"}
        with patch("db.get_session", return_value=admin_user), \
             patch("db.get_user", return_value={"username": "noura", "role": "viewer", "can_download": False}), \
             patch("db.set_user_can_download") as mock_set, \
             patch("db.log_action") as mock_log:
            
            res = client.post("/api/users", json={"action": "toggle_download_permission", "username": "noura", "can_download": True})
            self.assertEqual(res.status_code, 200)
            data = res.get_json()
            self.assertTrue(data.get("ok"))
            self.assertTrue(data.get("can_download"))
            mock_set.assert_called_once_with("noura", True)
            mock_log.assert_called_once()
            self.assertEqual(mock_log.call_args[0][1], "DOWNLOAD_PERM_UPDATE")

if __name__ == "__main__":
    unittest.main()
