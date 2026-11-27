"""Offline unit tests for BULK-SUBTAB-ASSIGN-01 (Bulk Assign Sub-Tab action)."""
import os
os.environ["DATABASE_URL"] = ""
os.environ["SECRET_KEY"] = "synthetic-test-secret-key-123456789012345678901234"
os.environ["MAGIC_SECRET"] = "synthetic-test-magic-secret-only"
os.environ["GEMINI_API_KEY"] = "synthetic-test-gemini-key-only"

import copy
import unittest
from contextlib import contextmanager
from unittest.mock import patch

import db
import i18n
from app import app


class _InMemoryRecordsCursor:
    """Simulates PostgreSQL JSONB atomic UPDATE on records(id, project_id, dt_id, data)."""

    def __init__(self, rows):
        self.rows = rows
        self.rowcount = 0
        self.executed = []

    def execute(self, sql, params=None):
        self.executed.append((sql, params))
        sql_norm = " ".join(sql.split()).lower()
        assert "where project_id = %s and dt_id = %s and id = any(%s)" in sql_norm, (
            f"Expected project+dt_id+id scoped WHERE clause, got: {sql}"
        )
        matched = 0
        if "jsonb_build_object(%s, to_jsonb(%s::text))" in sql_norm:
            col_key, val, pid, dt_id, id_list = params
            id_set = set(id_list)
            for row in self.rows:
                if row["project_id"] == pid and row["dt_id"] == dt_id and row["id"] in id_set:
                    row["data"] = dict(row.get("data") or {})
                    row["data"][col_key] = val
                    matched += 1
        elif "coalesce(data, '{}'::jsonb) - %s" in sql_norm:
            col_key, pid, dt_id, id_list = params
            id_set = set(id_list)
            for row in self.rows:
                if row["project_id"] == pid and row["dt_id"] == dt_id and row["id"] in id_set:
                    row["data"] = dict(row.get("data") or {})
                    row["data"].pop(col_key, None)
                    matched += 1
        else:
            raise AssertionError(f"Unexpected SQL: {sql}")
        self.rowcount = matched


class TestBulkSubtabAssign(unittest.TestCase):
    def setUp(self):
        app.config["TESTING"] = True
        self.client = app.test_client()
        self.editor_user = {"username": "editor1", "role": "user"}
        self.viewer_user = {"username": "viewer1", "role": "viewer"}

    def test_requires_login(self):
        with (
            patch("app.current_user", return_value=None),
            patch("auth.current_user", return_value=None),
        ):
            resp = self.client.post(
                "/api/records/P1/PTW/bulk-subtab",
                json={"record_ids": ["r1"], "subtab": "Hot Work Permit"},
            )
        self.assertEqual(resp.status_code, 403)
        self.assertEqual(resp.get_json().get("error"), "LOGIN_REQUIRED")

    def test_readonly_user_rejected_403(self):
        with (
            patch("app.current_user", return_value=self.viewer_user),
            patch("app.can_view_project", return_value=True),
            patch("auth.current_user", return_value=self.viewer_user),
            patch("blueprints.records.current_user", return_value=self.viewer_user),
            patch("blueprints.records._resolve_project", return_value={"id": "P1"}),
            patch("blueprints.records.can_edit", return_value=False),
            patch("blueprints.records.db.log_action") as mock_log,
        ):
            resp = self.client.post(
                "/api/records/P1/PTW/bulk-subtab",
                json={"record_ids": ["r1", "r2"], "subtab": "Hot Work Permit"},
            )
        self.assertEqual(resp.status_code, 403)
        self.assertEqual(resp.get_json().get("error"), "Forbidden")
        mock_log.assert_called_once()

    def test_doctype_access_denied_403(self):
        with (
            patch("app.current_user", return_value=self.editor_user),
            patch("app.can_view_project", return_value=True),
            patch("auth.current_user", return_value=self.editor_user),
            patch("blueprints.records.current_user", return_value=self.editor_user),
            patch("blueprints.records._resolve_project", return_value={"id": "P1"}),
            patch("blueprints.records.can_edit", return_value=True),
            patch("blueprints.records.can_access_doc_type", return_value=False),
        ):
            resp = self.client.post(
                "/api/records/P1/PTW/bulk-subtab",
                json={"record_ids": ["r1"], "subtab": "Hot Work Permit"},
            )
        self.assertEqual(resp.status_code, 403)
        self.assertIn("access denied", resp.get_json().get("error", "").lower())

    def test_invalid_subtab_rejected_400(self):
        sub_cfg = {
            "col_key": "legacySubtype",
            "list_name": "hse_ptw_subtype",
            "items": ["Hot Work Permit", "Cold Work Permit"],
            "is_ptw": True,
        }
        with (
            patch("app.current_user", return_value=self.editor_user),
            patch("app.can_view_project", return_value=True),
            patch("auth.current_user", return_value=self.editor_user),
            patch("blueprints.records.current_user", return_value=self.editor_user),
            patch("blueprints.records._resolve_project", return_value={"id": "P1"}),
            patch("blueprints.records.can_edit", return_value=True),
            patch("blueprints.records.can_access_doc_type", return_value=True),
            patch("blueprints.records.db.get_doc_type", return_value={"id": "PTW", "code": "PTW"}),
            patch("blueprints.records.db.get_doc_type_subtype_config", return_value=sub_cfg),
            patch("blueprints.records.db.bulk_assign_records_subtab") as mock_bulk,
        ):
            # Unknown sub-tab
            r1 = self.client.post(
                "/api/records/P1/PTW/bulk-subtab",
                json={"record_ids": ["rec-1"], "subtab": "Nonexistent Subtab"},
            )
            self.assertEqual(r1.status_code, 400)
            self.assertFalse(r1.get_json().get("ok"))

            # Empty sub-tab without clear=True
            r2 = self.client.post(
                "/api/records/P1/PTW/bulk-subtab",
                json={"record_ids": ["rec-1"], "subtab": ""},
            )
            self.assertEqual(r2.status_code, 400)

            # Empty record_ids
            r3 = self.client.post(
                "/api/records/P1/PTW/bulk-subtab",
                json={"record_ids": [], "subtab": "Hot Work Permit"},
            )
            self.assertEqual(r3.status_code, 400)

            mock_bulk.assert_not_called()

    def test_doctype_without_subtabs_rejected_400(self):
        with (
            patch("app.current_user", return_value=self.editor_user),
            patch("app.can_view_project", return_value=True),
            patch("auth.current_user", return_value=self.editor_user),
            patch("blueprints.records.current_user", return_value=self.editor_user),
            patch("blueprints.records._resolve_project", return_value={"id": "P1"}),
            patch("blueprints.records.can_edit", return_value=True),
            patch("blueprints.records.can_access_doc_type", return_value=True),
            patch("blueprints.records.db.get_doc_type", return_value={"id": "LTR", "code": "LTR"}),
            patch("blueprints.records.db.get_doc_type_subtype_config", return_value=None),
        ):
            resp = self.client.post(
                "/api/records/P1/LTR/bulk-subtab",
                json={"record_ids": ["rec-1"], "subtab": "Hot Work Permit"},
            )
        self.assertEqual(resp.status_code, 400)
        self.assertIn("no sub-tabs", resp.get_json().get("error", "").lower())

    def test_cross_project_and_cross_doctype_isolation_and_field_preservation(self):
        initial_rows = [
            {
                "id": "r1",
                "project_id": "P1",
                "dt_id": "PTW",
                "data": {
                    "docNo": "PEM-HWP-001",
                    "legacySubtype": "Cold Work Permit",
                    "status": "Open",
                    "issuedDate": "2026-01-10",
                    "fileLocation": "https://drive.google.com/file/d/abc123/view",
                    "drive_link": "https://drive.google.com/drive/folders/xyz",
                    "attachments": [
                        {"slot": "submit", "name": "permit.pdf", "file_id": "abc123"}
                    ],
                },
            },
            {
                "id": "r2",
                "project_id": "P1",
                "dt_id": "PTW",
                "data": {
                    "docNo": "PEM-HWP-002",
                    "status": "Closed",
                    "fileLocation": "https://drive.google.com/file/d/def456/view",
                    "attachments": [
                        {"slot": "replied", "name": "signed.pdf", "file_id": "def456"}
                    ],
                },
            },
            {
                "id": "r_other_proj",
                "project_id": "P2",
                "dt_id": "PTW",
                "data": {
                    "docNo": "OTHER-001",
                    "legacySubtype": "Excavation Permit",
                    "fileLocation": "https://drive.google.com/file/d/p2/view",
                },
            },
            {
                "id": "r_other_dt",
                "project_id": "P1",
                "dt_id": "OBS",
                "data": {
                    "docNo": "OBS-001",
                    "legacySubtype": "Unsafe Act",
                    "attachments": [{"name": "obs.jpg"}],
                },
            },
        ]
        snapshot_other_proj = copy.deepcopy(initial_rows[2]["data"])
        snapshot_other_dt = copy.deepcopy(initial_rows[3]["data"])
        fake_cursor = _InMemoryRecordsCursor(initial_rows)

        @contextmanager
        def _fake_db():
            yield fake_cursor

        sub_cfg = {
            "col_key": "legacySubtype",
            "list_name": "hse_ptw_subtype",
            "items": ["Hot Work Permit", "Cold Work Permit"],
            "is_ptw": True,
        }
        with (
            patch("app.current_user", return_value=self.editor_user),
            patch("app.can_view_project", return_value=True),
            patch("auth.current_user", return_value=self.editor_user),
            patch("blueprints.records.current_user", return_value=self.editor_user),
            patch("blueprints.records._resolve_project", return_value={"id": "P1"}),
            patch("blueprints.records.can_edit", return_value=True),
            patch("blueprints.records.can_access_doc_type", return_value=True),
            patch("blueprints.records.db.get_doc_type", return_value={"id": "PTW", "code": "PTW"}),
            patch("blueprints.records.db.get_doc_type_subtype_config", return_value=sub_cfg),
            patch("db.DB", _fake_db),
            patch("db.invalidate_dashboard_cache") as mock_inv,
            patch("blueprints.records.db.log_action") as mock_log,
        ):
            # Pass IDs from P1/PTW plus IDs from P2/PTW and P1/OBS; also test case-insensitive canonicalization
            resp = self.client.post(
                "/api/records/P1/PTW/bulk-subtab",
                json={
                    "record_ids": ["r1", "r2", "r_other_proj", "r_other_dt", "r1"],
                    "subtab": "hot work permit",
                },
            )
            self.assertEqual(resp.status_code, 200)
            body = resp.get_json()
            self.assertTrue(body["ok"])
            self.assertEqual(body["updated"], 2)
            self.assertEqual(body["subtab"], "Hot Work Permit")
            self.assertEqual(body["field"], "legacySubtype")

            # Verify P1/PTW rows updated ONLY legacySubtype and preserved attachments, fileLocation, drive_link, status
            self.assertEqual(initial_rows[0]["data"]["legacySubtype"], "Hot Work Permit")
            self.assertEqual(initial_rows[0]["data"]["fileLocation"], "https://drive.google.com/file/d/abc123/view")
            self.assertEqual(initial_rows[0]["data"]["drive_link"], "https://drive.google.com/drive/folders/xyz")
            self.assertEqual(
                initial_rows[0]["data"]["attachments"],
                [{"slot": "submit", "name": "permit.pdf", "file_id": "abc123"}],
            )
            self.assertEqual(initial_rows[0]["data"]["status"], "Open")
            self.assertEqual(initial_rows[0]["data"]["issuedDate"], "2026-01-10")

            self.assertEqual(initial_rows[1]["data"]["legacySubtype"], "Hot Work Permit")
            self.assertEqual(initial_rows[1]["data"]["fileLocation"], "https://drive.google.com/file/d/def456/view")
            self.assertEqual(
                initial_rows[1]["data"]["attachments"],
                [{"slot": "replied", "name": "signed.pdf", "file_id": "def456"}],
            )

            # Verify cross-project and cross-doctype rows are untouched
            self.assertEqual(initial_rows[2]["data"], snapshot_other_proj)
            self.assertEqual(initial_rows[3]["data"], snapshot_other_dt)

            mock_inv.assert_called_once_with("P1")
            mock_log.assert_called_once()

            # Now test explicit clear mode preserves attachments & fileLocation
            resp_clear = self.client.post(
                "/api/records/P1/PTW/bulk-subtab",
                json={"record_ids": ["r1"], "subtab": "", "clear": True},
            )
            self.assertEqual(resp_clear.status_code, 200)
            self.assertEqual(resp_clear.get_json()["updated"], 1)
            self.assertNotIn("legacySubtype", initial_rows[0]["data"])
            self.assertEqual(initial_rows[0]["data"]["fileLocation"], "https://drive.google.com/file/d/abc123/view")
            self.assertEqual(
                initial_rows[0]["data"]["attachments"],
                [{"slot": "submit", "name": "permit.pdf", "file_id": "abc123"}],
            )

    def test_i18n_keys_parity_for_bulk_subtab(self):
        required_keys = [
            "reg.bulk_assign_subtab",
            "reg.bulk_subtab_modal_title",
            "reg.bulk_subtab_summary",
            "reg.bulk_subtab_label",
            "reg.bulk_subtab_placeholder",
            "reg.bulk_subtab_clear_opt",
            "reg.bulk_subtab_apply",
            "reg.bulk_subtab_select_req",
            "reg.bulk_subtab_confirm",
            "reg.bulk_subtab_confirm_clear",
            "reg.bulk_subtab_assigned",
            "reg.bulk_subtab_cleared",
            "reg.bulk_subtab_err",
        ]
        en_dict = i18n.TRANSLATIONS["en"]
        ar_dict = i18n.TRANSLATIONS["ar"]
        for k in required_keys:
            self.assertIn(k, en_dict, f"Missing EN i18n key: {k}")
            self.assertIn(k, ar_dict, f"Missing AR i18n key: {k}")
            self.assertTrue(en_dict[k].strip())
            self.assertTrue(ar_dict[k].strip())
            for eastern_digit in "٠١٢٣٤٥٦٧٨٩":
                self.assertNotIn(eastern_digit, ar_dict[k])

    def test_frontend_bulk_subtab_markup_and_handlers_present(self):
        with open(os.path.join(os.path.dirname(__file__), "html_render.py"), encoding="utf-8") as f:
            src = f.read()
        for token in (
            'id="btn-bulk-subtab"',
            'id="bulk-subtab-modal"',
            'id="bulk-subtab-select"',
            "function openBulkSubtabModal()",
            "async function submitBulkSubtab()",
            "/bulk-subtab",
        ):
            self.assertIn(token, src, f"Expected token missing in html_render.py: {token}")


if __name__ == "__main__":
    unittest.main()
