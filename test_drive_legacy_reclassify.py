"""Regression tests for DRIVE-LEGACY-ROLE-FIX-01:
Reclassify legacy single-slot Drive links located inside Replied folders from submit to reply.

Offline unit tests only (zero network / live DB calls):
1. Legacy record with drive_link / fileLocation / un-typed submission pointing to a file inside
   Replied is promoted to approval (role='reply') and cleared from submit during sync/reclassification.
2. Record with a genuine file in Submit stays submission ('Subm') untouched.
3. Running sync / reclassification twice is idempotent (no duplicate Reply badges or extra writes).
4. Parent-folder inspection and POST /api/drive/reclassify-legacy/<pid> & /api/drive/sync/<pid>
   work safely and enforce permissions.
"""

import json
import os
import re
import unittest
from unittest.mock import patch

os.environ.setdefault("DATABASE_URL", "")
os.environ.setdefault("SECRET_KEY", "synthetic-test-secret-key-drive-legacy-role")
os.environ.setdefault("MAGIC_SECRET", "synthetic-test-magic-secret-drive-legacy-role")
os.environ.setdefault("GEMINI_API_KEY", "synthetic-test-gemini-key-drive-legacy-role")
os.environ.setdefault("ALLOW_INSECURE_COOKIES_FOR_LOCAL_DEV", "1")

import app as app_mod
import attachments
import db
import drive_service


class _FakeExec:
    def __init__(self, payload):
        self._payload = payload

    def execute(self):
        return self._payload


class _FakeDriveFiles:
    def __init__(self, items):
        self._items = list(items)
        self._by_id = {it["id"]: it for it in self._items}
        self.get_calls = []

    def list(self, q="", fields=None, pageSize=1000, pageToken=None, **kwargs):
        parent_ids = re.findall(r"'([^']+)'\s+in\s+parents", q or "")
        matched = [
            it for it in self._items
            if any(p in (it.get("parents") or []) for p in parent_ids) and not it.get("trashed", False)
        ]
        return _FakeExec({"files": matched})

    def get(self, fileId="", fields=None, **kwargs):
        self.get_calls.append(fileId)
        item = self._by_id.get(fileId)
        if item is None:
            raise RuntimeError(f"404 not found: {fileId}")
        return _FakeExec(dict(item))


class _FakeDriveService:
    def __init__(self, items):
        self._files = _FakeDriveFiles(items)

    def files(self):
        return self._files


class TestLegacyReclassifyPureAttachments(unittest.TestCase):
    def test_01_promote_legacy_drive_link_and_file_location_to_reply(self):
        rec = {
            "docNo": "DS-058",
            "status": "A - Approved",
            "fileLocation": "https://drive.google.com/file/d/REP_FILE_58/view",
            "drive_link": "https://drive.google.com/file/d/REP_FILE_58/view",
        }
        promoted, changed = attachments.promote_legacy_link_to_reply(
            rec,
            matched_item={
                "id": "REP_FILE_58",
                "name": "DS-058_Replied.pdf",
                "mimeType": "application/pdf",
                "webViewLink": "https://drive.google.com/file/d/REP_FILE_58/view",
            },
        )
        self.assertTrue(changed)
        self.assertNotIn("fileLocation", promoted)
        self.assertNotIn("drive_link", promoted)
        self.assertNotIn("submission", promoted.get("attachments", {}))
        app_entry = promoted["attachments"]["approval"]
        self.assertEqual(app_entry["role"], "reply")
        self.assertTrue(app_entry["slot_verified"])
        self.assertEqual(app_entry["file_id"], "REP_FILE_58")
        self.assertEqual(app_entry["url"], "https://drive.google.com/file/d/REP_FILE_58/view")

        norm = attachments.normalize_attachments(promoted)
        self.assertIsNone(norm["submission"])
        self.assertIsNotNone(norm["approval"])
        self.assertEqual(norm["approval"]["file_id"], "REP_FILE_58")

        # Idempotency check
        promoted2, changed2 = attachments.promote_legacy_link_to_reply(
            promoted,
            matched_item={
                "id": "REP_FILE_58",
                "name": "DS-058_Replied.pdf",
                "mimeType": "application/pdf",
                "webViewLink": "https://drive.google.com/file/d/REP_FILE_58/view",
            },
        )
        self.assertFalse(changed2)
        self.assertEqual(promoted2, promoted)

    def test_02_merge_sync_links_clears_duplicate_legacy_and_submission_when_approval_only(self):
        rec = {
            "docNo": "DS-059",
            "fileLocation": "https://drive.google.com/file/d/REP_FILE_59/view",
            "driveLink": "https://drive.google.com/file/d/REP_FILE_59/view",
            "attachments": {
                "submission": {
                    "url": "https://drive.google.com/file/d/REP_FILE_59/view",
                    "legacy": True,
                }
            },
        }
        merged = attachments.merge_sync_links(
            rec,
            {"approval": "https://drive.google.com/file/d/REP_FILE_59/view"},
        )
        self.assertNotIn("fileLocation", merged)
        self.assertNotIn("driveLink", merged)
        self.assertNotIn("submission", merged["attachments"])
        self.assertEqual(merged["attachments"]["approval"]["role"], "reply")
        self.assertEqual(merged["attachments"]["approval"]["file_id"], "REP_FILE_59")

    def test_03_merge_sync_links_preserves_distinct_submission_link(self):
        rec = {
            "docNo": "DS-060",
            "fileLocation": "https://drive.google.com/file/d/GENUINE_SUB_60/view",
            "attachments": {
                "submission": {
                    "url": "https://drive.google.com/file/d/GENUINE_SUB_60/view",
                    "source": "drive_sync",
                }
            },
        }
        merged = attachments.merge_sync_links(
            rec,
            {"approval": "https://drive.google.com/file/d/REP_FILE_60/view"},
        )
        self.assertEqual(merged["fileLocation"], "https://drive.google.com/file/d/GENUINE_SUB_60/view")
        self.assertEqual(merged["attachments"]["submission"]["url"], "https://drive.google.com/file/d/GENUINE_SUB_60/view")
        self.assertEqual(merged["attachments"]["approval"]["url"], "https://drive.google.com/file/d/REP_FILE_60/view")


class TestDriveSyncLegacyReclassification(unittest.TestCase):
    def _build_env(self):
        projects_rows = [
            {"id": "PEM-058", "data": {"drive_folder_id": "ROOT_CFC"}},
        ]
        records_store = {
            # Legacy record whose file is in Replied/ (filename matches docNo)
            "REC-REP-1": {
                "id": "REC-REP-1",
                "project_id": "PEM-058",
                "data": {
                    "docNo": "DS-001",
                    "status": "A - Approved",
                    "fileLocation": "https://drive.google.com/file/d/FILE_REP_1/view",
                    "drive_link": "https://drive.google.com/file/d/FILE_REP_1/view",
                },
            },
            # Legacy record whose file is in Replied/ but filename does NOT match docNo (un-typed submission)
            "REC-REP-2": {
                "id": "REC-REP-2",
                "project_id": "PEM-058",
                "data": {
                    "docNo": "DS-002",
                    "status": "B - Approved As Noted",
                    "fileLocation": "https://drive.google.com/file/d/FILE_REP_2/view",
                    "attachments": {
                        "submission": {
                            "url": "https://drive.google.com/file/d/FILE_REP_2/view",
                            "legacy": True,
                        }
                    },
                },
            },
            # Genuine Submit record: must stay in submission untouched
            "REC-SUB-3": {
                "id": "REC-SUB-3",
                "project_id": "PEM-058",
                "data": {
                    "docNo": "DS-003",
                    "status": "Under Review",
                    "fileLocation": "https://drive.google.com/file/d/FILE_SUB_3/view",
                },
            },
            # Legacy link outside scanned tree whose parent folder on Drive is named 'REPLIED'
            "REC-EXT-4": {
                "id": "REC-EXT-4",
                "project_id": "PEM-058",
                "data": {
                    "docNo": "DS-004",
                    "status": "A - Approved",
                    "drive_link": "https://drive.google.com/file/d/FILE_EXT_REP_4/view",
                },
            },
        }
        drive_items = [
            {
                "id": "F_DS",
                "name": "2. Document Submittal (DS)",
                "mimeType": "application/vnd.google-apps.folder",
                "parents": ["ROOT_CFC"],
                "webViewLink": "https://drive.google.com/drive/folders/F_DS",
            },
            {
                "id": "F_SUB",
                "name": "Submit",
                "mimeType": "application/vnd.google-apps.folder",
                "parents": ["F_DS"],
                "webViewLink": "https://drive.google.com/drive/folders/F_SUB",
            },
            {
                "id": "F_REP",
                "name": "Replied",
                "mimeType": "application/vnd.google-apps.folder",
                "parents": ["F_DS"],
                "webViewLink": "https://drive.google.com/drive/folders/F_REP",
            },
            {
                "id": "FILE_REP_1",
                "name": "DS-001_Stamped.pdf",
                "mimeType": "application/pdf",
                "parents": ["F_REP"],
                "webViewLink": "https://drive.google.com/file/d/FILE_REP_1/view",
            },
            {
                "id": "FILE_REP_2",
                "name": "Consultant_Comment_Sheet_Final.pdf",
                "mimeType": "application/pdf",
                "parents": ["F_REP"],
                "webViewLink": "https://drive.google.com/file/d/FILE_REP_2/view",
            },
            {
                "id": "FILE_SUB_3",
                "name": "DS-003_For_Approval.pdf",
                "mimeType": "application/pdf",
                "parents": ["F_SUB"],
                "webViewLink": "https://drive.google.com/file/d/FILE_SUB_3/view",
            },
            # External Replied folder (not under ROOT_CFC, discovered via parent inspection)
            {
                "id": "F_EXT_REP",
                "name": "REPLIED",
                "mimeType": "application/vnd.google-apps.folder",
                "parents": ["F_OLD_ARCHIVE"],
                "webViewLink": "https://drive.google.com/drive/folders/F_EXT_REP",
            },
            {
                "id": "FILE_EXT_REP_4",
                "name": "Legacy_Reply_004.pdf",
                "mimeType": "application/pdf",
                "parents": ["F_EXT_REP"],
                "webViewLink": "https://drive.google.com/file/d/FILE_EXT_REP_4/view",
            },
        ]
        return projects_rows, records_store, drive_items

    def test_04_process_drive_folder_reclassifies_replied_preserves_submit_and_is_idempotent(self):
        projects_rows, records_store, drive_items = self._build_env()
        exe_calls = []

        def fake_q(sql, params=(), one=False):
            s = " ".join(sql.split())
            if "FROM projects" in s:
                return projects_rows
            if "FROM records" in s:
                if "docNo" in s:
                    pid_arg, doc_arg = params[0], params[1]
                    matched = [
                        {"id": r["id"], "data": r["data"]}
                        for r in records_store.values()
                        if r["project_id"] == pid_arg and r["data"].get("docNo") == doc_arg
                    ]
                    return (matched[0] if matched else None) if one else matched
                pid_arg = params[0]
                matched = [
                    {"id": r["id"], "data": r["data"]}
                    for r in records_store.values()
                    if r["project_id"] == pid_arg
                ]
                return (matched[0] if matched else None) if one else matched
            return None if one else []

        def fake_exe(sql, params=()):
            new_data_json, rec_id, pid_arg = params
            exe_calls.append((rec_id, pid_arg))
            if rec_id in records_store and records_store[rec_id]["project_id"] == pid_arg:
                records_store[rec_id]["data"] = json.loads(new_data_json)
            return 1

        svc = _FakeDriveService(drive_items)
        with patch("db.q", side_effect=fake_q), \
             patch("db.exe", side_effect=fake_exe), \
             patch.object(drive_service, "get_drive_service", return_value=svc):
            drive_service.process_drive_folder("ROOT_CFC", project_id="PEM-058")

        # 1) REC-REP-1 promoted to approval (role='reply') and cleared from fileLocation/drive_link/submission
        d1 = records_store["REC-REP-1"]["data"]
        self.assertNotIn("fileLocation", d1)
        self.assertNotIn("drive_link", d1)
        self.assertNotIn("submission", d1.get("attachments", {}))
        self.assertEqual(d1["attachments"]["approval"]["role"], "reply")
        self.assertEqual(d1["attachments"]["approval"]["file_id"], "FILE_REP_1")

        # 2) REC-REP-2 (filename didn't match docNo) promoted via BFS replied_items_by_id match
        d2 = records_store["REC-REP-2"]["data"]
        self.assertNotIn("fileLocation", d2)
        self.assertNotIn("submission", d2.get("attachments", {}))
        self.assertEqual(d2["attachments"]["approval"]["role"], "reply")
        self.assertEqual(d2["attachments"]["approval"]["file_id"], "FILE_REP_2")

        # 3) REC-SUB-3 (genuine Submit file) stays in submission untouched
        d3 = records_store["REC-SUB-3"]["data"]
        self.assertEqual(d3["attachments"]["submission"]["url"], "https://drive.google.com/file/d/FILE_SUB_3/view")
        self.assertNotIn("approval", d3["attachments"])

        # 4) REC-EXT-4 (outside scanned tree, parent named 'REPLIED') promoted via parent inspection
        d4 = records_store["REC-EXT-4"]["data"]
        self.assertNotIn("drive_link", d4)
        self.assertNotIn("submission", d4.get("attachments", {}))
        self.assertEqual(d4["attachments"]["approval"]["role"], "reply")
        self.assertEqual(d4["attachments"]["approval"]["file_id"], "FILE_EXT_REP_4")

        # 5) Idempotency: running reclassification a second time makes 0 changes and 0 DB updates
        snap_after_first = json.dumps(records_store, sort_keys=True)
        exe_calls.clear()
        with patch("db.q", side_effect=fake_q), \
             patch("db.exe", side_effect=fake_exe), \
             patch.object(drive_service, "get_drive_service", return_value=svc):
            summary2 = drive_service.reclassify_project_legacy_reply_links(
                "PEM-058", service=svc, inspect_parents=True
            )
        self.assertEqual(summary2["reclassified"], 0)
        self.assertEqual(exe_calls, [])
        self.assertEqual(json.dumps(records_store, sort_keys=True), snap_after_first)

        # Running full process_drive_folder a second time also keeps exact same normalized attachments
        with patch("db.q", side_effect=fake_q), \
             patch("db.exe", side_effect=fake_exe), \
             patch.object(drive_service, "get_drive_service", return_value=svc):
            drive_service.process_drive_folder("ROOT_CFC", project_id="PEM-058")
        for rid in ("REC-REP-1", "REC-REP-2", "REC-EXT-4"):
            norm = attachments.normalize_attachments(records_store[rid]["data"])
            self.assertIsNone(norm["submission"])
            self.assertIsNotNone(norm["approval"])
            self.assertEqual(norm["approval"]["role"], "reply")
        norm3 = attachments.normalize_attachments(records_store["REC-SUB-3"]["data"])
        self.assertIsNotNone(norm3["submission"])
        self.assertIsNone(norm3["approval"])

    def test_05_api_endpoints_reclassify_legacy_and_sync_option(self):
        client = app_mod.app.test_client()
        client.set_cookie("dcr_token", "tok_editor")
        editor = {"username": "dc_editor", "role": "editor"}

        # Forbidden when can_edit is False
        with patch("db.get_session", return_value=editor), \
             patch("app.can_view_project", return_value=True), \
             patch("app.can_edit", return_value=False):
            r_forb = client.post("/api/drive/reclassify-legacy/PEM-058")
            self.assertEqual(r_forb.status_code, 403)

        # Allowed when can_edit is True
        fake_summary = {"scanned": 10, "reclassified": 3, "skipped": 7, "errors": 0, "record_ids": ["R1", "R2", "R3"]}
        with patch("db.get_session", return_value=editor), \
             patch("app.can_view_project", return_value=True), \
             patch("app.can_edit", return_value=True), \
             patch("drive_service.reclassify_project_legacy_reply_links", return_value=fake_summary) as mock_rec:
            r_ok = client.post("/api/drive/reclassify-legacy/PEM-058")
            self.assertEqual(r_ok.status_code, 200)
            body = r_ok.get_json()
            self.assertTrue(body["ok"])
            self.assertEqual(body["reclassified"], 3)
            mock_rec.assert_called_once_with("PEM-058", inspect_parents=True)

        # /api/drive/sync/<pid> defaults reclassify_legacy_links=True and passes it to process_drive_folder
        started_kwargs = []

        class _FakeThread:
            def __init__(self, target=None, args=(), kwargs=None):
                started_kwargs.append((target, args, kwargs or {}))

            def start(self):
                pass

        with patch("db.get_session", return_value=editor), \
             patch("app.can_view_project", return_value=True), \
             patch("app.can_edit", return_value=True), \
             patch("db.get_project", return_value={"id": "PEM-058", "drive_folder_id": "ROOT_CFC"}), \
             patch("db.update_project_settings"), \
             patch("drive_routing.provision_project_drive_structure"), \
             patch("threading.Thread", side_effect=_FakeThread):
            r_sync = client.post("/api/drive/sync/PEM-058", json={"folder_id": "ROOT_CFC"})
            self.assertEqual(r_sync.status_code, 200)
            self.assertTrue(r_sync.get_json().get("reclassify_legacy_links"))
            self.assertEqual(started_kwargs[-1][2].get("reclassify_legacy_links"), True)

            r_sync_off = client.post(
                "/api/drive/sync/PEM-058",
                json={"folder_id": "ROOT_CFC", "reclassify_legacy_links": False},
            )
            self.assertEqual(r_sync_off.status_code, 200)
            self.assertFalse(r_sync_off.get_json().get("reclassify_legacy_links"))
            self.assertEqual(started_kwargs[-1][2].get("reclassify_legacy_links"), False)


if __name__ == "__main__":
    unittest.main()
