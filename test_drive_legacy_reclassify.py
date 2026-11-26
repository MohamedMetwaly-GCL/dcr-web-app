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


class TestDriveLegacyRoleFix02(unittest.TestCase):
    """Regression tests for DRIVE-LEGACY-ROLE-FIX-02:
    1. Stale legacy file_id (not in Submit) whose docNo matches a PDF in Replied is promoted to Reply and clears Subm.
    2. Verified Submit upload (slot_verified=True) or file inside Submit is NEVER overwritten by Replied.
    3. Status-less doc types (like EI) render neutral File (drive.badge_file) instead of Subm, while status-enabled
       doc types (like DS) keep Subm.
    """

    def test_06_stale_legacy_file_id_matched_in_replied_promotes_to_reply_and_clears_subm(self):
        projects_rows = [{"id": "PEM-058", "data": {"drive_folder_id": "ROOT_CFC"}}]
        records_store = {
            "REC-STALE-260": {
                "id": "REC-STALE-260",
                "project_id": "PEM-058",
                "data": {
                    "docNo": "DS-CY002P608-00260 REV00",
                    "status": "B - Approved As Noted",
                    "fileLocation": "https://drive.google.com/file/d/OLD_STALE_COPY_260/view",
                    "attachments": {
                        "submission": {
                            "url": "https://drive.google.com/file/d/OLD_STALE_COPY_260/view",
                            "legacy": True,
                        }
                    },
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
                "id": "NEW_REPLIED_260",
                "name": "DS-CY002P608-00260 REV00.pdf",
                "mimeType": "application/pdf",
                "parents": ["F_REP"],
                "webViewLink": "https://drive.google.com/file/d/NEW_REPLIED_260/view",
            },
        ]

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
            if rec_id in records_store and records_store[rec_id]["project_id"] == pid_arg:
                records_store[rec_id]["data"] = json.loads(new_data_json)
            return 1

        svc = _FakeDriveService(drive_items)
        with patch("db.q", side_effect=fake_q), \
             patch("db.exe", side_effect=fake_exe), \
             patch.object(drive_service, "get_drive_service", return_value=svc):
            drive_service.process_drive_folder("ROOT_CFC", project_id="PEM-058")

        d = records_store["REC-STALE-260"]["data"]
        self.assertNotIn("fileLocation", d)
        self.assertNotIn("submission", d.get("attachments", {}))
        app_entry = d["attachments"]["approval"]
        self.assertEqual(app_entry["role"], "reply")
        self.assertTrue(app_entry["slot_verified"])
        self.assertEqual(app_entry["file_id"], "NEW_REPLIED_260")
        self.assertEqual(app_entry["url"], "https://drive.google.com/file/d/NEW_REPLIED_260/view")
        self.assertEqual(app_entry["file_name"], "DS-CY002P608-00260 REV00.pdf")

        norm = attachments.normalize_attachments(d)
        self.assertIsNone(norm["submission"])
        self.assertIsNotNone(norm["approval"])
        self.assertEqual(norm["approval"]["file_id"], "NEW_REPLIED_260")

    def test_07_verified_submit_or_file_in_submit_never_overwritten_by_replied(self):
        projects_rows = [{"id": "PEM-058", "data": {"drive_folder_id": "ROOT_CFC"}}]
        records_store = {
            # Case A: slot_verified=True DCR Submit upload must NEVER be cleared even when Replied has same docNo
            "REC-VERIFIED-SUB": {
                "id": "REC-VERIFIED-SUB",
                "project_id": "PEM-058",
                "data": {
                    "docNo": "DS-CY002P608-00261 REV00",
                    "status": "A - Approved",
                    "attachments": {
                        "submission": {
                            "file_id": "DCR_UPLOADED_SUB_261",
                            "url": "https://drive.google.com/file/d/DCR_UPLOADED_SUB_261/view",
                            "slot_verified": True,
                            "uploaded_by": "dc_user",
                            "uploaded_at": "2026-11-26T10:00:00Z",
                        }
                    },
                },
            },
            # Case B: File confirmed inside Submit folder must NEVER be overwritten by Replied
            "REC-IN-SUBMIT-FOLDER": {
                "id": "REC-IN-SUBMIT-FOLDER",
                "project_id": "PEM-058",
                "data": {
                    "docNo": "DS-CY002P608-00262 REV00",
                    "status": "B - Approved As Noted",
                    "fileLocation": "https://drive.google.com/file/d/GENUINE_SUB_262/view",
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
                "id": "REP_261",
                "name": "DS-CY002P608-00261 REV00.pdf",
                "mimeType": "application/pdf",
                "parents": ["F_REP"],
                "webViewLink": "https://drive.google.com/file/d/REP_261/view",
            },
            {
                "id": "GENUINE_SUB_262",
                "name": "DS-CY002P608-00262 REV00.pdf",
                "mimeType": "application/pdf",
                "parents": ["F_SUB"],
                "webViewLink": "https://drive.google.com/file/d/GENUINE_SUB_262/view",
            },
            {
                "id": "REP_262",
                "name": "DS-CY002P608-00262 REV00.pdf",
                "mimeType": "application/pdf",
                "parents": ["F_REP"],
                "webViewLink": "https://drive.google.com/file/d/REP_262/view",
            },
        ]

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
            if rec_id in records_store and records_store[rec_id]["project_id"] == pid_arg:
                records_store[rec_id]["data"] = json.loads(new_data_json)
            return 1

        svc = _FakeDriveService(drive_items)
        with patch("db.q", side_effect=fake_q), \
             patch("db.exe", side_effect=fake_exe), \
             patch.object(drive_service, "get_drive_service", return_value=svc):
            drive_service.process_drive_folder("ROOT_CFC", project_id="PEM-058")

        # Case A: verified submission preserved alongside synced approval
        d261 = records_store["REC-VERIFIED-SUB"]["data"]
        self.assertEqual(d261["attachments"]["submission"]["file_id"], "DCR_UPLOADED_SUB_261")
        self.assertTrue(d261["attachments"]["submission"]["slot_verified"])
        self.assertEqual(d261["attachments"]["approval"]["url"], "https://drive.google.com/file/d/REP_261/view")

        # Case B: genuine Submit file preserved alongside synced approval
        d262 = records_store["REC-IN-SUBMIT-FOLDER"]["data"]
        self.assertEqual(d262["attachments"]["submission"]["url"], "https://drive.google.com/file/d/GENUINE_SUB_262/view")
        self.assertEqual(d262["attachments"]["approval"]["url"], "https://drive.google.com/file/d/REP_262/view")

    def test_08_neutral_file_badge_for_statusless_doc_types_and_subm_for_status_types(self):
        import html_render
        import i18n

        # 1) Verify i18n keys for EN & AR
        self.assertEqual(i18n.t("drive.badge_file", "en"), "File")
        self.assertEqual(i18n.t("drive.badge_file", "ar"), "ملف")

        # 2) Verify rendered register page JS wires tabHasStatusField() and drive.badge_file
        user = {"username": "dc_editor", "role": "editor"}
        proj = {"id": "PEM-058", "name": "Test Project", "code": "PEM-058", "workspace_id": "ENG", "status_colors": {}}
        with patch("html_render.can_edit", return_value=True), \
             patch("html_render.is_project_admin", return_value=True), \
             patch("html_render.get_allowed_doc_type_filter", return_value=None), \
             patch("db.can_user_download", return_value=True), \
             patch("db.get_user_workspaces", return_value=[{"id": "ENG", "code": "ENG", "name": "Engineering"}]), \
             patch("db.get_projects_for_user", return_value=[proj]), \
             patch("db.get_doc_types", return_value=[]), \
             patch("db.get_logo", return_value=None):
            rendered_en = html_render.render_register(user, proj, active_ws_id="ENG", lang="en")
            rendered_ar = html_render.render_register(user, proj, active_ws_id="ENG", lang="ar")

        self.assertIn("function tabHasStatusField(", rendered_en)
        self.assertIn("const _hasStatusField = tabHasStatusField();", rendered_en)
        self.assertIn("else if(!_hasStatusField){", rendered_en)
        self.assertIn("window.t('drive.badge_file')", rendered_en)
        self.assertIn('"drive.badge_file": "File"', rendered_en)
        self.assertIn('"drive.badge_file": "ملف"', rendered_ar)

        # 3) Verify status-less EI vs status-enabled DS logic matches exact JS helper semantics
        m_fn = re.search(r"function tabHasStatusField\(dtObj, cols\)\{.*?\n\}", rendered_en, re.DOTALL)
        self.assertIsNotNone(m_fn, "tabHasStatusField JS helper must be present in rendered HTML")
        fn_body = m_fn.group(0)
        self.assertIn("Array.isArray(dt.fields)", fn_body)
        self.assertIn("isStatusLikeField(c)", fn_body)


class TestDriveLegacyRoleFix03(unittest.TestCase):
    """Regression tests for DRIVE-LEGACY-ROLE-FIX-03:
    1. Filenames with REV variants (e.g. DS-CY002P608-00260_REV00.pdf, DS-CY002P608-00260 REV01.pdf,
       MS-CY002P608-00021-REV 02.pdf) in Replied/ match DB docNo without cross-revision collision.
    2. Legacy/imported record with uploaded_at set (no slot_verified=True and NOT in Submit/)
       is promoted cleanly to Reply and clears Subm when matched in Replied/.
    3. Genuine Submit/ file stays Subm untouched and is reported in reclassify_skipped_sample.
    4. Sync diagnostics (reclassified_count + reclassify_skipped_sample) are returned by
       process_drive_folder, /api/drive/sync/<pid>, and /api/drive/reclassify-legacy/<pid>.
    """

    def test_09_rev_filename_variants_and_legacy_uploaded_at_reclassification_with_diagnostics(self):
        projects_rows = [{"id": "PEM-058", "data": {"drive_folder_id": "ROOT_CFC"}}]
        records_store = {
            # 1) Legacy record with uploaded_at (no slot_verified=True), Drive file uses underscore _REV00
            "REC-DS-260-R00": {
                "id": "REC-DS-260-R00",
                "project_id": "PEM-058",
                "data": {
                    "docNo": "DS-CY002P608-00260 REV00",
                    "status": "B - Approved As Noted",
                    "attachments": {
                        "submission": {
                            "file_id": "LEGACY_IMPORTED_260_R00",
                            "url": "https://drive.google.com/file/d/LEGACY_IMPORTED_260_R00/view",
                            "uploaded_at": "2026-01-15T08:30:00Z",
                            "uploaded_by": "legacy_migrator",
                        }
                    },
                },
            },
            # 2) Separate revision REV01 on same base docNo -> must match ONLY REV01 PDF
            "REC-DS-260-R01": {
                "id": "REC-DS-260-R01",
                "project_id": "PEM-058",
                "data": {
                    "docNo": "DS-CY002P608-00260 REV01",
                    "status": "A - Approved",
                    "fileLocation": "https://drive.google.com/file/d/OLD_STALE_260_R01/view",
                },
            },
            # 3) MS record with hyphen/space REV variant in Drive filename + legacy uploaded_at
            "REC-MS-021-R02": {
                "id": "REC-MS-021-R02",
                "project_id": "PEM-058",
                "data": {
                    "docNo": "MS-CY002P608-00021 REV02",
                    "status": "A - Approved",
                    "fileLocation": "https://drive.google.com/file/d/LEGACY_MS_021/view",
                    "attachments": {
                        "submission": {
                            "file_id": "LEGACY_MS_021",
                            "url": "https://drive.google.com/file/d/LEGACY_MS_021/view",
                            "uploaded_at": "2026-02-01T12:00:00Z",
                        }
                    },
                },
            },
            # 4) Genuine Submit/ file -> must stay Subm untouched and appear in reclassify_skipped_sample
            "REC-GENUINE-SUB": {
                "id": "REC-GENUINE-SUB",
                "project_id": "PEM-058",
                "data": {
                    "docNo": "DS-CY002P608-00999 REV00",
                    "status": "Under Review",
                    "fileLocation": "https://drive.google.com/file/d/GENUINE_SUB_999/view",
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
                "id": "REP_DS_260_R00",
                "name": "DS-CY002P608-00260_REV00.pdf",
                "mimeType": "application/pdf",
                "parents": ["F_REP"],
                "webViewLink": "https://drive.google.com/file/d/REP_DS_260_R00/view",
            },
            {
                "id": "REP_DS_260_R01",
                "name": "DS-CY002P608-00260 REV01.pdf",
                "mimeType": "application/pdf",
                "parents": ["F_REP"],
                "webViewLink": "https://drive.google.com/file/d/REP_DS_260_R01/view",
            },
            {
                "id": "REP_MS_021_R02",
                "name": "MS-CY002P608-00021-REV 02.pdf",
                "mimeType": "application/pdf",
                "parents": ["F_REP"],
                "webViewLink": "https://drive.google.com/file/d/REP_MS_021_R02/view",
            },
            {
                "id": "GENUINE_SUB_999",
                "name": "DS-CY002P608-00999 REV00.pdf",
                "mimeType": "application/pdf",
                "parents": ["F_SUB"],
                "webViewLink": "https://drive.google.com/file/d/GENUINE_SUB_999/view",
            },
        ]

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
            if rec_id in records_store and records_store[rec_id]["project_id"] == pid_arg:
                records_store[rec_id]["data"] = json.loads(new_data_json)
            return 1

        svc = _FakeDriveService(drive_items)
        with patch("db.q", side_effect=fake_q), \
             patch("db.exe", side_effect=fake_exe), \
             patch.object(drive_service, "get_drive_service", return_value=svc):
            sync_result = drive_service.process_drive_folder("ROOT_CFC", project_id="PEM-058")

        # Verify diagnostics returned by process_drive_folder
        self.assertIsInstance(sync_result, dict)
        self.assertEqual(sync_result["reclassified_count"], 3)
        self.assertEqual(len(sync_result["reclassify_skipped_sample"]), 1)
        self.assertEqual(sync_result["reclassify_skipped_sample"][0]["docNo"], "DS-CY002P608-00999 REV00")
        self.assertEqual(sync_result["reclassify_skipped_sample"][0]["reason"], "confirmed_in_submit_folder")

        # 1) DS-CY002P608-00260 REV00 (legacy uploaded_at, underscore _REV00) promoted to Reply, Subm cleared
        d_r00 = records_store["REC-DS-260-R00"]["data"]
        self.assertNotIn("submission", d_r00.get("attachments", {}))
        self.assertEqual(d_r00["attachments"]["approval"]["file_id"], "REP_DS_260_R00")
        self.assertEqual(d_r00["attachments"]["approval"]["role"], "reply")
        self.assertTrue(d_r00["attachments"]["approval"]["slot_verified"])

        # 2) DS-CY002P608-00260 REV01 matched ONLY REV01 PDF (no collision with REV00)
        d_r01 = records_store["REC-DS-260-R01"]["data"]
        self.assertNotIn("fileLocation", d_r01)
        self.assertNotIn("submission", d_r01.get("attachments", {}))
        self.assertEqual(d_r01["attachments"]["approval"]["file_id"], "REP_DS_260_R01")

        # 3) MS-CY002P608-00021 REV02 (hyphen/space REV variant + legacy uploaded_at) promoted to Reply, Subm cleared
        d_ms = records_store["REC-MS-021-R02"]["data"]
        self.assertNotIn("fileLocation", d_ms)
        self.assertNotIn("submission", d_ms.get("attachments", {}))
        self.assertEqual(d_ms["attachments"]["approval"]["file_id"], "REP_MS_021_R02")
        self.assertEqual(d_ms["attachments"]["approval"]["role"], "reply")

        # 4) Genuine Submit/ file stays Subm untouched
        d_sub = records_store["REC-GENUINE-SUB"]["data"]
        self.assertEqual(d_sub["attachments"]["submission"]["url"], "https://drive.google.com/file/d/GENUINE_SUB_999/view")
        self.assertNotIn("approval", d_sub["attachments"])

        # 5) /api/drive/sync/<pid> exposes cached diagnostics (reclassified_count & reclassify_skipped_sample)
        client = app_mod.app.test_client()
        client.set_cookie("dcr_token", "tok_editor")
        editor = {"username": "dc_editor", "role": "editor"}

        class _NoopThread:
            def __init__(self, target=None, args=(), kwargs=None):
                pass
            def start(self):
                pass

        with patch("db.get_session", return_value=editor), \
             patch("app.can_view_project", return_value=True), \
             patch("app.can_edit", return_value=True), \
             patch("db.get_project", return_value={"id": "PEM-058", "drive_folder_id": "ROOT_CFC"}), \
             patch("db.update_project_settings"), \
             patch("drive_routing.provision_project_drive_structure"), \
             patch("threading.Thread", side_effect=_NoopThread):
            r_sync = client.post("/api/drive/sync/PEM-058", json={"folder_id": "ROOT_CFC"})
            self.assertEqual(r_sync.status_code, 200)
            body = r_sync.get_json()
            self.assertEqual(body["reclassified_count"], 3)
            self.assertEqual(len(body["reclassify_skipped_sample"]), 1)
            self.assertEqual(body["reclassify_skipped_sample"][0]["reason"], "confirmed_in_submit_folder")


if __name__ == "__main__":
    unittest.main()
