"""Offline unit & Node VM tests for DRIVE-ARCHIVE-04:
- Drive move_file & _DCR_Deleted archive folder provisioning + collision timestamping
- drive_service.process_drive_folder ignoring _DCR_Deleted + supportsAllDrives=True
- 3-mode DELETE /api/drive/attachment/<record_id>/<slot> (archive|unlink|trash),
  drive_sync URL ID extraction, cross-record shared reference protection,
  overwrite-to-archive with Trash fallback, graceful 403/404 fallback, and RBAC
- Bilingual 3-option #drive-unlink-modal & localized toasts in html_render.py (Node VM)
"""

import json
import os
import re
import subprocess
import tempfile
import unittest
from unittest.mock import patch

os.environ.setdefault("DATABASE_URL", "")
os.environ.setdefault("SECRET_KEY", "synthetic-test-secret-key-drive-archive-04")
os.environ.setdefault("MAGIC_SECRET", "synthetic-test-magic-secret-drive-archive-04")
os.environ.setdefault("GEMINI_API_KEY", "synthetic-test-gemini-key-drive-archive-04")
os.environ.setdefault("ALLOW_INSECURE_COOKIES_FOR_LOCAL_DEV", "1")

import app as app_mod
import blueprints.drive_uploads as du_mod
import db
import drive_routing
import drive_service
import drive_upload
import html_render
import i18n


class _Resp:
    def __init__(self, status_code, body=None, headers=None):
        self.status_code = status_code
        self._body = body
        self.headers = headers or {}
        if body is None:
            self.text = ""
        elif isinstance(body, str):
            self.text = body
        else:
            self.text = json.dumps(body)

    def json(self):
        return json.loads(self.text) if self.text else {}


class _FakeHTTP:
    """Scriptable fake requests-like client recording GET/POST/PATCH/DELETE calls."""

    def __init__(self, handler=None):
        self.calls = []
        self._handler = handler

    def _call(self, method, url, **kwargs):
        self.calls.append((method, url, kwargs))
        if self._handler:
            return self._handler(method, url, kwargs)
        return _Resp(200, {})

    def get(self, url, **kwargs):
        return self._call("GET", url, **kwargs)

    def post(self, url, **kwargs):
        return self._call("POST", url, **kwargs)

    def patch(self, url, **kwargs):
        return self._call("PATCH", url, **kwargs)

    def delete(self, url, **kwargs):
        return self._call("DELETE", url, **kwargs)


class _InMemoryDriveHTTP:
    """In-memory Drive REST mock backing RestDriveAdapter and drive_upload.* calls."""

    def __init__(self, items=None, fail_patch_status=None):
        self.items = {}
        for it in (items or []):
            self.items[it["id"]] = dict(it)
        self.fail_patch_status = fail_patch_status
        self.calls = []
        self._seq = 100

    def get(self, url, params=None, headers=None, timeout=30):
        params = params or {}
        self.calls.append(("GET", url, params))
        if url == drive_upload.FILES_URL:
            q = params.get("q", "")
            parent_ids = re.findall(r"'([^']+)'\s+in\s+parents", q)
            only_folders = drive_routing.FOLDER_MIME in q
            out = []
            for it in self.items.values():
                if it.get("trashed"):
                    continue
                if parent_ids and not any(p in (it.get("parents") or []) for p in parent_ids):
                    continue
                if only_folders and it.get("mimeType") != drive_routing.FOLDER_MIME:
                    continue
                out.append(dict(it))
            return _Resp(200, {"files": out})
        fid = url.rsplit("/", 1)[-1]
        if fid not in self.items:
            return _Resp(404, {"error": {"message": f"File not found: {fid}"}})
        return _Resp(200, dict(self.items[fid]))

    def post(self, url, params=None, headers=None, data=None, timeout=30):
        params = params or {}
        body = json.loads(data) if data else {}
        self.calls.append(("POST", url, params, body))
        self._seq += 1
        new_id = f"CREATED_{self._seq}"
        rec = {
            "id": new_id,
            "name": body.get("name", "Untitled"),
            "mimeType": body.get("mimeType", "application/octet-stream"),
            "parents": list(body.get("parents") or []),
            "trashed": False,
        }
        self.items[new_id] = rec
        return _Resp(200, dict(rec))

    def patch(self, url, params=None, headers=None, data=None, timeout=30):
        params = params or {}
        body = json.loads(data) if data else {}
        self.calls.append(("PATCH", url, params, body))
        if self.fail_patch_status is not None:
            return _Resp(
                self.fail_patch_status,
                {"error": {"message": f"Simulated HTTP {self.fail_patch_status}"}},
            )
        fid = url.rsplit("/", 1)[-1]
        if fid not in self.items:
            return _Resp(404, {"error": {"message": f"File not found: {fid}"}})
        rec = self.items[fid]
        if "name" in body and body["name"]:
            rec["name"] = body["name"]
        if "trashed" in body:
            rec["trashed"] = bool(body["trashed"])
        parents = list(rec.get("parents") or [])
        if params.get("removeParents"):
            rem = set(p.strip() for p in str(params["removeParents"]).split(",") if p.strip())
            parents = [p for p in parents if p not in rem]
        if params.get("addParents"):
            for ap in str(params["addParents"]).split(","):
                ap = ap.strip()
                if ap and ap not in parents:
                    parents.append(ap)
        rec["parents"] = parents
        return _Resp(200, dict(rec))


class TestDriveMoveAndArchiveFolder(unittest.TestCase):
    def setUp(self):
        drive_routing.invalidate_subfolder_caches()

    def test_01_move_file_adds_and_removes_parents_with_supports_all_drives(self):
        mem = _InMemoryDriveHTTP([
            {"id": "ROOT", "name": "ProjectRoot", "mimeType": drive_routing.FOLDER_MIME, "parents": []},
            {"id": "SUB", "name": "Submit", "mimeType": drive_routing.FOLDER_MIME, "parents": ["ROOT"]},
            {"id": "DEL", "name": "_DCR_Deleted", "mimeType": drive_routing.FOLDER_MIME, "parents": ["ROOT"]},
            {"id": "F1", "name": "DS-001.pdf", "mimeType": "application/pdf", "parents": ["SUB"], "trashed": False},
        ])
        moved = drive_upload.move_file("TOK", "F1", "DEL", new_name="20261126-100000__DS-001.pdf", http=mem)
        self.assertEqual(moved["parents"], ["DEL"])
        self.assertEqual(moved["name"], "20261126-100000__DS-001.pdf")

        patch_calls = [c for c in mem.calls if c[0] == "PATCH"]
        self.assertEqual(len(patch_calls), 1)
        _, url, params, body = patch_calls[0]
        self.assertTrue(url.endswith("/F1"))
        self.assertEqual(params["addParents"], "DEL")
        self.assertEqual(params["removeParents"], "SUB")
        self.assertEqual(params["supportsAllDrives"], "true")
        self.assertEqual(body, {"name": "20261126-100000__DS-001.pdf"})

    def test_02_ensure_deleted_archive_folder_and_collision_timestamp_prefix(self):
        mem = _InMemoryDriveHTTP([
            {"id": "ROOT", "name": "ProjectRoot", "mimeType": drive_routing.FOLDER_MIME, "parents": []},
            {"id": "SUB", "name": "Submit", "mimeType": drive_routing.FOLDER_MIME, "parents": ["ROOT"]},
            {"id": "F1", "name": "DS-001.pdf", "mimeType": "application/pdf", "parents": ["SUB"], "trashed": False},
            {"id": "F2", "name": "DS-001.pdf", "mimeType": "application/pdf", "parents": ["SUB"], "trashed": False},
        ])
        svc = drive_upload.RestDriveAdapter("TOK", http=mem)

        # First call creates _DCR_Deleted under ROOT
        del_folder = drive_routing.ensure_deleted_archive_folder(svc, "ROOT")
        self.assertEqual(del_folder["name"], "_DCR_Deleted")
        # Second call reuses the exact same folder (idempotent)
        del_folder_again = drive_routing.ensure_deleted_archive_folder(svc, "ROOT")
        self.assertEqual(del_folder_again["id"], del_folder["id"])

        # Archive F1 (no collision -> keeps original name)
        res1 = drive_routing.archive_drive_item(svc, "ROOT", "F1")
        self.assertTrue(res1["ok"])
        self.assertFalse(res1["renamed"])
        self.assertEqual(mem.items["F1"]["parents"], [del_folder["id"]])
        self.assertEqual(mem.items["F1"]["name"], "DS-001.pdf")

        # Archive F2 with same name -> collision detected -> prefixed YYYYMMDD-HHMMSS__
        res2 = drive_routing.archive_drive_item(svc, "ROOT", "F2")
        self.assertTrue(res2["ok"])
        self.assertTrue(res2["renamed"])
        self.assertEqual(mem.items["F2"]["parents"], [del_folder["id"]])
        self.assertRegex(mem.items["F2"]["name"], r"^\d{8}-\d{6}(?:-\d{3})?__DS-001\.pdf$")


class TestDriveServiceSkipsDcrDeleted(unittest.TestCase):
    def test_03_process_drive_folder_ignores_dcr_deleted_and_passes_all_drives_flags(self):
        captured_list_kwargs = []
        items = [
            {
                "id": "DT_DS",
                "name": "DS - Shop Drawings",
                "mimeType": drive_routing.FOLDER_MIME,
                "parents": ["ROOT"],
                "webViewLink": "https://drive.google.com/drive/folders/DT_DS",
            },
            {
                "id": "SUB_F",
                "name": "Submit",
                "mimeType": drive_routing.FOLDER_MIME,
                "parents": ["DT_DS"],
                "webViewLink": "https://drive.google.com/drive/folders/SUB_F",
            },
            {
                "id": "LIVE_FILE",
                "name": "DS-001.pdf",
                "mimeType": "application/pdf",
                "parents": ["SUB_F"],
                "webViewLink": "https://drive.google.com/file/d/LIVE_FILE/view",
            },
            {
                "id": "ARCHIVE_DIR",
                "name": "_DCR_Deleted",
                "mimeType": drive_routing.FOLDER_MIME,
                "parents": ["ROOT"],
                "webViewLink": "https://drive.google.com/drive/folders/ARCHIVE_DIR",
            },
            {
                "id": "ARCHIVED_SUB",
                "name": "Submit",
                "mimeType": drive_routing.FOLDER_MIME,
                "parents": ["ARCHIVE_DIR"],
                "webViewLink": "https://drive.google.com/drive/folders/ARCHIVED_SUB",
            },
            {
                "id": "ARCHIVED_FILE",
                "name": "DS-999.pdf",
                "mimeType": "application/pdf",
                "parents": ["ARCHIVED_SUB"],
                "webViewLink": "https://drive.google.com/file/d/ARCHIVED_FILE/view",
            },
        ]

        class _Exec:
            def __init__(self, payload):
                self._payload = payload

            def execute(self):
                return self._payload

        class _Files:
            def list(self, **kwargs):
                captured_list_kwargs.append(kwargs)
                q = kwargs.get("q", "")
                parent_ids = re.findall(r"'([^']+)'\s+in\s+parents", q)
                matched = [it for it in items if any(p in it["parents"] for p in parent_ids)]
                return _Exec({"files": matched})

        class _Svc:
            def files(self):
                return _Files()

        updated_calls = []
        with patch("drive_service.get_drive_service", return_value=_Svc()), \
             patch("db.update_record_link", side_effect=lambda doc_no, link, sync_links=None, project_id=None: updated_calls.append((doc_no, link, sync_links, project_id))):
            drive_service.process_drive_folder("ROOT", root_folder_name="Root", project_id="P1")

        # Verify supportsAllDrives and includeItemsFromAllDrives were passed
        self.assertTrue(len(captured_list_kwargs) >= 1)
        for kw in captured_list_kwargs:
            self.assertTrue(kw.get("supportsAllDrives"))
            self.assertTrue(kw.get("includeItemsFromAllDrives"))
            self.assertNotIn("ARCHIVE_DIR", kw.get("q", ""))

        # Only DS-001 was linked; DS-999 inside _DCR_Deleted was never traversed or linked
        linked_docs = [c[0] for c in updated_calls]
        self.assertEqual(linked_docs, ["DS-001"])


class TestAttachmentRemoval3ModesAndSafety(unittest.TestCase):
    def setUp(self):
        drive_routing.invalidate_subfolder_caches()
        self.client = app_mod.app.test_client()
        self.client.set_cookie("dcr_token", "tok_editor")
        self.editor_user = {"username": "dc_editor", "role": "editor"}
        self.viewer_user = {"username": "viewer_u", "role": "viewer"}

    def test_04_resolve_entry_drive_target_extracts_ids_from_all_url_shapes(self):
        self.assertEqual(
            du_mod._resolve_entry_drive_target({"folder_id": "FLD_1", "file_id": "FIL_1"}),
            ("FLD_1", True),
        )
        self.assertEqual(
            du_mod._resolve_entry_drive_target({"file_id": "FIL_2", "is_folder": False}),
            ("FIL_2", False),
        )
        self.assertEqual(
            du_mod._resolve_entry_drive_target({
                "url": "https://drive.google.com/drive/folders/SYNC_DIR_99?usp=sharing",
                "source": "drive_sync",
            }),
            ("SYNC_DIR_99", True),
        )
        self.assertEqual(
            du_mod._resolve_entry_drive_target({
                "url": "https://drive.google.com/file/d/SYNC_FILE_88/view?usp=drivesdk",
                "source": "drive_sync",
            }),
            ("SYNC_FILE_88", False),
        )
        self.assertEqual(
            du_mod._resolve_entry_drive_target({
                "url": "https://drive.google.com/open?id=SYNC_OPEN_77",
                "source": "drive_sync",
            }),
            ("SYNC_OPEN_77", False),
        )

    def _run_delete(self, rec, query="", json_body=None, mem=None, project_drive_id="ROOT", other_refs=0, can_edit_val=True):
        mem = mem if mem is not None else _InMemoryDriveHTTP([
            {"id": "ROOT", "name": "ProjRoot", "mimeType": drive_routing.FOLDER_MIME, "parents": []},
            {"id": "SUB", "name": "Submit", "mimeType": drive_routing.FOLDER_MIME, "parents": ["ROOT"]},
            {"id": "F_ATTACH", "name": "DS-001.pdf", "mimeType": "application/pdf", "parents": ["SUB"], "trashed": False},
            {"id": "SYNC_FILE_88", "name": "DS-002.pdf", "mimeType": "application/pdf", "parents": ["SUB"], "trashed": False},
        ])
        cleared = []
        with patch.object(du_mod, "_token_provider", return_value="TOK"), \
             patch.object(du_mod, "_http_client", mem), \
             patch("db.get_session", return_value=self.editor_user), \
             patch("blueprints.drive_uploads.current_user", return_value=self.editor_user), \
             patch("blueprints.drive_uploads.can_edit", return_value=can_edit_val), \
             patch("blueprints.drive_uploads.can_access_doc_type", return_value=True), \
             patch("db.get_record_by_id", return_value=rec), \
             patch("db.get_project", return_value={"id": "P1", "drive_folder_id": project_drive_id}), \
             patch("db.count_other_records_referencing_drive_item", return_value=other_refs), \
             patch("db.clear_record_attachment", side_effect=lambda rid, sl: cleared.append((rid, sl)) or True), \
             patch("db.get_record_attachments", return_value={}), \
             patch("db.q", return_value=[{"id": rec["_id"]}]), \
             patch("db.log_action"):
            url = f"/api/drive/attachment/{rec['_id']}/submission{query}"
            resp = self.client.delete(url, json=json_body)
        return resp, mem, cleared

    def test_05_delete_mode_archive_default_and_drive_sync_extraction(self):
        rec = {
            "_id": "R1",
            "_project_id": "P1",
            "_dt_id": "DS",
            "docNo": "DS-002",
            "fileLocation": "https://drive.google.com/file/d/SYNC_FILE_88/view",
            "attachments": {
                "submission": {
                    "url": "https://drive.google.com/file/d/SYNC_FILE_88/view",
                    "source": "drive_sync",
                }
            },
        }
        # Omit mode -> defaults to 'archive'
        resp, mem, cleared = self._run_delete(rec)
        self.assertEqual(resp.status_code, 200)
        body = resp.get_json()
        self.assertTrue(body["ok"])
        self.assertEqual(body["mode"], "archive")
        self.assertEqual(body["drive_action"], "archived")
        self.assertEqual(cleared, [("R1", "submission")])

        # Verify file was moved into newly created _DCR_Deleted under ROOT (NOT trashed)
        del_folders = [it for it in mem.items.values() if it["name"] == "_DCR_Deleted"]
        self.assertEqual(len(del_folders), 1)
        self.assertEqual(mem.items["SYNC_FILE_88"]["parents"], [del_folders[0]["id"]])
        self.assertFalse(mem.items["SYNC_FILE_88"]["trashed"])

    def test_06_delete_mode_unlink_leaves_drive_untouched_and_warns_for_drive_sync(self):
        rec = {
            "_id": "R2",
            "_project_id": "P1",
            "_dt_id": "DS",
            "docNo": "DS-002",
            "attachments": {
                "submission": {
                    "url": "https://drive.google.com/file/d/SYNC_FILE_88/view",
                    "source": "drive_sync",
                }
            },
        }
        resp, mem, cleared = self._run_delete(rec, query="?mode=unlink")
        self.assertEqual(resp.status_code, 200)
        body = resp.get_json()
        self.assertEqual(body["mode"], "unlink")
        self.assertTrue(body.get("sync_relink_warning"))
        self.assertEqual(cleared, [("R2", "submission")])
        # Zero Drive HTTP calls
        self.assertEqual(mem.calls, [])
        self.assertEqual(mem.items["SYNC_FILE_88"]["parents"], ["SUB"])

    def test_07_delete_mode_trash_moves_item_to_drive_trash(self):
        rec = {
            "_id": "R3",
            "_project_id": "P1",
            "_dt_id": "DS",
            "docNo": "DS-001",
            "attachments": {
                "submission": {
                    "file_id": "F_ATTACH",
                    "file_name": "DS-001.pdf",
                    "url": "https://drive.google.com/file/d/F_ATTACH/view",
                }
            },
        }
        resp, mem, cleared = self._run_delete(rec, query="?mode=trash")
        self.assertEqual(resp.status_code, 200)
        body = resp.get_json()
        self.assertEqual(body["mode"], "trash")
        self.assertEqual(body["drive_action"], "trashed")
        self.assertTrue(mem.items["F_ATTACH"]["trashed"])
        self.assertEqual(cleared, [("R3", "submission")])

    def test_08_cross_record_shared_reference_downgrades_to_unlink_only(self):
        rec = {
            "_id": "R4",
            "_project_id": "P1",
            "_dt_id": "DS",
            "docNo": "DS-001",
            "attachments": {
                "submission": {
                    "file_id": "F_ATTACH",
                    "file_name": "DS-001.pdf",
                    "url": "https://drive.google.com/file/d/F_ATTACH/view",
                }
            },
        }
        resp, mem, cleared = self._run_delete(rec, query="?mode=trash", other_refs=2)
        self.assertEqual(resp.status_code, 200)
        body = resp.get_json()
        self.assertEqual(body["mode"], "unlink")
        self.assertTrue(body.get("shared_reference_preserved"))
        self.assertTrue(bool(body.get("drive_warning")))
        self.assertEqual(cleared, [("R4", "submission")])
        # Drive item was neither moved nor trashed
        self.assertEqual(mem.calls, [])
        self.assertFalse(mem.items["F_ATTACH"]["trashed"])
        self.assertEqual(mem.items["F_ATTACH"]["parents"], ["SUB"])

    def test_09_graceful_403_and_404_fallback_still_unlinks_db_with_warning(self):
        rec = {
            "_id": "R5",
            "_project_id": "P1",
            "_dt_id": "DS",
            "docNo": "DS-001",
            "attachments": {
                "submission": {
                    "file_id": "F_ATTACH",
                    "file_name": "DS-001.pdf",
                    "url": "https://drive.google.com/file/d/F_ATTACH/view",
                }
            },
        }
        for code in (403, 404):
            with self.subTest(http_status=code):
                mem = _InMemoryDriveHTTP([
                    {"id": "ROOT", "name": "ProjRoot", "mimeType": drive_routing.FOLDER_MIME, "parents": []},
                    {"id": "DEL", "name": "_DCR_Deleted", "mimeType": drive_routing.FOLDER_MIME, "parents": ["ROOT"]},
                    {"id": "SUB", "name": "Submit", "mimeType": drive_routing.FOLDER_MIME, "parents": ["ROOT"]},
                    {"id": "F_ATTACH", "name": "DS-001.pdf", "mimeType": "application/pdf", "parents": ["SUB"], "trashed": False},
                ], fail_patch_status=code)
                resp, _, cleared = self._run_delete(rec, query="?mode=archive", mem=mem)
                self.assertEqual(resp.status_code, 200)
                body = resp.get_json()
                self.assertTrue(body["ok"])
                self.assertIn(str(code), body.get("drive_warning", ""))
                self.assertEqual(cleared, [("R5", "submission")])

    def test_10_rbac_blocks_unauthorized_user_and_rejects_invalid_mode(self):
        rec = {"_id": "R6", "_project_id": "P1", "_dt_id": "DS", "attachments": {}}
        resp_forbid, _, _ = self._run_delete(rec, query="?mode=archive", can_edit_val=False)
        self.assertEqual(resp_forbid.status_code, 403)

        resp_bad, _, _ = self._run_delete(rec, query="?mode=nuke")
        self.assertEqual(resp_bad.status_code, 400)

    def test_11_upload_complete_overwrite_archives_to_dcr_deleted_and_falls_back_to_trash(self):
        # Case A: Project has drive_folder_id -> replaced file archived to _DCR_Deleted
        mem_arch = _InMemoryDriveHTTP([
            {"id": "ROOT", "name": "ProjRoot", "mimeType": drive_routing.FOLDER_MIME, "parents": []},
            {"id": "SUB", "name": "Submit", "mimeType": drive_routing.FOLDER_MIME, "parents": ["ROOT"]},
            {"id": "OLD_F", "name": "DS-001.pdf", "mimeType": "application/pdf", "parents": ["SUB"], "trashed": False},
            {"id": "NEW_F", "name": "DS-001.pdf", "mimeType": "application/pdf", "parents": ["SUB"], "trashed": False, "size": "1234", "webViewLink": "https://drive.google.com/file/d/NEW_F/view"},
        ])
        rec = {
            "_id": "R7",
            "_project_id": "P1",
            "_dt_id": "DS",
            "docNo": "DS-001",
            "attachments": {
                "submission": {"file_id": "OLD_F", "file_name": "DS-001.pdf", "url": "https://drive.google.com/file/d/OLD_F/view"}
            },
        }
        ticket = du_mod.sign_upload_ticket({
            "pid": "P1", "dt_id": "DS", "ws_id": "ENG", "record_id": "R7",
            "slot": "submission", "folder_id": "SUB", "file_name": "DS-001.pdf",
            "file_size": 1234, "mime_type": "application/pdf", "user": "dc_editor",
            "session_url": "https://upload.example/s1",
        })
        with patch.object(du_mod, "_token_provider", return_value="TOK"), \
             patch.object(du_mod, "_http_client", mem_arch), \
             patch("db.get_session", return_value=self.editor_user), \
             patch("blueprints.drive_uploads.current_user", return_value=self.editor_user), \
             patch("blueprints.drive_uploads.can_edit", return_value=True), \
             patch("blueprints.drive_uploads.can_access_doc_type", return_value=True), \
             patch("db.get_record_by_id", return_value=rec), \
             patch("db.get_project", return_value={"id": "P1", "workspace_id": "ENG", "drive_folder_id": "ROOT"}), \
             patch("db.count_other_records_referencing_drive_item", return_value=0), \
             patch("db.set_record_attachment", return_value=True), \
             patch("db.get_record_attachments", return_value={}), \
             patch("db.q", return_value=[{"id": "R7"}]), \
             patch("db.log_action"):
            r = self.client.post("/api/drive/upload-complete", json={"upload_ticket": ticket, "file_id": "NEW_F", "replace_existing": True})
        self.assertEqual(r.status_code, 200)
        del_folders = [it for it in mem_arch.items.values() if it["name"] == "_DCR_Deleted"]
        self.assertEqual(len(del_folders), 1)
        self.assertEqual(mem_arch.items["OLD_F"]["parents"], [del_folders[0]["id"]])
        self.assertFalse(mem_arch.items["OLD_F"]["trashed"])

        # Case B: Project has NO drive_folder_id -> falls back to Trash
        mem_trash = _InMemoryDriveHTTP([
            {"id": "SUB", "name": "Submit", "mimeType": drive_routing.FOLDER_MIME, "parents": []},
            {"id": "OLD_F", "name": "DS-001.pdf", "mimeType": "application/pdf", "parents": ["SUB"], "trashed": False},
            {"id": "NEW_F", "name": "DS-001.pdf", "mimeType": "application/pdf", "parents": ["SUB"], "trashed": False, "size": "1234", "webViewLink": "https://drive.google.com/file/d/NEW_F/view"},
        ])
        with patch.object(du_mod, "_token_provider", return_value="TOK"), \
             patch.object(du_mod, "_http_client", mem_trash), \
             patch("db.get_session", return_value=self.editor_user), \
             patch("blueprints.drive_uploads.current_user", return_value=self.editor_user), \
             patch("blueprints.drive_uploads.can_edit", return_value=True), \
             patch("blueprints.drive_uploads.can_access_doc_type", return_value=True), \
             patch("db.get_record_by_id", return_value=rec), \
             patch("db.get_project", return_value={"id": "P1", "workspace_id": "ENG", "drive_folder_id": ""}), \
             patch("db.count_other_records_referencing_drive_item", return_value=0), \
             patch("db.set_record_attachment", return_value=True), \
             patch("db.get_record_attachments", return_value={}), \
             patch("db.q", return_value=[{"id": "R7"}]), \
             patch("db.log_action"):
            r2 = self.client.post("/api/drive/upload-complete", json={"upload_ticket": ticket, "file_id": "NEW_F", "replace_existing": True})
        self.assertEqual(r2.status_code, 200)
        self.assertTrue(mem_trash.items["OLD_F"]["trashed"])

    def test_12_db_count_other_records_referencing_drive_item_sql_params(self):
        captured = []
        with patch("db.q", side_effect=lambda sql, params=(), one=False: captured.append((sql, params, one)) or {"cnt": 3}):
            cnt = db.count_other_records_referencing_drive_item(
                "P1", "REC_1", item_id="FILE_ABC_123", url="https://drive.google.com/file/d/FILE_ABC_123/view"
            )
        self.assertEqual(cnt, 3)
        self.assertEqual(len(captured), 1)
        sql, params, one = captured[0]
        self.assertTrue(one)
        self.assertIn("project_id = %s", sql)
        self.assertIn("id != %s", sql)
        self.assertIn("STRPOS", sql)
        self.assertEqual(params[0], "P1")
        self.assertEqual(params[1], "REC_1")
        self.assertIn("FILE_ABC_123", params)


class TestDriveUnlinkModalFrontendNodeVM(unittest.TestCase):
    def test_13_i18n_en_ar_parity_and_node_vm_modal_flow(self):
        # 1. Verify EN/AR dictionary key parity for all new unlink keys
        en_keys = set(i18n.TRANSLATIONS["en"].keys())
        ar_keys = set(i18n.TRANSLATIONS["ar"].keys())
        self.assertEqual(en_keys, ar_keys)
        for k in (
            "form.unlink_modal_title",
            "form.unlink_mode_archive_title",
            "form.unlink_mode_unlink_title",
            "form.unlink_mode_trash_title",
            "form.unlinked_archived",
            "form.unlinked_only",
            "form.unlinked_trashed",
            "form.unlinked_drive_warning",
        ):
            self.assertIn(k, en_keys)
            self.assertIn(k, ar_keys)

        # 2. Extract openDriveUnlinkModal, _showDriveUnlinkToast, and mountDriveAttachmentsBox from rendered HTML
        src_path = os.path.join(os.path.dirname(__file__), "html_render.py")
        with open(src_path, "r", encoding="utf-8") as fh:
            py_src = fh.read()

        # Render the JS functions by unescaping {{ and }} in the extracted JS block
        start_idx = py_src.index("function openDriveUnlinkModal(")
        end_idx = py_src.index("function mountDriveAttachmentsBox(", start_idx)
        # Find the matching end of mountDriveAttachmentsBox by locating the next top-level function
        next_fn = py_src.find("\nfunction ", end_idx + 10)
        js_block = py_src[start_idx:next_fn].replace("{{", "{").replace("}}", "}")

        en_json = json.dumps(i18n.TRANSLATIONS["en"], ensure_ascii=False)
        ar_json = json.dumps(i18n.TRANSLATIONS["ar"], ensure_ascii=False)

        driver = f"""
const assert = require('assert');
const EN = {en_json};
const AR = {ar_json};
let activeLang = 'en';
global.window = {{
  t: (k, vars) => {{
    const dict = activeLang === 'ar' ? AR : EN;
    let s = dict[k] || k;
    if (vars) {{
      for (const [vk, vv] of Object.entries(vars)) {{
        s = s.replace(new RegExp('\\\\{{' + vk + '\\\\}}', 'g'), String(vv));
      }}
    }}
    return s;
  }}
}};

function mkEl(tag) {{
  const el = {{
    tagName: String(tag || 'div').toUpperCase(),
    id: '',
    className: '',
    type: '',
    name: '',
    value: '',
    checked: false,
    disabled: false,
    textContent: '',
    innerHTML: '',
    title: '',
    dir: '',
    style: {{}},
    dataset: {{}},
    attrs: {{}},
    children: [],
    parentNode: null,
    listeners: {{}},
    classList: {{
      _set: new Set(),
      add(c) {{ this._set.add(c); }},
      remove(c) {{ this._set.delete(c); }},
      contains(c) {{ return this._set.has(c); }}
    }},
    setAttribute(k, v) {{ this.attrs[k] = String(v); }},
    getAttribute(k) {{ return this.attrs[k]; }},
    appendChild(child) {{
      child.parentNode = this;
      this.children.push(child);
      return child;
    }},
    removeChild(child) {{
      const i = this.children.indexOf(child);
      if (i >= 0) this.children.splice(i, 1);
      child.parentNode = null;
    }},
    addEventListener(ev, fn) {{ (this.listeners[ev] = this.listeners[ev] || []).push(fn); }},
    removeEventListener(ev, fn) {{
      const arr = this.listeners[ev] || [];
      const i = arr.indexOf(fn);
      if (i >= 0) arr.splice(i, 1);
    }},
    querySelector(sel) {{
      if (sel.startsWith('#')) {{
        const tid = sel.slice(1);
        return findBy(this, n => n.id === tid);
      }}
      return null;
    }}
  }};
  return el;
}}

function findBy(root, pred) {{
  if (!root) return null;
  if (pred(root)) return root;
  for (const ch of (root.children || [])) {{
    const hit = findBy(ch, pred);
    if (hit) return hit;
  }}
  return null;
}}

function findAll(root, pred, out = []) {{
  if (!root) return out;
  if (pred(root)) out.push(root);
  for (const ch of (root.children || [])) findAll(ch, pred, out);
  return out;
}}

const body = mkEl('body');
const docEl = mkEl('html');
global.document = {{
  body,
  documentElement: docEl,
  createElement: mkEl,
  getElementById: (id) => findBy(body, n => n.id === id),
  addEventListener: () => {{}},
  removeEventListener: () => {{}},
}};

const toasts = [];
global.toast = (msg, typ) => toasts.push({{ msg, typ }});
let nativeConfirmCalls = 0;
global.confirm = () => {{ nativeConfirmCalls++; return true; }};

{js_block}

(async () => {{
  // Test 1: openDriveUnlinkModal defaults to 'archive'
  const p1 = openDriveUnlinkModal({{ slot: 'submission', slotTitle: 'Submission', itemName: 'DS-001.pdf', isSync: true }});
  const modal1 = document.getElementById('drive-unlink-modal');
  assert(modal1, '#drive-unlink-modal must be mounted');
  assert.strictEqual(modal1.getAttribute('role'), 'dialog');
  assert.strictEqual(modal1.getAttribute('aria-modal'), 'true');
  const radios1 = findAll(modal1, n => n.tagName === 'INPUT' && n.type === 'radio');
  assert.strictEqual(radios1.length, 3);
  assert.strictEqual(radios1[0].value, 'archive');
  assert.strictEqual(radios1[0].checked, true);
  assert.strictEqual(radios1[1].value, 'unlink');
  assert.strictEqual(radios1[2].value, 'trash');
  // Confirm default ('archive')
  document.getElementById('drive-unlink-confirm-btn').onclick();
  assert.strictEqual(await p1, 'archive');
  assert.strictEqual(document.getElementById('drive-unlink-modal'), null);

  // Test 2: Select 'trash' in Dark + Arabic RTL mode
  activeLang = 'ar';
  docEl.dir = 'rtl';
  body.classList.add('dark');
  const p2 = openDriveUnlinkModal({{ slot: 'approval', slotTitle: 'Replied', itemName: 'DS-001_Reply.pdf' }});
  const modal2 = document.getElementById('drive-unlink-modal');
  assert.strictEqual(modal2.dir, 'rtl');
  const title2 = document.getElementById('drive-unlink-modal-title');
  assert.strictEqual(title2.textContent, AR['form.unlink_modal_title']);
  const trashLbl = findBy(modal2, n => n.dataset && n.dataset.unlinkMode === 'trash');
  assert(trashLbl, 'trash option card must exist');
  trashLbl.onclick();
  document.getElementById('drive-unlink-confirm-btn').onclick();
  assert.strictEqual(await p2, 'trash');

  // Test 3: Cancel resolves null
  activeLang = 'en';
  docEl.dir = 'ltr';
  const p3 = openDriveUnlinkModal({{ slot: 'submission', slotTitle: 'Submission', itemName: 'DS-001.pdf' }});
  document.getElementById('drive-unlink-cancel-btn').onclick();
  assert.strictEqual(await p3, null);

  // Test 4: _showDriveUnlinkToast outcomes
  toasts.length = 0;
  _showDriveUnlinkToast({{ ok: true, mode: 'archive' }}, 'archive', {{ slot: 'submission' }});
  _showDriveUnlinkToast({{ ok: true, mode: 'unlink', sync_relink_warning: true }}, 'unlink', {{ slot: 'submission' }});
  _showDriveUnlinkToast({{ ok: true, mode: 'trash' }}, 'trash', {{ slot: 'submission' }});
  _showDriveUnlinkToast({{ ok: true, mode: 'unlink', shared_reference_preserved: true }}, 'archive', {{ slot: 'submission' }});
  _showDriveUnlinkToast({{ ok: true, mode: 'archive', drive_warning: 'HTTP 403 denied' }}, 'archive', {{ slot: 'submission' }});
  assert.strictEqual(toasts.length, 5);
  assert(toasts[0].msg.includes('_DCR_Deleted') && toasts[0].typ === 'ok');
  assert(toasts[1].msg.includes('Submit/Replied') && toasts[1].typ === 'wn');
  assert(toasts[2].msg.includes('Trash') && toasts[2].typ === 'ok');
  assert(toasts[3].msg.includes('another record') && toasts[3].typ === 'wn');
  assert(toasts[4].msg.includes('HTTP 403 denied') && toasts[4].typ === 'wn');

  // Verify native confirm was never called and source uses openDriveUnlinkModal + ?mode=
  assert.strictEqual(nativeConfirmCalls, 0);
  console.log('NODE_VM_OK');
}})().catch(err => {{ console.error(err); process.exit(1); }});
"""
        with tempfile.NamedTemporaryFile("w", suffix=".js", delete=False, encoding="utf-8") as tf:
            tf.write(driver)
            tmp_js = tf.name
        try:
            proc = subprocess.run(["node", tmp_js], capture_output=True, text=True, timeout=15)
            self.assertEqual(proc.returncode, 0, f"Node VM failed:\nSTDOUT: {proc.stdout}\nSTDERR: {proc.stderr}")
            self.assertIn("NODE_VM_OK", proc.stdout)
        finally:
            try:
                os.unlink(tmp_js)
            except OSError:
                pass

        # Also verify mountDriveAttachmentsBox no longer uses native confirm() for unlink and sends mode param
        fn_only = py_src[end_idx:next_fn]
        self.assertNotIn("confirm(msg)", fn_only)
        self.assertNotIn("confirm(window.t('noc.confirm_remove_archived'", fn_only)
        self.assertIn("openDriveUnlinkModal(", fn_only)
        self.assertIn("'?mode='+encodeURIComponent(mode)", fn_only)
        self.assertIn("'&mode='+encodeURIComponent(mode)", fn_only)


if __name__ == "__main__":
    unittest.main()
