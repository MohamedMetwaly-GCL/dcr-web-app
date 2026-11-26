"""Offline unit tests for DRIVE-DELETE-CLEANUP-05:
Safe Smart Drive Folder Cleanup/Archival on Project, DocType Tab, and Sub-Tab Deletion.
"""

import copy
import json
import os
import re
import unittest
from unittest.mock import patch

os.environ.setdefault("DATABASE_URL", "")
os.environ.setdefault("SECRET_KEY", "synthetic-test-secret-key-drive-delete-cleanup-05")
os.environ.setdefault("MAGIC_SECRET", "synthetic-test-magic-secret-drive-delete-cleanup-05")
os.environ.setdefault("GEMINI_API_KEY", "synthetic-test-gemini-key-drive-delete-cleanup-05")
os.environ.setdefault("ALLOW_INSECURE_COOKIES_FOR_LOCAL_DEV", "1")

import app as app_mod
import drive_routing
import drive_upload


class _Resp:
    def __init__(self, status_code, body=None):
        self.status_code = status_code
        self.text = "" if body is None else (body if isinstance(body, str) else json.dumps(body))

    def json(self):
        return json.loads(self.text) if self.text else {}


class _InMemoryDriveHTTP:
    """In-memory Drive REST mock backing RestDriveAdapter for delete/cleanup tests."""

    def __init__(self, items=None, fail_on_get=False):
        self.items = {it["id"]: dict(it) for it in (items or [])}
        self.fail_on_get = fail_on_get
        self.calls = []
        self._seq = 500

    def get(self, url, params=None, headers=None, timeout=30):
        params = params or {}
        self.calls.append(("GET", url, params))
        if self.fail_on_get:
            return _Resp(503, {"error": {"message": "Simulated Drive unreachable"}})
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
            return _Resp(404, {"error": {"message": f"Not found: {fid}"}})
        return _Resp(200, dict(self.items[fid]))

    def post(self, url, params=None, headers=None, data=None, timeout=30):
        params = params or {}
        body = json.loads(data) if data else {}
        self.calls.append(("POST", url, params, body))
        self._seq += 1
        new_id = f"DRV_NEW_{self._seq}"
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
        fid = url.rsplit("/", 1)[-1]
        if fid not in self.items:
            return _Resp(404, {"error": {"message": f"Not found: {fid}"}})
        rec = self.items[fid]
        if "name" in body and body["name"]:
            rec["name"] = body["name"]
        if "trashed" in body:
            rec["trashed"] = bool(body["trashed"])
        parents = list(rec.get("parents") or [])
        if params.get("removeParents"):
            rem = {p.strip() for p in str(params["removeParents"]).split(",") if p.strip()}
            parents = [p for p in parents if p not in rem]
        if params.get("addParents"):
            for ap in str(params["addParents"]).split(","):
                ap = ap.strip()
                if ap and ap not in parents:
                    parents.append(ap)
        rec["parents"] = parents
        return _Resp(200, dict(rec))


def _fld(fid, name, parent=None, trashed=False):
    return {
        "id": fid,
        "name": name,
        "mimeType": drive_routing.FOLDER_MIME,
        "parents": [parent] if parent else [],
        "trashed": trashed,
    }


def _file(fid, name, parent):
    return {
        "id": fid,
        "name": name,
        "mimeType": "application/pdf",
        "parents": [parent],
        "trashed": False,
    }


class TestDriveDeleteCleanup(unittest.TestCase):
    def setUp(self):
        drive_routing.invalidate_subfolder_caches()
        app_mod.app.config["TESTING"] = True
        self.client = app_mod.app.test_client()
        self.client.set_cookie("dcr_token", "tok_test")
        self.super_user = {"username": "admin", "role": "superadmin"}

    def test_01_project_delete_trashes_empty_archives_nonempty_skips_legacy(self):
        http = _InMemoryDriveHTTP([
            _fld("MASTER", "DCR Master"),
            _fld("WS_ENG", "ENG - Engineering", "MASTER"),
            # Auto-provisioned empty project (only empty subfolders)
            _fld("PROJ_EMPTY", "EP - Empty Proj", "WS_ENG"),
            _fld("PE_DT", "DS - Document Submittal", "PROJ_EMPTY"),
            _fld("PE_SUB", "Submit", "PE_DT"),
            # Auto-provisioned non-empty project (has nested PDF)
            _fld("PROJ_FULL", "FP - Full Proj", "WS_ENG"),
            _fld("PF_DT", "DS - Document Submittal", "PROJ_FULL"),
            _fld("PF_SUB", "Submit", "PF_DT"),
            _file("FILE_1", "submittal.pdf", "PF_SUB"),
            # Legacy manual project (auto_provisioned_drive=False)
            _fld("PROJ_LEGACY", "Legacy Client Root", "MASTER"),
            _file("FILE_LEG", "keep.pdf", "PROJ_LEGACY"),
        ])
        svc = drive_upload.RestDriveAdapter("TOK", http=http)

        projects = {
            "EP": {
                "id": "EP", "name": "Empty Proj", "code": "EP", "workspace_id": "ENG",
                "drive_folder_id": "PROJ_EMPTY", "auto_provisioned_drive": True,
            },
            "FP": {
                "id": "FP", "name": "Full Proj", "code": "FP", "workspace_id": "ENG",
                "drive_folder_id": "PROJ_FULL", "auto_provisioned_drive": True,
            },
            "LP": {
                "id": "LP", "name": "Legacy Proj", "code": "LP", "workspace_id": "ENG",
                "drive_folder_id": "PROJ_LEGACY", "auto_provisioned_drive": False,
            },
        }
        deleted_pids = []

        with patch("db.get_session", return_value=self.super_user), \
             patch.object(drive_routing, "_service_provider", return_value=svc), \
             patch("db.get_master_drive_root_id", return_value="MASTER"), \
             patch("db.get_project", side_effect=lambda pid: copy.deepcopy(projects.get(pid))), \
             patch("db.delete_project", side_effect=lambda pid: deleted_pids.append(pid)):

            r1 = self.client.post("/api/projects/delete/EP")
            self.assertEqual(r1.status_code, 200)
            j1 = r1.get_json()
            self.assertTrue(j1["ok"])
            self.assertEqual(j1["drive_cleanup"]["action"], "trashed")
            self.assertTrue(http.items["PROJ_EMPTY"]["trashed"])

            r2 = self.client.post("/api/projects/delete/FP")
            self.assertEqual(r2.status_code, 200)
            j2 = r2.get_json()
            self.assertTrue(j2["ok"])
            self.assertEqual(j2["drive_cleanup"]["action"], "archived")
            archive_id = j2["drive_cleanup"]["archive_folder_id"]
            self.assertEqual(http.items[archive_id]["name"], drive_routing.DELETED_ARCHIVE_FOLDER_NAME)
            self.assertEqual(http.items[archive_id]["parents"], ["MASTER"])
            self.assertEqual(http.items["PROJ_FULL"]["parents"], [archive_id])
            self.assertFalse(http.items["PROJ_FULL"]["trashed"])

            r3 = self.client.post("/api/projects/delete/LP")
            self.assertEqual(r3.status_code, 200)
            j3 = r3.get_json()
            self.assertTrue(j3["ok"])
            self.assertTrue(j3["drive_cleanup"]["skipped"])
            self.assertEqual(j3["drive_cleanup"]["reason"], "not_auto_provisioned")
            self.assertFalse(http.items["PROJ_LEGACY"]["trashed"])
            self.assertEqual(http.items["PROJ_LEGACY"]["parents"], ["MASTER"])

        self.assertEqual(deleted_pids, ["EP", "FP", "LP"])

    def test_02_doc_type_delete_trashes_empty_archives_nonempty_and_skips_ambiguous(self):
        http = _InMemoryDriveHTTP([
            _fld("PROJ", "P1 - Project"),
            # Empty DocType folder with empty Submit/Replied
            _fld("DT_EMPTY", "RFI - Request for Info", "PROJ"),
            _fld("DT_EMPTY_SUB", "Submit", "DT_EMPTY"),
            _fld("DT_EMPTY_REP", "Replied", "DT_EMPTY"),
            # Non-empty DocType folder with a file in Replied
            _fld("DT_FULL", "DS - Document Submittal", "PROJ"),
            _fld("DT_FULL_SUB", "Submit", "DT_FULL"),
            _fld("DT_FULL_REP", "Replied", "DT_FULL"),
            _file("F_REP", "reply.pdf", "DT_FULL_REP"),
            # Ambiguous DocType folders (two folders matching 'TQ')
            _fld("DT_AMB_1", "TQ - Technical Query", "PROJ"),
            _fld("DT_AMB_2", "TQ - Technical Query (Old)", "PROJ"),
        ])
        svc = drive_upload.RestDriveAdapter("TOK", http=http)

        proj = {"id": "P1", "name": "Project", "code": "P1", "workspace_id": "ENG", "drive_folder_id": "PROJ"}
        dts = {
            "RFI": {"id": "RFI", "code": "RFI", "name": "Request for Info"},
            "DS": {"id": "DS", "code": "DS", "name": "Document Submittal"},
            "TQ": {"id": "TQ", "code": "TQ", "name": "Technical Query"},
        }
        deleted_dts = []

        with patch("db.get_session", return_value=self.super_user), \
             patch.object(drive_routing, "_service_provider", return_value=svc), \
             patch("db.get_project", return_value=proj), \
             patch("db.get_doc_type", side_effect=lambda pid, dt_id: copy.deepcopy(dts.get(dt_id))), \
             patch("db.delete_doc_type", side_effect=lambda pid, dt_id: deleted_dts.append(dt_id)):

            r1 = self.client.delete("/api/doc_types/P1/RFI")
            self.assertEqual(r1.status_code, 200)
            self.assertEqual(r1.get_json()["drive_cleanup"]["action"], "trashed")
            self.assertTrue(http.items["DT_EMPTY"]["trashed"])

            r2 = self.client.delete("/api/doc_types/P1/DS")
            self.assertEqual(r2.status_code, 200)
            j2 = r2.get_json()
            self.assertEqual(j2["drive_cleanup"]["action"], "archived")
            arc_id = j2["drive_cleanup"]["archive_folder_id"]
            self.assertEqual(http.items[arc_id]["parents"], ["PROJ"])
            self.assertEqual(http.items["DT_FULL"]["parents"], [arc_id])
            self.assertFalse(http.items["DT_FULL"]["trashed"])

            r3 = self.client.delete("/api/doc_types/P1/TQ")
            self.assertEqual(r3.status_code, 200)
            j3 = r3.get_json()
            self.assertTrue(j3["drive_cleanup"]["skipped"])
            self.assertEqual(j3["drive_cleanup"]["reason"], "ambiguous")
            self.assertFalse(http.items["DT_AMB_1"]["trashed"])
            self.assertFalse(http.items["DT_AMB_2"]["trashed"])

        self.assertEqual(deleted_dts, ["RFI", "DS", "TQ"])

    def test_03_subtab_delete_cleans_up_without_affecting_siblings_and_is_non_blocking(self):
        http = _InMemoryDriveHTTP([
            _fld("PROJ", "P1 - Project"),
            _fld("DT_PTW", "PTW - Permit to Work", "PROJ"),
            # Empty sub-tab
            _fld("SUB_HOT", "Hot Work", "DT_PTW"),
            _fld("SUB_HOT_S", "Submit", "SUB_HOT"),
            # Non-empty sub-tab
            _fld("SUB_CONF", "Confined Space", "DT_PTW"),
            _fld("SUB_CONF_S", "Submit", "SUB_CONF"),
            _file("F_CONF", "permit.pdf", "SUB_CONF_S"),
            # Sibling sub-tab that must stay untouched
            _fld("SUB_EXC", "Excavation", "DT_PTW"),
            _fld("SUB_EXC_S", "Submit", "SUB_EXC"),
            _file("F_EXC", "excavation.pdf", "SUB_EXC_S"),
        ])
        svc = drive_upload.RestDriveAdapter("TOK", http=http)

        proj = {"id": "P1", "name": "Project", "code": "P1", "workspace_id": "HSE", "drive_folder_id": "PROJ"}
        dt_ptw = {"id": "PTW", "code": "PTW", "name": "Permit to Work"}
        removed_items = []

        with patch("db.get_session", return_value=self.super_user), \
             patch.object(drive_routing, "_service_provider", return_value=svc), \
             patch("db.get_project", return_value=proj), \
             patch("db.resolve_doc_types_for_subtype_list", return_value=["PTW"]), \
             patch("db.get_doc_type", return_value=dt_ptw), \
             patch("db.set_doc_type_drive_cache"), \
             patch("db.remove_list_item", side_effect=lambda pid, ln, it: removed_items.append((ln, it))):

            # Empty sub-tab -> trashed
            r1 = self.client.delete("/api/lists/P1", json={
                "list_name": "hse_ptw_subtype", "item": "Hot Work"
            })
            self.assertEqual(r1.status_code, 200)
            self.assertEqual(r1.get_json()["drive_cleanup"]["action"], "trashed")
            self.assertTrue(http.items["SUB_HOT"]["trashed"])

            # Non-empty sub-tab -> archived to <Project>/_DCR_Deleted
            r2 = self.client.delete("/api/lists/P1", json={
                "list_name": "hse_ptw_subtype", "item": "Confined Space"
            })
            self.assertEqual(r2.status_code, 200)
            j2 = r2.get_json()
            self.assertEqual(j2["drive_cleanup"]["action"], "archived")
            arc_id = j2["drive_cleanup"]["archive_folder_id"]
            self.assertEqual(http.items[arc_id]["parents"], ["PROJ"])
            self.assertEqual(http.items["SUB_CONF"]["parents"], [arc_id])
            self.assertFalse(http.items["SUB_CONF"]["trashed"])

            # Sibling sub-tab untouched
            self.assertFalse(http.items["SUB_EXC"]["trashed"])
            self.assertEqual(http.items["SUB_EXC"]["parents"], ["DT_PTW"])

            # Non-blocking on Drive API failure
            with patch.object(drive_routing, "cleanup_subtype_drive_folder", side_effect=RuntimeError("Drive down")):
                r3 = self.client.delete("/api/lists/P1", json={
                    "list_name": "hse_ptw_subtype", "item": "Excavation"
                })
                self.assertEqual(r3.status_code, 200)
                j3 = r3.get_json()
                self.assertTrue(j3["ok"])
                self.assertFalse(j3["drive_cleanup"]["ok"])
                self.assertEqual(j3["drive_cleanup"]["action"], "error")

        self.assertEqual(
            removed_items,
            [("hse_ptw_subtype", "Hot Work"), ("hse_ptw_subtype", "Confined Space"), ("hse_ptw_subtype", "Excavation")],
        )


if __name__ == "__main__":
    unittest.main()
