"""Offline unit and Node VM tests for DRIVE-AUTO-03:
Standardized Drive Folder Hierarchy, Unified Submit/Replied Cycle & Automatic Drive Folder Provisioning.

Covers all 5 Acceptance Criteria with zero live PostgreSQL or Google Drive network access:
1. Auto-provisioning on workspace, project, document type (with/without has_reply_cycle), and sub-tab creation,
   plus graceful no-op when drive_folder_id is empty or Drive credentials are unconfigured.
2. Safe in-place Drive folder rename (PATCH files/{id}) for document types and sub-tabs without duplicates,
   and strict preservation of Drive folders/files on DCR delete.
3. Frontend (dtHasReplyCycle) & backend (doc_type_has_reply_cycle / doc_type_allows_approval_slot) parity
   across ENG, HSE, O&M, and custom workspaces (CVI, SI, EI, PCQ, SVP, PTW, and custom tabs),
   including removal of the hardcoded ws_id != "ENG" block in drive_uploads.py.
4. POST /api/projects/<pid>/init-drive-folders permission enforcement (CAN_EDIT) and full idempotent provisioning.
5. Generalized subtype folders, concurrency locking (_get_subfolder_lock), segment sanitization,
   UI controls (#dt-has-reply-cycle, #btn-init-drive-folders), and EN/AR bilingual parity.
"""

from concurrent.futures import ThreadPoolExecutor
import copy
import json
import os
from pathlib import Path
import re
import subprocess
import tempfile
import threading
import unittest
from unittest.mock import patch

os.environ.setdefault("DATABASE_URL", "")
os.environ.setdefault("SECRET_KEY", "synthetic-test-secret-key-drive-auto-03")
os.environ.setdefault("MAGIC_SECRET", "synthetic-test-magic-secret-drive-auto-03")
os.environ.setdefault("GEMINI_API_KEY", "synthetic-test-gemini-key-drive-auto-03")
os.environ.setdefault("ALLOW_INSECURE_COOKIES_FOR_LOCAL_DEV", "1")

import app as app_mod
import attachments
import blueprints.drive_uploads as du_mod
import db
import drive_routing
import drive_upload
import html_render
import i18n


class ThreadSafeFakeDrive:
    """Thread-safe in-memory Google Drive v3 fake supporting list, get, create, update (PATCH), and delete/trash tracking."""

    def __init__(self, tree=None):
        self._lock = threading.Lock()
        self.items = {}
        self.calls = {"list": 0, "get": 0, "create": 0, "update": 0, "delete": 0, "trash": 0}
        self._seq = 0
        if tree:
            self._load(tree, None)

    def _load(self, tree, parent):
        for name, sub in tree.items():
            fid = "id_" + re.sub(r"[^A-Za-z0-9_-]+", "_", name)
            self.items[fid] = {
                "id": fid,
                "name": name,
                "mimeType": drive_routing.FOLDER_MIME,
                "parent": parent,
                "trashed": False,
            }
            if isinstance(sub, dict):
                self._load(sub, fid)

    def id_of(self, name):
        return "id_" + re.sub(r"[^A-Za-z0-9_-]+", "_", name)

    def children_of(self, parent_id):
        with self._lock:
            return [
                dict(it)
                for it in self.items.values()
                if it["parent"] == parent_id and not it["trashed"]
            ]

    def child_by_name(self, parent_id, name):
        low = str(name or "").strip().lower()
        for ch in self.children_of(parent_id):
            if ch["name"].strip().lower() == low:
                return ch
        return None

    def files(self):
        outer = self

        class _Exec:
            def __init__(self, fn):
                self._fn = fn

            def execute(self):
                return self._fn()

        class _Files:
            def list(self, q, fields=None, pageSize=None, pageToken=None, **kw):
                def _run():
                    assert kw.get("supportsAllDrives") is True
                    pid = q.split("'")[1]
                    only_folders = drive_routing.FOLDER_MIME in q
                    with outer._lock:
                        outer.calls["list"] += 1
                        res = [
                            {"id": it["id"], "name": it["name"], "mimeType": it["mimeType"]}
                            for it in outer.items.values()
                            if it["parent"] == pid and not it["trashed"] and (not only_folders or it["mimeType"] == drive_routing.FOLDER_MIME)
                        ]
                    return {"files": res}

                return _Exec(_run)

            def get(self, fileId, fields=None, **kw):
                def _run():
                    assert kw.get("supportsAllDrives") is True
                    with outer._lock:
                        outer.calls["get"] += 1
                        it = outer.items.get(fileId)
                        if not it:
                            raise RuntimeError("404 not found")
                        return {"id": fileId, "name": it["name"], "mimeType": it["mimeType"], "trashed": it["trashed"]}

                return _Exec(_run)

            def create(self, body, fields=None, **kw):
                def _run():
                    assert kw.get("supportsAllDrives") is True
                    parent = (body.get("parents") or [None])[0]
                    with outer._lock:
                        outer.calls["create"] += 1
                        outer._seq += 1
                        fid = f"drv_fld_{outer._seq}"
                        outer.items[fid] = {
                            "id": fid,
                            "name": body["name"],
                            "mimeType": body.get("mimeType") or drive_routing.FOLDER_MIME,
                            "parent": parent,
                            "trashed": False,
                        }
                    return {"id": fid, "name": body["name"]}

                return _Exec(_run)

            def update(self, fileId, body, fields=None, **kw):
                def _run():
                    assert kw.get("supportsAllDrives") is True
                    with outer._lock:
                        it = outer.items.get(fileId)
                        if not it:
                            raise RuntimeError("404 not found")
                        if body.get("trashed"):
                            outer.calls["trash"] += 1
                            it["trashed"] = True
                        if "name" in body:
                            outer.calls["update"] += 1
                            it["name"] = body["name"]
                        return {"id": fileId, "name": it["name"], "trashed": it["trashed"]}

                return _Exec(_run)

            def delete(self, fileId, **kw):
                def _run():
                    with outer._lock:
                        outer.calls["delete"] += 1
                        outer.items.pop(fileId, None)
                    return {}

                return _Exec(_run)

        return _Files()


class FakeDriveRestHttp:
    """HTTP transport backed by ThreadSafeFakeDrive so RestDriveAdapter & drive_uploads endpoints run end-to-end offline."""

    def __init__(self, drive: ThreadSafeFakeDrive):
        self.drive = drive
        self.sessions = []

    class _Resp:
        def __init__(self, status_code, payload=None, headers=None):
            self.status_code = status_code
            self._payload = payload if payload is not None else {}
            self.headers = headers or {}
            self.text = json.dumps(self._payload)

        def json(self):
            return self._payload

    def get(self, url, params=None, headers=None, timeout=None):
        if "/drive/v3/files/" in url:
            fid = url.rsplit("/", 1)[-1].split("?", 1)[0]
            try:
                data = self.drive.files().get(fid, supportsAllDrives=True).execute()
                return self._Resp(200, data)
            except Exception:
                return self._Resp(404, {"error": {"message": "Not found"}})
        q = (params or {}).get("q", "")
        res = self.drive.files().list(q=q, supportsAllDrives=True).execute()
        return self._Resp(200, res)

    def post(self, url, params=None, json=None, headers=None, data=None, timeout=None):
        import json as _json
        body = json if json is not None else (_json.loads(data) if data else {})
        if "upload/drive/v3/files" in url:
            parent = (body.get("parents") or [""])[0]
            sess_url = f"https://fake-drive.local/upload/{len(self.sessions) + 1}"
            self.sessions.append({"name": body.get("name"), "parent": parent, "session_url": sess_url})
            return self._Resp(200, {}, headers={"Location": sess_url})
        res = self.drive.files().create(body=body, supportsAllDrives=True).execute()
        return self._Resp(200, res)

    def patch(self, url, params=None, json=None, headers=None, data=None, timeout=None):
        import json as _json
        fid = url.rsplit("/", 1)[-1].split("?", 1)[0]
        body = json if json is not None else (_json.loads(data) if data else {})
        try:
            res = self.drive.files().update(fileId=fid, body=body, supportsAllDrives=True).execute()
            return self._Resp(200, res)
        except Exception:
            return self._Resp(404, {"error": {"message": "Not found"}})

    def delete(self, url, params=None, headers=None, timeout=None):
        fid = url.rsplit("/", 1)[-1].split("?", 1)[0]
        self.drive.files().delete(fileId=fid, supportsAllDrives=True).execute()
        return self._Resp(204, {})


class InMemoryDcrStore:
    """In-memory project / doc_types / lists / columns store matching db.py interfaces."""

    def __init__(self):
        self.workspaces = {}
        self.projects = {}
        self.doc_types = {}      # pid -> {dt_id: dict}
        self.columns = {}        # (pid, dt_id) -> [col_dict]
        self.lists = {}          # pid -> {list_name: [{"value": ..., "meta": ...}]}

    def add_project(self, pid, name=None, code=None, workspace_id="ENG", drive_folder_id=""):
        self.projects[pid] = {
            "id": pid,
            "name": name or pid,
            "code": code or pid,
            "workspace_id": workspace_id,
            "drive_folder_id": drive_folder_id,
        }
        self.doc_types.setdefault(pid, {})
        self.lists.setdefault(pid, {})

    def add_dt(self, pid, dt_id, code, name, has_reply_cycle=None, sort_order=0, subtypes_list_name=None, subtypes=None):
        dt = {
            "id": dt_id,
            "project_id": pid,
            "code": code,
            "name": name,
            "sort_order": sort_order,
        }
        if isinstance(has_reply_cycle, bool):
            dt["has_reply_cycle"] = has_reply_cycle
        self.doc_types.setdefault(pid, {})[dt_id] = dt
        cols = [
            {"col_key": "docNo", "label": "Document No.", "col_type": "docno", "list_name": None, "visible": True, "sort_order": 0},
            {"col_key": "fileLocation", "label": "Drive Link", "col_type": "url", "list_name": None, "visible": True, "sort_order": 99},
        ]
        if subtypes_list_name:
            cols.insert(1, {
                "col_key": "legacySubtype",
                "label": "Subtype",
                "col_type": "dropdown",
                "list_name": subtypes_list_name,
                "visible": True,
                "sort_order": 1,
            })
            self.lists.setdefault(pid, {})[subtypes_list_name] = [{"value": s, "meta": None} for s in (subtypes or ())]
        self.columns[(pid, dt_id)] = cols

    def get_project(self, pid):
        p = self.projects.get(pid)
        return copy.deepcopy(p) if p else None

    def get_projects(self, project_ids=None, workspace_id=None):
        out = []
        for p in self.projects.values():
            if workspace_id and str(p.get("workspace_id") or "ENG").upper() != str(workspace_id).upper():
                continue
            if project_ids is not None and p["id"] not in project_ids:
                continue
            out.append(copy.deepcopy(p))
        return out

    def get_doc_types(self, pid):
        dts = list((self.doc_types.get(pid) or {}).values())
        dts.sort(key=lambda d: (d.get("sort_order", 0), d["id"]))
        return copy.deepcopy(dts)

    def get_doc_type(self, pid, dt_id):
        dt = (self.doc_types.get(pid) or {}).get(dt_id)
        return copy.deepcopy(dt) if dt else None

    def get_columns(self, pid, dt_id, visible_only=False):
        cols = self.columns.get((pid, dt_id), [])
        if visible_only:
            cols = [c for c in cols if c.get("visible", True)]
        return copy.deepcopy(cols)

    def get_lists(self, pid):
        raw = self.lists.get(pid) or {}
        return {ln: [it["value"] for it in items] for ln, items in raw.items()}

    def set_doc_type_drive_cache(self, pid, dt_id, patch):
        dt = (self.doc_types.get(pid) or {}).get(dt_id)
        if not dt:
            return False
        allowed = {
            "drive_folder_id",
            "drive_folder_name",
            "drive_slot_folders",
            "drive_direction_folders",
            "drive_subtype_folders",
            "folder_synced_at",
        }
        for k, v in (patch or {}).items():
            if k in allowed:
                dt[k] = copy.deepcopy(v)
        return True


class TestDriveAutoProvisioningAndRouting(unittest.TestCase):
    def setUp(self):
        drive_routing.invalidate_subfolder_caches()

    def test_01_graceful_skip_when_no_drive_folder_or_unconfigured_credentials(self):
        store = InMemoryDcrStore()
        store.add_project("PRJ-NO-DRV", workspace_id="ENG", drive_folder_id="")
        store.add_dt("PRJ-NO-DRV", "DS", "DS", "Document Submittal", has_reply_cycle=True)

        # 1. Empty drive_folder_id -> skipped cleanly without touching Drive
        res_no_folder = drive_routing.provision_project_drive_structure("PRJ-NO-DRV", service=None, store=store)
        self.assertEqual(res_no_folder["ok"], True)
        self.assertEqual(res_no_folder["skipped"], True)
        self.assertEqual(res_no_folder["reason"], "no_drive_folder")

        # 2. drive_folder_id set, but no Drive credentials available -> skipped cleanly without error
        store.add_project("PRJ-OFFLINE", workspace_id="ENG", drive_folder_id="ROOT_OFFLINE_1")
        store.add_dt("PRJ-OFFLINE", "DS", "DS", "Document Submittal", has_reply_cycle=True)
        with patch.object(drive_routing, "get_default_write_service", return_value=None):
            res_unconf = drive_routing.provision_project_drive_structure("PRJ-OFFLINE", store=store)
            self.assertEqual(res_unconf["ok"], True)
            self.assertEqual(res_unconf["skipped"], True)
            self.assertEqual(res_unconf["reason"], "drive_not_configured")

            ren_dt = drive_routing.rename_doc_type_drive_folder("PRJ-OFFLINE", "DS", "DS", "Old", "DS", "New", store=store)
            self.assertTrue(ren_dt["ok"] and ren_dt["skipped"])
            ren_sub = drive_routing.rename_subtype_drive_folder("PRJ-OFFLINE", "DS", "OldSub", "NewSub", store=store)
            self.assertTrue(ren_sub["ok"] and ren_sub["skipped"])

    def test_02_provision_full_project_tree_dual_flat_noc_ltr_and_subtypes_idempotently(self):
        drive = ThreadSafeFakeDrive({"ENG Root": {"05 - RFI": {"Reply": {}}}})
        store = InMemoryDcrStore()
        store.add_project(
            "PEM-100",
            workspace_id="ENG",
            drive_folder_id=f"https://drive.google.com/drive/folders/{drive.id_of('ENG Root')}?usp=sharing",
        )
        store.add_dt("PEM-100", "DS", "DS", "Document Submittal", has_reply_cycle=True, sort_order=0)
        store.add_dt("PEM-100", "CVI", "CVI", "Confirmation of Verbal Instruction", has_reply_cycle=True, sort_order=1)
        store.add_dt("PEM-100", "RFI", "RFI", "Request for Information", has_reply_cycle=True, sort_order=2)
        store.add_dt("PEM-100", "PR", "PR", "Purchase Requisition", has_reply_cycle=False, sort_order=3)
        store.add_dt("PEM-100", "NOC", "NOC", "Notice of Change", has_reply_cycle=True, sort_order=4)
        store.add_dt("PEM-100", "LTR", "LTR", "Letters", has_reply_cycle=False, sort_order=5)
        store.add_dt(
            "PEM-100", "SD", "SD", "Shop Drawing",
            has_reply_cycle=True, sort_order=6,
            subtypes_list_name="eng_sd_subtype",
            subtypes=["HVAC", "Electrical / Power"],
        )

        res1 = drive_routing.provision_project_drive_structure("PEM-100", service=drive, store=store)
        self.assertTrue(res1["ok"], res1)
        self.assertFalse(res1["skipped"])
        self.assertEqual(res1["errors"], [])

        root_id = drive.id_of("ENG Root")
        # DS -> Submit & Replied
        ds_fld = drive.child_by_name(root_id, "DS - Document Submittal")
        self.assertIsNotNone(ds_fld)
        self.assertIsNotNone(drive.child_by_name(ds_fld["id"], "Submit"))
        self.assertIsNotNone(drive.child_by_name(ds_fld["id"], "Replied"))

        # CVI -> Submit & Replied
        cvi_fld = drive.child_by_name(root_id, "CVI - Confirmation of Verbal Instruction")
        self.assertIsNotNone(cvi_fld)
        self.assertIsNotNone(drive.child_by_name(cvi_fld["id"], "Submit"))
        self.assertIsNotNone(drive.child_by_name(cvi_fld["id"], "Replied"))

        # RFI -> reuses pre-existing '05 - RFI' and 'Reply', creates only missing 'Submit' (no duplicate 'Replied')
        rfi_fld = drive.child_by_name(root_id, "05 - RFI")
        self.assertIsNotNone(rfi_fld)
        rfi_children = {c["name"] for c in drive.children_of(rfi_fld["id"])}
        self.assertEqual(rfi_children, {"Submit", "Reply"})

        # PR (has_reply_cycle=False) -> flat folder, NO Submit or Replied
        pr_fld = drive.child_by_name(root_id, "PR - Purchase Requisition")
        self.assertIsNotNone(pr_fld)
        self.assertEqual(drive.children_of(pr_fld["id"]), [])

        # NOC -> Part A, Part B, Part C, Part D
        noc_fld = drive.child_by_name(root_id, "NOC - Notice of Change")
        self.assertIsNotNone(noc_fld)
        self.assertEqual({c["name"] for c in drive.children_of(noc_fld["id"])}, {"Part A", "Part B", "Part C", "Part D"})

        # LTR -> Sent & Received
        ltr_fld = drive.child_by_name(root_id, "LTR - Letters")
        self.assertIsNotNone(ltr_fld)
        self.assertEqual({c["name"] for c in drive.children_of(ltr_fld["id"])}, {"Sent", "Received"})

        # SD with subtypes -> HVAC/{Submit,Replied} and Electrical - Power/{Submit,Replied}
        sd_fld = drive.child_by_name(root_id, "SD - Shop Drawing")
        self.assertIsNotNone(sd_fld)
        hvac_fld = drive.child_by_name(sd_fld["id"], "HVAC")
        elec_fld = drive.child_by_name(sd_fld["id"], "Electrical - Power")
        self.assertIsNotNone(hvac_fld)
        self.assertIsNotNone(elec_fld)
        self.assertEqual({c["name"] for c in drive.children_of(hvac_fld["id"])}, {"Submit", "Replied"})
        self.assertEqual({c["name"] for c in drive.children_of(elec_fld["id"])}, {"Submit", "Replied"})

        # Second run is 100% idempotent (0 created, all existing)
        creates_after_first = drive.calls["create"]
        res2 = drive_routing.provision_project_drive_structure("PEM-100", service=drive, store=store)
        self.assertTrue(res2["ok"])
        self.assertEqual(res2["created"], 0)
        self.assertGreater(res2["existing"], 0)
        self.assertEqual(drive.calls["create"], creates_after_first)

    def test_03_generalized_subtypes_sanitization_and_parallel_concurrency_locks(self):
        drive = ThreadSafeFakeDrive({"Root": {}})
        store = InMemoryDcrStore()
        store.add_project("HSE-1", workspace_id="HSE", drive_folder_id=drive.id_of("Root"))
        store.add_dt("HSE-1", "OBS", "OBS", "Safety / Observations", has_reply_cycle=True)
        store.add_dt("HSE-1", "FEI", "FEI", "Fire & Equipment", has_reply_cycle=False)

        # Parallel uploads to OBS / Unsafe / Condition / DOC../001 must create exactly ONE of each folder
        def _worker(idx):
            slot = "submission" if idx % 2 == 0 else "approval"
            return drive_routing.resolve_upload_destination(
                "HSE-1", "OBS", slot, drive, store,
                create_if_missing=True,
                flat=False,
                subtype="Unsafe / Condition..",
                subfolder_name="OBS-001/REV00",
            )

        with ThreadPoolExecutor(max_workers=8) as pool:
            results = list(pool.map(_worker, range(12)))

        root_children = drive.children_of(drive.id_of("Root"))
        self.assertEqual(len(root_children), 1)
        self.assertEqual(root_children[0]["name"], "OBS - Safety - Observations")

        obs_children = drive.children_of(root_children[0]["id"])
        self.assertEqual(len(obs_children), 1)
        self.assertEqual(obs_children[0]["name"], "Unsafe - Condition")

        slot_children = drive.children_of(obs_children[0]["id"])
        self.assertEqual({c["name"] for c in slot_children}, {"Submit", "Replied"})
        for slot_ch in slot_children:
            doc_children = drive.children_of(slot_ch["id"])
            self.assertEqual(len(doc_children), 1)
            self.assertEqual(doc_children[0]["name"], "OBS-001 - REV00")
        self.assertEqual(len({r["doc_type_folder_id"] for r in results}), 1)

        # Flat doc type (FEI, has_reply_cycle=False) with subtype routes directly to <FEI>/<Subtype>/ (no Submit/Replied)
        r_fei = drive_routing.resolve_upload_destination(
            "HSE-1", "FEI", "submission", drive, store,
            create_if_missing=True,
            flat=True,
            subtype="CO2 / Extinguisher",
        )
        fei_sub = drive.items[r_fei["folder_id"]]
        self.assertEqual(fei_sub["name"], "CO2 - Extinguisher")
        self.assertEqual(drive.children_of(r_fei["folder_id"]), [])


class TestAutoTriggersRenameInitAndPermissions(unittest.TestCase):
    def setUp(self):
        drive_routing.invalidate_subfolder_caches()
        self.client = app_mod.app.test_client()
        self.client.set_cookie("dcr_token", "tok_test")

    def test_04_crud_auto_triggers_workspace_project_doctype_subtab_rename_and_safe_delete(self):
        drive = ThreadSafeFakeDrive({
            "WS Root": {},
            "Proj Root": {},
        })
        http = FakeDriveRestHttp(drive)
        svc = drive_upload.RestDriveAdapter("FAKE_TOKEN", http=http)

        # Shared state backing mocked db calls for Flask route tests
        proj_state = {
            "WS-P1": {"id": "WS-P1", "name": "QA Project", "code": "QAP1", "workspace_id": "QAQC", "drive_folder_id": drive.id_of("WS Root")},
            "ENG-P1": {"id": "ENG-P1", "name": "Eng Project", "code": "EP1", "workspace_id": "ENG", "drive_folder_id": ""},
        }
        dt_state = {
            "WS-P1": {
                "QIR": {"id": "QIR", "project_id": "WS-P1", "code": "QIR", "name": "Quality Inspection", "has_reply_cycle": True, "sort_order": 0},
            },
            "ENG-P1": {
                "DS": {"id": "DS", "project_id": "ENG-P1", "code": "DS", "name": "Document Submittal", "has_reply_cycle": True, "sort_order": 0},
            },
        }
        lists_state = {"ENG-P1": {}, "WS-P1": {}}
        cols_state = {"ENG-P1": {}, "WS-P1": {}}

        super_user = {"username": "admin", "role": "superadmin"}

        def fake_get_project(pid):
            p = proj_state.get(pid)
            return copy.deepcopy(p) if p else None

        def fake_get_projects(project_ids=None, workspace_id=None):
            out = []
            for p in proj_state.values():
                if workspace_id and p["workspace_id"].upper() != str(workspace_id).upper():
                    continue
                if project_ids is not None and p["id"] not in project_ids:
                    continue
                out.append(copy.deepcopy(p))
            return out

        def fake_save_project(pid, name, code, extra, workspace_id=None):
            proj_state[pid] = {
                "id": pid,
                "name": name,
                "code": code,
                "workspace_id": workspace_id or proj_state.get(pid, {}).get("workspace_id", "ENG"),
                **copy.deepcopy(extra or {}),
            }

        def fake_get_doc_types(pid):
            return copy.deepcopy(list((dt_state.get(pid) or {}).values()))

        def fake_get_doc_type(pid, dt_id):
            d = (dt_state.get(pid) or {}).get(dt_id)
            return copy.deepcopy(d) if d else None

        def fake_add_doc_type(pid, code, name, override=None, has_reply_cycle=None):
            dt_state.setdefault(pid, {})[code] = {
                "id": code,
                "project_id": pid,
                "code": code,
                "name": name,
                "sort_order": len(dt_state.get(pid, {})),
                "has_reply_cycle": True if has_reply_cycle is None else bool(has_reply_cycle),
                "expected_reply_override": override or {},
            }

        def fake_save_dt_meta(pid, dt_id, expected_reply_override=None, has_reply_cycle=None):
            d = (dt_state.get(pid) or {}).get(dt_id)
            if d:
                if expected_reply_override is not None:
                    d["expected_reply_override"] = expected_reply_override
                if isinstance(has_reply_cycle, bool):
                    d["has_reply_cycle"] = has_reply_cycle

        def fake_set_cache(pid, dt_id, patch_d):
            d = (dt_state.get(pid) or {}).get(dt_id)
            if d:
                d.update(copy.deepcopy(patch_d))
            return True

        def fake_get_lists(pid):
            return copy.deepcopy(lists_state.get(pid) or {})

        def fake_get_cols(pid, dt_id, visible_only=False):
            return copy.deepcopy((cols_state.get(pid) or {}).get(dt_id) or [])

        with patch("db.get_session", return_value=super_user), \
             patch.object(drive_routing, "_service_provider", return_value=svc), \
             patch("db.get_project", side_effect=fake_get_project), \
             patch("db.get_projects", side_effect=fake_get_projects), \
             patch("db.save_project", side_effect=fake_save_project), \
             patch("db.get_doc_types", side_effect=fake_get_doc_types), \
             patch("db.get_doc_type", side_effect=fake_get_doc_type), \
             patch("db.add_doc_type", side_effect=fake_add_doc_type), \
             patch("db.save_doc_type_metadata", side_effect=fake_save_dt_meta), \
             patch("db.set_doc_type_drive_cache", side_effect=fake_set_cache), \
             patch("db.get_lists", side_effect=fake_get_lists), \
             patch("db.get_columns", side_effect=fake_get_cols), \
             patch("db.clear_project_drive_caches"), \
             patch("db.project_has_doc_type", return_value=False), \
             patch("db.log_action"):

            # 1) New Workspace auto-provisions linked project with drive_folder_id
            with patch("db.create_workspace"), patch("db.q", return_value={"m": 8}):
                r_ws = self.client.post("/api/workspaces", json={
                    "action": "add",
                    "code": "QAQC",
                    "name": "Quality Assurance",
                    "name_ar": "الجودة",
                })
                self.assertEqual(r_ws.status_code, 200, r_ws.get_json())
                qir_fld = drive.child_by_name(drive.id_of("WS Root"), "QIR - Quality Inspection")
                self.assertIsNotNone(qir_fld)
                self.assertEqual({c["name"] for c in drive.children_of(qir_fld["id"])}, {"Submit", "Replied"})

            # 2a) Saving project WITHOUT drive_folder_id gracefully skips without error
            r_save_empty = self.client.post("/api/project/ENG-P1", json={
                "name": "Eng Project", "code": "EP1", "drive_folder_id": ""
            })
            self.assertEqual(r_save_empty.status_code, 200)
            self.assertTrue(r_save_empty.get_json()["drive_provision"]["skipped"])

            # 2b) Saving project WITH drive_folder_id auto-provisions its doc types
            r_save_drv = self.client.post("/api/project/ENG-P1", json={
                "name": "Eng Project", "code": "EP1", "drive_folder_id": drive.id_of("Proj Root")
            })
            self.assertEqual(r_save_drv.status_code, 200)
            self.assertGreater(r_save_drv.get_json()["drive_provision"]["created"], 0)
            ds_fld = drive.child_by_name(drive.id_of("Proj Root"), "DS - Document Submittal")
            self.assertIsNotNone(ds_fld)
            self.assertEqual({c["name"] for c in drive.children_of(ds_fld["id"])}, {"Submit", "Replied"})

            # 3) Adding custom tabs: one WITH reply cycle (TQ -> Submit & Replied), one WITHOUT (LOG -> flat)
            with patch("db.q", return_value=None):
                r_dt_dual = self.client.post("/api/doc_types/ENG-P1", json={
                    "code": "TQ", "name": "Technical Query", "has_reply_cycle": True
                })
                self.assertEqual(r_dt_dual.status_code, 200)
                tq_fld = drive.child_by_name(drive.id_of("Proj Root"), "TQ - Technical Query")
                self.assertIsNotNone(tq_fld)
                self.assertEqual({c["name"] for c in drive.children_of(tq_fld["id"])}, {"Submit", "Replied"})

                r_dt_flat = self.client.post("/api/doc_types/ENG-P1", json={
                    "code": "LOG", "name": "Daily Site Log", "has_reply_cycle": False
                })
                self.assertEqual(r_dt_flat.status_code, 200)
                log_fld = drive.child_by_name(drive.id_of("Proj Root"), "LOG - Daily Site Log")
                self.assertIsNotNone(log_fld)
                self.assertEqual(drive.children_of(log_fld["id"]), [])

            # 4) Rename custom tab ('LOG' -> 'SLOG - Site Logbook') & enable has_reply_cycle=True:
            #    renames existing folder via PATCH in-place (no duplicate!) and provisions Submit & Replied inside it
            def fake_q_rename(sql, params=(), one=False):
                s = " ".join(sql.split())
                if "id<>%s" in s:
                    return None
                if "FROM doc_types WHERE id=%s" in s:
                    return copy.deepcopy(dt_state["ENG-P1"]["LOG"])
                return None

            def fake_exe_rename(sql, params=()):
                if "UPDATE doc_types SET code=%s, name=%s" in sql:
                    nc, nn, did, pid_p = params
                    dt_state[pid_p][did]["code"] = nc
                    dt_state[pid_p][did]["name"] = nn
                return 1

            updates_before = drive.calls["update"]
            with patch("db.q", side_effect=fake_q_rename), patch("db.exe", side_effect=fake_exe_rename):
                r_ren_dt = self.client.patch("/api/doc_types/ENG-P1/LOG", json={
                    "code": "SLOG", "name": "Site Logbook", "has_reply_cycle": True
                })
                self.assertEqual(r_ren_dt.status_code, 200)
                body_ren = r_ren_dt.get_json()
                self.assertTrue(body_ren["drive_rename"]["renamed"])
                self.assertEqual(drive.calls["update"], updates_before + 1)
                # Same folder ID was renamed in place — no old 'LOG - Daily Site Log' duplicate
                self.assertEqual(drive.items[log_fld["id"]]["name"], "SLOG - Site Logbook")
                self.assertIsNone(drive.child_by_name(drive.id_of("Proj Root"), "LOG - Daily Site Log"))
                self.assertEqual({c["name"] for c in drive.children_of(log_fld["id"])}, {"Submit", "Replied"})

            # 5) Add Sub-Tab ('Civil Works') to TQ -> provisions TQ/<Civil Works>/{Submit, Replied}
            def fake_add_list_item(pid, ln, item, meta=None):
                lists_state.setdefault(pid, {}).setdefault(ln, []).append(item)
                return {"item": item, "meta": meta}

            with patch("db.add_list_item", side_effect=fake_add_list_item), \
                 patch("db.ensure_doc_type_subtype_column", return_value=True):
                cols_state["ENG-P1"]["TQ"] = [{
                    "col_key": "legacySubtype", "label": "Subtype",
                    "col_type": "dropdown", "list_name": "eng_tq_subtype", "visible": True,
                }]
                r_sub_add = self.client.post("/api/lists/ENG-P1", json={
                    "list_name": "eng_tq_subtype", "item": "Civil Works", "dt_id": "TQ"
                })
                self.assertEqual(r_sub_add.status_code, 200)
                civil_fld = drive.child_by_name(tq_fld["id"], "Civil Works")
                self.assertIsNotNone(civil_fld)
                self.assertEqual({c["name"] for c in drive.children_of(civil_fld["id"])}, {"Submit", "Replied"})

            # 6) Rename Sub-Tab ('Civil Works' -> 'Structural Works') -> renames in-place via PATCH without duplicates
            def fake_rename_list_item(pid, ln, old_i, new_i, meta=None):
                arr = lists_state.get(pid, {}).get(ln, [])
                if old_i in arr:
                    arr[arr.index(old_i)] = new_i
                return 1

            updates_before_sub = drive.calls["update"]
            with patch("db.rename_list_item", side_effect=fake_rename_list_item), \
                 patch("db.resolve_doc_types_for_subtype_list", return_value=["TQ"]):
                r_sub_ren = self.client.patch("/api/lists/ENG-P1", json={
                    "list_name": "eng_tq_subtype",
                    "old_item": "Civil Works",
                    "new_item": "Structural Works",
                })
                self.assertEqual(r_sub_ren.status_code, 200)
                self.assertTrue(r_sub_ren.get_json()["drive_rename"]["renamed"])
                self.assertEqual(drive.calls["update"], updates_before_sub + 1)
                self.assertEqual(drive.items[civil_fld["id"]]["name"], "Structural Works")
                self.assertIsNone(drive.child_by_name(tq_fld["id"], "Civil Works"))
                self.assertEqual(
                    dt_state["ENG-P1"]["TQ"]["drive_subtype_folders"]["structural works"]["id"],
                    civil_fld["id"],
                )

            # 7) Deleting an empty sub-tab and empty tab trashes their empty Drive folders without hard-deleting
            folder_count_before = len(drive.items)
            with patch("db.remove_list_item"), patch("db.delete_doc_type"):
                r_del_sub = self.client.delete("/api/lists/ENG-P1", json={
                    "list_name": "eng_tq_subtype", "item": "Structural Works"
                })
                self.assertEqual(r_del_sub.status_code, 200)
                self.assertEqual(r_del_sub.get_json()["drive_cleanup"]["action"], "trashed")
                r_del_dt = self.client.delete("/api/doc_types/ENG-P1/TQ")
                self.assertEqual(r_del_dt.status_code, 200)
                self.assertEqual(r_del_dt.get_json()["drive_cleanup"]["action"], "trashed")

            self.assertEqual(drive.calls["delete"], 0)
            self.assertEqual(len(drive.items), folder_count_before)
            self.assertTrue(drive.items[tq_fld["id"]]["trashed"])
            self.assertTrue(drive.items[civil_fld["id"]]["trashed"])

    def test_05_init_drive_folders_endpoint_permissions_and_provisioning(self):
        drive = ThreadSafeFakeDrive({"Init Root": {}})
        http = FakeDriveRestHttp(drive)
        svc = drive_upload.RestDriveAdapter("FAKE_TOKEN", http=http)

        proj = {"id": "PRJ-INIT", "name": "Init Proj", "code": "PINIT", "workspace_id": "ENG", "drive_folder_id": drive.id_of("Init Root")}
        dts = {
            "SI": {"id": "SI", "project_id": "PRJ-INIT", "code": "SI", "name": "Site Instructions", "has_reply_cycle": True, "sort_order": 0},
            "MOM": {"id": "MOM", "project_id": "PRJ-INIT", "code": "MOM", "name": "Minutes of Meetings", "has_reply_cycle": False, "sort_order": 1},
        }

        # 1. Viewer (can_edit=False) is rejected with 403 Forbidden
        viewer = {"username": "viewer1", "role": "viewer"}
        with patch("db.get_session", return_value=viewer), \
             patch("app.can_view_project", return_value=True), \
             patch("blueprints.projects.can_edit", return_value=False), \
             patch("db.log_action"):
            r_forbidden = self.client.post("/api/projects/PRJ-INIT/init-drive-folders", json={})
            self.assertEqual(r_forbidden.status_code, 403)
            self.assertEqual(drive.calls["create"], 0)

        # 2. Editor/Admin (can_edit=True) on project without drive_folder_id -> 400
        editor = {"username": "editor1", "role": "editor"}
        with patch("db.get_session", return_value=editor), \
             patch("app.can_view_project", return_value=True), \
             patch("blueprints.projects.can_edit", return_value=True), \
             patch("db.get_project", return_value={**proj, "drive_folder_id": ""}):
            r_missing_id = self.client.post("/api/projects/PRJ-INIT/init-drive-folders", json={})
            self.assertEqual(r_missing_id.status_code, 400)
            self.assertEqual(r_missing_id.get_json()["reason"], "no_drive_folder")

        # 3. Editor/Admin on configured project provisions missing folders and returns accurate counts
        with patch("db.get_session", return_value=editor), \
             patch("app.can_view_project", return_value=True), \
             patch("blueprints.projects.can_edit", return_value=True), \
             patch.object(drive_routing, "_service_provider", return_value=svc), \
             patch("db.get_project", return_value=proj), \
             patch("db.get_doc_types", side_effect=lambda _pid: copy.deepcopy(list(dts.values()))), \
             patch("db.get_doc_type", side_effect=lambda _pid, k: copy.deepcopy(dts.get(k))), \
             patch("db.get_lists", return_value={}), \
             patch("db.get_columns", return_value=[]), \
             patch("db.set_doc_type_drive_cache", side_effect=lambda _p, k, patch_d: dts[k].update(patch_d) or True), \
             patch("db.log_action"):
            r_ok = self.client.post("/api/projects/PRJ-INIT/init-drive-folders", json={})
            self.assertEqual(r_ok.status_code, 200, r_ok.get_json())
            data1 = r_ok.get_json()
            self.assertTrue(data1["ok"])
            # SI + Submit + Replied (3) + MOM (1) = 4 created
            self.assertEqual(data1["created"], 4)
            self.assertEqual(data1["existing"], 0)
            self.assertEqual(data1["errors"], [])

            # Re-running immediately creates 0 and verifies 4 existing
            r_ok2 = self.client.post("/api/projects/PRJ-INIT/init-drive-folders", json={})
            self.assertEqual(r_ok2.status_code, 200)
            data2 = r_ok2.get_json()
            self.assertEqual(data2["created"], 0)
            self.assertEqual(data2["existing"], 4)

    def test_06_custom_workspace_allows_both_submission_and_approval_uploads(self):
        drive = ThreadSafeFakeDrive({"Custom WS Root": {}})
        http = FakeDriveRestHttp(drive)
        editor = {"username": "qa_editor", "role": "editor"}
        proj = {"id": "QA-01", "code": "QA01", "name": "QA Dept Project", "workspace_id": "QAQC", "drive_folder_id": drive.id_of("Custom WS Root")}
        dt_dual = {"id": "NCRQ", "code": "NCRQ", "name": "Quality Non-Conformance", "has_reply_cycle": True}
        dt_flat = {"id": "MEMO", "code": "MEMO", "name": "Internal Quality Note", "has_reply_cycle": False}
        rec = {"_id": "rec-qa-1", "_project_id": "QA-01", "_dt_id": "NCRQ", "docNo": "QA-NCRQ-001"}

        with patch("db.get_session", return_value=editor), \
             patch("app.can_view_project", return_value=True), \
             patch.object(du_mod, "current_user", return_value=editor), \
             patch.object(du_mod, "can_edit", return_value=True), \
             patch.object(du_mod, "can_access_doc_type", return_value=True), \
             patch.object(du_mod, "_token_provider", return_value="FAKE_TOKEN"), \
             patch.object(du_mod, "_http_client", http), \
             patch("db.get_project", return_value=proj), \
             patch("db.get_record_by_id", return_value=rec), \
             patch("db.set_doc_type_drive_cache", side_effect=lambda _p, _d, patch_d: dt_dual.update(patch_d) or True):

            # Custom workspace + has_reply_cycle=True -> BOTH submission (Submit) and approval (Replied) succeed!
            with patch("db.get_doc_type", side_effect=lambda _p, _d: copy.deepcopy(dt_dual)):
                r_sub = self.client.post("/api/drive/upload-session", json={
                    "project_id": "QA-01", "dt_id": "NCRQ", "record_id": "rec-qa-1",
                    "slot": "submission", "file_name": "ncr_sub.pdf", "file_size": 512,
                })
                self.assertEqual(r_sub.status_code, 200, r_sub.get_json())
                self.assertEqual(drive.items[r_sub.get_json()["folder_id"]]["name"], "Submit")

                r_app = self.client.post("/api/drive/upload-session", json={
                    "project_id": "QA-01", "dt_id": "NCRQ", "record_id": "rec-qa-1",
                    "slot": "approval", "file_name": "ncr_reply.pdf", "file_size": 512,
                })
                self.assertEqual(r_app.status_code, 200, r_app.get_json())
                self.assertEqual(drive.items[r_app.get_json()["folder_id"]]["name"], "Replied")

            # Custom workspace + has_reply_cycle=False -> submission routes flat, approval rejected (400)
            with patch("db.get_doc_type", return_value=dt_flat):
                r_flat_sub = self.client.post("/api/drive/upload-session", json={
                    "project_id": "QA-01", "dt_id": "MEMO",
                    "slot": "submission", "file_name": "memo.pdf", "file_size": 256,
                })
                self.assertEqual(r_flat_sub.status_code, 200)
                self.assertEqual(drive.items[r_flat_sub.get_json()["folder_id"]]["name"], "MEMO - Internal Quality Note")

                r_flat_app = self.client.post("/api/drive/upload-session", json={
                    "project_id": "QA-01", "dt_id": "MEMO",
                    "slot": "approval", "file_name": "memo_rep.pdf", "file_size": 256,
                })
                self.assertEqual(r_flat_app.status_code, 400)


class TestReplyCycleParityAndUiControls(unittest.TestCase):
    def _render_register(self, ws_id="ENG", lang="en"):
        user = {"username": "tester", "role": "superadmin", "can_download": True}
        proj = {"id": "P-01", "code": "P01", "name": "Test Project", "workspace_id": ws_id, "drive_folder_id": "ROOT_1"}
        dts = [
            {"id": "CVI", "code": "CVI", "name": "Confirmation of Verbal Instruction", "has_reply_cycle": True},
            {"id": "PR", "code": "PR", "name": "Requisition", "has_reply_cycle": False},
        ]
        with patch("html_render.can_edit", return_value=True), \
             patch("html_render.is_project_admin", return_value=True), \
             patch("html_render.get_allowed_doc_type_filter", return_value=None), \
             patch("db.can_user_download", return_value=True), \
             patch("db.get_projects_for_user", return_value=[proj]), \
             patch("db.get_doc_types", return_value=dts), \
             patch("db.get_logo", return_value=None), \
             patch("db.get_user_workspaces", return_value=[{"id": ws_id, "code": ws_id, "name": ws_id, "name_ar": ws_id, "icon": "📁"}]):
            return html_render.render_register(user, proj, active_ws_id=ws_id, lang=lang)

    def test_07_frontend_and_backend_reply_cycle_parity_and_ui_controls_en_ar(self):
        # Verify bilingual dictionary parity and rendered HTML controls in EN & AR
        self.assertEqual(set(i18n.TRANSLATIONS["en"].keys()), set(i18n.TRANSLATIONS["ar"].keys()))
        self.assertEqual(i18n.t("reg.dt_has_reply_cycle", "en"), "Has Reply Cycle (Submit & Replied)")
        self.assertEqual(i18n.t("reg.dt_has_reply_cycle", "ar"), "دورة رد واعتماد (Submit & Replied)")
        self.assertEqual(i18n.t("reg.proj_init_drv_btn", "en"), "Initialize Drive Folders")
        self.assertEqual(i18n.t("reg.proj_init_drv_btn", "ar"), "تهيئة مجلدات Drive")

        html_en = self._render_register(ws_id="ENG", lang="en")
        html_ar = self._render_register(ws_id="ENG", lang="ar")
        self.assertIn('id="dt-has-reply-cycle"', html_en)
        self.assertIn("Has Reply Cycle (Submit & Replied)", html_en)
        self.assertIn("دورة رد واعتماد (Submit & Replied)", html_ar)
        self.assertIn("btn-init-drive-folders", html_en)
        self.assertIn("initProjectDriveFolders", html_en)

        # Build matrix of test cases in Python and verify 100% agreement with JS dtHasReplyCycle in Node VM
        workspaces = ["ENG", "HSE", "OM", "QAQC"]
        test_dts = [
            # Built-in reply-cycle registers highlighted in prompt
            {"code": "CVI", "name": "Confirmation of Verbal Instruction"},
            {"code": "SI", "name": "Site Instructions"},
            {"code": "EI", "name": "Engineer's Instruction"},
            {"code": "PCQ", "name": "Potential Change Questionnaire"},
            {"code": "SVP", "name": "Safety Violation with Penalty"},
            {"code": "PTW", "name": "Permit to Work"},
            {"code": "DS", "name": "Document Submittal"},
            {"code": "SD", "name": "Shop Drawing Submittal"},
            {"code": "MS", "name": "Material Submittal"},
            {"code": "IR", "name": "Inspection Request"},
            {"code": "MIR", "name": "Material Inspection Request"},
            {"code": "RFI", "name": "Request for Information"},
            {"code": "ABD", "name": "As Built Drawing"},
            {"code": "WN", "name": "Work Notification"},
            {"code": "NCR", "name": "Non-Conformance Report"},
            {"code": "NCN", "name": "Non-Conformance Notice"},
            {"code": "NOC", "name": "Notice of Change"},
            {"code": "OBS", "name": "Safety Observations"},
            {"code": "INC", "name": "Incidents"},
            {"code": "JSA", "name": "Job Safety Analysis"},
            # Built-in single-slot (flat) registers
            {"code": "PR", "name": "Requisition"},
            {"code": "MOM", "name": "Minutes of Meetings"},
            {"code": "IOM", "name": "Internal Office Memo"},
            {"code": "QS", "name": "Quantity Surveyor"},
            {"code": "LTR", "name": "Letters"},
            {"code": "TBT", "name": "Toolbox Talks"},
            {"code": "FEI", "name": "Fire and Equipment Inspections"},
            {"code": "CEV", "name": "Contractor Evaluations"},
            {"code": "EVAC", "name": "Evacuation Drills"},
            {"code": "MHR", "name": "Man-hours"},
            {"code": "WKR", "name": "Worker Register"},
            {"code": "WO", "name": "Work Orders"},
            {"code": "PPM", "name": "Preventive Maintenance"},
            {"code": "CM", "name": "Corrective Maintenance"},
            {"code": "MR", "name": "Maintenance Requests"},
            {"code": "SLA", "name": "Service Level Reports"},
            {"code": "COR", "name": "Correspondences"},
            # Custom tabs with explicit boolean flag
            {"code": "CUS_DUAL", "name": "Custom Vendor Package", "has_reply_cycle": True},
            {"code": "CUS_FLAT", "name": "Custom Archive Log", "has_reply_cycle": False},
            {"code": "INC", "name": "Custom Reply Incident", "has_reply_cycle": True},
        ]

        expected_matrix = []
        for ws in workspaces:
            for dt in test_dts:
                py_val = attachments.doc_type_has_reply_cycle(dt, workspace_id=ws)
                # Explicit check for CVI, SI, EI, PCQ, SVP, PTW across all workspaces
                if dt["code"] in ("CVI", "SI", "EI", "PCQ", "SVP", "PTW"):
                    self.assertTrue(py_val, f"Expected True for {dt['code']} in {ws}")
                if dt["code"] == "CUS_DUAL":
                    self.assertTrue(py_val, f"Expected True for CUS_DUAL in {ws}")
                if dt["code"] == "CUS_FLAT":
                    self.assertFalse(py_val, f"Expected False for CUS_FLAT in {ws}")
                expected_matrix.append({"ws": ws, "dt": dt, "expected": py_val})

        scripts = re.findall(r"<script(?:\s+[^>]*)?>([\s\S]*?)</script>", html_en, re.IGNORECASE)
        main_script = next(s for s in scripts if "function dtHasReplyCycle(" in s and "function saveDocType(" in s)

        node_harness = r"""
const fs = require('fs');
const vm = require('vm');
const assert = require('assert');

const mainJs = fs.readFileSync(process.argv[2], 'utf8');
const matrix = JSON.parse(fs.readFileSync(process.argv[3], 'utf8'));

function mkEl(tag) {
  const children = [];
  const el = {
    tagName: String(tag || 'DIV').toUpperCase(),
    id: '',
    value: '',
    checked: false,
    disabled: false,
    textContent: '',
    style: { cssText: '', display: '' },
    dataset: {},
    children,
    classList: { add(){}, remove(){}, toggle(){}, contains(){ return false; } },
    appendChild(c) { children.push(c); c.parentElement = el; return c; },
    insertBefore(c) { children.push(c); c.parentElement = el; return c; },
    remove() {},
    addEventListener() {},
    querySelector() { return null; },
    querySelectorAll() { return []; },
  };
  Object.defineProperty(el, 'innerHTML', {
    get() { return this._html || ''; },
    set(v) { this._html = String(v || ''); },
  });
  return el;
}

const byId = {};
[
  'tabsbar', 'ptw-subtabs', 'toolbar', 'btn-add-subtab', 'thead', 'tbody', 'srchbox',
  'dt-modal', 'dt-modal-title', 'dt-save-btn', 'dt-edit-id', 'dt-code', 'dt-name',
  'dt-has-reply-cycle', 'dt-er-use', 'dt-er-rev0', 'dt-er-rev', 'dt-er-fields',
  'proj-body', 'proj-modal', 'proj-drive-id', 'proj-drive-init-status',
  's-clock', 's-total', 's-show', 's-ov', 'empty'
].forEach(id => { const e = mkEl('div'); e.id = id; byId[id] = e; });

const toasts = [];
const apiCalls = [];
const fetchCalls = [];

const documentStub = {
  documentElement: { dir: 'ltr', lang: 'en', getAttribute: () => 'en' },
  body: mkEl('body'),
  createElement: (tag) => mkEl(tag),
  getElementById: (id) => {
    if (!byId[id]) { const e = mkEl('div'); e.id = id; byId[id] = e; }
    return byId[id];
  },
  querySelector: () => null,
  querySelectorAll: () => [],
  addEventListener: () => {},
};

const ctx = {
  console, Math, Date, RegExp, Map, Set, Array, Object, String, Number, Boolean, JSON, Promise, Error,
  encodeURIComponent, decodeURIComponent, parseInt, parseFloat, isNaN,
  setTimeout: () => 0, setInterval: () => 0, requestAnimationFrame: () => 0,
  URLSearchParams: function() { return { get: () => null }; },
  URL: { createObjectURL: () => '' },
  location: { search: '', reload: () => {} },
  document: documentStub,
  window: {
    DCR_LANG: 'en',
    t: (k, params) => {
      if (k === 'reg.proj_init_drv_ok') return `Ready: ${params.created} created, ${params.existing} verified`;
      if (k === 'reg.dt_added_drive') return `Added (${params.created} created)`;
      return k;
    },
    tCol: (_k, l) => l,
    tDt: (_c, n) => n,
    tStatus: (s) => s,
    matchMedia: () => ({ matches: false }),
    addEventListener: () => {},
  },
  escHtml: (s) => String(s || ''),
  confirm: () => true,
  apiFetch: async (url, opts) => {
    apiCalls.push({ url, opts: opts ? JSON.parse(opts.body || '{}') : null, method: opts?.method || 'GET' });
    if (String(url).startsWith('/api/project/')) return { id: 'P-01', code: 'P01', name: 'Test Project', drive_folder_id: 'ROOT_1' };
    if (String(url).startsWith('/api/doc_types/') && opts && opts.method === 'POST') {
      return { ok: true, drive_provision: { ok: true, skipped: false, created: 3, existing: 0, errors: [] } };
    }
    if (String(url).startsWith('/api/doc_types/')) return [{ id: 'CUS', code: 'CUS', name: 'Custom', has_reply_cycle: false }];
    if (String(url).startsWith('/api/counts/')) return {};
    if (String(url).startsWith('/api/records/')) return { records: [], columns: [], count: 0 };
    return {};
  },
  fetch: async (url, opts) => {
    fetchCalls.push({ url, body: opts?.body ? JSON.parse(opts.body) : null });
    return {
      ok: true,
      status: 200,
      blob: async () => null,
      json: async () => ({ ok: true, created: 12, existing: 5, errors: [] }),
    };
  },
};
ctx.window.document = documentStub;
ctx.globalThis = ctx;

const api = vm.runInNewContext(
  'function bindDirectionalInput(){} function bindSmartTextarea(){} function openM(){} function closeM(){} function initRz(){}\n' +
  'function toast(msg, kind){ toasts.push({ msg, kind }); }\n' +
  mainJs +
  '\n;({ dtHasReplyCycle, setDocTypeModalMode, saveDocType, editProject, initProjectDriveFolders });',
  Object.assign(ctx, { toasts })
);

(async () => {
  // 1. Check 100% parity between Python doc_type_has_reply_cycle and JS dtHasReplyCycle across all workspaces
  for (const item of matrix) {
    const actual = api.dtHasReplyCycle(item.dt, item.ws);
    assert.strictEqual(
      actual,
      item.expected,
      `Mismatch for ws=${item.ws} dt=${JSON.stringify(item.dt)}: JS=${actual} vs PY=${item.expected}`
    );
  }

  // 2. Test Add Document Type modal toggle + payload + Drive auto-provision toast
  api.setDocTypeModalMode(null);
  assert.strictEqual(byId['dt-has-reply-cycle'].checked, true, 'Default checked in ENG workspace');
  byId['dt-code'].value = 'CUSF';
  byId['dt-name'].value = 'Custom Flat Log';
  byId['dt-has-reply-cycle'].checked = false;
  await api.saveDocType();
  const postCall = apiCalls.find(c => c.method === 'POST' && c.url.includes('/api/doc_types/'));
  assert.ok(postCall, 'saveDocType must POST to /api/doc_types/');
  assert.strictEqual(postCall.opts.code, 'CUSF');
  assert.strictEqual(postCall.opts.has_reply_cycle, false);
  assert.ok(toasts.some(t => t.msg === 'Added (3 created)'), 'Must show Drive auto-provision toast');

  // 3. Test Edit Document Type modal loads stored has_reply_cycle boolean accurately
  api.setDocTypeModalMode({ id: 'PR', code: 'PR', name: 'Requisition', has_reply_cycle: false });
  assert.strictEqual(byId['dt-has-reply-cycle'].checked, false);
  api.setDocTypeModalMode({ id: 'CVI', code: 'CVI', name: 'CVI', has_reply_cycle: true });
  assert.strictEqual(byId['dt-has-reply-cycle'].checked, true);

  // 4. Test editProject renders btn-init-drive-folders and initProjectDriveFolders calls endpoint & updates status
  await api.editProject();
  const drvSectionHtml = byId['proj-body'].children.map(c => c.innerHTML).join('\n');
  assert.ok(drvSectionHtml.includes('btn-init-drive-folders'), 'Project Settings must render Initialize Drive Folders button');
  byId['proj-drive-id'].value = 'https://drive.google.com/drive/folders/FOLDER_XYZ_99?usp=sharing';
  const fakeBtn = mkEl('button');
  fakeBtn.innerHTML = '📁 Initialize Drive Folders';
  fetchCalls.length = 0;
  await api.initProjectDriveFolders(fakeBtn);
  assert.strictEqual(fetchCalls.length, 1);
  assert.strictEqual(fetchCalls[0].url, '/api/projects/P-01/init-drive-folders');
  assert.strictEqual(fetchCalls[0].body.drive_folder_id, 'FOLDER_XYZ_99');
  assert.strictEqual(byId['proj-drive-init-status'].textContent, 'Ready: 12 created, 5 verified');
  assert.strictEqual(fakeBtn.disabled, false);
})().catch(err => {
  console.error(err);
  process.exit(1);
});
"""

        with tempfile.TemporaryDirectory() as tmpdir:
            js_path = Path(tmpdir) / "main.js"
            mat_path = Path(tmpdir) / "matrix.json"
            har_path = Path(tmpdir) / "harness.js"
            js_path.write_text(main_script, encoding="utf-8")
            mat_path.write_text(json.dumps(expected_matrix), encoding="utf-8")
            har_path.write_text(node_harness, encoding="utf-8")
            proc = subprocess.run(
                ["node", str(har_path), str(js_path), str(mat_path)],
                capture_output=True,
                text=True,
                timeout=20,
            )
            self.assertEqual(
                proc.returncode,
                0,
                f"Node parity/UI test failed:\nSTDOUT:\n{proc.stdout}\nSTDERR:\n{proc.stderr}",
            )


class TestFastBackgroundProjectCreation(unittest.TestCase):
    """Regression coverage for DRIVE-PERF-I18N-FIX-06: fast project creation + background subfolder provisioning."""

    def test_post_create_project_provisions_root_sync_and_subfolders_in_background(self):
        drive = ThreadSafeFakeDrive({"MasterRoot": {}})
        master_root_id = drive.id_of("MasterRoot")

        projects = {}
        doc_types = {}

        def _create_project(pid, name, code, creator="", workspace_id="ENG", client="", **kw):
            projects[pid] = {
                "id": pid,
                "name": name,
                "code": code,
                "workspace_id": workspace_id,
                "client": client,
                "drive_folder_id": kw.get("drive_folder_id", ""),
            }
            doc_types[pid] = [
                {"id": "SD", "code": "SD", "name": "Shop Drawing", "has_reply_cycle": True, "drive_slot_folders": {}},
                {"id": "IR", "code": "IR", "name": "Inspection Request", "has_reply_cycle": True, "drive_slot_folders": {}},
            ]

        def _get_project(pid):
            return copy.deepcopy(projects.get(pid))

        def _update_project_settings(pid, patch_dict):
            if pid in projects:
                projects[pid].update(patch_dict)

        def _get_doc_types(pid):
            return copy.deepcopy(doc_types.get(pid, []))

        def _set_dt_cache(pid, dt_id, patch_dict):
            for dt in doc_types.get(pid, []):
                if dt["id"] == dt_id:
                    dt.update(patch_dict)

        service_calls = []

        class _ThreadLocalDriveWrapper:
            def __init__(self, inner):
                self.inner = inner
                self.thread_id = threading.get_ident()

            def files(self):
                return self.inner.files()

        def _make_service():
            wrapper = _ThreadLocalDriveWrapper(drive)
            service_calls.append(wrapper)
            return wrapper

        super_user = {"username": "admin", "role": "superadmin"}
        client = app_mod.app.test_client()
        client.set_cookie("dcr_token", "synthetic-test-token")
        ajax_headers = {"X-Requested-With": "XMLHttpRequest"}

        with patch.object(drive_routing, "_service_provider", _make_service), \
             patch("db.get_session", return_value=super_user), \
             patch("db.get_master_drive_root_id", return_value=master_root_id), \
             patch("db.create_project", side_effect=_create_project), \
             patch("db.get_project", side_effect=_get_project), \
             patch("db.update_project_settings", side_effect=_update_project_settings), \
             patch("db.get_workspace", return_value={"id": "ENG", "code": "ENG", "name": "Engineering", "drive_folder_id": ""}), \
             patch("db.update_workspace", return_value=None), \
             patch("db.get_doc_types", side_effect=_get_doc_types), \
             patch("db.get_lists", return_value={}), \
             patch("db.set_doc_type_drive_cache", side_effect=_set_dt_cache):
            resp = client.post(
                "/api/projects/create",
                headers=ajax_headers,
                json={
                    "id": "FAST01",
                    "name": "Fast Plant",
                    "code": "FP01",
                    "workspace_id": "ENG",
                    "drive_folder_id": "",
                },
            )
            self.assertEqual(resp.status_code, 200)
            payload = resp.get_json()
            self.assertTrue(payload["ok"])
            self.assertTrue(payload["drive_folder_id"])
            self.assertTrue(payload["drive_provisioning_background"])
            self.assertTrue(payload["drive_provision"]["drive_provisioning_background"])
            self.assertFalse(payload["drive_provision"]["skipped"])
            self.assertEqual(projects["FAST01"]["drive_folder_id"], payload["drive_folder_id"])
            self.assertTrue(projects["FAST01"].get("auto_provisioned_drive"))

            # Wait deterministically for the background daemon thread to complete
            drive_routing.wait_for_background_provisioning(timeout=5.0)

        # Must have built at least 2 distinct service instances: foreground thread + background thread
        self.assertGreaterEqual(len(service_calls), 2)
        self.assertNotEqual(service_calls[0].thread_id, service_calls[1].thread_id)
        self.assertIsNot(service_calls[0], service_calls[1])

        # Verify DocType + Submit/Replied subfolders were created in background
        proj_fid = payload["drive_folder_id"]
        sd_folder = drive.child_by_name(proj_fid, "SD - Shop Drawing")
        self.assertIsNotNone(sd_folder)
        self.assertIsNotNone(drive.child_by_name(sd_folder["id"], "Submit"))
        self.assertIsNotNone(drive.child_by_name(sd_folder["id"], "Replied"))
        ir_folder = drive.child_by_name(proj_fid, "IR - Inspection Request")
        self.assertIsNotNone(ir_folder)
        self.assertIsNotNone(drive.child_by_name(ir_folder["id"], "Submit"))
        self.assertIsNotNone(drive.child_by_name(ir_folder["id"], "Replied"))


if __name__ == "__main__":
    unittest.main()
