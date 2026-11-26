"""Offline unit and Node VM tests for DRIVE-MASTER-ROOT-05:
Zero-Touch Drive Hierarchy — Global Master Root Folder, Auto Workspace/Department Folders,
Auto Project Folders & Full Tree Provisioning.

Covers all 4 Acceptance Criteria + UI/i18n parity with zero live PostgreSQL or Drive network access:
1. Configuring Master Root ID (`Gas Chill Workspace Records`) + creating/initializing a Department
   automatically creates `<Master Root>/<WS_ID> - <Name>` on Drive and supports in-place rename.
2. Creating a Project inside a Department with blank `drive_folder_id` automatically creates
   `<Master Root>/<Department>/<Project>`, persists `drive_folder_id` + `auto_provisioned_drive=True`,
   provisions all DocTypes, Subtypes, and `Submit`/`Replied` folders, and supports in-place rename.
3. Existing projects with a manually configured `drive_folder_id` keep their existing folder unchanged
   and are not renamed on project rename.
4. `GET/POST /api/admin/drive-master-root` and `POST /api/admin/init-all-drive-folders` enforce Admin
   authorization (`403` for non-admins) and provision all departments, projects, and internal folders
   idempotently in one click.
5. Admin & Project Settings UI controls (`#admin-master-drive-root-id`, `#btn-init-all-drive`,
   `#proj-drive-id` auto placeholder) work in Node VM with 100% EN/AR i18n parity.
"""

from concurrent.futures import ThreadPoolExecutor
import copy
import json
import os
import re
import subprocess
import tempfile
import unittest
from unittest.mock import patch

os.environ.setdefault("DATABASE_URL", "")
os.environ.setdefault("SECRET_KEY", "synthetic-test-secret-key-drive-master-05")
os.environ.setdefault("MAGIC_SECRET", "synthetic-test-magic-secret-drive-master-05")
os.environ.setdefault("GEMINI_API_KEY", "synthetic-test-gemini-key-drive-master-05")
os.environ.setdefault("ALLOW_INSECURE_COOKIES_FOR_LOCAL_DEV", "1")

import app as app_mod
import db
import drive_routing
import html_render
import i18n
from test_drive_auto_provisioning import ThreadSafeFakeDrive, InMemoryDcrStore


class MasterRootInMemoryStore(InMemoryDcrStore):
    """Extends InMemoryDcrStore with app_settings + workspace CRUD + Master Drive Root helpers."""

    def __init__(self):
        super().__init__()
        self.settings = {}
        self.audit = []

    def get_setting(self, key, default=None):
        return copy.deepcopy(self.settings.get(key, default))

    def save_setting(self, key, value):
        self.settings[key] = copy.deepcopy(value)

    def get_master_drive_root_info(self):
        raw = self.settings.get(db.MASTER_DRIVE_ROOT_SETTING_KEY, "")
        db_val = drive_routing.extract_drive_folder_id(raw)
        if db_val:
            return {"master_root_id": db_val, "source": "db"}
        env_val = drive_routing.extract_drive_folder_id(os.environ.get("DRIVE_MASTER_ROOT_ID", "").strip())
        if env_val:
            return {"master_root_id": env_val, "source": "env"}
        return {"master_root_id": "", "source": "none"}

    def get_master_drive_root_id(self):
        return self.get_master_drive_root_info()["master_root_id"]

    def set_master_drive_root_id(self, folder_id):
        clean_id = drive_routing.extract_drive_folder_id(folder_id)
        self.settings[db.MASTER_DRIVE_ROOT_SETTING_KEY] = clean_id
        return clean_id

    def get_workspace_drive_folder_id(self, ws_id):
        ws_key = str(ws_id or "").strip().upper()
        mapping = self.settings.get(db.WORKSPACE_DRIVE_FOLDERS_SETTING_KEY) or {}
        return drive_routing.extract_drive_folder_id(mapping.get(ws_key) or "")

    def set_workspace_drive_folder_id(self, ws_id, folder_id):
        ws_key = str(ws_id or "").strip().upper()
        mapping = dict(self.settings.get(db.WORKSPACE_DRIVE_FOLDERS_SETTING_KEY) or {})
        clean_id = drive_routing.extract_drive_folder_id(folder_id)
        if clean_id:
            mapping[ws_key] = clean_id
        else:
            mapping.pop(ws_key, None)
        self.settings[db.WORKSPACE_DRIVE_FOLDERS_SETTING_KEY] = mapping
        return clean_id

    def get_workspaces(self, active_only=True):
        out = list(self.workspaces.values())
        if active_only:
            out = [w for w in out if w.get("is_active", True)]
        return copy.deepcopy(out)

    def get_workspace(self, ws_id):
        w = self.workspaces.get(str(ws_id or "").strip().upper())
        return copy.deepcopy(w) if w else None

    def create_workspace(self, ws_id, code, name, name_ar="", icon="📁", ws_type="general", sort_order=0, description=""):
        wid = str(ws_id or "").strip().upper()
        self.workspaces[wid] = {
            "id": wid,
            "code": str(code or wid).strip().upper(),
            "name": name,
            "name_ar": name_ar,
            "icon": icon,
            "type": ws_type,
            "sort_order": sort_order,
            "description": description,
            "is_active": True,
        }

    def update_workspace(self, ws_id, code=None, name=None, name_ar=None, icon=None, sort_order=None, is_active=None, description=None):
        wid = str(ws_id or "").strip().upper()
        ws = self.workspaces.get(wid)
        if not ws:
            return
        if code is not None:
            ws["code"] = code
        if name is not None:
            ws["name"] = name
        if name_ar is not None:
            ws["name_ar"] = name_ar
        if icon is not None:
            ws["icon"] = icon
        if sort_order is not None:
            ws["sort_order"] = sort_order
        if is_active is not None:
            ws["is_active"] = bool(is_active)
        if description is not None:
            ws["description"] = description

    def create_project(self, pid, name, code, creator=None, workspace_id="ENG", client=None, drive_folder_id=None):
        self.add_project(pid, name=name, code=code, workspace_id=workspace_id, drive_folder_id=drive_folder_id or "")
        if client:
            self.projects[pid]["client"] = client
        ws_up = str(workspace_id or "ENG").strip().upper()
        if ws_up == "HSE":
            self.add_dt(pid, "PTW", "PTW", "Permit to Work", has_reply_cycle=True,
                        subtypes_list_name="hse_ptw_subtype", subtypes=["Hot Work", "Confined Space"])
            self.add_dt(pid, "OBS", "OBS", "Safety Observation", has_reply_cycle=True)
        else:
            self.add_dt(pid, "DS", "DS", "Document Submittal", has_reply_cycle=True)
            self.add_dt(pid, "SD", "SD", "Shop Drawing Submittal", has_reply_cycle=True)

    def save_project(self, pid, name, code, extra, workspace_id=None):
        proj = self.projects.get(pid)
        if not proj:
            raise db.ProjectOwnershipError("missing")
        old_fid = drive_routing.extract_drive_folder_id(proj.get("drive_folder_id"))
        new_fid = drive_routing.extract_drive_folder_id(extra.get("drive_folder_id")) if "drive_folder_id" in extra else old_fid
        keep_auto = proj.get("auto_provisioned_drive") if (old_fid and new_fid == old_fid and "auto_provisioned_drive" not in extra) else extra.get("auto_provisioned_drive")
        proj["name"] = name
        proj["code"] = code
        for k, v in (extra or {}).items():
            if k not in ("id", "name", "code", "workspace_id"):
                proj[k] = v
        if keep_auto is not None:
            proj["auto_provisioned_drive"] = bool(keep_auto)

    def update_project_settings(self, pid, settings_dict):
        proj = self.projects.get(pid)
        if not proj:
            return False
        proj.update(settings_dict or {})
        return True

    def clear_project_drive_caches(self, pid):
        for dt in (self.doc_types.get(pid) or {}).values():
            dt.pop("drive_folder_id", None)
            dt.pop("drive_slot_folders", None)
            dt.pop("drive_subtype_folders", None)

    def log_action(self, *args, **kwargs):
        self.audit.append((args, kwargs))


class TestDriveMasterRoot(unittest.TestCase):
    def setUp(self):
        drive_routing.invalidate_subfolder_caches()
        self._orig_env = os.environ.pop("DRIVE_MASTER_ROOT_ID", None)
        self.drive = ThreadSafeFakeDrive({"Gas Chill Workspace Records": {}, "Legacy Manual Project": {}})
        self.master_root_id = self.drive.id_of("Gas Chill Workspace Records")
        self.legacy_root_id = self.drive.id_of("Legacy Manual Project")
        self.store = MasterRootInMemoryStore()
        self.store.create_workspace("ENG", "ENG", "Engineering", "الهندسة")
        self.store.create_workspace("HSE", "HSE", "Health & Safety", "السلامة والصحة المهنية")
        self.store.create_workspace("OM", "OM", "Operations & Maintenance", "التشغيل والصيانة")
        self.client = app_mod.app.test_client()
        self.client.set_cookie("dcr_token", "synthetic-test-token")
        self.ajax = {"X-Requested-With": "XMLHttpRequest"}

    def tearDown(self):
        drive_routing.invalidate_subfolder_caches()
        if self._orig_env is not None:
            os.environ["DRIVE_MASTER_ROOT_ID"] = self._orig_env
        else:
            os.environ.pop("DRIVE_MASTER_ROOT_ID", None)

    def _patch_db_with_store(self):
        return patch.multiple(
            db,
            get_setting=self.store.get_setting,
            save_setting=self.store.save_setting,
            get_master_drive_root_info=self.store.get_master_drive_root_info,
            get_master_drive_root_id=self.store.get_master_drive_root_id,
            set_master_drive_root_id=self.store.set_master_drive_root_id,
            get_workspace_drive_folder_id=self.store.get_workspace_drive_folder_id,
            set_workspace_drive_folder_id=self.store.set_workspace_drive_folder_id,
            get_workspaces=self.store.get_workspaces,
            get_workspace=self.store.get_workspace,
            create_workspace=self.store.create_workspace,
            update_workspace=self.store.update_workspace,
            get_project=self.store.get_project,
            get_projects=self.store.get_projects,
            create_project=self.store.create_project,
            save_project=self.store.save_project,
            update_project_settings=self.store.update_project_settings,
            clear_project_drive_caches=self.store.clear_project_drive_caches,
            get_doc_types=self.store.get_doc_types,
            get_doc_type=self.store.get_doc_type,
            get_columns=self.store.get_columns,
            get_lists=self.store.get_lists,
            set_doc_type_drive_cache=self.store.set_doc_type_drive_cache,
            log_action=self.store.log_action,
            resolve_project_alias=self.store.get_project,
        )

    # ── AC 1: Master Root Storage, Env Fallback, Workspace Auto-Provision & Rename ──
    def test_ac1_master_root_db_and_env_fallback(self):
        settings_mem = {}
        with patch.object(db, "get_setting", side_effect=lambda k, d=None: settings_mem.get(k, d)), \
             patch.object(db, "save_setting", side_effect=lambda k, v: settings_mem.__setitem__(k, v)):
            # Empty -> none
            self.assertEqual(db.get_master_drive_root_info(), {"master_root_id": "", "source": "none"})
            self.assertEqual(db.get_master_drive_root_id(), "")

            # Env fallback
            os.environ["DRIVE_MASTER_ROOT_ID"] = "https://drive.google.com/drive/folders/ENV_ROOT_123?usp=sharing"
            self.assertEqual(db.get_master_drive_root_info(), {"master_root_id": "ENV_ROOT_123", "source": "env"})
            self.assertEqual(db.get_master_drive_root_id(), "ENV_ROOT_123")

            # DB setting overrides env and sanitizes URL
            saved = db.set_master_drive_root_id("https://drive.google.com/drive/folders/DB_ROOT_456/view")
            self.assertEqual(saved, "DB_ROOT_456")
            self.assertEqual(db.get_master_drive_root_info(), {"master_root_id": "DB_ROOT_456", "source": "db"})
            self.assertEqual(db.get_master_drive_root_id(), "DB_ROOT_456")

            # Workspace drive folder cache round-trip
            self.assertEqual(db.get_workspace_drive_folder_id("HSE"), "")
            db.set_workspace_drive_folder_id("hse", "https://drive.google.com/drive/folders/WS_HSE_789")
            self.assertEqual(db.get_workspace_drive_folder_id("HSE"), "WS_HSE_789")

    def test_ac1_workspace_auto_create_concurrency_and_rename(self):
        self.store.set_master_drive_root_id(self.master_root_id)

        # Concurrent creation must not duplicate `<Master Root>/HSE - Health & Safety`
        def _worker():
            return drive_routing.ensure_workspace_drive_folder(self.drive, "HSE", "Health & Safety", store=self.store)

        with ThreadPoolExecutor(max_workers=6) as pool:
            ids = list(pool.map(lambda _: _worker(), range(6)))

        self.assertEqual(len(set(ids)), 1)
        ws_fid = ids[0]
        self.assertTrue(ws_fid)
        children = self.drive.children_of(self.master_root_id)
        self.assertEqual(len(children), 1)
        self.assertEqual(children[0]["name"], "HSE - Health & Safety")
        self.assertEqual(self.store.get_workspace_drive_folder_id("HSE"), ws_fid)

        # Rename workspace in-place via PATCH
        ren = drive_routing.rename_workspace_drive_folder(self.drive, "HSE", "QHSE & Safety", store=self.store)
        self.assertTrue(ren["ok"])
        self.assertTrue(ren["renamed"])
        self.assertEqual(ren["folder_id"], ws_fid)
        children_after = self.drive.children_of(self.master_root_id)
        self.assertEqual(len(children_after), 1)
        self.assertEqual(children_after[0]["id"], ws_fid)
        self.assertEqual(children_after[0]["name"], "HSE - QHSE & Safety")

    def test_ac1_workspace_api_hooks_create_and_rename(self):
        self.store.set_master_drive_root_id(self.master_root_id)
        super_user = {"username": "admin", "role": "superadmin"}
        with self._patch_db_with_store(), \
             patch.object(drive_routing, "_service_provider", lambda: self.drive), \
             patch("db.get_session", return_value=super_user):
            r_add = self.client.post("/api/workspaces", headers=self.ajax, json={
                "action": "add",
                "id": "QA",
                "code": "QA",
                "name": "Quality Assurance",
                "name_ar": "ضمان الجودة",
            })
            self.assertEqual(r_add.status_code, 200)
            body_add = r_add.get_json()
            self.assertTrue(body_add["ok"])
            ws_fid = body_add["drive_folder_id"]
            self.assertTrue(ws_fid)
            ws_folder = self.drive.child_by_name(self.master_root_id, "QA - Quality Assurance")
            self.assertIsNotNone(ws_folder)
            self.assertEqual(ws_folder["id"], ws_fid)

            r_upd = self.client.post("/api/workspaces", headers=self.ajax, json={
                "action": "update",
                "id": "QA",
                "name": "Quality Control",
            })
            self.assertEqual(r_upd.status_code, 200)
            body_upd = r_upd.get_json()
            self.assertTrue(body_upd["ok"])
            self.assertTrue(body_upd["drive_rename"]["renamed"])
            self.assertIsNotNone(self.drive.child_by_name(self.master_root_id, "QA - Quality Control"))
            self.assertIsNone(self.drive.child_by_name(self.master_root_id, "QA - Quality Assurance"))

    # ── AC 2: Project Auto-Provisioning with Blank drive_folder_id & Rename ──
    def test_ac2_create_project_blank_drive_id_provisions_full_tree_and_renames(self):
        self.store.set_master_drive_root_id(self.master_root_id)
        super_user = {"username": "admin", "role": "superadmin"}
        with self._patch_db_with_store(), \
             patch.object(drive_routing, "_service_provider", lambda: self.drive), \
             patch("db.get_session", return_value=super_user):
            res = self.client.post("/api/projects/create", headers=self.ajax, json={
                "id": "P_HSE_1",
                "name": "Cairo Plant",
                "code": "CP01",
                "workspace_id": "HSE",
                "drive_folder_id": "",
            })
            self.assertEqual(res.status_code, 200)
            data = res.get_json()
            self.assertTrue(data["ok"])
            self.assertFalse(data["drive_provision"]["skipped"])

            # Verify `<Master Root>/HSE - Health & Safety/CP01 - Cairo Plant`
            ws_folder = self.drive.child_by_name(self.master_root_id, "HSE - Health & Safety")
            self.assertIsNotNone(ws_folder)
            proj_folder = self.drive.child_by_name(ws_folder["id"], "CP01 - Cairo Plant")
            self.assertIsNotNone(proj_folder)

            # Verify persisted drive_folder_id & auto_provisioned_drive flag
            proj = self.store.get_project("P_HSE_1")
            self.assertEqual(proj["drive_folder_id"], proj_folder["id"])
            self.assertTrue(proj.get("auto_provisioned_drive"))

            # Verify DocTypes, Subtypes (Hot Work / Confined Space), and Submit/Replied folders
            ptw_folder = self.drive.child_by_name(proj_folder["id"], "PTW - Permit to Work")
            self.assertIsNotNone(ptw_folder)
            hot_work = self.drive.child_by_name(ptw_folder["id"], "Hot Work")
            self.assertIsNotNone(hot_work)
            self.assertIsNotNone(self.drive.child_by_name(hot_work["id"], "Submit"))
            self.assertIsNotNone(self.drive.child_by_name(hot_work["id"], "Replied"))

            obs_folder = self.drive.child_by_name(proj_folder["id"], "OBS - Safety Observation")
            self.assertIsNotNone(obs_folder)
            self.assertIsNotNone(self.drive.child_by_name(obs_folder["id"], "Submit"))
            self.assertIsNotNone(self.drive.child_by_name(obs_folder["id"], "Replied"))

            # Rename project via POST /api/project/<pid> -> renames auto-provisioned folder in-place
            r_save = self.client.post("/api/project/P_HSE_1", headers=self.ajax, json={
                "name": "Cairo Plant Phase 2",
                "code": "CP01",
                "drive_folder_id": proj_folder["id"],
            })
            self.assertEqual(r_save.status_code, 200)
            self.assertTrue(r_save.get_json()["drive_rename"]["renamed"])
            renamed_proj = self.drive.child_by_name(ws_folder["id"], "CP01 - Cairo Plant Phase 2")
            self.assertIsNotNone(renamed_proj)
            self.assertEqual(renamed_proj["id"], proj_folder["id"])
            self.assertEqual(len(self.drive.children_of(ws_folder["id"])), 1)

    def test_ac2_project_init_endpoint_works_with_blank_drive_id_when_master_root_set(self):
        self.store.set_master_drive_root_id(self.master_root_id)
        self.store.add_project("P_BLANK", name="Delta Hub", code="DH1", workspace_id="ENG", drive_folder_id="")
        self.store.add_dt("P_BLANK", "DS", "DS", "Document Submittal", has_reply_cycle=True)
        super_user = {"username": "admin", "role": "superadmin"}
        with self._patch_db_with_store(), \
             patch.object(drive_routing, "_service_provider", lambda: self.drive), \
             patch("db.get_session", return_value=super_user):
            r = self.client.post("/api/projects/P_BLANK/init-drive-folders", headers=self.ajax, json={"drive_folder_id": ""})
            self.assertEqual(r.status_code, 200)
            body = r.get_json()
            self.assertTrue(body["ok"])
            self.assertTrue(body["drive_folder_id"])
            self.assertEqual(self.store.get_project("P_BLANK")["drive_folder_id"], body["drive_folder_id"])

    # ── AC 3: Existing Projects With Manual drive_folder_id Stay Unchanged ──
    def test_ac3_legacy_manual_drive_folder_preserved_and_not_renamed(self):
        self.store.set_master_drive_root_id(self.master_root_id)
        self.store.add_project("P_LEG", name="Legacy Mall", code="LM1", workspace_id="ENG", drive_folder_id=self.legacy_root_id)
        self.store.add_dt("P_LEG", "DS", "DS", "Document Submittal", has_reply_cycle=True)

        # ensure_project_drive_folder returns existing manual ID unchanged
        fid = drive_routing.ensure_project_drive_folder(self.drive, "P_LEG", store=self.store)
        self.assertEqual(fid, self.legacy_root_id)
        self.assertEqual(self.drive.children_of(self.master_root_id), [])
        self.assertFalse(self.store.get_project("P_LEG").get("auto_provisioned_drive"))

        # provision_project_drive_structure provisions inside the existing legacy folder, NOT Master Root
        prov = drive_routing.provision_project_drive_structure("P_LEG", service=self.drive, store=self.store)
        self.assertTrue(prov["ok"])
        self.assertEqual(prov["drive_folder_id"], self.legacy_root_id)
        self.assertIsNotNone(self.drive.child_by_name(self.legacy_root_id, "DS - Document Submittal"))
        self.assertEqual(self.drive.children_of(self.master_root_id), [])

        # Renaming a manual project must NOT rename its Drive folder
        ren = drive_routing.rename_project_drive_folder(self.drive, "P_LEG", "Legacy Mall Renamed", store=self.store)
        self.assertTrue(ren["skipped"])
        self.assertFalse(ren["renamed"])
        self.assertEqual(self.drive.items[self.legacy_root_id]["name"], "Legacy Manual Project")

    # ── AC 4: Admin Authorization (403 for Non-Admins) & One-Click Init All ──
    def test_ac4_admin_endpoints_enforce_403_and_provision_all(self):
        self.store.add_project("P_AUTO_1", name="Auto Project", code="AP1", workspace_id="ENG", drive_folder_id="")
        self.store.add_dt("P_AUTO_1", "DS", "DS", "Document Submittal", has_reply_cycle=True)
        self.store.add_project("P_MANUAL_1", name="Manual O&M", code="MO1", workspace_id="OM", drive_folder_id=self.legacy_root_id)
        self.store.add_dt("P_MANUAL_1", "WO", "WO", "Work Orders", has_reply_cycle=False)

        viewer = {"username": "viewer1", "role": "viewer"}
        editor = {"username": "editor1", "role": "editor"}
        super_user = {"username": "admin", "role": "superadmin"}

        with self._patch_db_with_store(), \
             patch.object(drive_routing, "_service_provider", lambda: self.drive), \
             patch.object(db, "is_user_workspace_admin", return_value=False):
            # Unauthenticated -> 403
            with patch("db.get_session", return_value=None):
                self.assertEqual(self.client.get("/api/admin/drive-master-root").status_code, 403)
                self.assertEqual(self.client.post("/api/admin/drive-master-root", headers=self.ajax, json={"master_root_id": self.master_root_id}).status_code, 403)
                self.assertEqual(self.client.post("/api/admin/init-all-drive-folders", headers=self.ajax).status_code, 403)

            # Viewer -> 403
            with patch("db.get_session", return_value=viewer):
                self.assertEqual(self.client.get("/api/admin/drive-master-root").status_code, 403)
                self.assertEqual(self.client.post("/api/admin/init-all-drive-folders", headers=self.ajax).status_code, 403)

            # Editor -> 403
            with patch("db.get_session", return_value=editor):
                self.assertEqual(self.client.post("/api/admin/drive-master-root", headers=self.ajax, json={"master_root_id": self.master_root_id}).status_code, 403)
                self.assertEqual(self.client.post("/api/admin/init-all-drive-folders", headers=self.ajax).status_code, 403)

            # Super Admin -> 200
            with patch("db.get_session", return_value=super_user):
                # Init-all before setting master root returns 400
                r_no_root = self.client.post("/api/admin/init-all-drive-folders", headers=self.ajax)
                self.assertEqual(r_no_root.status_code, 400)

                # Save master root via URL
                r_save = self.client.post("/api/admin/drive-master-root", headers=self.ajax, json={
                    "master_root_id": f"https://drive.google.com/drive/folders/{self.master_root_id}?usp=sharing"
                })
                self.assertEqual(r_save.status_code, 200)
                self.assertEqual(r_save.get_json()["master_root_id"], self.master_root_id)
                self.assertEqual(r_save.get_json()["source"], "db")

                # GET returns saved ID & source
                r_get = self.client.get("/api/admin/drive-master-root")
                self.assertEqual(r_get.status_code, 200)
                self.assertEqual(r_get.get_json(), {"ok": True, "master_root_id": self.master_root_id, "source": "db"})

                # First init-all run creates 3 workspace folders, 1 auto project folder (manual stays existing), and internal folders
                r_init1 = self.client.post("/api/admin/init-all-drive-folders", headers=self.ajax)
                self.assertEqual(r_init1.status_code, 200)
                b1 = r_init1.get_json()
                self.assertTrue(b1["ok"])
                self.assertEqual(b1["workspaces_created"], 3)
                self.assertEqual(b1["workspaces_existing"], 0)
                self.assertEqual(b1["projects_created"], 1)
                self.assertEqual(b1["projects_existing"], 1)
                self.assertGreater(b1["folders_created"], 0)
                self.assertEqual(b1["errors"], [])

                # Verify manual project kept its legacy folder ID
                self.assertEqual(self.store.get_project("P_MANUAL_1")["drive_folder_id"], self.legacy_root_id)
                # Verify auto project got provisioned under ENG workspace folder
                eng_ws_folder = self.drive.child_by_name(self.master_root_id, "ENG - Engineering")
                self.assertIsNotNone(eng_ws_folder)
                auto_proj_folder = self.drive.child_by_name(eng_ws_folder["id"], "AP1 - Auto Project")
                self.assertIsNotNone(auto_proj_folder)
                self.assertEqual(self.store.get_project("P_AUTO_1")["drive_folder_id"], auto_proj_folder["id"])

                # Second init-all run is 100% idempotent (zero new creations)
                r_init2 = self.client.post("/api/admin/init-all-drive-folders", headers=self.ajax)
                self.assertEqual(r_init2.status_code, 200)
                b2 = r_init2.get_json()
                self.assertTrue(b2["ok"])
                self.assertEqual(b2["workspaces_created"], 0)
                self.assertEqual(b2["workspaces_existing"], 3)
                self.assertEqual(b2["projects_created"], 0)
                self.assertEqual(b2["projects_existing"], 2)
                self.assertEqual(b2["folders_created"], 0)
                self.assertGreater(b2["folders_existing"], 0)

    # ── UI & Bilingual EN/AR Verification (Node VM) ──
    def test_ui_controls_i18n_parity_and_node_vm(self):
        required_keys = [
            "admin.master_drive_title",
            "admin.master_drive_desc",
            "admin.master_drive_ph",
            "admin.master_drive_save",
            "admin.master_drive_saving",
            "admin.master_drive_saved",
            "admin.master_drive_cleared",
            "admin.master_drive_save_err",
            "admin.master_drive_src_db",
            "admin.master_drive_src_env",
            "admin.master_drive_src_none",
            "admin.btn_init_all_drive",
            "admin.init_all_drive_ing",
            "admin.init_all_drive_ok",
            "admin.init_all_drive_partial",
            "admin.init_all_drive_err",
            "reg.proj_drv_auto_ph",
            "reg.proj_drv_auto_hint",
        ]
        for k in required_keys:
            en_val = i18n.TRANSLATIONS["en"].get(k)
            ar_val = i18n.TRANSLATIONS["ar"].get(k)
            self.assertTrue(en_val, f"Missing EN key: {k}")
            self.assertTrue(ar_val, f"Missing AR key: {k}")
            self.assertNotEqual(en_val, ar_val, f"AR key not translated: {k}")

        # Verify expected DOM IDs & hooks exist in html_render.py
        src = html_render.SHARED_JS
        self.assertIn("admin-master-drive-root-id", src)
        self.assertIn("btn-init-all-drive", src)
        self.assertIn("saveMasterDriveRoot", src)
        self.assertIn("initAllDriveFolders", src)

        # Extract saveMasterDriveRoot & initAllDriveFolders and run them in Node VM
        m = re.search(
            r"(async function saveMasterDriveRoot\(btn\)[\s\S]*?\nasync function initAllDriveFolders\(btn\)[\s\S]*?\n\})",
            src,
        )
        self.assertIsNotNone(m, "Could not extract Master Drive JS handlers from SHARED_JS")
        js_code = m.group(1)

        node_script = f"""
const vm = require('vm');
const calls = [];
const elements = {{
  'admin-master-drive-root-id': {{ value: 'https://drive.google.com/drive/folders/MASTER_XYZ_999?usp=sharing' }},
  'admin-master-drive-status': {{ style: {{}}, textContent: '' }},
  'admin-master-drive-source': {{ textContent: '' }},
}};
const toasts = [];
const ctx = {{
  _ADMIN_DATA: {{ masterDrive: {{ master_root_id: '', source: 'none' }} }},
  document: {{ getElementById: (id) => elements[id] || null }},
  toast: (msg, type) => toasts.push({{ msg, type }}),
  _admT: (k, vars, fb) => {{
    let out = fb || k;
    if (vars) {{
      for (const [vk, vv] of Object.entries(vars)) out = out.replaceAll('{{' + vk + '}}', String(vv));
    }}
    return out;
  }},
  fetch: async (url, opts) => {{
    calls.push({{ url, body: JSON.parse(opts.body || '{{}}') }});
    if (url === '/api/admin/drive-master-root') {{
      return {{ ok: true, status: 200, json: async () => ({{ ok: true, master_root_id: 'MASTER_XYZ_999', source: 'db' }}) }};
    }}
    if (url === '/api/admin/init-all-drive-folders') {{
      return {{ ok: true, status: 200, json: async () => ({{
        ok: true, workspaces_created: 2, workspaces_existing: 1,
        projects_created: 3, projects_existing: 1,
        folders_created: 18, folders_existing: 4, errors: []
      }}) }};
    }}
    return {{ ok: false, status: 500, json: async () => ({{ ok: false, error: 'unexpected' }}) }};
  }},
}};
vm.createContext(ctx);
vm.runInContext({json.dumps(js_code)}, ctx);
(async () => {{
  const saveBtn = {{ disabled: false, textContent: 'Save' }};
  await ctx.saveMasterDriveRoot(saveBtn);
  const initBtn = {{ disabled: false, innerHTML: 'Init', textContent: 'Init' }};
  await ctx.initAllDriveFolders(initBtn);
  console.log(JSON.stringify({{
    calls,
    inputVal: elements['admin-master-drive-root-id'].value,
    statusText: elements['admin-master-drive-status'].textContent,
    sourceText: elements['admin-master-drive-source'].textContent,
    toasts,
    saveBtnDisabled: saveBtn.disabled,
    initBtnDisabled: initBtn.disabled,
  }}));
}})();
"""
        with tempfile.NamedTemporaryFile("w", suffix=".js", delete=False, encoding="utf-8") as tf:
            tf.write(node_script)
            tf_path = tf.name
        try:
            proc = subprocess.run(["node", tf_path], capture_output=True, text=True, check=True, timeout=15)
        finally:
            os.unlink(tf_path)

        res = json.loads(proc.stdout.strip())
        self.assertEqual(res["calls"][0]["url"], "/api/admin/drive-master-root")
        self.assertEqual(res["calls"][0]["body"]["master_root_id"], "MASTER_XYZ_999")
        self.assertEqual(res["calls"][1]["url"], "/api/admin/init-all-drive-folders")
        self.assertEqual(res["inputVal"], "MASTER_XYZ_999")
        self.assertIn("2 new / 1 verified", res["statusText"])
        self.assertIn("3 new / 1 verified", res["statusText"])
        self.assertIn("18 new / 4 verified", res["statusText"])
        self.assertFalse(res["saveBtnDisabled"])
        self.assertFalse(res["initBtnDisabled"])

    def test_db_get_setting_jsonb_string_roundtrip_and_api(self):
        """Regression test: psycopg2 returns JSONB string scalars as Python str;
        db.get_setting must not crash with JSONDecodeError and clear master_root_id."""
        pg_jsonb_store = {}

        def fake_exe(sql, params=()):
            if "INSERT INTO app_settings" in sql:
                key, json_text = params
                # PostgreSQL ::jsonb parses json_text, and psycopg2 returns decoded Python value
                pg_jsonb_store[key] = json.loads(json_text)

        def fake_q(sql, params=(), one=False):
            if "FROM app_settings" in sql:
                key = params[0]
                if key not in pg_jsonb_store:
                    return None if one else []
                row = {"value": copy.deepcopy(pg_jsonb_store[key])}
                return row if one else [row]
            return None if one else []

        admin_user = {"username": "admin", "role": "superadmin"}
        with patch.object(db, "exe", side_effect=fake_exe), \
             patch.object(db, "q", side_effect=fake_q), \
             patch("db.get_session", return_value=admin_user), \
             patch.object(db, "is_user_workspace_admin", return_value=False), \
             patch.object(db, "log_action", side_effect=self.store.log_action):
            # Save via POST /api/admin/drive-master-root with full Drive URL
            r_post = self.client.post(
                "/api/admin/drive-master-root",
                json={"master_root_id": "https://drive.google.com/drive/folders/1AbCdEfGhIjKlMnOpQrStUvWxYz?usp=sharing"},
                headers=self.ajax,
            )
            self.assertEqual(r_post.status_code, 200)
            body_post = r_post.get_json()
            self.assertEqual(body_post, {
                "ok": True,
                "master_root_id": "1AbCdEfGhIjKlMnOpQrStUvWxYz",
                "source": "db",
            })

            # GET /api/admin/drive-master-root must return the persisted ID (survives modal reopen)
            r_get = self.client.get("/api/admin/drive-master-root", headers=self.ajax)
            self.assertEqual(r_get.status_code, 200)
            self.assertEqual(r_get.get_json(), {
                "ok": True,
                "master_root_id": "1AbCdEfGhIjKlMnOpQrStUvWxYz",
                "source": "db",
            })

            # Numeric-looking string IDs and list/dict settings also round-trip cleanly
            db.save_setting("holidays", ["2026-01-01"])
            self.assertEqual(db.get_setting("holidays"), ["2026-01-01"])
            db.save_setting("numeric_str", "123456")
            self.assertEqual(db.get_setting("numeric_str"), "123456")

            # Explicit clear via empty string returns source="none"
            r_clear = self.client.post(
                "/api/admin/drive-master-root",
                json={"master_root_id": ""},
                headers=self.ajax,
            )
            self.assertEqual(r_clear.status_code, 200)
            self.assertEqual(r_clear.get_json(), {
                "ok": True,
                "master_root_id": "",
                "source": "none",
            })

    def test_save_master_drive_root_node_vm_clear_and_mismatch_guard(self):
        """Regression test: saveMasterDriveRoot preserves ID, handles clear status,
        and rejects silent empty response when a non-empty ID was submitted."""
        src = html_render.SHARED_JS
        m = re.search(
            r"(async function saveMasterDriveRoot\(btn\)[\s\S]*?\nasync function initAllDriveFolders\(btn\))",
            src,
        )
        self.assertIsNotNone(m)
        js_code = m.group(1).rsplit("\nasync function initAllDriveFolders", 1)[0]

        node_script = f"""
const vm = require('vm');
const elements = {{
  'admin-master-drive-root-id': {{ value: '1AbCdEfGhIjKlMnOpQrStUvWxYz' }},
  'admin-master-drive-status': {{ style: {{}}, textContent: '' }},
  'admin-master-drive-source': {{ textContent: 'Not configured' }},
}};
let nextResp = null;
const ctx = {{
  _ADMIN_DATA: {{ masterDrive: {{ master_root_id: '', source: 'none' }} }},
  document: {{ getElementById: (id) => elements[id] || null }},
  toast: () => {{}},
  _admT: (k, vars, fb) => fb || k,
  fetch: async () => ({{ ok: true, status: 200, json: async () => nextResp }}),
}};
vm.createContext(ctx);
vm.runInContext({json.dumps(js_code)}, ctx);
(async () => {{
  const btn = {{ disabled: false, textContent: 'Save' }};
  // 1. Silent empty response when non-empty ID submitted -> error, does NOT clear input
  nextResp = {{ ok: true, master_root_id: '', source: 'none' }};
  await ctx.saveMasterDriveRoot(btn);
  const errStatus = elements['admin-master-drive-status'].textContent;
  const valAfterErr = elements['admin-master-drive-root-id'].value;

  // 2. Successful save -> preserves ID, sets Source: Admin Settings, shows saved
  nextResp = {{ ok: true, master_root_id: '1AbCdEfGhIjKlMnOpQrStUvWxYz', source: 'db' }};
  await ctx.saveMasterDriveRoot(btn);
  const savedStatus = elements['admin-master-drive-status'].textContent;
  const savedSrc = elements['admin-master-drive-source'].textContent;
  const valAfterSave = elements['admin-master-drive-root-id'].value;

  // 3. Explicit clear -> clears input, sets Not configured, shows cleared
  elements['admin-master-drive-root-id'].value = '';
  nextResp = {{ ok: true, master_root_id: '', source: 'none' }};
  await ctx.saveMasterDriveRoot(btn);
  const clearedStatus = elements['admin-master-drive-status'].textContent;
  const clearedSrc = elements['admin-master-drive-source'].textContent;

  console.log(JSON.stringify({{
    errStatus, valAfterErr,
    savedStatus, savedSrc, valAfterSave,
    clearedStatus, clearedSrc,
  }}));
}})();
"""
        with tempfile.NamedTemporaryFile("w", suffix=".js", delete=False, encoding="utf-8") as tf:
            tf.write(node_script)
            tf_path = tf.name
        try:
            proc = subprocess.run(["node", tf_path], capture_output=True, text=True, check=True, timeout=15)
        finally:
            os.unlink(tf_path)

        res = json.loads(proc.stdout.strip())
        self.assertIn("Failed to save Master Drive Root", res["errStatus"])
        self.assertEqual(res["valAfterErr"], "1AbCdEfGhIjKlMnOpQrStUvWxYz")
        self.assertIn("saved", res["savedStatus"])
        self.assertEqual(res["savedSrc"], "Source: Admin Settings")
        self.assertEqual(res["valAfterSave"], "1AbCdEfGhIjKlMnOpQrStUvWxYz")
        self.assertIn("cleared", res["clearedStatus"])
        self.assertEqual(res["clearedSrc"], "Not configured")


if __name__ == "__main__":
    unittest.main()
