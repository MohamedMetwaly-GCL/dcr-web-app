"""Unit and integration tests for PTW sub-tabs, per-subtype auto-numbering, and Drive subfolder routing."""

import json
import os
from pathlib import Path
import re
import subprocess
import tempfile
import unittest
from unittest.mock import patch

os.environ.setdefault("DATABASE_URL", "")
os.environ.setdefault("SECRET_KEY", "synthetic-test-secret-key-only")
os.environ.setdefault("MAGIC_SECRET", "synthetic-test-magic-secret-only")
os.environ.setdefault("GEMINI_API_KEY", "synthetic-test-gemini-key-only")

import app as app_mod
import blueprints.drive_uploads as du_mod
import db
import drive_routing
import hse_defaults
import html_render
import i18n


APPROVED_8_PTW_SUBTYPES = (
    "Hot Work Permit",
    "Cold Work Permit",
    "Work at Height Permit",
    "Confined Space Permit",
    "Excavation Permit",
    "Electrical / Isolation Permit",
    "Lifting Operations Permit",
    "Radiography Permit",
)

EXPECTED_8_CODES = {
    "Hot Work Permit": "HWP",
    "Cold Work Permit": "CWP",
    "Work at Height Permit": "WAH",
    "Confined Space Permit": "CSP",
    "Excavation Permit": "EXP",
    "Electrical / Isolation Permit": "EIP",
    "Lifting Operations Permit": "LOP",
    "Radiography Permit": "RGP",
}


class FakeDrive:
    """In-memory Drive stub compatible with drive_routing and RestDriveAdapter tests."""

    def __init__(self, tree):
        self.items = {}
        self.calls = {"list": 0, "get": 0, "create": 0}
        self._seq = 0
        self._load(tree, None)

    def _load(self, tree, parent):
        for name, sub in tree.items():
            fid = "id_" + name.replace(" ", "_").replace("/", "_")
            self.items[fid] = {"id": fid, "name": name, "parent": parent, "trashed": False}
            self._load(sub, fid)

    def id_of(self, name):
        return "id_" + name.replace(" ", "_").replace("/", "_")

    def files(self):
        outer = self

        class _Exec:
            def __init__(self, val=None, exc=None):
                self.val, self.exc = val, exc

            def execute(self):
                if self.exc:
                    raise self.exc
                return self.val

        class _Files:
            def list(self, q, fields=None, pageSize=None, pageToken=None, **kw):
                outer.calls["list"] += 1
                pid = q.split("'")[1]
                res = [
                    {"id": it["id"], "name": it["name"]}
                    for it in outer.items.values()
                    if it["parent"] == pid and not it["trashed"]
                ]
                return _Exec({"files": res})

            def get(self, fileId, fields=None, **kw):
                outer.calls["get"] += 1
                it = outer.items.get(fileId)
                if not it:
                    return _Exec(exc=Exception("404"))
                return _Exec({"id": fileId, "mimeType": drive_routing.FOLDER_MIME, "trashed": it["trashed"]})

            def create(self, body, fields=None, **kw):
                outer.calls["create"] += 1
                outer._seq += 1
                fid = f"new_{outer._seq}"
                outer.items[fid] = {
                    "id": fid,
                    "name": body["name"],
                    "parent": body["parents"][0],
                    "trashed": False,
                }
                return _Exec({"id": fid, "name": body["name"]})

        return _Files()


class FakeStore:
    def __init__(self, project, dts):
        self.project, self.dts, self.patches = project, dts, []

    def get_project(self, pid):
        return self.project

    def get_doc_type(self, pid, dt_id):
        return self.dts.get(dt_id)

    def set_doc_type_drive_cache(self, pid, dt_id, patch):
        self.patches.append(patch)
        self.dts[dt_id].update(patch)


class TestPtwSubtypesDefaultsAndNumbering(unittest.TestCase):
    def test_01_active_ptw_subtypes_exact_8_without_maintenance_work(self):
        self.assertEqual(tuple(hse_defaults.HSE_SUBTYPES["PTW"].values()), APPROVED_8_PTW_SUBTYPES)
        self.assertEqual(hse_defaults.HSE_LISTS["hse_ptw_subtype"], APPROVED_8_PTW_SUBTYPES)
        self.assertNotIn("Maintenance Work", hse_defaults.HSE_LISTS["hse_ptw_subtype"])
        self.assertEqual(hse_defaults.PTW_SUBTYPE_CODES, EXPECTED_8_CODES)

    def test_02_normalize_legacy_subtype_formats(self):
        self.assertEqual(hse_defaults.normalize_ptw_subtype_value("Hot Work Permit"), "Hot Work Permit")
        self.assertEqual(hse_defaults.normalize_ptw_subtype_value(["Cold Work Permit"]), "Cold Work Permit")
        self.assertEqual(hse_defaults.normalize_ptw_subtype_value('["Work at Height Permit"]'), "Work at Height Permit")
        self.assertEqual(
            hse_defaults.normalize_ptw_subtype_value("Excavation Permit, Hot Work Permit"),
            "Excavation Permit",
        )
        self.assertEqual(
            hse_defaults.normalize_ptw_subtype_value("Electrical / Isolation Permit"),
            "Electrical / Isolation Permit",
        )
        self.assertEqual(hse_defaults.normalize_ptw_subtype_value(""), "")
        self.assertEqual(hse_defaults.normalize_ptw_subtype_value(None), "")

    def test_03_per_subtype_independent_doc_no_generation(self):
        cfc_proj = {"id": "HSE-CFC-P2", "code": "HSE-CFC-P2", "workspace_id": "HSE"}
        # Legacy records should not interfere or collide with new per-subtype numbering
        legacy_recs = [
            {"docNo": f"PEM-058-{i:03d}", "legacySubtype": "Hot Work Permit"}
            for i in range(1, 22)
        ]

        # First per-subtype record starts cleanly at 001
        r_hwp1 = hse_defaults.generate_next_ptw_doc_no("HSE-CFC-P2", "Hot Work Permit", cfc_proj, legacy_recs)
        self.assertEqual(r_hwp1["next"], "PEM058-HWP-001")
        self.assertEqual(r_hwp1["subtype_code"], "HWP")

        r_cwp1 = hse_defaults.generate_next_ptw_doc_no("HSE-CFC-P2", "Cold Work Permit", cfc_proj, legacy_recs)
        self.assertEqual(r_cwp1["next"], "PEM058-CWP-001")

        # Add HWP-001 and HWP-002; CWP must stay 001 independently while HWP advances to 003
        recs_with_hwp = legacy_recs + [
            {"docNo": "PEM058-HWP-001", "legacySubtype": "Hot Work Permit"},
            {"docNo": "PEM058-HWP-002", "legacySubtype": "Hot Work Permit"},
        ]
        r_hwp3 = hse_defaults.generate_next_ptw_doc_no("HSE-CFC-P2", "Hot Work Permit", cfc_proj, recs_with_hwp)
        self.assertEqual(r_hwp3["next"], "PEM058-HWP-003")

        r_cwp_still_1 = hse_defaults.generate_next_ptw_doc_no("HSE-CFC-P2", "Cold Work Permit", cfc_proj, recs_with_hwp)
        self.assertEqual(r_cwp_still_1["next"], "PEM058-CWP-001")

        r_wah1 = hse_defaults.generate_next_ptw_doc_no("HSE-CFC-P2", "Work at Height Permit", cfc_proj, recs_with_hwp)
        self.assertEqual(r_wah1["next"], "PEM058-WAH-001")

        # Adapts dynamically to other projects and explicit ptw_prefix / custom list metadata
        anopc_proj = {"id": "HSE-ANOPC", "code": "HSE-ANOPC", "workspace_id": "HSE"}
        r_anopc = hse_defaults.generate_next_ptw_doc_no("HSE-ANOPC", "Confined Space Permit", anopc_proj, [])
        self.assertEqual(r_anopc["next"], "ANOPC-CSP-001")

        custom_proj = {"id": "HSE-NEW", "code": "HSE-NEW", "ptw_prefix": "PEM099", "workspace_id": "HSE"}
        custom_meta = {"Solar Work Permit": "SWP", "Marine Diving Permit": "MDP"}
        r_custom = hse_defaults.generate_next_ptw_doc_no(
            "HSE-NEW", "Solar Work Permit", custom_proj, [], subtype_meta_map=custom_meta
        )
        self.assertEqual(r_custom["next"], "PEM099-SWP-001")

        # Fallback code derivation works cleanly even without explicit metadata
        r_fallback = hse_defaults.generate_next_ptw_doc_no("HSE-CFC-P2", "Demolition Permit", cfc_proj, [])
        self.assertEqual(r_fallback["next"], "PEM058-DEP-001")

    def test_04_api_next_doc_no_endpoint_ptw_per_subtype(self):
        client = app_mod.app.test_client()
        recs = [
            {"id": "rec-leg-1", "docNo": "PEM-058-001", "legacySubtype": "Maintenance Work", "fileLocation": "https://drive.google.com/file/d/leg1/view"},
            {"id": "rec-hwp-1", "docNo": "PEM058-HWP-001", "legacySubtype": "Hot Work Permit"},
        ]
        meta_list = {
            "hse_ptw_subtype": [
                {"value": "Hot Work Permit", "meta": "HWP"},
                {"value": "Cold Work Permit", "meta": "CWP"},
                {"value": "Solar Work Permit", "meta": "SWP"},
            ]
        }
        proj = {"id": "HSE-CFC-P2", "code": "HSE-CFC-P2", "name": "CFC Phase 2", "workspace_id": "HSE"}
        user = {"username": "tester", "role": "superadmin"}
        with patch.object(app_mod, "current_user", return_value=user), \
             patch.object(app_mod, "can_view_project", return_value=True), \
             patch("app.can_access_doc_type", return_value=True), \
             patch("db.get_doc_type", return_value={"id": "PTW", "code": "PTW", "name": "Permit to Work"}), \
             patch("db.get_project", return_value=proj), \
             patch("db.get_records", return_value=recs), \
             patch("db.get_lists_with_meta", return_value=meta_list):
            r_hwp = client.get("/api/next_doc_no/HSE-CFC-P2/PTW?subtype=Hot%20Work%20Permit")
            self.assertEqual(r_hwp.status_code, 200)
            self.assertEqual(r_hwp.get_json()["next"], "PEM058-HWP-002")

            r_cwp = client.get("/api/next_doc_no/HSE-CFC-P2/PTW?subtype=Cold%20Work%20Permit")
            self.assertEqual(r_cwp.status_code, 200)
            self.assertEqual(r_cwp.get_json()["next"], "PEM058-CWP-001")

            r_swp = client.get("/api/next_doc_no/HSE-CFC-P2/PTW?subtype=Solar%20Work%20Permit")
            self.assertEqual(r_swp.status_code, 200)
            self.assertEqual(r_swp.get_json()["next"], "PEM058-SWP-001")

            r_empty = client.get("/api/next_doc_no/HSE-CFC-P2/PTW")
            self.assertEqual(r_empty.status_code, 200)
            self.assertEqual(r_empty.get_json()["next"], "")

    def test_05_lists_api_and_sync_preserve_records_and_remove_maintenance(self):
        executed = []

        def fake_exe(sql, params=()):
            executed.append((" ".join(sql.split()), params))

        def fake_q(sql, params=(), one=False):
            s = " ".join(sql.split())
            if "FROM projects" in s:
                return [{"id": "HSE-CFC-P2"}]
            if "FROM doc_types" in s:
                return {"id": "PTW", "data": {"legacy_source": "x", "legacy_subtypes": {"maintenance": "Maintenance Work"}}}
            return {"n": 8} if one else []

        with patch("db.exe", side_effect=fake_exe), \
             patch("db.q", side_effect=fake_q):
            db._sync_all_hse_ptw_subtypes()

        sql_blob = "\n".join(s for s, _ in executed).lower()
        self.assertIn("delete from dropdown_lists", sql_blob)
        self.assertNotIn("records", sql_blob)
        self.assertIn(("hse_ptw_subtype", "Maintenance Work"), [p for _, p in executed])

        # Adding via db.add_list_item / API automatically stores parsed/derived code meta
        client = app_mod.app.test_client()
        add_calls = []

        def capture_add(sql, params=()):
            add_calls.append((" ".join(sql.split()), params))

        user = {"username": "tester", "role": "superadmin"}
        with patch.object(app_mod, "current_user", return_value=user), \
             patch.object(app_mod, "can_view_project", return_value=True), \
             patch("blueprints.lists.can_edit", return_value=True), \
             patch("db.q", return_value={"n": 8}), \
             patch("db.exe", side_effect=capture_add), \
             patch("db.invalidate_dashboard_cache"):
            resp = client.post(
                "/api/lists/HSE-CFC-P2",
                json={"list_name": "hse_ptw_subtype", "item": "Solar Work Permit | SWP"},
            )
            self.assertEqual(resp.status_code, 200)
            body = resp.get_json()
            self.assertTrue(body["ok"])
            self.assertEqual(body["item"], "Solar Work Permit")
            self.assertEqual(body["meta"], "SWP")

    def test_06_first_subtype_auto_wires_legacy_subtype_column(self):
        client = app_mod.app.test_client()
        user = {"username": "tester", "role": "superadmin"}
        executed = []

        def fake_exe(sql, params=()):
            executed.append((" ".join(sql.split()), params))

        # Case 1: JSA has no legacySubtype column yet -> ensure_doc_type_subtype_column inserts it
        def fake_q_no_col(sql, params=(), one=False):
            s = " ".join(sql.split())
            if "FROM columns_config" in s:
                return None
            return {"n": 1} if one else []

        with patch.object(app_mod, "current_user", return_value=user), \
             patch.object(app_mod, "can_view_project", return_value=True), \
             patch("blueprints.lists.can_edit", return_value=True), \
             patch("db.q", side_effect=fake_q_no_col), \
             patch("db.exe", side_effect=fake_exe), \
             patch("db.invalidate_dashboard_cache"):
            resp = client.post(
                "/api/lists/HSE-CFC-P2",
                json={"list_name": "hse_jsa_subtype", "item": "Hot Work JSA", "dt_id": "JSA"},
            )
            self.assertEqual(resp.status_code, 200)
            body = resp.get_json()
            self.assertTrue(body["ok"])
            self.assertEqual(body["item"], "Hot Work JSA")
            self.assertTrue(body["column_wired"])

        sql_blob = "\n".join(s for s, _ in executed).lower()
        self.assertIn("insert into columns_config", sql_blob)
        self.assertNotIn("records", sql_blob)
        self.assertIn(("HSE-CFC-P2", "JSA", "hse_jsa_subtype"), [p for _, p in executed])

        # Case 2: Column already exists -> column_wired is False and no columns_config INSERT runs
        executed.clear()
        def fake_q_has_col(sql, params=(), one=False):
            s = " ".join(sql.split())
            if "FROM columns_config" in s:
                return {"col_key": "legacySubtype"}
            return {"n": 2} if one else []

        with patch.object(app_mod, "current_user", return_value=user), \
             patch.object(app_mod, "can_view_project", return_value=True), \
             patch("blueprints.lists.can_edit", return_value=True), \
             patch("db.q", side_effect=fake_q_has_col), \
             patch("db.exe", side_effect=fake_exe), \
             patch("db.invalidate_dashboard_cache"):
            resp2 = client.post(
                "/api/lists/HSE-CFC-P2",
                json={"list_name": "hse_jsa_subtype", "item": "Lifting JSA", "dt_id": "JSA"},
            )
            self.assertEqual(resp2.status_code, 200)
            self.assertFalse(resp2.get_json()["column_wired"])
        self.assertNotIn("insert into columns_config", "\n".join(s for s, _ in executed).lower())


class TestPtwDriveSubfolderRouting(unittest.TestCase):
    def test_01_drive_routing_ptw_subtype_auto_creates_and_reuses_folders(self):
        drive = FakeDrive({"HSE Root": {}})
        store = FakeStore(
            {"drive_folder_id": drive.id_of("HSE Root"), "workspace_id": "HSE"},
            {"PTW": {"id": "PTW", "code": "PTW", "name": "Permit to Work"}},
        )

        # 1. Hot Work Permit creates HSE Root / PTW / Hot Work Permit
        res1 = drive_routing.resolve_upload_destination(
            "HSE-CFC-P2", "PTW", "submission", drive, store,
            create_if_missing=True, flat=False, subtype="Hot Work Permit",
        )
        ptw_id = res1["doc_type_folder_id"]
        hwp_id = res1["folder_id"]
        self.assertEqual(drive.items[ptw_id]["name"], "PTW")
        self.assertEqual(drive.items[ptw_id]["parent"], drive.id_of("HSE Root"))
        self.assertEqual(drive.items[hwp_id]["name"], "Hot Work Permit")
        self.assertEqual(drive.items[hwp_id]["parent"], ptw_id)

        # No Submit or Replied folders created for PTW
        ptw_children = {it["name"] for it in drive.items.values() if it["parent"] == ptw_id}
        self.assertNotIn("Submit", ptw_children)
        self.assertNotIn("Replied", ptw_children)

        # 2. Second upload (even approval/closeout slot) for same subtype reuses existing folder ID
        creates_before = drive.calls["create"]
        res2 = drive_routing.resolve_upload_destination(
            "HSE-CFC-P2", "PTW", "approval", drive, store,
            create_if_missing=True, flat=False, subtype="Hot Work Permit",
        )
        self.assertEqual(res2["folder_id"], hwp_id)
        self.assertEqual(drive.calls["create"], creates_before)

        # 3. Cold Work Permit creates separate sibling folder under PTW
        res3 = drive_routing.resolve_upload_destination(
            "HSE-CFC-P2", "PTW", "submission", drive, store,
            create_if_missing=True, flat=False, subtype="Cold Work Permit",
        )
        cwp_id = res3["folder_id"]
        self.assertEqual(drive.items[cwp_id]["name"], "Cold Work Permit")
        self.assertEqual(drive.items[cwp_id]["parent"], ptw_id)
        self.assertNotEqual(cwp_id, hwp_id)

        # 4. Missing subtype (edge case) safely falls back to parent PTW folder
        res_empty = drive_routing.resolve_upload_destination(
            "HSE-CFC-P2", "PTW", "submission", drive, store,
            create_if_missing=True, flat=False, subtype="",
        )
        self.assertEqual(res_empty["folder_id"], ptw_id)

        # 4b. Dynamically added custom permit subtype routes to PTW/<New Subtype Name>/
        res_marine = drive_routing.resolve_upload_destination(
            "HSE-CFC-P2", "PTW", "submission", drive, store,
            create_if_missing=True, flat=False, subtype="Marine Work Permit",
        )
        self.assertEqual(drive.items[res_marine["folder_id"]]["name"], "Marine Work Permit")
        self.assertEqual(drive.items[res_marine["folder_id"]]["parent"], ptw_id)

        # 5. Reuses existing "PTW - Permit to Work" folder if present in legacy Drive layout
        legacy_drive = FakeDrive({"HSE Root Legacy": {"PTW - Permit to Work": {}}})
        legacy_store = FakeStore(
            {"drive_folder_id": legacy_drive.id_of("HSE Root Legacy"), "workspace_id": "HSE"},
            {"PTW": {"id": "PTW", "code": "PTW", "name": "Permit to Work"}},
        )
        res_leg = drive_routing.resolve_upload_destination(
            "HSE-CFC-P2", "PTW", "submission", legacy_drive, legacy_store,
            create_if_missing=True, flat=False, subtype="Electrical / Isolation Permit",
        )
        self.assertEqual(res_leg["doc_type_folder_id"], legacy_drive.id_of("PTW - Permit to Work"))
        self.assertEqual(legacy_drive.items[res_leg["folder_id"]]["name"], "Electrical / Isolation Permit")
        self.assertEqual(
            legacy_drive.items[res_leg["folder_id"]]["parent"],
            legacy_drive.id_of("PTW - Permit to Work"),
        )

    def test_02_upload_session_endpoint_passes_ptw_subtype_and_preserves_legacy_links(self):
        drive = FakeDrive({"HSE Root": {}})
        created_sessions = []

        class _Resp:
            def __init__(self, code, payload=None, headers=None):
                self.status_code = code
                self._payload = payload or {}
                self.headers = headers or {}
                self.text = json.dumps(self._payload)

            def json(self):
                return self._payload

        class _FakeHttp:
            def get(self, url, params=None, headers=None, timeout=None):
                if "/drive/v3/files/" in url:
                    fid = url.rsplit("/", 1)[-1].split("?", 1)[0]
                    r = drive.files().get(fid)
                    try:
                        return _Resp(200, r.execute())
                    except Exception:
                        return _Resp(404, {"error": {"message": "Not found"}})
                q = (params or {}).get("q", "")
                res = drive.files().list(q=q).execute()
                return _Resp(200, res)

            def post(self, url, params=None, json=None, headers=None, data=None, timeout=None):
                import json as _json
                body = json if json is not None else (_json.loads(data) if data else {})
                if "upload/drive/v3/files" in url:
                    parent = (body.get("parents") or [""])[0]
                    created_sessions.append({"name": body.get("name"), "parent": parent})
                    return _Resp(200, {}, headers={"Location": "https://fake-drive/upload/1"})
                res = drive.files().create(body=body).execute()
                return _Resp(200, res)

        client = app_mod.app.test_client()
        proj = {"id": "HSE-CFC-P2", "code": "HSE-CFC-P2", "drive_folder_id": drive.id_of("HSE Root"), "workspace_id": "HSE"}
        dt = {"id": "PTW", "code": "PTW", "name": "Permit to Work"}
        rec = {
            "_id": "rec-1",
            "_project_id": "HSE-CFC-P2",
            "_dt_id": "PTW",
            "docNo": "PEM058-HWP-001",
            "legacySubtype": "Hot Work Permit",
        }
        user = {"username": "tester", "role": "superadmin"}
        dt_patches = []

        with patch.object(app_mod, "current_user", return_value=user), \
             patch.object(app_mod, "can_view_project", return_value=True), \
             patch.object(du_mod, "current_user", return_value=user), \
             patch.object(du_mod, "can_edit", return_value=True), \
             patch.object(du_mod, "can_access_doc_type", return_value=True), \
             patch.object(du_mod, "_token_provider", return_value="FAKE_TOKEN"), \
             patch.object(du_mod, "_http_client", _FakeHttp()), \
             patch("db.get_project", return_value=proj), \
             patch("db.get_doc_type", return_value=dt), \
             patch("db.get_record_by_id", return_value=rec), \
             patch("db.set_doc_type_drive_cache", side_effect=lambda p, d, patch_d: dt_patches.append(patch_d)):
            r = client.post("/api/drive/upload-session", json={
                "project_id": "HSE-CFC-P2",
                "dt_id": "PTW",
                "record_id": "rec-1",
                "slot": "submission",
                "file_name": "permit.pdf",
                "file_size": 1024,
            })
            self.assertEqual(r.status_code, 200, r.get_json())
            data = r.get_json()
            self.assertTrue(data["ok"])
            fid = data["folder_id"]
            self.assertEqual(drive.items[fid]["name"], "Hot Work Permit")
            self.assertEqual(drive.items[drive.items[fid]["parent"]]["name"], "PTW")
            self.assertEqual(created_sessions[-1]["parent"], fid)


class TestPtwUiRenderingAndJsBehavior(unittest.TestCase):
    def _render_register(self, lang="en"):
        user = {"id": 1, "username": "tester", "role": "admin", "workspace_role": "admin", "can_download": True}
        proj = {"id": "HSE-CFC-P2", "code": "HSE-CFC-P2", "name": "CFC Phase 2", "workspace_id": "HSE"}
        dts = [
            {"id": "PTW", "code": "PTW", "name": "Permit to Work"},
            {"id": "OBS", "code": "OBS", "name": "Safety Observations"},
        ]
        workspaces = [{"id": "HSE", "code": "HSE", "name": "HSE", "name_ar": "السلامة", "icon": "🛡️"}]
        with patch("html_render.can_edit", return_value=True), \
             patch("html_render.is_project_admin", return_value=True), \
             patch("html_render.get_allowed_doc_type_filter", return_value=None), \
             patch("db.can_user_download", return_value=True), \
             patch("db.get_projects_for_user", return_value=[proj]), \
             patch("db.get_doc_types", return_value=dts), \
             patch("db.get_logo", return_value=None), \
             patch("db.get_user_workspaces", return_value=workspaces):
            return html_render.render_register(user, proj, active_ws_id="HSE", lang=lang)

    def test_01_rendered_html_contains_ptw_subtabs_and_single_select_and_notes_fix(self):
        for lang in ("en", "ar"):
            html = self._render_register(lang=lang)
            self.assertIn('id="ptw-subtabs"', html)
            self.assertIn('id="btn-add-subtab"', html)
            self.assertIn('id="add-subtype-modal"', html)
            self.assertIn('id="add-subtype-name"', html)
            self.assertIn('id="add-subtype-code"', html)
            self.assertIn("openAddSubtypeModal", html)
            self.assertIn("openSubtabContextMenu", html)
            self.assertIn("openEditSubtypeModal", html)
            self.assertIn("moveSubtypeTab", html)
            self.assertIn("deleteSubtypeTab", html)
            self.assertIn("subtab-ctx-menu", html)
            self.assertIn("saveNewSubtype", html)
            self.assertIn("getActiveSubtypeConfig", html)
            self.assertIn("getOrBuildSubtypeModalConfig", html)
            self.assertIn("DEFAULT_PTW_SUBTYPE_CODES", html)
            expected_btn_lbl = "+ تابة فرعية" if lang == "ar" else "+ Sub-Tab"
            self.assertIn(expected_btn_lbl, html)
            self.assertIn("updatePtwSubTabs", html)
            self.assertIn("setPtwSubtypeFilter", html)
            self.assertIn("isHsePtwTab", html)
            self.assertIn("updatePtwDocNoForSubtype", html)
            self.assertIn("placeholder:col.placeholder||''", html)
            self.assertIn("ta.placeholder=longTextMeta.placeholder||''", html)
            self.assertLess(html.find('id="tabsbar"'), html.find('id="ptw-subtabs"'))
            self.assertLess(html.find('id="ptw-subtabs"'), html.find('id="toolbar"'))

    def test_02_i18n_en_ar_parity(self):
        self.assertEqual(set(i18n.TRANSLATIONS["en"].keys()), set(i18n.TRANSLATIONS["ar"].keys()))
        for key in (
            "hse.ptw.subtabs_aria",
            "hse.ptw.select_subtype",
            "hse.ptw.no_subtype_fallback",
            "hse.ptw.list_code_ph",
            "hse.ptw.edit_code_prompt",
            "hse.ptw.edit_code_btn",
            "hse.subtab.add_btn_title",
            "hse.subtab.toolbar_btn",
            "hse.subtab.toolbar_btn_title",
            "hse.subtab.modal_title_ptw",
            "hse.subtab.modal_title_generic",
            "hse.subtab.name_label",
            "hse.subtab.code_label",
            "hse.subtab.added_toast",
            "hse.subtab.ctx_edit",
            "hse.subtab.ctx_move_left",
            "hse.subtab.ctx_move_right",
            "hse.subtab.ctx_delete",
            "hse.subtab.edit_modal_title",
            "hse.subtab.confirm_delete",
            "hse.subtab.updated_toast",
            "hse.subtab.deleted_toast",
        ):
            self.assertTrue(i18n.TRANSLATIONS["en"][key].strip())
            self.assertTrue(i18n.TRANSLATIONS["ar"][key].strip())

    def test_03_node_js_syntax_and_dom_behavior(self):
        html = self._render_register(lang="ar")
        scripts = re.findall(r"<script(?:\s+[^>]*)?>([\s\S]*?)</script>", html, re.IGNORECASE)
        main_scripts = [s for s in scripts if "updatePtwSubTabs" in s and "buildForm" in s]
        self.assertEqual(len(main_scripts), 1)
        main_script = main_scripts[0]

        node_script = r"""
const fs = require('fs');
const vm = require('vm');
const assert = require('assert');

const mainJs = fs.readFileSync(process.argv[2], 'utf8');

function mkEl(tag) {
  const listeners = {};
  const attrs = {};
  const children = [];
  const classSet = new Set();
  const el = {
    tagName: String(tag || 'DIV').toUpperCase(),
    id: '',
    value: '',
    placeholder: '',
    innerHTML: '',
    textContent: '',
    style: { cssText: '' },
    dataset: {},
    children,
    classList: {
      add: (...cs) => cs.forEach(c => classSet.add(c)),
      remove: (...cs) => cs.forEach(c => classSet.delete(c)),
      toggle: (c, force) => { if (force === undefined) { classSet.has(c) ? classSet.delete(c) : classSet.add(c); } else if (force) classSet.add(c); else classSet.delete(c); },
      contains: (c) => classSet.has(c),
    },
    set className(v) { classSet.clear(); String(v || '').split(/\s+/).filter(Boolean).forEach(c => classSet.add(c)); },
    get className() { return [...classSet].join(' '); },
    setAttribute: (k, v) => { attrs[k] = String(v); },
    getAttribute: (k) => attrs[k] || null,
    appendChild: (c) => { children.push(c); c.parentElement = el; return c; },
    insertBefore: (c) => { children.push(c); c.parentElement = el; return c; },
    remove: () => {
      if (el.parentElement && Array.isArray(el.parentElement.children)) {
        const i = el.parentElement.children.indexOf(el);
        if (i >= 0) el.parentElement.children.splice(i, 1);
      }
      if (el.id && byId[el.id] === el) delete byId[el.id];
    },
    addEventListener: (ev, fn) => { (listeners[ev] = listeners[ev] || []).push(fn); },
    dispatchEvent: (ev) => { (listeners[ev.type] || []).forEach(fn => fn(ev)); },
    querySelector: (sel) => String(sel || '').startsWith('#') ? null : mkEl('span'),
    querySelectorAll: () => [],
    closest: () => null,
    get isConnected() { return true; },
  };
  Object.defineProperty(el, 'innerHTML', {
    get() { return this._inner || ''; },
    set(v) { this._inner = String(v || ''); if (!v) children.length = 0; }
  });
  return el;
}

const byId = {};
[
  'tabsbar', 'ptw-subtabs', 'toolbar', 'btn-add-subtab', 'thead', 'tbody', 'srchbox',
  'rec-form', 'rec-title', 'rec-modal', 'rec-save-btn', 'rec-cancel-btn',
  'add-subtype-modal', 'add-subtype-title', 'add-subtype-name',
  'add-subtype-code-wrap', 'add-subtype-code', 'add-subtype-docno-preview', 'add-subtype-save-btn',
  's-clock', 's-total', 's-show', 's-ov', 'empty'
].forEach(id => {
  const e = mkEl('div'); e.id = id; byId[id] = e;
});

function walkEls(root, out = []) {
  if (!root || !root.children) return out;
  for (const ch of root.children) {
    out.push(ch);
    walkEls(ch, out);
  }
  return out;
}

const bodyChildren = [];
const bodyEl = {
  children: bodyChildren,
  classList: { contains: () => false },
  appendChild: (c) => {
    bodyChildren.push(c);
    c.parentElement = bodyEl;
    if (c.id) byId[c.id] = c;
    return c;
  },
};

const documentStub = {
  documentElement: { dir: 'rtl', lang: 'ar' },
  body: bodyEl,
  createElement: (tag) => mkEl(tag),
  getElementById: (id) => {
    if (byId[id]) return byId[id];
    const all = walkEls(byId['rec-form']).concat(walkEls(byId['thead']));
    return all.find(e => e.id === id) || null;
  },
  querySelector: (sel) => {
    const m = String(sel || '').match(/\[data-col-key="([^"]+)"\]/);
    if (m) {
      return walkEls(byId['thead']).find(e => e.dataset && e.dataset.colKey === m[1]) || null;
    }
    return null;
  },
  querySelectorAll: () => [],
  addEventListener: () => {},
};

const ptwColsMock = [
  { col_key: 'docNo', label: 'Register No.', col_type: 'docno', visible: true },
  { col_key: 'legacySubtype', label: 'Subtype', col_type: 'dropdown', list_name: 'hse_ptw_subtype', visible: true },
  { col_key: 'title', label: 'Subject', col_type: 'text', visible: true },
  { col_key: 'status', label: 'Status', col_type: 'dropdown', list_name: 'hse_doc_status', visible: true },
  { col_key: 'remarks', label: 'Notes', col_type: 'long_text', visible: true },
  { col_key: 'fileLocation', label: 'Drive Link', col_type: 'url', visible: true },
];

const mockServerLists = {
  hse_doc_status: ['Open', 'Closed'],
  hse_ptw_subtype: [
    'Hot Work Permit',
    'Cold Work Permit',
    'Work at Height Permit',
    'Confined Space Permit',
    'Excavation Permit',
    'Electrical / Isolation Permit',
    'Lifting Operations Permit',
    'Radiography Permit',
  ],
  hse_obs_subtype: ['Unsafe Act', 'Unsafe Condition', 'Near Miss'],
};
const mockServerMeta = {
  hse_ptw_subtype: [
    { value: 'Hot Work Permit', meta: 'HWP' },
    { value: 'Cold Work Permit', meta: 'CWP' },
  ],
};
const postedLists = [];
const patchedLists = [];
const reorderedLists = [];
const deletedLists = [];
const confirmCalls = [];

const mockColsByDt = {
  PTW: ptwColsMock,
  OBS: [
    { col_key: 'docNo', label: 'Register No.', col_type: 'docno', visible: true },
    { col_key: 'legacySubtype', label: 'Subtype', col_type: 'dropdown', list_name: 'hse_obs_subtype', visible: true },
    { col_key: 'title', label: 'Subject', col_type: 'text', visible: true },
  ],
  JSA: [
    { col_key: 'docNo', label: 'Register No.', col_type: 'docno', visible: true },
    { col_key: 'title', label: 'Subject', col_type: 'text', visible: true },
  ],
  TBT: [
    { col_key: 'docNo', label: 'Register No.', col_type: 'docno', visible: true },
  ],
};

const ctx = {
  console,
  Math,
  Date,
  RegExp,
  Map,
  Set,
  Array,
  Object,
  String,
  Number,
  Boolean,
  JSON,
  Promise,
  Error,
  encodeURIComponent,
  decodeURIComponent,
  parseInt,
  parseFloat,
  isNaN,
  setTimeout: () => 0,
  setInterval: () => 0,
  requestAnimationFrame: () => 0,
  URLSearchParams: function() { return { get: () => null }; },
  location: { search: '' },
  document: documentStub,
  CustomEvent: function(type, init) { this.type = type; this.detail = init && init.detail; },
  Event: function(type) { this.type = type; },
  CSS: { escape: s => s },
  window: {
    t: (k) => k === 'reg.filter_all' ? 'الكل' : k,
    tCol: (_k, l) => l,
    tDt: (_c, n) => n,
    tStatus: (s) => s,
    matchMedia: () => ({ matches: false }),
    addEventListener: () => {},
  },
  escHtml: (s) => String(s || ''),
  confirm: (msg) => { confirmCalls.push(msg); return true; },
  apiFetch: async (url, opts) => {
    const u = String(url);
    if (u.startsWith('/api/columns/')) {
      const dt = u.split('/').pop();
      return mockColsByDt[dt] || apiObj.state.cols;
    }
    if (u.startsWith('/api/records/')) {
      const dt = u.split('/').pop().split('?')[0];
      return { records: apiObj.state.recs || [], columns: mockColsByDt[dt] || apiObj.state.cols, count: (apiObj.state.recs || []).length };
    }
    if (u.startsWith('/api/col_width/')) return {};
    if (u.includes('/api/next_doc_no/')) {
      const m = u.match(/subtype=([^&]+)/);
      const sub = m ? decodeURIComponent(m[1]) : '';
      return { next: apiObj.computeNextPtwDocNoClient(sub) };
    }
    if (u.includes('/reorder') && opts && opts.method === 'POST') {
      const body = JSON.parse(opts.body || '{}');
      reorderedLists.push(body);
      mockServerLists[body.list_name] = [...(body.order || [])];
      return { ok: true };
    }
    if (u.startsWith('/api/lists/') && opts && opts.method === 'PATCH') {
      const body = JSON.parse(opts.body || '{}');
      patchedLists.push(body);
      const arr = mockServerLists[body.list_name] || [];
      const i = arr.indexOf(body.old_item);
      if (i >= 0) arr[i] = body.new_item;
      if (body.meta) {
        mockServerMeta[body.list_name] = mockServerMeta[body.list_name] || [];
        const mi = mockServerMeta[body.list_name].find(x => x.value === body.old_item || x.value === body.new_item);
        if (mi) { mi.value = body.new_item; mi.meta = body.meta; }
        else mockServerMeta[body.list_name].push({ value: body.new_item, meta: body.meta });
      }
      return { ok: true, renamed: 1 };
    }
    if (u.startsWith('/api/lists/') && opts && opts.method === 'DELETE') {
      const body = JSON.parse(opts.body || '{}');
      deletedLists.push(body);
      mockServerLists[body.list_name] = (mockServerLists[body.list_name] || []).filter(x => x !== body.item);
      return { ok: true };
    }
    if (u.startsWith('/api/lists/') && opts && opts.method === 'POST') {
      const body = JSON.parse(opts.body || '{}');
      postedLists.push(body);
      mockServerLists[body.list_name] = mockServerLists[body.list_name] || [];
      mockServerLists[body.list_name].push(body.item);
      if (body.meta) {
        mockServerMeta[body.list_name] = mockServerMeta[body.list_name] || [];
        mockServerMeta[body.list_name].push({ value: body.item, meta: body.meta });
      }
      let columnWired = false;
      if (body.dt_id && mockColsByDt[body.dt_id] && !mockColsByDt[body.dt_id].some(c => c.col_key === 'legacySubtype')) {
        mockColsByDt[body.dt_id].splice(1, 0, {
          col_key: 'legacySubtype', label: 'Subtype', col_type: 'dropdown', list_name: body.list_name, visible: true
        });
        columnWired = true;
      }
      return { ok: true, item: body.item, meta: body.meta || null, list_name: body.list_name, column_wired: columnWired };
    }
    if (u.startsWith('/api/lists/')) return JSON.parse(JSON.stringify(mockServerLists));
    if (u.startsWith('/api/lists_meta/')) return JSON.parse(JSON.stringify(mockServerMeta));
    return null;
  },
  fetch: async () => ({ ok: true, json: async () => ({}) }),
};
ctx.window.document = documentStub;
ctx.globalThis = ctx;

// Stub helper functions defined in _shared_js_html / earlier script tag
const apiObj = vm.runInNewContext(
  'function bindDirectionalInput(){} function bindSmartTextarea(){} function openM(){} function closeM(){} function toast(){} function initRz(){} function renderNocSummary(){} function updateLTRQuickBar(){} function isLTRTab(){return false;} function isNOCTab(){return false;} function isPRTab(){return false;} function isSDOrABDTab(){return false;} function tabHasReplyCycle(){return false;} function getOrCreateFormSection(f){return f;} function compactSingleFieldFormSections(){} function mountIncBodyMapSection(){} function mountDriveAttachmentsBox(){} function getDynamicFieldSpan(){return "";} function getDynamicFormSection(){return "Basic";} function getOrderedRecordFormCols(cols){return cols;} function isCalculatedFormField(){return false;} function isCurrencyField(){return false;} function _normalizeRowAttachments(){return {};} function _getReplyStatusTheme(){return {color:"#333",bg:"#fff",bd:"#ccc"};} function buildMS(k,_o,init){const d=document.createElement("div");d.className="ms-con";d.id="f-"+k;d.dataset={value:init||""};return d;}\n' +
  mainJs +
  '\n;({ state, getLongTextMeta, getActiveSubtypeConfig, updatePtwSubTabs, setPtwSubtypeFilter, renderRows, computeNextPtwDocNoClient, buildForm, getBatchTargetDrivePath, openAddSubtypeModal, openSubtabContextMenu, openEditSubtypeModal, moveSubtypeTab, deleteSubtypeTab, onAddSubtypeNameInput, onAddSubtypeCodeInput, saveNewSubtype });',
  ctx
);

(async () => {
  // Check getLongTextMeta placeholder never undefined
  const meta = apiObj.getLongTextMeta({ col_key: 'remarks', label: 'Notes', col_type: 'long_text' });
  assert.strictEqual(meta.placeholder, '', 'long_text placeholder must be empty string, not undefined');

  // Setup state as PTW tab on HSE-CFC-P2
  apiObj.state.tab = 'PTW';
  apiObj.state.cols = ptwColsMock;
  apiObj.state.allTabCols = ptwColsMock;
  apiObj.state.dtList = [
    { id: 'PTW', code: 'PTW', name: 'Permit to Work' },
    { id: 'OBS', code: 'OBS', name: 'Safety Observations' },
    { id: 'JSA', code: 'JSA', name: 'Job Safety Analysis' },
    { id: 'TBT', code: 'TBT', name: 'Toolbox Talks' },
  ];
  apiObj.state.lists = JSON.parse(JSON.stringify(mockServerLists));
  apiObj.state.listsMeta = JSON.parse(JSON.stringify(mockServerMeta));
  apiObj.state.recs = [
    { _id: '1', docNo: 'PEM-058-001', legacySubtype: 'Maintenance Work' },
    { _id: '2', docNo: 'PEM058-HWP-001', legacySubtype: 'Hot Work Permit' },
  ];

  apiObj.updatePtwSubTabs();
  assert.strictEqual(byId['btn-add-subtab'].style.display, 'none', 'Toolbar + Sub-Tab button must stay hidden when PTW sub-tab strip is visible');
  const subBtns = byId['ptw-subtabs'].children;
  assert.strictEqual(subBtns.length, 10, 'All + 8 permit types + "+" button must appear as sub-tabs');
  assert.strictEqual(subBtns[0].dataset.subtype, '');
  assert.strictEqual(subBtns[1].dataset.subtype, 'Hot Work Permit');
  assert.strictEqual(subBtns[1].children[2].textContent, '1');
  assert.strictEqual(subBtns[2].dataset.subtype, 'Cold Work Permit');
  assert.strictEqual(subBtns[2].children[2].textContent, '0');
  assert.strictEqual(subBtns[9].id, 'ptw-subtab-add-btn', 'Last item on sub-tab strip must be "+" button');
  assert.strictEqual(subBtns[9].textContent, '+');

  // Sub-tab filtering restricts visibleRows cleanly and All restores all records
  apiObj.setPtwSubtypeFilter('Hot Work Permit');
  assert.strictEqual(apiObj.state.ptwSubtypeFilter, 'Hot Work Permit');
  assert.strictEqual(JSON.stringify(apiObj.state.visibleRows.map(r => r._id)), JSON.stringify(['2']));
  assert.strictEqual(apiObj.getBatchTargetDrivePath('submission'), 'HSE-CFC-P2 ❯ PTW ❯ Hot Work Permit');

  apiObj.setPtwSubtypeFilter('');
  assert.strictEqual(apiObj.state.visibleRows.length, 2);

  // Verify per-subtype auto-numbering calculation in JS
  assert.strictEqual(apiObj.computeNextPtwDocNoClient('Hot Work Permit'), 'PEM058-HWP-002');
  assert.strictEqual(apiObj.computeNextPtwDocNoClient('Cold Work Permit'), 'PEM058-CWP-001');
  assert.strictEqual(apiObj.computeNextPtwDocNoClient('Work at Height Permit'), 'PEM058-WAH-001');

  // Build Add form with Cold Work Permit sub-tab pre-selected
  apiObj.state.editId = null;
  apiObj.state.ptwSubtypeFilter = 'Cold Work Permit';
  await apiObj.buildForm(null, { defaultSubtype: 'Cold Work Permit', suggestedDocNo: 'PEM058-CWP-001' });

  const subSel = documentStub.getElementById('f-legacySubtype');
  assert.ok(subSel, '#f-legacySubtype must exist');
  assert.strictEqual(subSel.tagName, 'SELECT', 'PTW legacySubtype must render as single-select <select>, not buildMS');
  assert.strictEqual(documentStub.getElementById('f-free-legacySubtype'), null, 'PTW single-select must not render free-text multi-select input');
  assert.strictEqual(subSel.value, 'Cold Work Permit');

  const docInp = documentStub.getElementById('f-docNo');
  assert.strictEqual(docInp.value, 'PEM058-CWP-001');

  const remarksTa = documentStub.getElementById('f-remarks');
  assert.strictEqual(remarksTa.tagName, 'TEXTAREA');
  assert.strictEqual(remarksTa.placeholder, '', 'Notes textarea placeholder must not be undefined');

  // Switching subtype dropdown updates #f-docNo to Hot Work Permit next seq
  subSel.value = 'Hot Work Permit';
  subSel.dispatchEvent({ type: 'change' });
  await new Promise(r => setTimeout(r, 0));
  assert.strictEqual(docInp.value, 'PEM058-HWP-002');

  // Editing legacy record preserves existing Doc No. and legacy Maintenance Work option
  apiObj.state.editId = '1';
  await apiObj.buildForm(apiObj.state.recs[0], {});
  const editSel = documentStub.getElementById('f-legacySubtype');
  const editDoc = documentStub.getElementById('f-docNo');
  assert.strictEqual(editDoc.value, 'PEM-058-001', 'Legacy Doc No. must not be overwritten');
  assert.ok(editSel.innerHTML.includes('Maintenance Work'), 'Legacy subtype option must remain selectable on edit');

  // Test "+" modal inside PTW: auto-suggests MWP for Marine Work Permit, saves via /api/lists, activates new sub-tab
  apiObj.openAddSubtypeModal();
  assert.strictEqual(byId['add-subtype-code-wrap'].style.display, '', 'PTW modal must show code input');
  byId['add-subtype-name'].value = 'Marine Work Permit';
  apiObj.onAddSubtypeNameInput();
  assert.strictEqual(byId['add-subtype-code'].value, 'MWP', 'Auto-suggested code for Marine Work Permit must be MWP');
  assert.ok(byId['add-subtype-docno-preview'].textContent.includes('PEM058-MWP-001'));
  await apiObj.saveNewSubtype();
  assert.strictEqual(postedLists[0].list_name, 'hse_ptw_subtype');
  assert.strictEqual(postedLists[0].item, 'Marine Work Permit');
  assert.strictEqual(postedLists[0].meta, 'MWP');
  assert.strictEqual(apiObj.state.ptwSubtypeFilter, 'Marine Work Permit', 'New subtype should become active');
  assert.strictEqual(byId['ptw-subtabs'].children.length, 11, 'All + 9 permit types + "+" button');
  assert.strictEqual(apiObj.computeNextPtwDocNoClient('Marine Work Permit'), 'PEM058-MWP-001');

  // Test right-click context menu on sub-tabs (All and + must not have contextmenu)
  assert.strictEqual(byId['ptw-subtabs'].children[0].oncontextmenu, undefined, 'All pill must not open context menu');
  assert.strictEqual(byId['ptw-subtabs'].children[10].oncontextmenu, undefined, '+ button must not open context menu');
  assert.strictEqual(typeof byId['ptw-subtabs'].children[1].oncontextmenu, 'function', 'Subtype pill must attach oncontextmenu');

  // RTL: index 0 (Hot Work Permit) is rightmost -> move-right disabled, move-left enabled (delta=+1)
  documentStub.documentElement.dir = 'rtl';
  byId['ptw-subtabs'].children[1].oncontextmenu({ preventDefault() {}, clientX: 120, clientY: 80 });
  let ctxMenu = byId['subtab-ctx-menu'];
  assert.ok(ctxMenu, '#subtab-ctx-menu must render on right-click');
  assert.strictEqual(ctxMenu.children.length, 4);
  assert.strictEqual(ctxMenu.children[1].dataset.action, 'move-left');
  assert.strictEqual(ctxMenu.children[1].getAttribute('aria-disabled'), null, 'RTL index 0 can move left');
  assert.strictEqual(ctxMenu.children[2].dataset.action, 'move-right');
  assert.strictEqual(ctxMenu.children[2].getAttribute('aria-disabled'), 'true', 'RTL index 0 cannot move right');

  // Click Move Left in RTL -> swaps Hot Work Permit (0) and Cold Work Permit (1)
  await ctxMenu.children[1].onclick({ stopPropagation() {} });
  assert.strictEqual(reorderedLists.length, 1);
  assert.strictEqual(reorderedLists[0].order[0], 'Cold Work Permit');
  assert.strictEqual(reorderedLists[0].order[1], 'Hot Work Permit');

  // LTR: index 0 (Cold Work Permit) is leftmost -> move-left disabled, move-right enabled (delta=+1)
  documentStub.documentElement.dir = 'ltr';
  byId['ptw-subtabs'].children[1].oncontextmenu({ preventDefault() {}, clientX: 120, clientY: 80 });
  ctxMenu = byId['subtab-ctx-menu'];
  assert.strictEqual(ctxMenu.children[1].getAttribute('aria-disabled'), 'true', 'LTR index 0 cannot move left');
  assert.strictEqual(ctxMenu.children[2].getAttribute('aria-disabled'), null, 'LTR index 0 can move right');
  await ctxMenu.children[2].onclick({ stopPropagation() {} });
  assert.strictEqual(reorderedLists[1].order[0], 'Hot Work Permit');

  // Edit Sub-Tab Settings via context menu: rename Marine Work Permit -> Marine Ops Permit + code MOP
  apiObj.setPtwSubtypeFilter('Marine Work Permit');
  byId['ptw-subtabs'].children[9].oncontextmenu({ preventDefault() {}, clientX: 200, clientY: 80 });
  byId['subtab-ctx-menu'].children[0].onclick({ stopPropagation() {} });
  assert.strictEqual(byId['add-subtype-name'].value, 'Marine Work Permit');
  assert.strictEqual(byId['add-subtype-code'].value, 'MWP');
  byId['add-subtype-name'].value = 'Marine Ops Permit';
  byId['add-subtype-code'].value = 'MOP';
  await apiObj.saveNewSubtype();
  assert.strictEqual(patchedLists.length, 1);
  assert.strictEqual(patchedLists[0].old_item, 'Marine Work Permit');
  assert.strictEqual(patchedLists[0].new_item, 'Marine Ops Permit');
  assert.strictEqual(patchedLists[0].meta, 'MOP');
  assert.strictEqual(apiObj.state.ptwSubtypeFilter, 'Marine Ops Permit', 'Active filter must update after rename');
  assert.strictEqual(apiObj.computeNextPtwDocNoClient('Marine Ops Permit'), 'PEM058-MOP-001');

  // Delete Sub-Tab via context menu: removes from list, resets active filter to All, keeps records intact
  const recCountBeforeDel = apiObj.state.recs.length;
  byId['ptw-subtabs'].children[9].oncontextmenu({ preventDefault() {}, clientX: 200, clientY: 80 });
  await byId['subtab-ctx-menu'].children[3].onclick({ stopPropagation() {} });
  assert.strictEqual(confirmCalls.length, 1);
  assert.strictEqual(deletedLists.length, 1);
  assert.strictEqual(deletedLists[0].item, 'Marine Ops Permit');
  assert.strictEqual(apiObj.state.ptwSubtypeFilter, '', 'Deleting active sub-tab resets filter to All');
  assert.strictEqual(apiObj.state.recs.length, recCountBeforeDel, 'Existing records must never be deleted');
  assert.strictEqual(byId['ptw-subtabs'].children.length, 10, 'Back to All + 8 default subtypes + "+" button');

  // Test generalized sub-tabs on non-PTW document type with subtypes (OBS)
  const obsCols = [
    { col_key: 'docNo', label: 'Register No.', col_type: 'docno', visible: true },
    { col_key: 'legacySubtype', label: 'Subtype', col_type: 'dropdown', list_name: 'hse_obs_subtype', visible: true },
    { col_key: 'title', label: 'Subject', col_type: 'text', visible: true },
  ];
  apiObj.state.tab = 'OBS';
  apiObj.state.cols = obsCols;
  apiObj.state.allTabCols = obsCols;
  apiObj.state.ptwSubtypeFilter = '';
  apiObj.state.filters = {};
  apiObj.state.recs = [
    { _id: 'o1', docNo: 'OBS-001', legacySubtype: 'Unsafe Act' },
    { _id: 'o2', docNo: 'OBS-002', legacySubtype: 'Near Miss' },
  ];
  apiObj.updatePtwSubTabs();
  assert.strictEqual(byId['ptw-subtabs'].style.display, 'flex', 'OBS must show generalized sub-tab strip');
  assert.strictEqual(byId['ptw-subtabs'].children.length, 5, 'OBS: All + 3 subtypes + "+" button');
  assert.strictEqual(byId['ptw-subtabs'].children[4].id, 'ptw-subtab-add-btn');

  // Filtering OBS by subtype
  apiObj.setPtwSubtypeFilter('Unsafe Act');
  assert.strictEqual(JSON.stringify(apiObj.state.visibleRows.map(r => r._id)), JSON.stringify(['o1']));

  // Opening Add form on OBS pre-selects active subtype in buildMS
  apiObj.state.editId = null;
  await apiObj.buildForm(null, { defaultSubtype: 'Unsafe Act', suggestedDocNo: 'OBS-003' });
  const obsMs = documentStub.getElementById('f-legacySubtype');
  assert.strictEqual(obsMs.dataset.value, 'Unsafe Act', 'OBS Add form must pre-select active sub-tab subtype');

  // Adding a new subtype on OBS hides PTW code input and saves cleanly
  apiObj.openAddSubtypeModal();
  assert.strictEqual(byId['add-subtype-code-wrap'].style.display, 'none', 'Non-PTW modal hides Doc No. code field');
  byId['add-subtype-name'].value = 'Environmental Hazard';
  await apiObj.saveNewSubtype();
  assert.strictEqual(postedLists[1].list_name, 'hse_obs_subtype');
  assert.strictEqual(postedLists[1].item, 'Environmental Hazard');
  assert.strictEqual(postedLists[1].meta, undefined);
  assert.strictEqual(byId['ptw-subtabs'].children.length, 6, 'OBS: All + 4 subtypes + "+" button');

  // Document type without subtypes (JSA): hides sub-tab bar & shows toolbar "+ Sub-Tab" button
  apiObj.state.tab = 'JSA';
  apiObj.state.cols = mockColsByDt.JSA;
  apiObj.state.allTabCols = mockColsByDt.JSA;
  apiObj.state.ptwSubtypeFilter = '';
  apiObj.state.filters = {};
  apiObj.state.recs = [];
  apiObj.updatePtwSubTabs();
  assert.strictEqual(byId['ptw-subtabs'].style.display, 'none', 'JSA without subtypes hides sub-tab strip');
  assert.strictEqual(byId['btn-add-subtab'].style.display, '', 'JSA without subtypes shows toolbar + Sub-Tab button');

  // Clicking toolbar "+ Sub-Tab" on JSA adds first subtype, auto-wires legacySubtype column, shows sub-tabs & hides toolbar button
  apiObj.openAddSubtypeModal();
  assert.strictEqual(byId['add-subtype-code-wrap'].style.display, 'none', 'JSA modal hides PTW code input');
  byId['add-subtype-name'].value = 'Hot Work JSA';
  await apiObj.saveNewSubtype();
  assert.strictEqual(postedLists[2].list_name, 'hse_jsa_subtype');
  assert.strictEqual(postedLists[2].item, 'Hot Work JSA');
  assert.strictEqual(postedLists[2].dt_id, 'JSA');
  assert.ok(apiObj.state.cols.some(c => c.col_key === 'legacySubtype'), 'JSA columns refreshed with auto-wired legacySubtype');
  assert.strictEqual(byId['ptw-subtabs'].style.display, 'flex', 'JSA sub-tab strip appears immediately after first subtype');
  assert.strictEqual(byId['btn-add-subtab'].style.display, 'none', 'Toolbar + Sub-Tab button hides once sub-tab strip is active');
  assert.strictEqual(byId['ptw-subtabs'].children.length, 3, 'JSA: All + Hot Work JSA + circular "+" button');
  assert.strictEqual(apiObj.state.ptwSubtypeFilter, 'Hot Work JSA', 'First subtype becomes active sub-tab');

  // Opening Add form on JSA immediately pre-selects the new subtype
  apiObj.state.editId = null;
  await apiObj.buildForm(null, { defaultSubtype: 'Hot Work JSA', suggestedDocNo: 'JSA-001' });
  const jsaMs = documentStub.getElementById('f-legacySubtype');
  assert.ok(jsaMs, 'JSA Add form must render auto-wired legacySubtype field');
  assert.strictEqual(jsaMs.dataset.value, 'Hot Work JSA', 'JSA Add form pre-selects active first subtype');

  // Document type without subtypes (TBT) hides sub-tab bar until a subtype list is added via Lists
  apiObj.state.tab = 'TBT';
  apiObj.state.cols = mockColsByDt.TBT;
  apiObj.state.allTabCols = mockColsByDt.TBT;
  apiObj.state.ptwSubtypeFilter = '';
  apiObj.state.recs = [];
  apiObj.updatePtwSubTabs();
  assert.strictEqual(byId['ptw-subtabs'].style.display, 'none', 'TBT without subtype list hides sub-tab bar');
  assert.strictEqual(byId['btn-add-subtab'].style.display, '', 'TBT without subtype list shows toolbar + Sub-Tab button');

  // Step C: adding hse_tbt_subtype via Lists automatically enables sub-tabs for TBT
  apiObj.state.lists.hse_tbt_subtype = ['Fire Safety Talk'];
  apiObj.updatePtwSubTabs();
  assert.strictEqual(byId['ptw-subtabs'].style.display, 'flex', 'Adding hse_tbt_subtype enables sub-tab bar seamlessly');
  assert.strictEqual(byId['btn-add-subtab'].style.display, 'none', 'Toolbar + Sub-Tab hides once TBT sub-tab bar appears');
  assert.strictEqual(byId['ptw-subtabs'].children.length, 3, 'TBT: All + 1 subtype + "+" button');
})().catch(e => { console.error(e); process.exit(1); });
"""
        with tempfile.TemporaryDirectory() as tmpdir:
            js_path = Path(tmpdir) / "register_main.js"
            harness_path = Path(tmpdir) / "verify.js"
            js_path.write_text(main_script, encoding="utf-8")
            harness_path.write_text(node_script, encoding="utf-8")
            proc = subprocess.run(
                ["node", str(harness_path), str(js_path)],
                capture_output=True,
                text=True,
                timeout=15,
            )
            self.assertEqual(proc.returncode, 0, f"Node JS check failed:\nSTDOUT:{proc.stdout}\nSTDERR:{proc.stderr}")


if __name__ == "__main__":
    unittest.main()
