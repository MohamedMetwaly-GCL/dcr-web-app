"""Finding A1 regression suite - cross-workspace project authorization.

Run through security_test_support.py with a fresh disposable PostgreSQL cluster.
"""

import json
import unittest
import uuid
from contextlib import ExitStack

from security_test_support import disposable_database


_STATE = {}


def setUpModule():
    stack = ExitStack()
    unittest.addModuleCleanup(stack.close)
    db = stack.enter_context(disposable_database())
    # Ensure HSE workspace exists and is active alongside default ENG and OM
    db.create_workspace("HSE", "HSE", "Health, Safety & Environment", "HSE", "🛡️", "hse", 3, "HSE Workspace")

    # Synthetic projects across ENG, HSE, OM
    projects = [
        ("SYN-ENG-1", "ENG"),
        ("SYN-ENG-2", "ENG"),
        ("SYN-HSE-1", "HSE"),
        ("SYN-HSE-2", "HSE"),
        ("SYN-OM-1",  "OM"),
    ]
    for p, ws in projects:
        db.create_project(p, "Synthetic " + p, p, workspace_id=ws)

    # Ensure each synthetic project has at least DS and RFI doc types for cross-type/restricted tests
    for p, _ in projects:
        for code in ("DS", "RFI"):
            if not db.get_doc_type(p, code):
                db.exe(
                    "INSERT INTO doc_types(id, project_id, name, code, sort_order) VALUES(%s,%s,%s,%s,99) ON CONFLICT DO NOTHING",
                    (code, p, code, code),
                )

    # Synthetic records in each project
    for p, _ in projects:
        db.save_record(p, "DS", f"REC-{p}-DS-1", {
            "docNo": f"{p}-DS-001", "title": f"Original {p} DS", "status": "Under Review",
        })
        db.save_record(p, "RFI", f"REC-{p}-RFI-1", {
            "docNo": f"{p}-RFI-001", "title": f"Original {p} RFI", "status": "Under Review",
        })

    # Synthetic users covering every A1 role / workspace / project-assignment combination
    pw = "Synthetic#" + uuid.uuid4().hex[:8]
    user_specs = {
        "syn_super":        ("superadmin", "ENG", {"ENG": "admin", "HSE": "admin", "OM": "admin"}, ()),
        "syn_eng_admin":    ("admin",      "ENG", {"ENG": "admin"}, ()),
        "syn_eng_admin_legacy": ("admin",  "ENG", {}, ()),  # primary ENG admin with no user_workspaces row
        "syn_hse_admin":    ("admin",      "HSE", {"HSE": "admin"}, ()),
        "syn_eng_hse_ed":   ("admin",      "ENG", {"ENG": "admin", "HSE": "editor"}, ("SYN-HSE-1",)),
        "syn_eng_hse_ed_dt":("admin",      "ENG", {"ENG": "admin", "HSE": "editor"}, ("SYN-HSE-1",)),
        "syn_eng_hse_vw":   ("admin",      "ENG", {"ENG": "admin", "HSE": "viewer"}, ("SYN-HSE-1",)),
        "syn_eng_hse_norow":("admin",      "ENG", {"ENG": "admin"}, ("SYN-HSE-1",)),  # cross-ws fallback trap
        "syn_ed_eng1":      ("editor",     "ENG", {"ENG": "editor"}, ("SYN-ENG-1",)),
        "syn_ed_eng_norow": ("editor",     "ENG", {"ENG": "editor"}, ("SYN-ENG-1", "SYN-HSE-1")),
        "syn_vw_eng1":      ("viewer",     "ENG", {"ENG": "viewer"}, ("SYN-ENG-1",)),
        "syn_ed_hse_promoted": ("editor",  "ENG", {"ENG": "editor", "HSE": "admin"}, ()),
    }
    tokens = {}
    for uname, (role, pws, ws_roles, projs) in user_specs.items():
        db.add_user(uname, pw, role, f"{uname}@synthetic.local", pws, False)
        for ws_id, ws_role in ws_roles.items():
            db.assign_user_workspace(uname, ws_id, ws_role)
        for pid_assign in projs:
            db.assign_project(uname, pid_assign)
        tokens[uname] = db.create_session(uname, role)

    # Restrict syn_eng_hse_ed_dt to DS only on SYN-HSE-1
    db.set_user_project_doc_types("syn_eng_hse_ed_dt", "SYN-HSE-1", ["DS"])

    import auth
    from app import app
    app.config["TESTING"] = True
    _STATE["tokens"] = tokens
    _STATE["app"] = app
    _STATE["db"] = db
    _STATE["auth"] = auth



class A1CrossWorkspaceBase(unittest.TestCase):
    @classmethod
    def client_for(cls, username=None):
        c = _STATE["app"].test_client()
        if username:
            c.set_cookie("dcr_token", _STATE["tokens"][username])
        return c

    @property
    def db(self):
        return _STATE["db"]

    @property
    def auth(self):
        return _STATE["auth"]

    def get_record_snapshot(self, rid):
        r = self.db.q("SELECT id, project_id, dt_id, data::text AS data FROM records WHERE id=%s", (rid,), one=True)
        return dict(r) if r else None

    def get_project_snapshot(self, pid):
        r = self.db.q("SELECT id, name, code, workspace_id, data::text AS data FROM projects WHERE id=%s", (pid,), one=True)
        return dict(r) if r else None


class TestA1CrossWorkspaceAuthorization(A1CrossWorkspaceBase):
    def test_authorized_eng_admin_access_is_limited_to_eng(self):
        c = self.client_for("syn_eng_admin")
        # Listing ENG shows all ENG projects with edit rights
        r_eng = c.get("/api/projects?workspace_id=ENG")
        self.assertEqual(r_eng.status_code, 200)
        self.assertEqual({(p["id"], p["can_edit"]) for p in r_eng.get_json()},
                         {("SYN-ENG-1", True), ("SYN-ENG-2", True)})

        # Unscoped listing only contains ENG projects
        r_all = c.get("/api/projects")
        self.assertEqual(r_all.status_code, 200)
        self.assertEqual({p["id"] for p in r_all.get_json()}, {"SYN-ENG-1", "SYN-ENG-2"})

        # HSE & OM listings are empty
        self.assertEqual(c.get("/api/projects?workspace_id=HSE").get_json(), [])
        self.assertEqual(c.get("/api/projects?workspace_id=OM").get_json(), [])

        # Can read and edit records/settings in ENG
        self.assertEqual(c.get("/api/project/SYN-ENG-1").status_code, 200)
        self.assertEqual(c.get("/api/records/SYN-ENG-2/DS").status_code, 200)
        new_id = "REC-ENG-CREATED-" + uuid.uuid4().hex[:6]
        w = c.post("/api/records/SYN-ENG-1/DS", json={"_id": new_id, "docNo": "SYN-NEW-1", "title": "ENG Admin Edit"})
        self.assertEqual(w.status_code, 200)
        stored = self.get_record_snapshot(new_id)
        self.assertIsNotNone(stored)
        self.assertEqual((stored["project_id"], json.loads(stored["data"])["title"]), ("SYN-ENG-1", "ENG Admin Edit"))

        # Denied on HSE & OM projects (direct read and write) and DB stays unchanged
        hse_rec_before = self.get_record_snapshot("REC-SYN-HSE-1-DS-1")
        om_proj_before = self.get_project_snapshot("SYN-OM-1")
        for target in ("SYN-HSE-1", "SYN-HSE-2", "SYN-OM-1"):
            with self.subTest(target=target):
                self.assertEqual(c.get(f"/api/project/{target}").status_code, 403)
                self.assertEqual(c.get(f"/api/records/{target}/DS").status_code, 403)
                self.assertEqual(c.get(f"/api/records_range/{target}/DS").status_code, 403)
                self.assertEqual(c.get(f"/api/next_doc_no/{target}/DS").status_code, 403)
                self.assertEqual(c.get(f"/api/counts/{target}").status_code, 403)
                self.assertEqual(c.get(f"/api/distribution/{target}").status_code, 403)
                self.assertEqual(c.get(f"/api/project_users/{target}").status_code, 403)
                self.assertEqual(c.get(f"/app?p={target}").status_code, 403)
                self.assertEqual(c.post(f"/api/records/{target}/DS",
                                        json={"docNo": "INTRUSION", "title": "Bad"}).status_code, 403)
                self.assertEqual(c.post(f"/api/project/{target}", json={"name": "Hijacked"}).status_code, 403)
                self.assertEqual(c.post(f"/api/project/{target}/settings", json={"custom_x": "bad"}).status_code, 403)
                self.assertEqual(c.post(f"/api/distribution/{target}",
                                        json={"doc_type_id": "DS", "event_type": "submit", "emails": ["x@x.com"]}).status_code, 403)
                self.assertEqual(c.post(f"/api/lists_meta/{target}",
                                        json={"list_name": "status", "item_value": "Open", "meta": "approved"}).status_code, 403)
                self.assertEqual(c.post(f"/api/col_width/{target}/DS",
                                        json={"col_key": "docNo", "width_px": 250}).status_code, 403)
                self.assertEqual(c.get(f"/api/audit-sequence?project_id={target}&dt_id=DS").status_code, 403)

        self.assertEqual(self.get_record_snapshot("REC-SYN-HSE-1-DS-1"), hse_rec_before)
        self.assertEqual(self.get_project_snapshot("SYN-OM-1"), om_proj_before)
        # Reorder attempting cross-workspace project must fail without modifying order
        self.assertEqual(c.post("/api/projects/reorder", json={"order": ["SYN-ENG-1", "SYN-HSE-1"]}).status_code, 403)

    def test_hse_admin_is_limited_to_hse_and_cannot_touch_eng_or_om(self):
        c = self.client_for("syn_hse_admin")
        self.assertEqual({p["id"] for p in c.get("/api/projects").get_json()}, {"SYN-HSE-1", "SYN-HSE-2"})
        self.assertEqual(c.get("/api/projects?workspace_id=ENG").get_json(), [])
        self.assertEqual(c.get("/api/project/SYN-HSE-1").status_code, 200)
        self.assertEqual(c.get("/api/records/SYN-HSE-2/DS").status_code, 200)

        # Can edit HSE project & records
        r_ok = c.post("/api/records/SYN-HSE-2/DS", json={"_id": "REC-SYN-HSE-2-DS-1", "docNo": "SYN-HSE-2-DS-001", "title": "HSE Admin Updated"})
        self.assertEqual(r_ok.status_code, 200)
        self.assertEqual(json.loads(self.get_record_snapshot("REC-SYN-HSE-2-DS-1")["data"])["title"], "HSE Admin Updated")

        # Denied on ENG and OM
        eng_before = self.get_record_snapshot("REC-SYN-ENG-1-DS-1")
        om_before = self.get_record_snapshot("REC-SYN-OM-1-DS-1")
        for target in ("SYN-ENG-1", "SYN-ENG-2", "SYN-OM-1"):
            self.assertEqual(c.get(f"/api/project/{target}").status_code, 403)
            self.assertEqual(c.get(f"/api/records/{target}/DS").status_code, 403)
            self.assertEqual(c.post(f"/api/records/{target}/DS", json={"docNo": "BAD", "title": "Bad"}).status_code, 403)
            self.assertEqual(c.post(f"/api/project/{target}/settings", json={"tamper": True}).status_code, 403)
        self.assertEqual(self.get_record_snapshot("REC-SYN-ENG-1-DS-1"), eng_before)
        self.assertEqual(self.get_record_snapshot("REC-SYN-OM-1-DS-1"), om_before)

    def test_eng_admin_with_hse_editor_role_assigned_vs_unassigned_and_non_admin_hse_routes(self):
        c = self.client_for("syn_eng_hse_ed")
        # Listing HSE returns ONLY assigned SYN-HSE-1 with can_edit=True
        hse_list = c.get("/api/projects?workspace_id=HSE")
        self.assertEqual(hse_list.status_code, 200)
        self.assertEqual([(p["id"], p["can_edit"]) for p in hse_list.get_json()], [("SYN-HSE-1", True)])

        # Can read and edit records in assigned SYN-HSE-1
        self.assertEqual(c.get("/api/project/SYN-HSE-1").status_code, 200)
        r_save = c.post("/api/records/SYN-HSE-1/DS", json={
            "_id": "REC-SYN-HSE-1-DS-1", "docNo": "SYN-HSE-1-DS-001", "title": "Edited by HSE Editor",
        })
        self.assertEqual(r_save.status_code, 200)
        self.assertEqual(json.loads(self.get_record_snapshot("REC-SYN-HSE-1-DS-1")["data"])["title"], "Edited by HSE Editor")

        # Cannot access unassigned SYN-HSE-2 at all (read or write)
        hse2_before = self.get_record_snapshot("REC-SYN-HSE-2-DS-1")
        self.assertEqual(c.get("/api/project/SYN-HSE-2").status_code, 403)
        self.assertEqual(c.get("/api/records/SYN-HSE-2/DS").status_code, 403)
        self.assertEqual(c.post("/api/records/SYN-HSE-2/DS", json={"docNo": "BAD", "title": "Bad"}).status_code, 403)
        self.assertEqual(self.get_record_snapshot("REC-SYN-HSE-2-DS-1"), hse2_before)

        # Even on assigned SYN-HSE-1, ENG admin privileges MUST NOT leak to HSE admin-only endpoints
        proj_before = self.get_project_snapshot("SYN-HSE-1")
        self.assertEqual(c.post("/api/project/SYN-HSE-1/settings", json={"hijack": True}).status_code, 403)
        self.assertEqual(c.post("/api/lists_meta/SYN-HSE-1", json={"list_name": "status", "item_value": "X", "meta": "approved"}).status_code, 403)
        self.assertEqual(c.post("/api/col_width/SYN-HSE-1/DS", json={"col_key": "docNo", "width_px": 300}).status_code, 403)
        self.assertEqual(c.post("/api/distribution/SYN-HSE-1", json={"doc_type_id": "DS", "event_type": "submit", "emails": ["a@b.c"]}).status_code, 403)
        self.assertEqual(c.get("/api/audit-sequence?project_id=SYN-HSE-1&dt_id=DS").status_code, 403)
        self.assertEqual(c.get("/api/audit?pid=SYN-HSE-1").get_json()["rows"], [])
        self.assertEqual(self.get_project_snapshot("SYN-HSE-1"), proj_before)

    def test_eng_admin_with_hse_editor_respects_hse_doc_type_restrictions(self):
        c = self.client_for("syn_eng_hse_ed_dt")
        # DS is allowed
        self.assertEqual(c.get("/api/records/SYN-HSE-1/DS").status_code, 200)
        self.assertEqual([d["id"] for d in c.get("/api/doc_types/SYN-HSE-1").get_json()], ["DS"])
        self.assertEqual([d["id"] for d in c.get("/api/doc_types/SYN-HSE-1?all=1").get_json()], ["DS"])

        # RFI is denied for read, write, delete, next_doc_no, and records_range
        rfi_before = self.get_record_snapshot("REC-SYN-HSE-1-RFI-1")
        self.assertEqual(c.get("/api/records/SYN-HSE-1/RFI").status_code, 403)
        self.assertEqual(c.get("/api/next_doc_no/SYN-HSE-1/RFI").status_code, 403)
        self.assertEqual(c.get("/api/records_range/SYN-HSE-1/RFI").status_code, 403)
        self.assertEqual(c.post("/api/records/SYN-HSE-1/RFI", json={"docNo": "BAD-RFI", "title": "Denied"}).status_code, 403)
        self.assertEqual(c.delete("/api/records/REC-SYN-HSE-1-RFI-1").status_code, 403)
        self.assertEqual(self.get_record_snapshot("REC-SYN-HSE-1-RFI-1"), rfi_before)

    def test_cross_workspace_role_fallback_does_not_leak_admin_or_editor_into_hse(self):
        # User is primary ENG admin, assigned SYN-HSE-1 in user_projects, with NO user_workspaces row for HSE
        self.assertEqual(self.db.get_user_workspace_role("syn_eng_hse_norow", "HSE"), "viewer")
        self.assertFalse(self.db.is_user_workspace_admin("syn_eng_hse_norow", "HSE"))
        self.assertFalse(self.auth.can_edit("SYN-HSE-1", {"username": "syn_eng_hse_norow", "role": "admin"}))

        c = self.client_for("syn_eng_hse_norow")
        self.assertEqual(c.get("/api/project/SYN-HSE-1").status_code, 200)
        self.assertEqual(c.post("/api/records/SYN-HSE-1/DS", json={"docNo": "BAD", "title": "Should Fail"}).status_code, 403)
        self.assertEqual(c.post("/api/project/SYN-HSE-1", json={"name": "Should Fail"}).status_code, 403)

        # User is ENG editor assigned SYN-HSE-1 without an HSE workspace row -> HSE role defaults to viewer
        self.assertEqual(self.db.get_user_workspace_role("syn_ed_eng_norow", "HSE"), "viewer")
        ce = self.client_for("syn_ed_eng_norow")
        self.assertEqual(ce.get("/api/project/SYN-HSE-1").status_code, 200)
        self.assertEqual(ce.post("/api/records/SYN-HSE-1/DS", json={"docNo": "BAD", "title": "No Edit"}).status_code, 403)

    def test_eng_admin_with_hse_viewer_role_is_read_only_on_assigned_hse_only(self):
        c = self.client_for("syn_eng_hse_vw")
        hse_list = c.get("/api/projects?workspace_id=HSE").get_json()
        self.assertEqual([(p["id"], p["can_edit"]) for p in hse_list], [("SYN-HSE-1", False)])
        self.assertEqual(c.get("/api/project/SYN-HSE-1").status_code, 200)
        self.assertEqual(c.get("/api/records/SYN-HSE-1/DS").status_code, 200)
        self.assertEqual(c.get("/api/project/SYN-HSE-2").status_code, 403)

        rec_before = self.get_record_snapshot("REC-SYN-HSE-1-DS-1")
        self.assertEqual(c.post("/api/records/SYN-HSE-1/DS", json={"_id": "REC-SYN-HSE-1-DS-1", "title": "Blocked"}).status_code, 403)
        self.assertEqual(c.delete("/api/records/REC-SYN-HSE-1-DS-1").status_code, 403)
        self.assertEqual(self.get_record_snapshot("REC-SYN-HSE-1-DS-1"), rec_before)

    def test_mixed_role_editor_promoted_to_hse_admin_works_without_leaking_eng(self):
        # users.role == 'editor' in ENG, granted HSE admin in user_workspaces
        self.assertEqual(self.db.get_user_workspace_role("syn_ed_hse_promoted", "HSE"), "admin")
        self.assertTrue(self.db.is_user_workspace_admin("syn_ed_hse_promoted", "HSE"))
        self.assertFalse(self.db.is_user_workspace_admin("syn_ed_hse_promoted", "ENG"))
        c = self.client_for("syn_ed_hse_promoted")
        self.assertEqual({p["id"] for p in c.get("/api/projects?workspace_id=HSE").get_json()}, {"SYN-HSE-1", "SYN-HSE-2"})
        self.assertEqual(c.get("/api/projects?workspace_id=ENG").get_json(), [])
        self.assertEqual(c.get("/api/project/SYN-ENG-1").status_code, 403)
        self.assertEqual(c.post("/api/project/SYN-HSE-2/settings", json={"hse_ok": True}).status_code, 200)
        self.assertTrue(json.loads(self.get_project_snapshot("SYN-HSE-2")["data"])["hse_ok"])

    def test_allowed_project_ids_and_global_search_never_cross_workspaces(self):
        # Legacy ENG admin with no user_workspaces row stays confined to ENG
        eng_legacy = {"username": "syn_eng_admin_legacy", "role": "admin"}
        self.assertEqual(set(self.auth.get_allowed_project_ids(eng_legacy)), {"SYN-ENG-1", "SYN-ENG-2"})
        self.assertEqual(self.auth.get_allowed_project_ids(eng_legacy, workspace_id="HSE"), [])
        self.assertEqual(self.auth.get_allowed_project_ids(eng_legacy, workspace_id="OM"), [])

        # Mixed ENG admin + HSE editor assigned SYN-HSE-1 requesting HSE scope must not receive ENG ids
        mixed = {"username": "syn_eng_hse_ed", "role": "admin"}
        self.assertEqual(self.auth.get_allowed_project_ids(mixed, workspace_id="HSE"), ["SYN-HSE-1"])

        # Global search as ENG admin with workspace_id=HSE returns zero results
        c = self.client_for("syn_eng_admin")
        r_search = c.get("/api/records/search/global?q=Original&workspace_id=HSE")
        self.assertEqual(r_search.status_code, 200)
        self.assertEqual(r_search.get_json()["results"], [])

        # Global search as ENG admin without workspace filter excludes HSE and OM records
        r_all = c.get("/api/records/search/global?q=Original")
        self.assertEqual(r_all.status_code, 200)
        pids_found = {row["project_id"] for row in r_all.get_json()["results"]}
        self.assertTrue(pids_found <= {"SYN-ENG-1", "SYN-ENG-2"})
        self.assertFalse(pids_found & {"SYN-HSE-1", "SYN-HSE-2", "SYN-OM-1"})

    def test_viewer_read_only_restrictions(self):
        c = self.client_for("syn_vw_eng1")
        listing = c.get("/api/projects").get_json()
        self.assertEqual([(p["id"], p["can_edit"]) for p in listing], [("SYN-ENG-1", False)])
        self.assertEqual(c.get("/api/project/SYN-ENG-1").status_code, 200)
        self.assertEqual(c.get("/api/records/SYN-ENG-1/DS").status_code, 200)
        self.assertEqual(c.get("/api/project/SYN-ENG-2").status_code, 403)
        self.assertEqual(c.get("/api/project/SYN-HSE-1").status_code, 403)

        before = self.get_record_snapshot("REC-SYN-ENG-1-DS-1")
        self.assertEqual(c.post("/api/records/SYN-ENG-1/DS", json={"_id": "REC-SYN-ENG-1-DS-1", "title": "Viewer Write"}).status_code, 403)
        self.assertEqual(c.delete("/api/records/REC-SYN-ENG-1-DS-1").status_code, 403)
        self.assertEqual(c.post("/api/project/SYN-ENG-1", json={"name": "Viewer Project Write"}).status_code, 403)
        self.assertEqual(self.get_record_snapshot("REC-SYN-ENG-1-DS-1"), before)

    def test_legitimate_global_superadmin_retains_full_cross_workspace_access(self):
        c = self.client_for("syn_super")
        all_ids = {p["id"] for p in c.get("/api/projects").get_json()}
        self.assertTrue({"SYN-ENG-1", "SYN-ENG-2", "SYN-HSE-1", "SYN-HSE-2", "SYN-OM-1"} <= all_ids)
        for pid in ("SYN-ENG-1", "SYN-HSE-2", "SYN-OM-1"):
            self.assertEqual(c.get(f"/api/project/{pid}").status_code, 200)
            self.assertEqual(c.get(f"/api/records/{pid}/DS").status_code, 200)
            rid = f"REC-{pid}-DS-1"
            r = c.post(f"/api/records/{pid}/DS", json={"_id": rid, "docNo": f"{pid}-DS-001", "title": f"Super Edited {pid}"})
            self.assertEqual(r.status_code, 200)
            self.assertEqual(json.loads(self.get_record_snapshot(rid)["data"])["title"], f"Super Edited {pid}")
            self.assertEqual(c.post(f"/api/project/{pid}/settings", json={"super_verified": True}).status_code, 200)
            self.assertTrue(json.loads(self.get_project_snapshot(pid)["data"])["super_verified"])
