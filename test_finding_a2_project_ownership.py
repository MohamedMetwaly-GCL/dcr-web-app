"""Finding A2 regression suite - project workspace ownership and ID collision.

Run through security_test_support.py with a fresh disposable PostgreSQL cluster.
"""

import json
import threading
import unittest
import uuid
from unittest.mock import patch
from contextlib import ExitStack

from security_test_support import disposable_database


_STATE = {}


def setUpModule():
    stack = ExitStack()
    unittest.addModuleCleanup(stack.close)
    db = stack.enter_context(disposable_database())
    # Ensure HSE workspace exists alongside default ENG and OM
    db.create_workspace("HSE", "HSE", "Health, Safety & Environment", "HSE", "🛡️", "hse", 3, "HSE Workspace")

    # Baseline synthetic projects across ENG, HSE, OM
    projects = [
        ("SYN-ENG-1", "ENG"),
        ("SYN-HSE-1", "HSE"),
        ("SYN-OM-1",  "OM"),
    ]
    for p, ws in projects:
        db.create_project(p, "Synthetic " + p, p, workspace_id=ws)

    # Preserve custom dashboard_order + extra metadata on SYN-ENG-1 and SYN-HSE-1
    db.save_project("SYN-ENG-1", "Synthetic SYN-ENG-1", "SYN-ENG-1", {
        "client": "ENG Client",
        "dashboard_order": ["DS", "RFI"],
    }, workspace_id="ENG")
    db.save_project("SYN-HSE-1", "Synthetic SYN-HSE-1", "SYN-HSE-1", {
        "client": "HSE Client",
        "dashboard_order": ["HSE_PTW"],
    }, workspace_id="HSE")

    for p, _ in projects:
        if not db.get_doc_type(p, "DS"):
            db.exe(
                "INSERT INTO doc_types(id, project_id, name, code, sort_order) VALUES(%s,%s,%s,%s,99) ON CONFLICT DO NOTHING",
                (code := "DS", p, code, code),
            )
        db.save_record(p, "DS", f"REC-{p}-DS-1", {
            "docNo": f"{p}-DS-001", "title": f"Original {p} DS", "status": "Under Review",
        })

    pw = "Synthetic#" + uuid.uuid4().hex[:8]
    user_specs = {
        "syn_super":        ("superadmin", "ENG", {"ENG": "admin", "HSE": "admin", "OM": "admin"}, ()),
        "syn_eng_admin":    ("admin",      "ENG", {"ENG": "admin"}, ()),
        "syn_hse_admin":    ("admin",      "HSE", {"HSE": "admin"}, ()),
        "syn_eng_hse_ed":   ("admin",      "ENG", {"ENG": "admin", "HSE": "editor"}, ("SYN-HSE-1",)),
        "syn_eng_hse_vw":   ("admin",      "ENG", {"ENG": "admin", "HSE": "viewer"}, ("SYN-HSE-1",)),
        "syn_ed_eng1":      ("editor",     "ENG", {"ENG": "editor"}, ("SYN-ENG-1",)),
        "syn_vw_eng1":      ("viewer",     "ENG", {"ENG": "viewer"}, ("SYN-ENG-1",)),
    }
    tokens = {}
    for uname, (role, pws, ws_roles, projs) in user_specs.items():
        db.add_user(uname, pw, role, f"{uname}@synthetic.local", pws, False)
        for ws_id, ws_role in ws_roles.items():
            db.assign_user_workspace(uname, ws_id, ws_role)
        for pid_assign in projs:
            db.assign_project(uname, pid_assign)
        tokens[uname] = db.create_session(uname, role)

    import auth
    from app import app
    app.config["TESTING"] = True
    _STATE["tokens"] = tokens
    _STATE["app"] = app
    _STATE["db"] = db
    _STATE["auth"] = auth



class A2ProjectOwnershipBase(unittest.TestCase):
    @classmethod
    def client_for(cls, username=None):
        c = _STATE["app"].test_client()
        if username:
            c.set_cookie("dcr_token", _STATE["tokens"][username])
        return c

    @property
    def db(self):
        return _STATE["db"]

    def snapshot_project_state(self, pid):
        proj = self.db.q(
            "SELECT id, name, code, workspace_id, data::text AS data FROM projects WHERE id=%s",
            (pid,),
            one=True,
        )
        doc_types = self.db.q(
            "SELECT id, project_id, name, code, sort_order FROM doc_types WHERE project_id=%s ORDER BY id",
            (pid,),
        )
        assignments = self.db.q(
            "SELECT username, project_id, allowed_doc_types::text AS allowed FROM user_projects WHERE project_id=%s ORDER BY username",
            (pid,),
        )
        records = self.db.q(
            "SELECT id, project_id, dt_id, data::text AS data FROM records WHERE project_id=%s ORDER BY id",
            (pid,),
        )
        return {
            "project": dict(proj) if proj else None,
            "doc_types": [dict(r) for r in doc_types],
            "assignments": [dict(r) for r in assignments],
            "records": [dict(r) for r in records],
        }


class TestA2ProjectOwnershipAndCollision(A2ProjectOwnershipBase):
    def test_legitimate_create_in_eng_om_hse_preserves_seeding_and_creator_assignment(self):
        c_super = self.client_for("syn_super")
        cases = [
            ("SYN-NEW-ENG", "ENG", "syn_eng_admin"),
            ("SYN-NEW-OM",  "OM",  "syn_super"),
            ("SYN-NEW-HSE", "HSE", "syn_hse_admin"),
        ]
        for pid, ws, actor in cases:
            c = self.client_for(actor)
            r = c.post("/api/projects/create", json={
                "id": pid,
                "name": f"Valid {ws} Project",
                "code": pid,
                "workspace_id": ws,
                "client": f"Client {ws}",
            })
            self.assertEqual(r.status_code, 200, (pid, r.get_json()))
            self.assertEqual(r.get_json(), {"ok": True})

            snap = self.snapshot_project_state(pid)
            self.assertEqual(snap["project"]["workspace_id"], ws)
            self.assertEqual(snap["project"]["name"], f"Valid {ws} Project")
            self.assertGreater(len(snap["doc_types"]), 0)
            self.assertIn(actor, [a["username"] for a in snap["assignments"]])

            # Verify creator sees project with can_edit=True in that workspace
            listed = c.get(f"/api/projects?workspace_id={ws}")
            self.assertEqual(listed.status_code, 200)
            by_id = {p["id"]: p for p in listed.get_json()}
            self.assertIn(pid, by_id)
            self.assertTrue(by_id[pid]["can_edit"])

        # Clean up newly created test projects via superadmin delete to keep fixture clean
        for pid, _ws, _actor in cases:
            rd = c_super.post(f"/api/projects/delete/{pid}")
            self.assertEqual(rd.status_code, 200)

    def test_legitimate_same_workspace_metadata_update_preserves_workspace_order_and_data(self):
        c = self.client_for("syn_eng_admin")
        before = self.snapshot_project_state("SYN-ENG-1")
        r = c.post("/api/project/SYN-ENG-1", json={
            "name": "SYN-ENG-1 Renamed",
            "code": "SYN-ENG-1",
            "client": "Updated ENG Client",
        })
        self.assertEqual(r.status_code, 200, r.get_json())
        self.assertEqual(r.get_json(), {"ok": True})

        after = self.snapshot_project_state("SYN-ENG-1")
        self.assertEqual(after["project"]["workspace_id"], "ENG")
        self.assertEqual(after["project"]["name"], "SYN-ENG-1 Renamed")
        pdata = json.loads(after["project"]["data"])
        self.assertEqual(pdata.get("client"), "Updated ENG Client")
        self.assertEqual(pdata.get("dashboard_order"), ["DS", "RFI"])
        self.assertEqual(after["doc_types"], before["doc_types"])
        self.assertEqual(after["assignments"], before["assignments"])
        self.assertEqual(after["records"], before["records"])

    def test_duplicate_id_same_workspace_rejected_without_overwriting_or_reseeding(self):
        before = self.snapshot_project_state("SYN-ENG-1")
        c = self.client_for("syn_eng_admin")
        r = c.post("/api/projects/create", json={
            "id": "SYN-ENG-1",
            "name": "Hijacked ENG Name",
            "code": "HIJACK",
            "workspace_id": "ENG",
            "client": "Attacker",
        })
        self.assertEqual(r.status_code, 409)
        self.assertFalse(r.get_json().get("ok"))
        self.assertEqual(self.snapshot_project_state("SYN-ENG-1"), before)

    def test_cross_workspace_create_collision_rejected_without_disclosure_or_modification(self):
        before_hse = self.snapshot_project_state("SYN-HSE-1")
        before_eng = self.snapshot_project_state("SYN-ENG-1")

        # ENG admin submits SYN-HSE-1 under workspace ENG -> must reject cleanly
        c_eng = self.client_for("syn_eng_admin")
        r1 = c_eng.post("/api/projects/create", json={
            "id": "SYN-HSE-1",
            "name": "Claimed as ENG",
            "code": "HIJACK",
            "workspace_id": "ENG",
            "client": "Attacker",
        })
        self.assertEqual(r1.status_code, 409)
        body1 = r1.get_json()
        self.assertEqual(body1, {"ok": False, "error": "Project ID already exists"})
        self.assertNotIn("HSE", json.dumps(body1))
        self.assertEqual(self.snapshot_project_state("SYN-HSE-1"), before_hse)
        self.assertIsNone(self.db.q("SELECT 1 FROM user_projects WHERE username=%s AND project_id=%s",
                                    ("syn_eng_admin", "SYN-HSE-1"), one=True))

        # HSE admin submits SYN-ENG-1 under workspace HSE -> must reject cleanly
        c_hse = self.client_for("syn_hse_admin")
        r2 = c_hse.post("/api/projects/create", json={
            "id": "SYN-ENG-1",
            "name": "Claimed as HSE",
            "code": "HIJACK",
            "workspace_id": "HSE",
            "client": "Attacker",
        })
        self.assertEqual(r2.status_code, 409)
        body2 = r2.get_json()
        self.assertEqual(body2, {"ok": False, "error": "Project ID already exists"})
        self.assertNotIn("ENG", json.dumps(body2))
        self.assertEqual(self.snapshot_project_state("SYN-ENG-1"), before_eng)

        # Superadmin creating with colliding ID also gets 409, never overwrites
        c_super = self.client_for("syn_super")
        r3 = c_super.post("/api/projects/create", json={
            "id": "SYN-OM-1",
            "name": "Super Overwrite",
            "code": "OVER",
            "workspace_id": "ENG",
        })
        self.assertEqual(r3.status_code, 409)
        self.assertEqual(self.db.q("SELECT workspace_id FROM projects WHERE id='SYN-OM-1'", one=True)["workspace_id"], "OM")

    def test_cross_workspace_save_and_delete_attempts_rejected(self):
        before_hse = self.snapshot_project_state("SYN-HSE-1")
        before_eng = self.snapshot_project_state("SYN-ENG-1")

        # ENG admin cannot save or delete SYN-HSE-1
        c_eng = self.client_for("syn_eng_admin")
        self.assertEqual(c_eng.post("/api/project/SYN-HSE-1", json={"name": "Owned", "workspace_id": "ENG"}).status_code, 403)
        self.assertEqual(c_eng.post("/api/projects/delete/SYN-HSE-1").status_code, 403)
        self.assertEqual(self.snapshot_project_state("SYN-HSE-1"), before_hse)

        # HSE admin cannot save or delete SYN-ENG-1
        c_hse = self.client_for("syn_hse_admin")
        self.assertEqual(c_hse.post("/api/project/SYN-ENG-1", json={"name": "Owned", "workspace_id": "HSE"}).status_code, 403)
        self.assertEqual(c_hse.post("/api/projects/delete/SYN-ENG-1").status_code, 403)
        self.assertEqual(self.snapshot_project_state("SYN-ENG-1"), before_eng)

        # Mixed-role user: ENG admin + HSE viewer on SYN-HSE-1 cannot save or reassign SYN-HSE-1
        c_mix_vw = self.client_for("syn_eng_hse_vw")
        self.assertEqual(c_mix_vw.post("/api/project/SYN-HSE-1", json={"name": "VwEdit", "workspace_id": "ENG"}).status_code, 403)
        self.assertEqual(self.snapshot_project_state("SYN-HSE-1"), before_hse)

        # Mixed-role user: ENG admin + HSE editor on SYN-HSE-1 can edit SYN-HSE-1 metadata in HSE,
        # but CANNOT move SYN-HSE-1 to ENG via workspace_id in payload
        c_mix_ed = self.client_for("syn_eng_hse_ed")
        r_move = c_mix_ed.post("/api/project/SYN-HSE-1", json={
            "name": "Moved To ENG",
            "code": "SYN-HSE-1",
            "workspace_id": "ENG",
        })
        self.assertEqual(r_move.status_code, 403)
        self.assertEqual(self.snapshot_project_state("SYN-HSE-1"), before_hse)

        # Same-workspace update without workspace_id succeeds and keeps HSE workspace
        r_ok = c_mix_ed.post("/api/project/SYN-HSE-1", json={
            "name": "HSE Editor Updated",
            "code": "SYN-HSE-1",
            "client": "HSE Updated Client",
        })
        self.assertEqual(r_ok.status_code, 200)
        self.assertEqual(self.db.q("SELECT workspace_id FROM projects WHERE id='SYN-HSE-1'", one=True)["workspace_id"], "HSE")

    def test_superadmin_and_anonymous_boundaries_on_project_save(self):
        before_om = self.snapshot_project_state("SYN-OM-1")
        # Anonymous user rejected before touching DB
        c_anon = self.client_for(None)
        self.assertIn(c_anon.post("/api/project/SYN-OM-1", json={"name": "Anon"}).status_code, (401, 403))
        self.assertIn(c_anon.post("/api/projects/create", json={"id": "SYN-ANON", "name": "A", "code": "A"}).status_code, (401, 403))
        self.assertEqual(self.snapshot_project_state("SYN-OM-1"), before_om)

        # Superadmin gets 404 on nonexistent project rather than silent upsert
        c_super = self.client_for("syn_super")
        r_missing = c_super.post("/api/project/SYN-DOES-NOT-EXIST", json={"name": "Phantom", "workspace_id": "ENG"})
        self.assertEqual(r_missing.status_code, 404)
        self.assertIsNone(self.db.q("SELECT 1 FROM projects WHERE id='SYN-DOES-NOT-EXIST'", one=True))

        # Superadmin cannot silently reassign workspace via /api/project/<pid>
        r_move = c_super.post("/api/project/SYN-OM-1", json={"name": "MoveOM", "code": "SYN-OM-1", "workspace_id": "ENG"})
        self.assertEqual(r_move.status_code, 403)
        self.assertEqual(self.snapshot_project_state("SYN-OM-1"), before_om)

        # Superadmin same-workspace edit succeeds and retains OM
        r_same = c_super.post("/api/project/SYN-OM-1", json={"name": "OM Updated", "code": "SYN-OM-1", "workspace_id": "OM"})
        self.assertEqual(r_same.status_code, 200)
        self.assertEqual(self.db.q("SELECT workspace_id, name FROM projects WHERE id='SYN-OM-1'", one=True),
                         {"workspace_id": "OM", "name": "OM Updated"})

    def test_db_layer_ownership_errors_and_atomic_stale_precheck_guard(self):
        before_hse = self.snapshot_project_state("SYN-HSE-1")

        # Direct db.create_project collision raises ProjectOwnershipError
        with self.assertRaises(self.db.ProjectOwnershipError):
            self.db.create_project("SYN-HSE-1", "Direct Collide", "SYN-HSE-1", creator="syn_eng_admin", workspace_id="ENG")
        self.assertEqual(self.snapshot_project_state("SYN-HSE-1"), before_hse)

        # Direct db.save_project workspace mismatch raises ProjectOwnershipError
        with self.assertRaises(self.db.ProjectOwnershipError):
            self.db.save_project("SYN-HSE-1", "Direct Move", "SYN-HSE-1", {"client": "X"}, workspace_id="ENG")
        self.assertEqual(self.snapshot_project_state("SYN-HSE-1"), before_hse)

        # Simulated TOCTOU / stale precheck: db.q returns None for stored project,
        # so the atomic SQL ON CONFLICT ... WHERE COALESCE(projects.workspace_id, 'ENG') = EXCLUDED.workspace_id
        # must still reject the cross-workspace overwrite at the PostgreSQL layer.
        orig_q = self.db.q
        with patch.object(self.db, "q", return_value=None):
            with self.assertRaises(self.db.ProjectOwnershipError):
                self.db.save_project("SYN-HSE-1", "Race Move", "SYN-HSE-1", {"client": "Race"}, workspace_id="ENG")
        self.assertIsNotNone(orig_q("SELECT id FROM projects WHERE id='SYN-HSE-1'", one=True))
        self.assertEqual(self.snapshot_project_state("SYN-HSE-1"), before_hse)

    def test_concurrent_project_creation_collision_allows_only_one_winner(self):
        race_pid = "SYN-RACE-A2"
        barrier = threading.Barrier(2)
        results = {}

        def worker(username, ws):
            c = self.client_for(username)
            barrier.wait(timeout=10)
            resp = c.post("/api/projects/create", json={
                "id": race_pid,
                "name": f"Race {ws}",
                "code": race_pid,
                "workspace_id": ws,
            })
            results[username] = (resp.status_code, resp.get_json())

        t1 = threading.Thread(target=worker, args=("syn_eng_admin", "ENG"))
        t2 = threading.Thread(target=worker, args=("syn_hse_admin", "HSE"))
        t1.start()
        t2.start()
        t1.join(timeout=15)
        t2.join(timeout=15)

        statuses = sorted(v[0] for v in results.values())
        self.assertEqual(statuses, [200, 409], results)

        winner_ws = "ENG" if results["syn_eng_admin"][0] == 200 else "HSE"
        winner_user = "syn_eng_admin" if winner_ws == "ENG" else "syn_hse_admin"
        loser_user = "syn_hse_admin" if winner_ws == "ENG" else "syn_eng_admin"

        snap = self.snapshot_project_state(race_pid)
        self.assertEqual(snap["project"]["workspace_id"], winner_ws)
        self.assertEqual([a["username"] for a in snap["assignments"]], [winner_user])
        self.assertIsNone(self.db.q("SELECT 1 FROM user_projects WHERE username=%s AND project_id=%s",
                                    (loser_user, race_pid), one=True))

        # Clean up race project via superadmin
        rc = self.client_for("syn_super").post(f"/api/projects/delete/{race_pid}")
        self.assertEqual(rc.status_code, 200)


if __name__ == "__main__":
    unittest.main()
