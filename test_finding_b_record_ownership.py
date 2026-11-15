"""Finding B regression suite - record ownership integrity on update.

Run through security_test_support.py with a fresh disposable PostgreSQL cluster.
"""

import json
import threading
import unittest
import uuid
from contextlib import ExitStack

from security_test_support import disposable_database


_STATE = {}


def setUpModule():
    stack = ExitStack()
    unittest.addModuleCleanup(stack.close)
    db = stack.enter_context(disposable_database())
    # ---- synthetic users / projects ------------------------------------------------------------
    for p, ws in (("SYN-ENG-A", "ENG"), ("SYN-ENG-B", "ENG"), ("SYN-HSE-A", "HSE")):
        db.create_project(p, "Synthetic " + p, p, workspace_id=ws)
    pw = "Synthetic#" + uuid.uuid4().hex[:8]
    users = {
        "syn_super":      ("superadmin", "ENG", "superadmin", ()),
        "syn_eng_admin":  ("admin", "ENG", "admin", ()),
        "syn_ed_a":       ("editor", "ENG", "editor", ("SYN-ENG-A",)),
        "syn_ed_ab":      ("editor", "ENG", "editor", ("SYN-ENG-A", "SYN-ENG-B")),
        "syn_ed_dsonly":  ("editor", "ENG", "editor", ("SYN-ENG-A",)),
        "syn_viewer_a":   ("viewer", "ENG", "viewer", ("SYN-ENG-A",)),
        "syn_hse_ed":     ("editor", "HSE", "editor", ("SYN-HSE-A",)),
    }
    tokens = {}
    for name, (role, ws, ws_role, projs) in users.items():
        db.add_user(name, pw, role, None, ws, False)
        db.assign_user_workspace(name, ws, ws_role)
        for p in projs:
            db.assign_project(name, p)
        tokens[name] = db.create_session(name, role)
    db.set_user_project_doc_types("syn_ed_dsonly", "SYN-ENG-A", ["DS"])
    _STATE["tokens"] = tokens

    from app import app
    app.config["TESTING"] = True
    _STATE["app"] = app
    _STATE["db"] = db



class RecordOwnershipBase(unittest.TestCase):
    @classmethod
    def client_for(cls, username=None):
        c = _STATE["app"].test_client()
        if username:
            c.set_cookie("dcr_token", _STATE["tokens"][username])
        return c

    @property
    def db(self):
        return _STATE["db"]

    def seed(self, pid, dt, rid, **fields):
        data = {"docNo": "SYN-%s-%s" % (dt, rid[-4:]), "title": "Original " + rid, "status": "Under Review"}
        data.update(fields)
        self.db.save_record(pid, dt, rid, data)
        return rid

    def row(self, rid):
        return self.db.q("SELECT id, project_id, dt_id, data::text AS data, updated_at::text AS updated_at "
                         "FROM records WHERE id=%s", (rid,), one=True)

    def count(self, pid=None):
        if pid:
            return self.db.q("SELECT count(*) AS n FROM records WHERE project_id=%s", (pid,), one=True)["n"]
        return self.db.q("SELECT count(*) AS n FROM records", one=True)["n"]

    def audit(self, username, action):
        return self.db.q("SELECT project_id, dt_id, record_id, detail FROM audit_log "
                         "WHERE username=%s AND action=%s ORDER BY id", (username, action))

    def rid(self, tag):
        return "syn-%s-%s" % (tag, uuid.uuid4().hex[:8])


class TestRouteOwnership(RecordOwnershipBase):
    # ---- legitimate flows stay working -------------------------------------------------------
    def test_authorized_update_in_own_project_and_type(self):
        rid = self.seed("SYN-ENG-A", "DS", self.rid("own"))
        r = self.client_for("syn_ed_a").post("/api/records/SYN-ENG-A/DS",
                                             json={"_id": rid, "docNo": "SYN-DS-0001", "title": "Edited"})
        self.assertEqual((r.status_code, r.get_json()["id"]), (200, rid))
        row = self.row(rid)
        self.assertEqual((row["project_id"], row["dt_id"]), ("SYN-ENG-A", "DS"))
        self.assertEqual(json.loads(row["data"])["title"], "Edited")
        self.assertTrue(any(a["record_id"] == rid for a in self.audit("syn_ed_a", "EDIT")))

    def test_create_without_id_generates_new_record(self):
        before = self.count("SYN-ENG-A")
        c = self.client_for("syn_ed_a")
        ids = []
        for n in (1, 2):
            r = c.post("/api/records/SYN-ENG-A/DS", json={"docNo": "SYN-NEW-%d" % n, "title": "New %d" % n})
            self.assertEqual(r.status_code, 200)
            ids.append(r.get_json()["id"])
        self.assertEqual(len(set(ids)), 2)
        self.assertEqual(self.count("SYN-ENG-A"), before + 2)
        for rid in ids:
            self.assertEqual((self.row(rid)["project_id"], self.row(rid)["dt_id"]), ("SYN-ENG-A", "DS"))

    def test_client_supplied_unused_id_still_creates_in_requested_scope(self):
        rid = self.rid("chosen")
        r = self.client_for("syn_ed_a").post("/api/records/SYN-ENG-A/DS", json={"_id": rid, "docNo": "SYN-CH-1"})
        self.assertEqual((r.status_code, r.get_json()["id"]), (200, rid))
        self.assertEqual(self.row(rid)["project_id"], "SYN-ENG-A")
        # ...but once it exists it can no longer be addressed through another project.
        r2 = self.client_for("syn_ed_ab").post("/api/records/SYN-ENG-B/DS", json={"_id": rid, "docNo": "X"})
        self.assertEqual(r2.status_code, 404)
        self.assertEqual(self.row(rid)["project_id"], "SYN-ENG-A")

    def test_noc_attachments_and_data_survive_a_legitimate_edit(self):
        from attachments import build_attachment_meta
        rid = self.seed("SYN-ENG-A", "NOC", self.rid("noc"), nocNo="SYN-NOC-001")
        atts = {slot: build_attachment_meta("fake-" + slot, slot + ".pdf", 10, "application/pdf", "syn")
                for slot in ("part_a", "part_b", "part_c", "part_d", "submission")}
        self.db.exe("UPDATE records SET data = data || %s::jsonb WHERE id=%s",
                    (json.dumps({"attachments": atts}), rid))
        r = self.client_for("syn_ed_a").post("/api/records/SYN-ENG-A/NOC", json={
            "_id": rid, "docNo": "SYN-NOC-001", "title": "NOC edited",
            "attachments": {"part_a": {"file_id": "forged"}}})  # client-sent attachments must be ignored
        self.assertEqual(r.status_code, 200)
        stored = json.loads(self.row(rid)["data"])
        self.assertEqual(stored["attachments"], atts)
        self.assertEqual(stored["title"], "NOC edited")

    def test_list_and_delete_unchanged(self):
        rid = self.seed("SYN-ENG-A", "DS", self.rid("del"))
        c = self.client_for("syn_ed_a")
        listed = c.get("/api/records/SYN-ENG-A/DS").get_json()["records"]
        self.assertTrue(any(x.get("_id") == rid for x in listed))
        self.assertEqual(c.delete("/api/records/" + rid).status_code, 200)
        self.assertIsNone(self.row(rid))

    # ---- mismatches are rejected, never moved or overwritten ---------------------------------
    def assert_rejected_and_untouched(self, resp, rid, snapshot, total_before):
        self.assertEqual(resp.status_code, 404, resp.get_data(as_text=True))
        self.assertNotIn("SYN-ENG-B", resp.get_data(as_text=True))  # no leak of the real location
        self.assertEqual(self.row(rid), snapshot)
        self.assertEqual(self.count(), total_before)

    def test_editor_of_both_projects_cannot_overwrite_or_move_across_projects(self):
        rid = self.seed("SYN-ENG-B", "DS", self.rid("xb"))
        snap, total = self.row(rid), self.count()
        r = self.client_for("syn_ed_ab").post("/api/records/SYN-ENG-A/DS",
                                              json={"_id": rid, "docNo": "HIJACK", "title": "Hijacked"})
        self.assert_rejected_and_untouched(r, rid, snap, total)
        denied = [a for a in self.audit("syn_ed_ab", "ACCESS DENIED") if a["record_id"] == rid]
        self.assertTrue(denied)
        self.assertEqual((denied[-1]["project_id"], denied[-1]["dt_id"]), ("SYN-ENG-A", "DS"))

    def test_editor_without_access_to_owning_project_is_rejected(self):
        rid = self.seed("SYN-ENG-B", "DS", self.rid("xb2"))
        snap, total = self.row(rid), self.count()
        r = self.client_for("syn_ed_a").post("/api/records/SYN-ENG-A/DS", json={"_id": rid, "title": "Hijack"})
        self.assert_rejected_and_untouched(r, rid, snap, total)

    def test_same_project_wrong_document_type_is_rejected(self):
        rid = self.seed("SYN-ENG-A", "DS", self.rid("wt"))
        snap, total = self.row(rid), self.count()
        r = self.client_for("syn_ed_a").post("/api/records/SYN-ENG-A/RFI", json={"_id": rid, "title": "Retyped"})
        self.assert_rejected_and_untouched(r, rid, snap, total)

    def test_doc_type_restricted_user_cannot_overwrite_a_type_they_are_denied(self):
        rid = self.seed("SYN-ENG-A", "RFI", self.rid("rfi"))
        snap, total = self.row(rid), self.count()
        c = self.client_for("syn_ed_dsonly")
        self.assertEqual(c.post("/api/records/SYN-ENG-A/RFI", json={"title": "x"}).status_code, 403)  # existing control
        r = c.post("/api/records/SYN-ENG-A/DS", json={"_id": rid, "title": "Smuggled through allowed type"})
        self.assert_rejected_and_untouched(r, rid, snap, total)

    def test_cross_workspace_is_rejected_for_editor_and_superadmin(self):
        rid = self.seed("SYN-ENG-A", "DS", self.rid("ws"))
        snap, total = self.row(rid), self.count()
        r = self.client_for("syn_hse_ed").post("/api/records/SYN-HSE-A/DS", json={"_id": rid, "title": "Cross-ws"})
        self.assert_rejected_and_untouched(r, rid, snap, total)
        r2 = self.client_for("syn_super").post("/api/records/SYN-HSE-A/DS", json={"_id": rid, "title": "Super move"})
        self.assert_rejected_and_untouched(r2, rid, snap, total)  # superadmin may not silently move either

    def test_workspace_admin_cannot_move_record_between_projects(self):
        rid = self.seed("SYN-ENG-B", "DS", self.rid("adm"))
        snap, total = self.row(rid), self.count()
        r = self.client_for("syn_eng_admin").post("/api/records/SYN-ENG-A/DS", json={"_id": rid, "title": "Move"})
        self.assert_rejected_and_untouched(r, rid, snap, total)

    def test_viewer_and_anonymous_cannot_update_at_all(self):
        rid = self.seed("SYN-ENG-A", "DS", self.rid("vw"))
        snap = self.row(rid)
        self.assertEqual(self.client_for("syn_viewer_a").post(
            "/api/records/SYN-ENG-A/DS", json={"_id": rid, "title": "v"}).status_code, 403)
        self.assertEqual(self.client_for().post(
            "/api/records/SYN-ENG-A/DS", json={"_id": rid, "title": "anon"}).status_code, 403)
        self.assertEqual(self.row(rid), snap)

    def test_non_string_record_ids_are_rejected_without_writing(self):
        total = self.count()
        c = self.client_for("syn_ed_a")
        for bad in (123, 1.5, True, ["x"], {"a": 1}):
            with self.subTest(bad=bad):
                self.assertEqual(c.post("/api/records/SYN-ENG-A/DS", json={"_id": bad, "docNo": "B"}).status_code, 400)
        self.assertEqual(self.count(), total)

    def test_stale_precheck_is_still_stopped_by_the_atomic_sql_guard(self):
        from unittest.mock import patch
        rid = self.seed("SYN-ENG-B", "DS", self.rid("stale"))
        snap, total = self.row(rid), self.count()
        with patch.object(self.db, "get_record_by_id", return_value=None):  # simulate a lost race
            r = self.client_for("syn_ed_ab").post("/api/records/SYN-ENG-A/DS", json={"_id": rid, "title": "Race"})
        self.assert_rejected_and_untouched(r, rid, snap, total)

    def test_concurrent_creates_of_the_same_new_id_have_exactly_one_owner(self):
        for _ in range(15):
            rid = self.rid("race")
            barrier = threading.Barrier(2)
            results = {}

            def worker(user, pid):
                c = self.client_for(user)
                barrier.wait()
                results[pid] = c.post("/api/records/%s/DS" % pid, json={"_id": rid, "docNo": pid, "title": pid}).status_code

            threads = [threading.Thread(target=worker, args=("syn_ed_ab", "SYN-ENG-A")),
                       threading.Thread(target=worker, args=("syn_ed_ab", "SYN-ENG-B"))]
            [t.start() for t in threads]
            [t.join() for t in threads]
            self.assertEqual(sorted(results.values()), [200, 404], results)
            winner = next(p for p, s in results.items() if s == 200)
            row = self.row(rid)
            self.assertEqual(row["project_id"], winner)
            self.assertEqual(json.loads(row["data"])["title"], winner)


class TestDbLayerGuard(RecordOwnershipBase):
    def test_save_record_raises_on_mismatch_and_leaves_row_untouched(self):
        rid = self.seed("SYN-ENG-B", "DS", self.rid("dbl"))
        snap = self.row(rid)
        for pid, dt in (("SYN-ENG-A", "DS"), ("SYN-ENG-B", "RFI"), ("SYN-HSE-A", "DS")):
            with self.subTest(pid=pid, dt=dt):
                with self.assertRaises(self.db.RecordOwnershipError):
                    self.db.save_record(pid, dt, rid, {"docNo": "X", "title": "Overwrite"})
                self.assertEqual(self.row(rid), snap)

    def test_save_record_same_scope_update_and_new_insert_succeed(self):
        rid = self.rid("dbs")
        self.db.save_record("SYN-ENG-A", "DS", rid, {"docNo": "SYN-1", "title": "v1"})
        self.db.save_record("SYN-ENG-A", "DS", rid, {"docNo": "SYN-1", "title": "v2"})
        self.assertEqual(json.loads(self.row(rid)["data"])["title"], "v2")

    def test_import_style_update_of_existing_doc_no_still_works(self):
        rid = self.seed("SYN-ENG-A", "DS", self.rid("imp"), docNo="SYN-IMP-0001")
        existing = self.db.get_record_by_doc_no("SYN-ENG-A", "DS", "SYN-IMP-0001")
        self.db.save_record("SYN-ENG-A", "DS", existing["_id"], {"docNo": "SYN-IMP-0001", "title": "Imported"})
        self.assertEqual(json.loads(self.row(rid)["data"])["title"], "Imported")


if __name__ == "__main__":
    unittest.main()
