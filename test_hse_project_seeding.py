"""HSE foundations checks for the existing offline disposable PostgreSQL runner."""

from contextlib import ExitStack
from pathlib import Path
import sys
import unittest
from unittest.mock import patch


class HseProjectSeedingTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        from security_test_support import disposable_database

        stack = ExitStack()
        cls.addClassCleanup(stack.close)
        cls.db = stack.enter_context(disposable_database())

    def test_hse_creation_seeds_catalog_and_creator_assignment(self):
        from hse_defaults import HSE_LISTS, HSE_SUBTYPES, hse_columns

        db = self.db
        from app import app

        db.add_user("syn_hse_01_admin", "synthetic-hse-01", primary_workspace="HSE")
        db.assign_user_workspace("syn_hse_01_admin", "HSE", "admin")
        client = app.test_client()
        client.set_cookie("dcr_token", db.create_session("syn_hse_01_admin", "viewer"))
        pid = "SYN-HSE-01-CATALOG"
        payload = {"id": pid, "name": "Synthetic catalog", "code": "HSE-01", "workspace_id": "HSE"}
        response = client.post("/api/projects/create", json=payload)
        self.assertEqual(response.status_code, 200)
        duplicate = client.post("/api/projects/create", json=payload)
        self.assertEqual(duplicate.status_code, 409)
        self.assertEqual(db.get_project(pid)["workspace_id"], "HSE")
        self.assertEqual([r["project_id"] for r in db.get_user_projects("syn_hse_01_admin")], [pid])

        doc_types = db.get_doc_types(pid)
        self.assertEqual([r["id"] for r in doc_types],
                         ["PTW", "OBS", "INC", "SVP", "TBT", "FEI", "JSA", "CEV", "EVAC", "MHR", "WKR"])
        for doc_type in doc_types:
            code = doc_type["id"]
            self.assertTrue(doc_type["legacy_source"])
            self.assertEqual(doc_type["legacy_subtypes"], HSE_SUBTYPES.get(code, {}))
            actual = [
                (c["col_key"], c["label"], c["col_type"], c["list_name"])
                for c in db.get_columns(pid, code, visible_only=True)
            ]
            self.assertEqual(actual, list(hse_columns(code)))
            self.assertEqual(db.count_records(pid, code), 0)
        self.assertEqual(db.get_lists(pid), {key: list(values) for key, values in HSE_LISTS.items()})
        self.assertEqual(db.get_lists(pid)["hse_doc_status"], ["Open", "Closed"])
        for doc_type in doc_types:
            col_keys = [c["col_key"] for c in db.get_columns(pid, doc_type["id"], visible_only=True)]
            self.assertIn("fileLocation", col_keys)

    def test_hse_records_and_document_grants_remain_project_scoped(self):
        db = self.db
        from app import app

        hse = "SYN-HSE-01-RECORD"
        other = "SYN-HSE-01-OTHER"
        om = "SYN-OM-01-RECORD"
        for pid, workspace in ((hse, "HSE"), (other, "HSE"), (om, "OM")):
            db.create_project(pid, pid, pid, workspace_id=workspace)
        db.add_user("syn_hse_01_editor", "synthetic-hse-01", role="editor", primary_workspace="HSE")
        db.assign_project("syn_hse_01_editor", hse)
        db.set_user_project_doc_types("syn_hse_01_editor", hse, ["INC", "JSA", "EVAC"])
        client = app.test_client()
        client.set_cookie("dcr_token", db.create_session("syn_hse_01_editor", "editor"))
        response = client.post(f"/api/records/{hse}/INC", json={
            "docNo": "SYN-HSE-INC-001", "title": "Synthetic incident record",
        })
        self.assertEqual(response.status_code, 200)
        records = client.get(f"/api/records/{hse}/INC").get_json()["records"]
        self.assertEqual([r["docNo"] for r in records], ["SYN-HSE-INC-001"])
        self.assertFalse(records[0].get("injuryClass"))
        self.assertFalse(records[0].get("status"))
        for code, field, sent in (
            ("JSA", "status", {}),
            ("JSA", "status", {"status": ""}),
            ("EVAC", "evaluationGrade", {}),
            ("EVAC", "evaluationGrade", {"evaluationGrade": ""}),
        ):
            with self.subTest(register=code, field=field, explicit_blank=bool(sent)):
                path = f"/api/records/{hse}/{code}"
                reference = f"SYN-{code}-{len(sent)}"
                created = client.post(path, json={"docNo": reference, "title": "Synthetic assessment", **sent})
                self.assertEqual(created.status_code, 200)
                loaded = client.get(path)
                self.assertEqual(loaded.status_code, 200)
                record = next(r for r in loaded.get_json()["records"] if r["docNo"] == reference)
                self.assertIn(record.get(field), (None, ""))
        for path in (f"/api/records/{hse}/PTW", f"/api/records/{other}/INC", f"/api/records/{om}/INC"):
            self.assertEqual(client.get(path).status_code, 403)

    def test_eng_and_om_keep_their_existing_seed_catalogs(self):
        db = self.db
        for workspace, catalog in (("ENG", db.DEFAULT_DOC_TYPES), ("OM", db.DEFAULT_OM_DOC_TYPES)):
            with self.subTest(workspace=workspace):
                pid = f"SYN-HSE-01-BASELINE-{workspace}"
                db.create_project(pid, pid, pid, workspace_id=workspace)
                types = db.get_doc_types(pid)
                self.assertEqual(
                    [(dt["id"], dt["name"], dt["sort_order"]) for dt in types], list(catalog)
                )
                self.assertTrue(all("legacy_source" not in dt for dt in types))
                lists = db.get_lists(pid)
                self.assertFalse(any(name.startswith("hse_") for name in lists))
                columns = db.DEFAULT_OM_COLS if workspace == "OM" else db._doc_type_col_specs(*catalog[0][:2])
                self.assertEqual(
                    [c["col_key"] for c in db.get_columns(pid, catalog[0][0])],
                    [c[0] for c in columns],
                )
                if workspace == "OM":
                    self.assertEqual(lists, db.DEFAULT_OM_LISTS)
                else:
                    for name, values in db.DEFAULT_LISTS.items():
                        self.assertTrue(set(values) <= set(lists[name]), name)

    def test_hse_seeding_rolls_back_with_project_creation(self):
        db = self.db
        pid = "SYN-HSE-01-ROLLBACK"
        seed = db._seed_hse_project

        def fail_after_seed(project_id):
            seed(project_id)
            raise RuntimeError("synthetic seed failure")

        with patch.object(db, "_seed_hse_project", side_effect=fail_after_seed):
            with self.assertRaisesRegex(RuntimeError, "synthetic seed failure"):
                db.create_project(pid, pid, pid, workspace_id="HSE")
        self.assertIsNone(db.get_project(pid))
        self.assertEqual(db.get_doc_types(pid), [])
        self.assertEqual(db.get_columns(pid, "INC"), [])
        self.assertEqual(db.get_lists(pid), {})
        db.create_project(pid, pid, pid, workspace_id="HSE")
        self.assertEqual(len(db.get_doc_types(pid)), 11)


if __name__ == "__main__":
    sys.path.insert(0, str(Path(__file__).resolve().parent))
    import security_test_support as runner

    module = Path(__file__).stem
    runner.SUITES = (module,)
    sys.argv.insert(1, module)
    raise SystemExit(runner.main())

