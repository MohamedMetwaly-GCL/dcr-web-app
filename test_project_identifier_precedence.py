"""Defensive project-resolution checks on a fresh offline PostgreSQL cluster."""

from contextlib import ExitStack
from pathlib import Path
import sys
import unittest
from unittest.mock import patch


class ProjectIdentifierPrecedenceTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        from security_test_support import disposable_database

        stack = ExitStack()
        cls.addClassCleanup(stack.close)
        cls.db = stack.enter_context(disposable_database())

        import auth
        from app import app

        cls.auth = auth
        cls.app = app
        for pid, code in (
            ("SYN-Exact", "Primary"),
            ("SYN-Other", "SYN-Exact"),
            ("SYN-Fold", "Shared Alias"),
            ("syn-fold", "Shared Alias"),
        ):
            cls.db.create_project(pid, f"Project {pid}", code, workspace_id="ENG")
        cls.db.exe(
            "INSERT INTO users (username, pw_hash, role, primary_workspace) "
            "VALUES ('syn_precedence', 'unused-synthetic-fixture', 'viewer', 'ENG')"
        )
        cls.db.exe(
            "INSERT INTO user_projects (username, project_id, allowed_doc_types) "
            "VALUES ('syn_precedence', 'Primary', '[\"DS\"]'::jsonb)"
        )
        cls.db.exe(
            "INSERT INTO doc_types (id, project_id, name, code) VALUES "
            "('DS', 'SYN-Exact', 'Synthetic DS', 'DS'), ('DS', 'SYN-Other', 'Synthetic DS', 'DS') "
            "ON CONFLICT DO NOTHING"
        )
        cls.db.save_record("SYN-Exact", "DS", "SYN-PRECEDENCE-ALLOWED", {"docNo": "SYN-ALLOWED"})
        cls.db.save_record("SYN-Other", "DS", "SYN-PRECEDENCE-OTHER", {"docNo": "SYN-OTHER"})

    def test_project_id_precedence(self):
        for raw, expected in (
            ("SYN-Exact", ["SYN-Exact"]),
            ("sYn-eXAct", ["SYN-Exact"]),
            ("SYN-Fold", ["SYN-Fold"]),
            ("sYn-FoLd", []),
            ("Shared Alias", []),
            ("Primary", ["SYN-Exact"]),
        ):
            with self.subTest(identifier=raw):
                self.assertEqual(self.db.resolve_project_identifiers([raw]), expected)

    def test_ambiguous_ids_do_not_consult_alias_index(self):
        class AliasLookupForbidden(dict):
            def get(self, key, default=None):
                raise AssertionError("Ambiguous IDs must return before alias lookup")

        exact, folded, _ = self.db._build_project_resolution_index()
        with patch.object(self.db, "_build_project_resolution_index",
                          return_value=(exact, folded, AliasLookupForbidden())):
            self.assertEqual(self.db.resolve_project_identifiers(["sYn-FoLd"]), [])
            self.assertFalse(self.db.resolve_user_project_grant("syn_precedence", "sYn-FoLd")["assigned"])

    def test_unique_legacy_grant_agrees_with_membership_and_record_scope(self):
        user = {"username": "syn_precedence", "role": "viewer"}
        self.assertEqual(self.auth.get_allowed_project_ids(user), ["SYN-Exact"])
        self.assertEqual(self.auth.get_allowed_doc_type_filter("SYN-Exact", user), ["DS"])
        self.assertEqual(self.auth.get_allowed_doc_type_filter("SYN-Other", user), [])

        client = self.app.test_client()
        client.set_cookie("dcr_token", self.db.create_session(user["username"], user["role"]))
        allowed = client.get("/api/records/SYN-Exact/DS")
        self.assertEqual(allowed.status_code, 200)
        self.assertEqual([record["docNo"] for record in allowed.get_json()["records"]], ["SYN-ALLOWED"])
        self.assertEqual(client.get("/api/records/SYN-Other/DS").status_code, 403)


if __name__ == "__main__":
    sys.path.insert(0, str(Path(__file__).resolve().parent))
    import security_test_support as runner

    name = Path(__file__).stem
    runner.SUITES = (name,)
    sys.argv.insert(1, name)
    raise SystemExit(runner.main())
