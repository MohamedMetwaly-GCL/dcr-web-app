"""Offline DCR authorization controls. Run with python -B -m unittest -v test_security_phase1_controls."""

import json
import os
import socket
import unittest
from contextlib import ExitStack
from copy import deepcopy
from unittest.mock import patch


class SecurityPhase1Controls(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        isolation = ExitStack()
        cls.addClassCleanup(isolation.close)
        isolation.enter_context(patch.dict(os.environ, {
            "SECRET_KEY": "synthetic-phase1-session-key",
            "MAGIC_SECRET": "synthetic-phase1-digest-key",
            "DATABASE_URL": "dbname=offline_phase1",
        }, clear=True))

        def blocked(*args, **kwargs):
            raise AssertionError("External I/O attempted in the offline security suite")

        for target in ("socket.socket.connect", "socket.socket.connect_ex",
                       "socket.socket.sendto", "socket.getaddrinfo",
                       "psycopg2.connect", "psycopg2._connect"):
            isolation.enter_context(patch(target, side_effect=blocked))

        import db

        isolation.enter_context(patch.object(db, "get_pool", side_effect=blocked))
        isolation.enter_context(patch.object(db, "init", side_effect=blocked))
        from app import app

        cls.db = db
        cls.app = app
        isolation.enter_context(patch.dict(app.config, {"TESTING": True}))

    def setUp(self):
        self.client = self.app.test_client()
        self.projects = {
            pid: {"id": pid, "workspace_id": pid.split("-")[0],
                  "name": pid, "code": pid, "data": {}}
            for pid in ("ENG-P1", "ENG-P2", "HSE-P1", "HSE-P2", "OM-P1")
        }
        self.login_as("editor", {"ENG": "editor"}, ("ENG-P1",))
        mocks = ExitStack()
        self.addCleanup(mocks.close)
        mocks.enter_context(patch.object(self.db, "q", side_effect=self.read_fixture))
        self.write = mocks.enter_context(patch.object(self.db, "exe"))
        self.seed = mocks.enter_context(patch.object(self.db, "_seed_project"))
        self.audit = mocks.enter_context(patch.object(self.db, "log_action"))

    def login_as(self, role, workspace_roles, project_ids=()):
        self.user = {"username": "synthetic-user", "role": role, "primary_workspace": "ENG"}
        self.workspace_roles = workspace_roles
        self.project_ids = project_ids
        self.client.set_cookie("dcr_token", "synthetic-session")

    def read_fixture(self, sql, params=(), one=False):
        # Reject unfamiliar queries so the suite cannot silently invent authorization data.
        query = " ".join(sql.lower().split())
        if "from sessions " in query:
            return deepcopy(self.user) if params == ("synthetic-session",) else None
        if "from users " in query:
            return deepcopy(self.user) if params == (self.user["username"],) else None
        if "from user_workspaces " in query:
            if "select 1" in query:
                return {"exists": 1} if "admin" in self.workspace_roles.values() else None
            if one:
                role = self.workspace_roles.get(params[1])
                return {"role": role} if role else None
            return [{"workspace_id": ws, "role": role} for ws, role in self.workspace_roles.items()]
        if "from user_projects " in query:
            if query.startswith("select allowed_doc_types from "):
                return {"allowed_doc_types": ["DS"]} if params[1] in self.project_ids else None
            return [{"project_id": pid} for pid in self.project_ids]
        if "from projects " in query:
            if one:
                return deepcopy(self.projects.get(params[0]))
            ids = list(params)
            projects = list(self.projects.values())
            if "coalesce(workspace_id" in query:
                workspace = ids.pop(0)
                projects = [p for p in projects if p["workspace_id"] == workspace]
            if "id in (" in query:
                projects = [p for p in projects if p["id"] in ids]
            return deepcopy(projects)
        if "from doc_types " in query:
            types = [{"id": code, "code": code, "name": code, "project_id": params[0]}
                     for code in ("DS", "RFI")]
            return next((dt for dt in types if dt["id"] == params[1]), None) if one else types
        if "from records " in query and params == ("R1",):
            return {"id": "R1", "project_id": "ENG-P1", "dt_id": "DS", "updated_at": None,
                    "data": {"title": "Original title", "attachments": {"submission": {"name": "sample.pdf"}}}}
        self.fail("Unrecognized read in the offline database fixture")

    def test_a1_superadmin_reads_each_workspace(self):
        self.login_as("superadmin", {})
        for project in ("ENG-P1", "HSE-P1", "OM-P1"):
            with self.subTest(project=project):
                self.assertEqual(self.client.get(f"/api/project/{project}").status_code, 200)

    def test_a1_eng_admin_lists_and_reads_own_project(self):
        self.login_as("admin", {"ENG": "admin"})
        response = self.client.get("/api/projects")
        self.assertEqual(response.status_code, 200)
        self.assertEqual({p["id"] for p in response.get_json()}, {"ENG-P1", "ENG-P2"})
        self.assertEqual(self.client.get("/api/project/ENG-P1").status_code, 200)

    def test_a1_mixed_role_viewer_can_read_assigned_project(self):
        self.login_as("admin", {"ENG": "admin", "HSE": "viewer"}, ("HSE-P1",))
        response = self.client.get("/api/projects?workspace_id=HSE")
        self.assertEqual(response.status_code, 200)
        self.assertEqual([p["id"] for p in response.get_json()], ["HSE-P1"])
        self.assertFalse(response.get_json()[0]["can_edit"])
        self.assertEqual(self.client.get("/api/project/HSE-P1").status_code, 200)
        self.assertEqual(self.client.post("/api/project/HSE-P1", json={"name": "New name"}).status_code, 403)
        self.write.assert_not_called()

    def test_a1_editor_cannot_read_unassigned_project(self):
        self.assertEqual(self.client.get("/api/project/ENG-P2").status_code, 403)

    def test_a1_viewer_can_read_but_cannot_save_assigned_project(self):
        self.login_as("viewer", {"ENG": "viewer"}, ("ENG-P1",))
        self.assertEqual(self.client.get("/api/project/ENG-P1").status_code, 200)
        self.assertEqual(self.client.post("/api/project/ENG-P1", json={"name": "New name"}).status_code, 403)
        self.write.assert_not_called()

    def test_a1_primary_workspace_admin_without_role_row(self):
        self.login_as("admin", {})
        self.assertTrue(self.db.is_user_workspace_admin(self.user["username"], "ENG"))
        self.assertFalse(self.db.is_user_workspace_admin(self.user["username"], "HSE"))

    def test_a2_admin_creates_project_in_own_workspace(self):
        self.login_as("admin", {"ENG": "admin"})
        response = self.client.post("/api/projects/create", json={
            "id": "ENG-NEW", "name": "Synthetic project", "code": "ENG-NEW", "workspace_id": "ENG",
        })
        self.assertEqual(response.status_code, 200)
        self.assertEqual(self.write.call_args_list[0].args[1],
                         ("ENG-NEW", "Synthetic project", "ENG-NEW", "{}", "ENG"))
        self.seed.assert_called_once_with("ENG-NEW", "ENG")
        self.assertEqual(self.write.call_args_list[1].args[1], (self.user["username"], "ENG-NEW", False))

    def test_a2_admin_cannot_create_in_unmanaged_workspace(self):
        self.login_as("admin", {"ENG": "admin"})
        response = self.client.post("/api/projects/create", json={
            "id": "HSE-NEW", "name": "Synthetic project", "code": "HSE-NEW", "workspace_id": "HSE",
        })
        self.assertEqual(response.status_code, 403)
        self.write.assert_not_called()
        self.seed.assert_not_called()

    def test_a2_metadata_edit_retains_workspace(self):
        self.login_as("admin", {"ENG": "admin"})
        response = self.client.post("/api/project/ENG-P1", json={"name": "Updated name", "code": "ENG-P1"})
        self.assertEqual(response.status_code, 200)
        self.assertEqual(self.write.call_args.args[1], ("ENG-P1", "Updated name", "ENG-P1", "{}", "ENG"))
        self.seed.assert_not_called()

    def test_a2_superadmin_can_create_om_project(self):
        self.login_as("superadmin", {})
        response = self.client.post("/api/projects/create", json={
            "id": "OM-NEW", "name": "Synthetic contract", "code": "OM-NEW", "workspace_id": "OM",
        })
        self.assertEqual(response.status_code, 200)
        self.assertEqual(self.write.call_args_list[0].args[1][-1], "OM")
        self.seed.assert_called_once_with("OM-NEW", "OM")

    def test_b_authorized_update_passes_identity_and_sanitized_data_to_database(self):
        response = self.client.post("/api/records/ENG-P1/DS", json={
            "_id": "R1", "title": "Updated title", "attachments": {"submission": {"name": "ignored.pdf"}},
        })
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.get_json()["id"], "R1")
        sql, params = self.write.call_args_list[0].args
        self.assertIn("INSERT INTO records", sql)
        self.assertEqual(params[:3], ("R1", "ENG-P1", "DS"))
        self.assertEqual(json.loads(params[3]), {"title": "Updated title"})
        self.assertEqual(self.write.call_args_list[1].args[1], ("ENG-P1",))
        self.assertEqual(self.audit.call_args.args[:5], ("synthetic-user", "EDIT", "ENG-P1", "DS", "R1"))

    def test_b_denied_type_route_stops_before_write(self):
        response = self.client.post("/api/records/ENG-P1/RFI", json={"title": "New RFI"})
        self.assertEqual(response.status_code, 403)
        self.write.assert_not_called()
        self.audit.assert_not_called()

    def test_a1_mixed_role_editor_can_edit_assigned_project(self):
        self.login_as("admin", {"ENG": "admin", "HSE": "editor"}, ("HSE-P1",))
        listing = self.client.get("/api/projects?workspace_id=HSE")
        self.assertEqual(listing.status_code, 200)
        self.assertEqual([(p["id"], p["can_edit"]) for p in listing.get_json()], [("HSE-P1", True)])
        self.assertEqual(self.client.get("/api/project/HSE-P1").status_code, 200)
        response = self.client.post("/api/project/HSE-P1", json={"name": "Updated name", "code": "HSE-P1"})
        self.assertEqual(response.status_code, 200)
        self.assertEqual(self.write.call_args.args[1], ("HSE-P1", "Updated name", "HSE-P1", "{}", "HSE"))

    def test_a2_create_sql_guards_workspace_on_conflict(self):
        # Capture the statement for a valid new project. No conflicting ID is submitted.
        self.db.create_project("ENG-NEW", "Synthetic project", "ENG-NEW", workspace_id="ENG")
        self.assertTrue(self.write.called)
        for call in self.write.call_args_list:
            sql = " ".join(call.args[0].lower().split())
            _, _, conflict = sql.partition("on conflict")
            _, _, assignments = conflict.partition("do update set")
            assignments, _, predicate = assignments.partition("where")
            if "workspace_id" in assignments:
                self.assertRegex(
                    predicate, r"\bprojects\.workspace_id\s*=\s*excluded\.workspace_id\b",
                    "Creation that sets workspace ownership on conflict must match the stored workspace",
                )

    def test_b_record_upsert_sql_requires_project_and_type_ownership(self):
        # Inspect SQL from an authorized update. The SQL is captured and never executed.
        response = self.client.post("/api/records/ENG-P1/DS", json={"_id": "R1", "title": "Updated title"})
        self.assertEqual(response.status_code, 200)
        statements = [" ".join(call.args[0].lower().split()) for call in self.write.call_args_list]
        writes = [sql for sql in statements if sql.startswith(("insert into records", "update records"))]
        self.assertTrue(writes, "Expected a captured record write")
        for sql in writes:
            if sql.startswith("update records"):
                predicate = sql.partition(" where ")[2]
                for column in ("project_id", "dt_id"):
                    self.assertRegex(predicate, rf"\b{column}\s*=\s*%s")
                continue
            _, _, conflict = sql.partition("on conflict")
            target, update, assignments = conflict.partition("do update")
            if not update:
                continue
            columns = {column.strip() for column in target.strip().strip("()").split(",")}
            if {"id", "project_id", "dt_id"} <= columns:
                continue
            predicate = assignments.partition(" where ")[2]
            for column in ("project_id", "dt_id"):
                self.assertRegex(
                    predicate, rf"\brecords\.{column}\s*=\s*excluded\.{column}\b",
                    f"The record conflict update must match stored {column}",
                )
