"""HSE Department Admin RBAC & User Management isolation tests."""

from contextlib import ExitStack
import os
import re
import unittest


class HseUserManagementRbacTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        if os.environ.get("DCR_TEST_DB_PREINJECTED") == "1":
            import db
            cls.db = db
        else:
            from security_test_support import disposable_database

            stack = ExitStack()
            cls.addClassCleanup(stack.close)
            cls.db = stack.enter_context(disposable_database())

        db = cls.db
        from app import app
        cls.app = app

        cls.hse_pid = "SYN-RBAC-HSE"
        cls.eng_pid = "SYN-RBAC-ENG"
        cls.om_pid = "SYN-RBAC-OM"
        db.create_project(cls.hse_pid, "HSE-RBAC-01", "Synthetic HSE RBAC", workspace_id="HSE")
        db.create_project(cls.eng_pid, "ENG-RBAC-01", "Synthetic ENG RBAC", workspace_id="ENG")
        db.create_project(cls.om_pid, "OM-RBAC-01", "Synthetic OM RBAC", workspace_id="OM")

        cls.sa_user = "syn_rbac_sa"
        cls.hse_admin = "syn_rbac_hse_admin"
        cls.hse_peer = "syn_rbac_hse_peer"
        cls.eng_only = "syn_rbac_eng_only"
        cls.multi_user = "syn_rbac_multi"

        db.add_user(cls.sa_user, "SynPass!Sa9", role="superadmin", primary_workspace="ENG")
        db.add_user(cls.hse_admin, "SynPass!Ha9", role="admin", primary_workspace="HSE")
        db.add_user(cls.hse_peer, "SynPass!Hp9", role="admin", primary_workspace="HSE")
        db.add_user(cls.eng_only, "SynPass!Eo9", role="viewer", primary_workspace="ENG")
        db.add_user(cls.multi_user, "SynPass!Mu9", role="editor", primary_workspace="ENG")

        sa_client = cls._client_for(cls.sa_user, "superadmin")
        # Assign syn_rbac_multi to ENG (editor) and HSE (viewer) + both ENG & HSE projects
        r = sa_client.post("/api/users", json={
            "action": "update_workspaces",
            "username": cls.multi_user,
            "workspaces": ["ENG", "HSE"],
        })
        assert r.status_code == 200, r.get_json()
        for ws_id, ws_role in (("ENG", "editor"), ("HSE", "viewer")):
            rr = sa_client.post("/api/users", json={
                "action": "update_workspace_role",
                "username": cls.multi_user,
                "workspace_id": ws_id,
                "role": ws_role,
            })
            assert rr.status_code == 200, rr.get_json()
        for pid in (cls.eng_pid, cls.hse_pid):
            ra = sa_client.post("/api/users", json={
                "action": "assign",
                "username": cls.multi_user,
                "project_id": pid,
            })
            assert ra.status_code == 200, ra.get_json()
        re_eng = sa_client.post("/api/users", json={
            "action": "assign",
            "username": cls.eng_only,
            "project_id": cls.eng_pid,
        })
        assert re_eng.status_code == 200, re_eng.get_json()

    @classmethod
    def _client_for(cls, username, role):
        client = cls.app.test_client()
        tok = cls.db.create_session(username, role)
        client.set_cookie("dcr_token", tok)
        return client

    def _find_user(self, users_list, username):
        for u in users_list:
            if u.get("username") == username:
                return u
        return None

    def test_01_hse_admin_full_subordinate_lifecycle(self):
        client = self._client_for(self.hse_admin, "admin")
        sub_name = "syn_rbac_sda"

        # 1. Verify whoami and workspace scope
        rw = client.get("/api/workspaces")
        self.assertEqual(rw.status_code, 200)
        ws_ids = [w["id"] for w in rw.get_json()]
        self.assertEqual(ws_ids, ["HSE"])

        rm = client.get("/api/whoami")
        self.assertEqual(rm.status_code, 200)
        who = rm.get_json()
        self.assertFalse(who["is_superadmin"])
        self.assertEqual([w["id"] for w in who["workspaces"]], ["HSE"])

        # 2. Create subordinate editor in HSE
        r = client.post("/api/users", json={
            "action": "add",
            "username": sub_name,
            "password": "SynUser!123",
            "role": "editor",
            "email": "sda.init@example.com",
        })
        self.assertEqual(r.status_code, 200, r.get_json())

        users = client.get("/api/users").get_json()
        sda = self._find_user(users, sub_name)
        self.assertIsNotNone(sda)
        self.assertEqual(sda["primary_workspace"], "HSE")
        self.assertEqual(sda["role"], "editor")
        self.assertEqual(sda["email"], "sda.init@example.com")
        self.assertEqual(
            [(w["workspace_id"], w["role"]) for w in sda["workspaces"]],
            [("HSE", "editor")],
        )

        # 3. Update email inline
        r = client.post("/api/users", json={
            "action": "update_email",
            "username": sub_name,
            "email": "sda.updated@example.com",
        })
        self.assertEqual(r.status_code, 200, r.get_json())
        sda = self._find_user(client.get("/api/users").get_json(), sub_name)
        self.assertEqual(sda["email"], "sda.updated@example.com")

        # 4. Assign HSE project & verify doc type permissions including [] (None)
        r = client.post("/api/users", json={
            "action": "assign",
            "username": sub_name,
            "project_id": self.hse_pid,
        })
        self.assertEqual(r.status_code, 200, r.get_json())

        r = client.post("/api/users", json={
            "action": "set_doc_type_permissions",
            "username": sub_name,
            "project_id": self.hse_pid,
            "doc_types": ["PTW", "OBS"],
        })
        self.assertEqual(r.status_code, 200, r.get_json())
        projs = client.get(f"/api/users/{sub_name}/projects").get_json()
        self.assertEqual(len(projs), 1)
        self.assertEqual(projs[0]["project_id"], self.hse_pid)
        self.assertEqual(projs[0]["allowed_doc_types"], ["PTW", "OBS"])

        r = client.post("/api/users", json={
            "action": "set_doc_type_permissions",
            "username": sub_name,
            "project_id": self.hse_pid,
            "doc_types": [],
        })
        self.assertEqual(r.status_code, 200, r.get_json())
        projs = client.get(f"/api/users/{sub_name}/projects").get_json()
        self.assertEqual(projs[0]["allowed_doc_types"], [])

        # 5. Create active session for sda, then demote HSE workspace role editor -> viewer
        sub_tok = self.db.create_session(sub_name, "editor")
        r = client.post("/api/users", json={
            "action": "update_workspace_role",
            "username": sub_name,
            "workspace_id": "HSE",
            "role": "viewer",
        })
        self.assertEqual(r.status_code, 200, r.get_json())
        sda_row = self.db.get_user(sub_name)
        self.assertEqual(sda_row["role"], "viewer")
        sess = self.db.get_session(sub_tok)
        self.assertEqual(sess["role"], "viewer")
        self.assertFalse(self.db.can_user_download(sub_name, self.hse_pid))

        # 6. Toggle download permission ON then OFF while viewer
        r = client.post("/api/users", json={
            "action": "toggle_download_permission",
            "username": sub_name,
            "can_download": True,
        })
        self.assertEqual(r.status_code, 200, r.get_json())
        self.assertTrue(self.db.can_user_download(sub_name, self.hse_pid))

        r = client.post("/api/users", json={
            "action": "toggle_download_permission",
            "username": sub_name,
            "can_download": False,
        })
        self.assertEqual(r.status_code, 200, r.get_json())
        self.assertFalse(self.db.can_user_download(sub_name, self.hse_pid))

        # 7. Promote back to editor -> downloads & session role sync back to editor
        r = client.post("/api/users", json={
            "action": "update_workspace_role",
            "username": sub_name,
            "workspace_id": "HSE",
            "role": "editor",
        })
        self.assertEqual(r.status_code, 200, r.get_json())
        self.assertEqual(self.db.get_user(sub_name)["role"], "editor")
        self.assertEqual(self.db.get_session(sub_tok)["role"], "editor")
        self.assertTrue(self.db.can_user_download(sub_name, self.hse_pid))

        # 8. Change password & verify login works with new password
        r = client.post("/api/users", json={
            "action": "change_password",
            "username": sub_name,
            "password": "SynUserReset!789",
        })
        self.assertEqual(r.status_code, 200, r.get_json())
        self.assertTrue(self.db.verify_pw(sub_name, "SynUserReset!789"))
        self.assertFalse(self.db.verify_pw(sub_name, "SynUser!123"))

        # 9. Unassign project and delete subordinate
        r = client.post("/api/users", json={
            "action": "unassign",
            "username": sub_name,
            "project_id": self.hse_pid,
        })
        self.assertEqual(r.status_code, 200, r.get_json())

        r = client.post("/api/users", json={
            "action": "delete",
            "username": sub_name,
        })
        self.assertEqual(r.status_code, 200, r.get_json())
        self.assertIsNone(self.db.get_user(sub_name))

    def test_02_hse_admin_blocks_escalation_self_peer_admin_and_superadmin(self):
        client = self._client_for(self.hse_admin, "admin")
        temp_sub = "syn_rbac_temp_sub"
        r = client.post("/api/users", json={
            "action": "add",
            "username": temp_sub,
            "password": "SynTemp!123",
            "role": "viewer",
        })
        self.assertEqual(r.status_code, 200, r.get_json())
        try:
            # Cannot create admin or superadmin
            for blocked_role in ("admin", "superadmin"):
                rc = client.post("/api/users", json={
                    "action": "add",
                    "username": f"syn_blocked_{blocked_role}",
                    "password": "SynTemp!123",
                    "role": blocked_role,
                })
                self.assertEqual(rc.status_code, 403, (blocked_role, rc.get_json()))
                self.assertIsNone(self.db.get_user(f"syn_blocked_{blocked_role}"))

            # Cannot promote subordinate to admin or superadmin
            for blocked_role in ("admin", "superadmin"):
                ru = client.post("/api/users", json={
                    "action": "update_role",
                    "username": temp_sub,
                    "role": blocked_role,
                })
                self.assertEqual(ru.status_code, 403, (blocked_role, ru.get_json()))

                rw = client.post("/api/users", json={
                    "action": "update_workspace_role",
                    "username": temp_sub,
                    "workspace_id": "HSE",
                    "role": blocked_role,
                })
                expected_status = 403 if blocked_role == "admin" else 400
                self.assertEqual(rw.status_code, expected_status, (blocked_role, rw.get_json()))

            # Cannot mutate self, peer HSE admin, or superadmin
            for target in (self.hse_admin, self.hse_peer, self.sa_user):
                for payload in (
                    {"action": "update_workspace_role", "username": target, "workspace_id": "HSE", "role": "viewer"},
                    {"action": "update_role", "username": target, "role": "viewer"},
                    {"action": "toggle_download_permission", "username": target, "can_download": False},
                    {"action": "change_password", "username": target, "password": "Hacked!123"},
                    {"action": "update_email", "username": target, "email": "hacked@example.com"},
                    {"action": "delete", "username": target},
                    {"action": "assign", "username": target, "project_id": self.hse_pid},
                    {"action": "set_doc_type_permissions", "username": target, "project_id": self.hse_pid, "doc_types": ["PTW"]},
                ):
                    res = client.post("/api/users", json=payload)
                    self.assertEqual(res.status_code, 403, (target, payload["action"], res.get_json()))

            # Cannot reassign workspace / department memberships
            for payload in (
                {"action": "assign_workspace", "username": temp_sub, "workspace_id": "ENG"},
                {"action": "update_workspaces", "username": temp_sub, "workspaces": ["HSE", "ENG"]},
                {"action": "update_department", "username": temp_sub, "primary_workspace": "ENG"},
            ):
                res = client.post("/api/users", json=payload)
                self.assertEqual(res.status_code, 403, (payload["action"], res.get_json()))
        finally:
            self._client_for(self.sa_user, "superadmin").post("/api/users", json={
                "action": "delete",
                "username": temp_sub,
            })

    def test_03_multi_dept_user_and_cross_dept_project_containment(self):
        # Workspace-aware download permission checks
        self.assertTrue(self.db.can_user_download(self.multi_user, self.eng_pid))
        self.assertFalse(self.db.can_user_download(self.multi_user, self.hse_pid))

        client = self._client_for(self.hse_admin, "admin")

        # GET /api/users shows multi_user only with HSE projects (no ENG leakage) and hides eng_only
        users = client.get("/api/users").get_json()
        self.assertIsNone(self._find_user(users, self.eng_only))
        self.assertIsNone(self._find_user(users, self.sa_user))

        multi = self._find_user(users, self.multi_user)
        self.assertIsNotNone(multi)
        self.assertEqual([p["project_id"] for p in multi["projects"]], [self.hse_pid])

        r_mp = client.get(f"/api/users/{self.multi_user}/projects")
        self.assertEqual(r_mp.status_code, 200)
        self.assertEqual([p["project_id"] for p in r_mp.get_json()], [self.hse_pid])

        # eng_only project list must be forbidden for HSE admin
        r_ep = client.get(f"/api/users/{self.eng_only}/projects")
        self.assertEqual(r_ep.status_code, 403)

        # HSE admin CAN manage HSE doc-types & HSE workspace role on multi_user
        r_dt = client.post("/api/users", json={
            "action": "set_doc_type_permissions",
            "username": self.multi_user,
            "project_id": self.hse_pid,
            "doc_types": ["PTW"],
        })
        self.assertEqual(r_dt.status_code, 200, r_dt.get_json())

        r_ws = client.post("/api/users", json={
            "action": "update_workspace_role",
            "username": self.multi_user,
            "workspace_id": "HSE",
            "role": "editor",
        })
        self.assertEqual(r_ws.status_code, 200, r_ws.get_json())
        self.assertTrue(self.db.can_user_download(self.multi_user, self.hse_pid))

        r_ws_back = client.post("/api/users", json={
            "action": "update_workspace_role",
            "username": self.multi_user,
            "workspace_id": "HSE",
            "role": "viewer",
        })
        self.assertEqual(r_ws_back.status_code, 200, r_ws_back.get_json())
        # Primary workspace is ENG (editor), so global users.role stays editor
        self.assertEqual(self.db.get_user(self.multi_user)["role"], "editor")

        # HSE admin CANNOT perform global mutations or ENG/OM mutations on multi_user
        for payload in (
            {"action": "update_workspace_role", "username": self.multi_user, "workspace_id": "ENG", "role": "viewer"},
            {"action": "update_role", "username": self.multi_user, "role": "viewer"},
            {"action": "toggle_download_permission", "username": self.multi_user, "can_download": True},
            {"action": "change_password", "username": self.multi_user, "password": "Hacked!123"},
            {"action": "update_email", "username": self.multi_user, "email": "hacked@example.com"},
            {"action": "delete", "username": self.multi_user},
            {"action": "assign", "username": self.multi_user, "project_id": self.eng_pid},
            {"action": "assign", "username": self.multi_user, "project_id": self.om_pid},
            {"action": "unassign", "username": self.multi_user, "project_id": self.eng_pid},
            {"action": "set_doc_type_permissions", "username": self.multi_user, "project_id": self.eng_pid, "doc_types": ["DWG"]},
        ):
            res = client.post("/api/users", json=payload)
            self.assertEqual(res.status_code, 403, (payload["action"], res.get_json()))

        # Confirm eng_only cannot be modified in any way by HSE admin
        for payload in (
            {"action": "assign", "username": self.eng_only, "project_id": self.hse_pid},
            {"action": "toggle_download_permission", "username": self.eng_only, "can_download": True},
            {"action": "change_password", "username": self.eng_only, "password": "Hacked!123"},
            {"action": "delete", "username": self.eng_only},
        ):
            res = client.post("/api/users", json=payload)
            self.assertEqual(res.status_code, 403, (payload["action"], res.get_json()))

    def test_04_superadmin_retains_global_authority(self):
        sa = self._client_for(self.sa_user, "superadmin")
        # Superadmin can toggle download on viewer in multi-dept or edit email globally
        r_em = sa.post("/api/users", json={
            "action": "update_email",
            "username": self.multi_user,
            "email": "sa.verified@example.com",
        })
        self.assertEqual(r_em.status_code, 200, r_em.get_json())
        self.assertEqual(self.db.get_user(self.multi_user)["email"], "sa.verified@example.com")

        # Superadmin can promote/demote workspace admin
        r_pr = sa.post("/api/users", json={
            "action": "add",
            "username": "syn_sa_managed_adm",
            "password": "SynAdm!123",
            "role": "admin",
            "primary_workspace": "HSE",
        })
        self.assertEqual(r_pr.status_code, 200, r_pr.get_json())
        try:
            r_dl = sa.post("/api/users", json={
                "action": "delete",
                "username": "syn_sa_managed_adm",
            })
            self.assertEqual(r_dl.status_code, 200, r_dl.get_json())
        finally:
            if self.db.get_user("syn_sa_managed_adm"):
                sa.post("/api/users", json={"action": "delete", "username": "syn_sa_managed_adm"})

    def test_05_user_management_ui_html_guards_and_badges(self):
        import html_render

        js = html_render.SHARED_JS
        # All Departments pill restricted to superadmin with >1 workspace
        self.assertIn("isSuper && availableWorkspaces.length > 1 && assignedWorkspaces.length === availableWorkspaces.length", js)
        # Add user role selector hides admin & superadmin from non-superadmin
        self.assertIn("${isSuper ? '<option value=\"admin\">Admin</option><option value=\"superadmin\">Super Admin</option>' : ''}", js)
        # No duplicate onblur on email-inp
        m = re.search(r"id=\"email-inp-[^\n]+", js)
        self.assertIsNotNone(m)
        self.assertEqual(len(re.findall(r"\bonblur=", m.group(0))), 1, m.group(0))
        # Dynamic workspace badges in Permissions modal + 0 types (None) + Array.isArray guard
        self.assertIn("Health, Safety & Environment (HSE)", js)
        self.assertIn("Operation & Maintenance (O&M)", js)
        self.assertIn("0 types (None)", js)
        self.assertIn("Array.isArray(pAccess.allowed_doc_types)", js)
        # UI action containment guards for subordinate vs peer/self/multi-dept
        self.assertIn("canManageThisWs", js)
        self.assertIn("canEditGlobalAccount", js)
        self.assertIn("isPeerOrHigherAdmin", js)


if __name__ == "__main__":
    unittest.main()
