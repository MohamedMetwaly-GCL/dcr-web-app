"""Comprehensive regression suite for the 5 final Codex review findings.

Run through security_test_support.py with a fresh disposable PostgreSQL cluster.
"""

import json
import logging
import re
import shutil
import subprocess
import unittest
import uuid
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from unittest.mock import MagicMock, patch
from contextlib import ExitStack

from security_test_support import disposable_database


_STATE = {}
_NODE_EXE = shutil.which("node") or (
    r"C:\Program Files\nodejs\node.exe" if Path(r"C:\Program Files\nodejs\node.exe").exists() else None
)



def setUpModule():
    stack = ExitStack()
    unittest.addModuleCleanup(stack.close)
    db = stack.enter_context(disposable_database())
    db.create_workspace("HSE", "HSE", "Health, Safety & Environment", "HSE", "🛡️", "hse", 3, "HSE Workspace")

    # ── Seed Projects for Finding 1 & Ambiguity Tests ────────────────────────────────────
    db.create_project("PRJ-OM-1", "Legacy Facility One", "LEG-CODE-1", workspace_id="OM")
    for dt_id, dt_code, dt_name in (
        ("WO", "WO", "Work Orders"),
        ("SLA", "SLA", "SLA Reports"),
        ("dt-ptw", "PTW", "Permit To Work"),
    ):
        db.exe(
            "INSERT INTO doc_types(id, project_id, name, code, sort_order) VALUES(%s,%s,%s,%s,10) ON CONFLICT DO NOTHING",
            (dt_id, "PRJ-OM-1", dt_name, dt_code),
        )
        db.add_column("PRJ-OM-1", dt_id, "docNo", "Document No.", "text")
        db.add_column("PRJ-OM-1", dt_id, "title", "Title", "text")
        db.add_column("PRJ-OM-1", dt_id, "status", "Status", "status")
        db.add_column("PRJ-OM-1", dt_id, "issuedDate", "Issued Date", "date")
        db.add_column("PRJ-OM-1", dt_id, "expectedReplyDate", "Expected Reply Date", "date")

    db.save_record("PRJ-OM-1", "WO", "OM1-WO-1", {
        "docNo": "WO-001", "title": "TOK_OM1_WO Visible WorkOrder", "status": "Open",
        "issuedDate": "2026-01-01", "expectedReplyDate": "2026-01-05",
    })
    db.save_record("PRJ-OM-1", "SLA", "OM1-SLA-1", {
        "docNo": "SLA-001", "title": "TOK_OM1_SLA Hidden SLA", "status": "Completed",
        "issuedDate": "2026-01-01", "expectedReplyDate": "2026-01-05",
    })
    db.save_record("PRJ-OM-1", "dt-ptw", "OM1-PTW-1", {
        "docNo": "PTW-001", "title": "TOK_OM1_PTW Hidden Permit", "status": "Pending",
        "issuedDate": "2026-01-01", "expectedReplyDate": "2026-01-05",
    })

    # Ambiguous project pair
    db.create_project("AMBIG-1", "Ambiguous Name", "AMBIG-C", workspace_id="ENG")
    db.create_project("AMBIG-2", "Ambiguous Name", "AMBIG-C", workspace_id="ENG")

    # ── Seed Projects for Finding 2 (Dashboards & AI) & Finding 3 (Explicit AI Selection) ──
    db.create_project("ENG-D2", "Alpha Engineering", "ENG-CODE-A", workspace_id="ENG")
    for dt_code, dt_name in (("MS", "Material Submittal"), ("RFI", "RFI"), ("NOC", "Notice of Change"), ("PR", "Purchase Request"), ("LTR", "Letters")):
        db.exe(
            "INSERT INTO doc_types(id, project_id, name, code, sort_order) VALUES(%s,%s,%s,%s,20) ON CONFLICT DO NOTHING",
            (dt_code, "ENG-D2", dt_name, dt_code),
        )
        db.add_column("ENG-D2", dt_code, "docNo", "Document No.", "text")
        db.add_column("ENG-D2", dt_code, "title", "Title", "text")
        db.add_column("ENG-D2", dt_code, "status", "Status", "status")
        db.add_column("ENG-D2", dt_code, "discipline", "Discipline", "text")
        db.add_column("ENG-D2", dt_code, "issuedDate", "Issued Date", "date")
        db.add_column("ENG-D2", dt_code, "expectedReplyDate", "Expected Reply Date", "date")

    db.save_record("ENG-D2", "MS", "D2-MS-1", {
        "docNo": "MS-D2-001", "title": "TOK_VIS_MS Chiller Plant", "status": "Approved",
        "discipline": "Mechanical", "issuedDate": "2026-01-02",
    })
    db.save_record("ENG-D2", "MS", "D2-MS-2", {
        "docNo": "MS-D2-002", "title": "TOK_VIS_MS Pending Valve", "status": "Under Review",
        "discipline": "Mechanical", "issuedDate": "2026-01-02", "expectedReplyDate": "2026-01-05",
    })
    db.save_record("ENG-D2", "RFI", "D2-RFI-1", {
        "docNo": "RFI-D2-009 REV 00", "title": "TOK_SECRET_RFI Electrical Shaft", "status": "Under Review",
        "discipline": "Electrical", "issuedDate": "2026-01-02", "expectedReplyDate": "2026-01-04",
    })
    db.save_record("ENG-D2", "RFI", "D2-RFI-2", {
        "docNo": "RFI-D2-009 REV 01", "title": "TOK_SECRET_RFI Revision Two", "status": "Approved",
        "discipline": "Electrical", "issuedDate": "2026-01-06",
    })
    db.save_record("ENG-D2", "NOC", "D2-NOC-1", {
        "docNo": "NOC-D2-001", "nocSubject": "TOK_SECRET_NOC Variation Order", "title": "TOK_SECRET_NOC Variation Order",
        "submittedCost": 987654.0, "finalApprovedCost": 543210.0, "partDStatus": "Approved",
    })
    db.save_record("ENG-D2", "PR", "D2-PR-1", {
        "docNo": "PR-D2-001", "title": "TOK_SECRET_PR Secret Pumps", "status": "Pending PO",
    })

    db.create_project("ENG-D2B", "Beta Engineering", "ENG-CODE-B", workspace_id="ENG")
    db.exe("INSERT INTO doc_types(id, project_id, name, code, sort_order) VALUES('MS','ENG-D2B','Material Submittal','MS',20) ON CONFLICT DO NOTHING")
    db.save_record("ENG-D2B", "MS", "D2B-MS-1", {"docNo": "MS-BETA-001", "title": "Beta Only Pump", "status": "Approved"})

    db.create_project("ENG-UNAUTH", "Secret Gamma Project", "UNAUTH-CODE", workspace_id="ENG")
    db.exe("INSERT INTO doc_types(id, project_id, name, code, sort_order) VALUES('MS','ENG-UNAUTH','Material Submittal','MS',20) ON CONFLICT DO NOTHING")
    db.save_record("ENG-UNAUTH", "MS", "UNAUTH-MS-1", {"docNo": "MS-UNAUTH-001", "title": "TOK_UNAUTH_SECRET", "status": "Approved"})

    # HSE workspace project with multiple doc types
    db.create_project("HSE-D2", "HSE Compliance Site", "HSE-C2", workspace_id="HSE")
    for dt_id, dt_name in (("HSE_PTW", "Permit to Work"), ("HSE_INC", "Incident Reports")):
        db.exe("INSERT INTO doc_types(id, project_id, name, code, sort_order) VALUES(%s,%s,%s,%s,10) ON CONFLICT DO NOTHING",
               (dt_id, "HSE-D2", dt_name, dt_id))
        db.add_column("HSE-D2", dt_id, "docNo", "Doc No", "text")
        db.add_column("HSE-D2", dt_id, "title", "Title", "text")
        db.add_column("HSE-D2", dt_id, "status", "Status", "status")
        db.add_column("HSE-D2", dt_id, "issuedDate", "Issued Date", "date")
        db.add_column("HSE-D2", dt_id, "expectedReplyDate", "Expected Reply Date", "date")
    db.save_record("HSE-D2", "HSE_PTW", "HSE-PTW-1", {"docNo": "PTW-HSE-1", "title": "TOK_HSE_VIS Excavation", "status": "Pending", "issuedDate": "2026-01-02", "expectedReplyDate": "2026-01-05"})
    db.save_record("HSE-D2", "HSE_INC", "HSE-INC-1", {"docNo": "INC-HSE-1", "title": "TOK_HSE_SECRET Incident", "status": "Open", "issuedDate": "2026-01-02", "expectedReplyDate": "2026-01-05"})

    # ── Seed Users ────────────────────────────────────────────────────────────────────────
    pw = "Syn5#" + uuid.uuid4().hex[:8]
    users = [
        ("f1_super",        "superadmin", "ENG", {"ENG": "admin", "OM": "admin", "HSE": "admin"}),
        ("u_leg_restr",     "editor",     "OM",  {"OM": "editor"}),
        ("u_leg_unrestr",   "editor",     "OM",  {"OM": "editor"}),
        ("u_leg_empty",     "editor",     "OM",  {"OM": "editor"}),
        ("u_coexist_mixed", "editor",     "OM",  {"OM": "editor"}),
        ("u_coexist_inter", "editor",     "OM",  {"OM": "editor"}),
        ("u_coexist_disj",  "editor",     "OM",  {"OM": "editor"}),
        ("u_ambig",         "editor",     "ENG", {"ENG": "editor"}),
        ("u_d2_multi",      "editor",     "ENG", {"ENG": "editor", "OM": "viewer", "HSE": "viewer"}),
        ("u_f3_ai",         "editor",     "ENG", {"ENG": "editor"}),
    ]
    for uname, role, pws, ws_roles in users:
        db.add_user(uname, pw, role, f"{uname}@synthetic.local", pws, False)
        for ws_id, wrole in ws_roles.items():
            db.assign_user_workspace(uname, ws_id, wrole)

    # Raw legacy & canonical user_projects assignments (deliberately NOT normalized!)
    db.exe("INSERT INTO user_projects(username, project_id, is_dc, allowed_doc_types) VALUES(%s,%s,%s,%s::jsonb)",
           ("u_leg_restr", "LEG-CODE-1", True, json.dumps(["WO"])))
    db.exe("INSERT INTO user_projects(username, project_id, is_dc, allowed_doc_types) VALUES(%s,%s,%s,NULL)",
           ("u_leg_unrestr", "Legacy Facility One", False))
    db.exe("INSERT INTO user_projects(username, project_id, is_dc, allowed_doc_types) VALUES(%s,%s,%s,%s::jsonb)",
           ("u_leg_empty", "LEG-CODE-1", False, json.dumps([])))

    # Coexisting canonical + legacy rows
    db.exe("INSERT INTO user_projects(username, project_id, is_dc, allowed_doc_types) VALUES(%s,%s,%s,NULL)",
           ("u_coexist_mixed", "PRJ-OM-1", False))
    db.exe("INSERT INTO user_projects(username, project_id, is_dc, allowed_doc_types) VALUES(%s,%s,%s,%s::jsonb)",
           ("u_coexist_mixed", "LEG-CODE-1", True, json.dumps(["WO", "SLA"])))

    db.exe("INSERT INTO user_projects(username, project_id, is_dc, allowed_doc_types) VALUES(%s,%s,%s,%s::jsonb)",
           ("u_coexist_inter", "PRJ-OM-1", False, json.dumps(["WO", "SLA"])))
    db.exe("INSERT INTO user_projects(username, project_id, is_dc, allowed_doc_types) VALUES(%s,%s,%s,%s::jsonb)",
           ("u_coexist_inter", "Legacy Facility One", False, json.dumps(["SLA", "PTW"])))

    db.exe("INSERT INTO user_projects(username, project_id, is_dc, allowed_doc_types) VALUES(%s,%s,%s,%s::jsonb)",
           ("u_coexist_disj", "PRJ-OM-1", False, json.dumps(["WO"])))
    db.exe("INSERT INTO user_projects(username, project_id, is_dc, allowed_doc_types) VALUES(%s,%s,%s,%s::jsonb)",
           ("u_coexist_disj", "LEG-CODE-1", False, json.dumps(["SLA"])))

    # Ambiguous alias assignment
    db.exe("INSERT INTO user_projects(username, project_id, is_dc, allowed_doc_types) VALUES(%s,%s,%s,%s::jsonb)",
           ("u_ambig", "Ambiguous Name", False, json.dumps(["MS"])))

    # Mixed user across ENG / OM / HSE with doc-type restrictions
    db.assign_project("u_d2_multi", "ENG-D2")
    db.set_user_project_doc_types("u_d2_multi", "ENG-D2", ["MS"])
    db.exe("INSERT INTO user_projects(username, project_id, is_dc, allowed_doc_types) VALUES(%s,%s,%s,%s::jsonb)",
           ("u_d2_multi", "LEG-CODE-1", False, json.dumps(["WO"])))
    db.assign_project("u_d2_multi", "HSE-D2")
    db.set_user_project_doc_types("u_d2_multi", "HSE-D2", ["HSE_PTW"])

    # Finding 3 AI selection user
    db.assign_project("u_f3_ai", "ENG-D2")
    db.assign_project("u_f3_ai", "ENG-D2B")

    tokens = {}
    for uname, role, _, _ in users:
        tokens[uname] = db.create_session(uname, role)

    import auth
    from app import app
    app.config["TESTING"] = True
    _STATE["tokens"] = tokens
    _STATE["app"] = app
    _STATE["db"] = db
    _STATE["auth"] = auth



class FiveFindingsBase(unittest.TestCase):
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


class TestFinding1LegacyAliasDocTypeGrants(FiveFindingsBase):
    def test_legacy_alias_restricted_unrestricted_empty_and_missing(self):
        # Restricted via legacy code LEG-CODE-1
        self.assertEqual(self.db.get_user_project_doc_types("u_leg_restr", "PRJ-OM-1"), ["WO"])
        self.assertTrue(self.db.user_is_dc("u_leg_restr", "PRJ-OM-1"))

        u_restr = self.db.get_user("u_leg_restr")
        self.assertEqual(self.auth.get_allowed_doc_type_filter("PRJ-OM-1", u_restr), ["WO"])
        self.assertTrue(self.auth.can_access_doc_type("PRJ-OM-1", "WO", u_restr))
        self.assertFalse(self.auth.can_access_doc_type("PRJ-OM-1", "SLA", u_restr))
        self.assertFalse(self.auth.can_access_doc_type("PRJ-OM-1", "dt-ptw", u_restr))

        c_restr = self.client_for("u_leg_restr")
        self.assertEqual([d["id"] for d in c_restr.get("/api/doc_types/PRJ-OM-1").get_json()], ["WO"])
        self.assertEqual(c_restr.get("/api/records/PRJ-OM-1/WO").status_code, 200)
        self.assertEqual(c_restr.get("/api/records/PRJ-OM-1/SLA").status_code, 403)
        self.assertEqual(c_restr.get("/api/records/PRJ-OM-1/dt-ptw").status_code, 403)
        self.assertEqual(c_restr.post("/api/records/PRJ-OM-1/SLA", json={"docNo": "SLA-X"}).status_code, 403)
        self.assertEqual(c_restr.post("/api/doc_types/PRJ-OM-1", json={"code": "ZZ", "name": "Blocked"}).status_code, 403)

        # Unrestricted NULL via legacy name "Legacy Facility One"
        self.assertIsNone(self.db.get_user_project_doc_types("u_leg_unrestr", "PRJ-OM-1"))
        u_unr = self.db.get_user("u_leg_unrestr")
        self.assertIsNone(self.auth.get_allowed_doc_type_filter("PRJ-OM-1", u_unr))
        self.assertTrue(self.auth.can_access_doc_type("PRJ-OM-1", "SLA", u_unr))
        self.assertTrue(self.auth.can_access_doc_type("PRJ-OM-1", "dt-ptw", u_unr))
        self.assertEqual(self.client_for("u_leg_unrestr").get("/api/records/PRJ-OM-1/dt-ptw").status_code, 200)

        # Explicit empty [] via legacy alias fails closed (NOT None!)
        self.assertEqual(self.db.get_user_project_doc_types("u_leg_empty", "PRJ-OM-1"), [])
        u_emp = self.db.get_user("u_leg_empty")
        self.assertEqual(self.auth.get_allowed_doc_type_filter("PRJ-OM-1", u_emp), [])
        self.assertFalse(self.auth.can_access_doc_type("PRJ-OM-1", "WO", u_emp))
        self.assertEqual(self.client_for("u_leg_empty").get("/api/records/PRJ-OM-1/WO").status_code, 403)

        # Missing row / unassigned / unknown / ambiguous project NEVER returns None
        self.assertEqual(self.db.get_user_project_doc_types("u_leg_restr", "ENG-D2"), [])
        self.assertEqual(self.db.get_user_project_doc_types("u_ambig", "AMBIG-1"), [])
        self.assertEqual(self.db.get_user_project_doc_types("u_ambig", "AMBIG-C"), [])
        self.assertEqual(self.auth.get_allowed_doc_type_filter("AMBIG-C", self.db.get_user("u_ambig")), [])
        self.assertEqual(self.auth.get_allowed_doc_type_filter("PRJ-OM-1", None), [])

    def test_coexisting_canonical_and_legacy_rows_conservative_policy_and_unmodified_storage(self):
        # NULL canonical + ["WO", "SLA"] legacy -> restricted wins (no broadening to None!)
        self.assertEqual(self.db.get_user_project_doc_types("u_coexist_mixed", "PRJ-OM-1"), ["WO", "SLA"])
        self.assertTrue(self.db.user_is_dc("u_coexist_mixed", "PRJ-OM-1"))
        u_mix = self.db.get_user("u_coexist_mixed")
        self.assertTrue(self.auth.can_access_doc_type("PRJ-OM-1", "WO", u_mix))
        self.assertTrue(self.auth.can_access_doc_type("PRJ-OM-1", "SLA", u_mix))
        self.assertFalse(self.auth.can_access_doc_type("PRJ-OM-1", "dt-ptw", u_mix))
        self.assertEqual(self.client_for("u_coexist_mixed").get("/api/records/PRJ-OM-1/dt-ptw").status_code, 403)

        # Overlapping ["WO", "SLA"] + ["SLA", "PTW"] -> intersection ["SLA"]
        self.assertEqual(self.db.get_user_project_doc_types("u_coexist_inter", "PRJ-OM-1"), ["SLA"])
        u_int = self.db.get_user("u_coexist_inter")
        self.assertFalse(self.auth.can_access_doc_type("PRJ-OM-1", "WO", u_int))
        self.assertTrue(self.auth.can_access_doc_type("PRJ-OM-1", "SLA", u_int))
        self.assertFalse(self.auth.can_access_doc_type("PRJ-OM-1", "dt-ptw", u_int))

        # Disjoint ["WO"] + ["SLA"] -> empty []
        self.assertEqual(self.db.get_user_project_doc_types("u_coexist_disj", "PRJ-OM-1"), [])
        self.assertFalse(self.auth.can_access_doc_type("PRJ-OM-1", "WO", self.db.get_user("u_coexist_disj")))

        # Stored legacy user_projects rows are preserved verbatim without migration
        stored = [r["project_id"] for r in self.db.get_user_projects("u_coexist_mixed")]
        self.assertEqual(sorted(stored), ["LEG-CODE-1", "PRJ-OM-1"])


class TestFinding2DashboardsAndAIScoping(FiveFindingsBase):
    def test_om_hse_dashboards_and_launchpad_exclude_restricted_doc_types(self):
        c_restr = self.client_for("u_d2_multi")
        c_super = self.client_for("f1_super")

        # O&M dashboard (/?ws=OM): u_d2_multi only has ["WO"] via LEG-CODE-1 (1 record), hides SLA & PTW
        om_html = c_restr.get("/?ws=OM").data.decode("utf-8")
        self.assertIn("WO", om_html)
        self.assertIn("Work Orders", om_html)
        self.assertNotIn("Service Level Reports", om_html)
        self.assertNotIn("SLA Reports", om_html)
        self.assertNotIn("Permit To Work", om_html)
        self.assertIn(">1<", re.sub(r"\s+", "", om_html))  # 1 visible total record, not 3!

        # Superadmin sees all 3 doc types in OM
        om_super = c_super.get("/?ws=OM").data.decode("utf-8")
        self.assertIn("Service Level Reports", om_super)
        self.assertIn("Permit To Work", om_super)

        # HSE dashboard (/?ws=HSE): u_d2_multi only has HSE_PTW, hides HSE_INC
        hse_html = c_restr.get("/?ws=HSE").data.decode("utf-8")
        self.assertIn("HSE_PTW", hse_html)
        self.assertNotIn("HSE_INC", hse_html)
        self.assertNotIn("Incident Reports", hse_html)

        # Launchpad (GET /) record counters must reflect ONLY permitted doc types per workspace!
        lp_html = c_restr.get("/").data.decode("utf-8")
        # Extract ENG (only MS = 2 records), OM (only WO = 1 record), HSE (only HSE_PTW = 1 record)
        for ws_id, expected_docs in (("ENG", "2"), ("OM", "1"), ("HSE", "1")):
            m = re.search(
                rf'data-ws-id="{ws_id}".*?<div class="hc-kpi-num">\s*(\d+)\s*</div>\s*<div class="hc-kpi-lbl">[^<]*</div>\s*</div>\s*<div class="hc-kpi-cell">\s*<div class="hc-kpi-num">\s*([\d,]+)\s*</div>',
                lp_html, re.S,
            )
            self.assertIsNotNone(m, f"Launchpad card for {ws_id} not found")
            self.assertEqual(m.group(2), expected_docs, f"Launchpad leaked hidden doc-type counts on {ws_id}: {m.group(2)}")

        # Warm dashboard_cache populated by superadmin must never leak into restricted user's stats
        c_super.get("/api/dashboard_stats?ws=ENG")
        r_stats = c_restr.get("/api/dashboard_stats?ws=ENG").get_json()
        eng_d2_stat = next(s for s in r_stats if s["id"] == "ENG-D2")
        self.assertEqual(eng_d2_stat["total"], 2)  # Only MS (2), NOT MS+RFI (4)!
        self.assertEqual([d["id"] for d in eng_d2_stat.get("dt_stats", [])], ["MS"])

    def test_all_ai_tools_and_endpoints_enforce_doc_type_scope(self):
        from blueprints import ai_tools, ai

        u_restr = self.db.get_user("u_d2_multi")
        dt_scope = self.auth.get_user_doc_type_scope(u_restr, ["ENG-D2"])

        # 1. search_submittals hides RFI, PR, NOC; allows MS
        s_all = ai_tools.search_submittals(["ENG-D2"], doc_type_scope=dt_scope)
        self.assertTrue(all(r["doc_type"] == "MS" for r in s_all["records"]))
        blob = json.dumps(s_all)
        self.assertNotIn("TOK_SECRET_RFI", blob)
        self.assertNotIn("TOK_SECRET_NOC", blob)
        self.assertNotIn("TOK_SECRET_PR", blob)
        s_rfi = ai_tools.search_submittals(["ENG-D2"], doc_type="RFI", doc_type_scope=dt_scope)
        self.assertEqual(s_rfi["count"], 0)

        # 2. get_project_summary counts only MS (2), not RFI/NOC/PR (total 6)
        summ = ai_tools.get_project_summary(["ENG-D2"], project_code="ENG-D2", doc_type_scope=dt_scope)
        self.assertEqual(summ["grand_total_documents"], 2)
        self.assertEqual(summ["grand_total_approved"], 1)

        # 3. get_noc_financials returns zeroed financials when NOC is excluded from scope
        noc = ai_tools.get_noc_financials(["ENG-D2"], doc_type_scope=dt_scope)
        self.assertEqual(noc["total_count"], 0)
        self.assertEqual(noc["total_submitted_cost"], 0.0)
        self.assertEqual(noc["total_approved_cost"], 0.0)
        self.assertEqual(noc["items"], [])

        # 4. get_overdue_submittals returns MS-D2-002 only, never RFI-D2-009
        ov = ai_tools.get_overdue_submittals(["ENG-D2"], doc_type_scope=dt_scope)
        self.assertEqual([r["docNo"] for r in ov["records"]], ["MS-D2-002"])

        # 5. get_document_history returns 0 revisions for hidden RFI-D2-009
        hist = ai_tools.get_document_history(["ENG-D2"], "RFI-D2-009", doc_type_scope=dt_scope)
        self.assertEqual(hist["total_revisions"], 0)
        self.assertEqual(hist["revisions"], [])

        # 6. get_registered_doc_types exposes only MS in dt_scope
        reg = ai_tools.get_registered_doc_types(doc_type_scope=dt_scope)
        self.assertEqual([d["code"] for d in reg], ["MS"])

        # 7. HTTP /api/ai/query deterministic fallback + agent path
        c_restr = self.client_for("u_d2_multi")
        with patch.object(ai, "_call_gemini_agent", return_value=(None, "UNAVAILABLE", [])):
            r_noc = c_restr.post("/api/ai/query", json={"prompt": "Show NOC variation orders and cost", "workspace_id": "ENG"}).get_json()["reply"]
            self.assertNotIn("987,654", r_noc)
            self.assertNotIn("543,210", r_noc)

            r_ov = c_restr.post("/api/ai/query", json={"prompt": "Show overdue submittals", "workspace_id": "ENG"}).get_json()["reply"]
            self.assertIn("MS-D2-002", r_ov)
            self.assertNotIn("RFI-D2-009", r_ov)

            r_sum = c_restr.post("/api/ai/query", json={"prompt": "KPI summary for ENG-D2", "workspace_id": "ENG"}).get_json()["reply"]
            self.assertIn("Total=2", r_sum)

        # Agent path: verify execute_tool_call enforces doc_type_scope when model issues tool calls
        res_tool = ai.execute_tool_call("search_submittals", {"doc_type": "RFI"}, ["ENG-D2"], doc_type_scope=dt_scope)
        self.assertEqual(res_tool["count"], 0)


class TestFinding3ExplicitAIProjectSelection(FiveFindingsBase):
    def test_explicit_project_id_and_alias_selection_and_authorization_bounds(self):
        from blueprints import ai_tools

        allowed = ["ENG-D2", "ENG-D2B"]
        # Explicit canonical ID (was previously broken by comparing str ID against dict!)
        sel_id = ai_tools._resolve_target_projects(allowed, "ENG-D2")
        self.assertEqual([p["id"] for p in sel_id], ["ENG-D2"])

        # Explicit legacy code & name
        self.assertEqual([p["id"] for p in ai_tools._resolve_target_projects(allowed, "ENG-CODE-B")], ["ENG-D2B"])
        self.assertEqual([p["id"] for p in ai_tools._resolve_target_projects(allowed, "Beta Engineering")], ["ENG-D2B"])

        # "all" preserves all allowed projects
        self.assertEqual({p["id"] for p in ai_tools._resolve_target_projects(allowed, "all")}, {"ENG-D2", "ENG-D2B"})

        # Unauthorized & ambiguous aliases fail closed
        self.assertEqual(ai_tools._resolve_target_projects(allowed, "ENG-UNAUTH"), [])
        self.assertEqual(ai_tools._resolve_target_projects(allowed, "UNAUTH-CODE"), [])
        self.assertEqual(ai_tools._resolve_target_projects(["ENG-D2", "AMBIG-1"], "AMBIG-C"), [])
        self.assertEqual(ai_tools._resolve_target_projects(["ENG-D2", "AMBIG-1"], "Ambiguous Name"), [])

        # Tool integration using explicit project_code
        noc = ai_tools.get_noc_financials(allowed, project_code="ENG-D2")
        self.assertEqual(noc["total_count"], 1)
        self.assertEqual(noc["by_project"][0]["project_code"], "ENG-CODE-A")


class TestFinding4DBPoolAndNestedTx(FiveFindingsBase):
    def test_setup_and_teardown_fault_injection_never_leaks_connections_or_credentials(self):
        secret_dsn = "postgresql://dcr_admin:SuperSecretPass99@db.internal:5432/prod_db"

        # 1. Cursor setup failure on standalone DB() closes & returns broken connection
        conn_cur_fail = MagicMock(closed=0)
        conn_cur_fail.cursor.side_effect = RuntimeError(f"cursor setup boom {secret_dsn}")
        fake_pool = MagicMock()
        fake_pool.getconn.return_value = conn_cur_fail
        with patch.object(self.db, "get_pool", return_value=fake_pool):
            with self.assertLogs("db", level=logging.ERROR) as cap:
                with self.assertRaises(RuntimeError):
                    with self.db.DB():
                        pass
            fake_pool.putconn.assert_called_once_with(conn_cur_fail, close=True)
            conn_cur_fail.rollback.assert_called_once()
            self.assertNotIn("SuperSecretPass99", "\n".join(cap.output))
            self.assertIn("***:***", "\n".join(cap.output))

        # 2. autocommit setup failure inside transaction() cleans up & marks tx aborted
        conn_ac_fail = MagicMock(closed=0)
        type(conn_ac_fail).autocommit = property(lambda s: False, MagicMock(side_effect=OSError("autocommit fail")))
        fake_pool2 = MagicMock()
        fake_pool2.getconn.return_value = conn_ac_fail
        with patch.object(self.db, "get_pool", return_value=fake_pool2):
            with self.assertLogs("db", level=logging.ERROR):
                with self.assertRaises(OSError):
                    with self.db.transaction():
                        with self.db.DB():
                            pass
            fake_pool2.putconn.assert_called_once_with(conn_ac_fail, close=True)

        # 3. Rollback + putconn failure on exit logs safely and closes raw connection
        conn_exit_fail = MagicMock(closed=0)
        conn_exit_fail.rollback.side_effect = RuntimeError(f"rb boom password=TopSecretPwd")
        fake_pool3 = MagicMock()
        fake_pool3.getconn.return_value = conn_exit_fail
        fake_pool3.putconn.side_effect = RuntimeError("putconn boom")
        with patch.object(self.db, "get_pool", return_value=fake_pool3):
            with self.assertLogs("db", level=logging.ERROR) as cap3:
                with self.assertRaises(ValueError):
                    with self.db.DB():
                        raise ValueError("app logic error")
            conn_exit_fail.close.assert_called_once()
            joined = "\n".join(cap3.output)
            self.assertIn("rb boom", joined)
            self.assertIn("putconn boom", joined)
            self.assertNotIn("TopSecretPwd", joined)

        # 4. Outer transaction commit failure triggers rollback and closes broken conn
        conn_cm_fail = MagicMock(closed=0)
        conn_cm_fail.commit.side_effect = RuntimeError("commit dead")
        fake_pool4 = MagicMock()
        fake_pool4.getconn.return_value = conn_cm_fail
        with patch.object(self.db, "get_pool", return_value=fake_pool4):
            with self.assertLogs("db", level=logging.ERROR):
                with self.assertRaises(RuntimeError):
                    with self.db.transaction():
                        with self.db.DB():
                            pass
            conn_cm_fail.rollback.assert_called_once()
            fake_pool4.putconn.assert_called_once_with(conn_cm_fail, close=True)

    def test_nested_transaction_abort_contract_and_create_project_recovery(self):
        # Nested failure caught by caller MUST roll back outer writes and forbid commit
        with self.assertRaises(self.db.TransactionAbortedError):
            with self.db.transaction():
                self.db.save_record("ENG-D2", "MS", "TX-OUTER-1", {"docNo": "TX-1", "title": "Outer", "status": "Open"})
                try:
                    with self.db.transaction():
                        self.db.save_record("ENG-D2", "MS", "TX-INNER-1", {"docNo": "TX-2", "title": "Inner", "status": "Open"})
                        raise ValueError("inner failure")
                except ValueError:
                    pass
                # Attempting further queries inside poisoned outer tx must also reject
                with self.assertRaises(self.db.TransactionAbortedError):
                    self.db.q("SELECT 1", one=True)

        self.assertIsNone(self.db.get_record_by_id("TX-OUTER-1"))
        self.assertIsNone(self.db.get_record_by_id("TX-INNER-1"))

        # create_project atomicity under _seed_project failure + clean retry
        pid = f"F4-ATOM-{uuid.uuid4().hex[:6]}"
        with patch.object(self.db, "_seed_project", side_effect=RuntimeError("seed boom")):
            with self.assertRaises(RuntimeError):
                self.db.create_project(pid, "Atomic Proj", pid, creator="u_f3_ai", workspace_id="ENG")
        self.assertIsNone(self.db.get_project(pid))
        # Retry succeeds cleanly with healthy pool
        self.db.create_project(pid, "Atomic Proj", pid, creator="u_f3_ai", workspace_id="ENG")
        self.assertIsNotNone(self.db.get_project(pid))
        self.db.delete_project(pid)


class TestFinding5MixedRoleProjectReorder(FiveFindingsBase):
    def test_interleaved_readonly_projects_preserved_and_atomic_jsonb_update(self):
        # Create interleaved multi-workspace lineup with unrelated metadata:
        # RO-0 (OM, order=0), ADM-1 (ENG, order=1), RO-2 (OM, order=2), ADM-3 (ENG, order=3)
        for pid, ws, order_val, extra in (
            ("MR-RO-0", "OM", 0, {"client": "OM Client 0", "keep_tag": "ro0"}),
            ("MR-ADM-1", "ENG", 1, {"client": "ENG Client 1", "keep_tag": "adm1"}),
            ("MR-RO-2", "OM", 2, {"client": "OM Client 2", "keep_tag": "ro2"}),
            ("MR-ADM-3", "ENG", 3, {"client": "ENG Client 3", "keep_tag": "adm3"}),
        ):
            self.db.create_project(pid, pid, pid, workspace_id=ws)
            self.db.exe(
                "UPDATE projects SET data=%s::jsonb WHERE id=%s",
                (json.dumps({**extra, "dashboard_order": order_val}), pid),
            )
        # User has ENG admin (sees MR-ADM-1, MR-ADM-3 + existing ENG projects) and OM viewer (sees MR-RO-0, MR-RO-2)
        u_mr = f"u_mr_{uuid.uuid4().hex[:6]}"
        self.db.add_user(u_mr, "Pw5#1234", "editor", f"{u_mr}@synthetic.local", "ENG", False)
        self.db.assign_user_workspace(u_mr, "ENG", "admin")
        self.db.assign_user_workspace(u_mr, "OM", "viewer")
        self.db.assign_project(u_mr, "MR-RO-0")
        self.db.assign_project(u_mr, "MR-RO-2")
        tok = self.db.create_session(u_mr, "editor")

        c = _STATE["app"].test_client()
        c.set_cookie("dcr_token", tok)

        # Ensure other ENG projects are pushed after MR-ADM-3 so baseline order starts clean
        eng_others = [p["id"] for p in self.db.get_projects_for_user(u_mr, "editor", workspace_id="ENG") if p["id"] not in ("MR-ADM-1", "MR-ADM-3")]
        for k, opid in enumerate(eng_others, start=10):
            self.db.exe("UPDATE projects SET data = jsonb_set(COALESCE(data,'{}'::jsonb), '{dashboard_order}', to_jsonb(%s::int), true) WHERE id=%s", (k, opid))

        # Before reorder: visible order of our 4 projects is [MR-RO-0, MR-ADM-1, MR-RO-2, MR-ADM-3]
        before = [p["id"] for p in c.get("/api/projects").get_json()[:4]]
        self.assertEqual(before, ["MR-RO-0", "MR-ADM-1", "MR-RO-2", "MR-ADM-3"])

        # Swap only the administrable projects across the interleaved RO cards (sending ONLY admin IDs)
        res = c.post("/api/projects/reorder", json={"order": ["MR-ADM-3", "MR-ADM-1"] + eng_others, "workspace_id": "ALL"})
        self.assertEqual(res.status_code, 200)
        body = res.get_json()
        self.assertTrue(body["ok"])
        self.assertEqual(body["order"][:4], ["MR-RO-0", "MR-ADM-3", "MR-RO-2", "MR-ADM-1"])

        # Reloading /api/projects preserves RO-0 and RO-2 exact positions + preserves unrelated metadata!
        after = [p["id"] for p in c.get("/api/projects").get_json()[:4]]
        self.assertEqual(after, ["MR-RO-0", "MR-ADM-3", "MR-RO-2", "MR-ADM-1"])

        for pid, expected_order, expected_tag in (
            ("MR-RO-0", 0, "ro0"),
            ("MR-ADM-3", 1, "adm3"),
            ("MR-RO-2", 2, "ro2"),
            ("MR-ADM-1", 3, "adm1"),
        ):
            pdata = self.db.get_project(pid)
            self.assertEqual(pdata.get("dashboard_order"), expected_order)
            self.assertEqual(pdata.get("keep_tag"), expected_tag)

        # Concurrent reorder calls serialize cleanly via FOR UPDATE + jsonb_set without losing unrelated JSONB keys
        def _run_reorder(seq):
            return self.db.reorder_projects(seq)

        with ThreadPoolExecutor(max_workers=2) as pool:
            f1 = pool.submit(_run_reorder, ["MR-ADM-1", "MR-ADM-3"])
            f2 = pool.submit(_run_reorder, ["MR-ADM-3", "MR-ADM-1"])
            f1.result()
            f2.result()
        self.assertEqual(self.db.get_project("MR-ADM-1").get("keep_tag"), "adm1")
        self.assertEqual(self.db.get_project("MR-ADM-3").get("keep_tag"), "adm3")

        # Cleanup temporary projects
        for pid in ("MR-RO-0", "MR-ADM-1", "MR-RO-2", "MR-ADM-3"):
            self.db.delete_project(pid)

    def test_move_project_js_in_node_preserves_interleaved_order_and_server_sync(self):
        if not _NODE_EXE:
            self.skipTest("Node.js runtime not available")

        c = self.client_for("f1_super")
        html = c.get("/?ws=ENG").data.decode("utf-8")
        scripts = re.findall(r"<script[^>]*>(.*?)</script>", html, re.S)
        dash_script = next((s for s in scripts if "function renderCards" in s and "function moveProject" in s), None)
        self.assertIsNotNone(dash_script)

        harness_js = r"""
const fs = require('fs');
const vm = require('vm');
const payload = JSON.parse(fs.readFileSync(0, 'utf8'));
const src = payload.script;
function extractFn(s, sig) {
  const start = s.indexOf(sig);
  if (start < 0) throw new Error('Missing ' + sig);
  let i = s.indexOf('{', start), depth = 0, inStr = false, q = '', esc = false;
  for (; i < s.length; i++) {
    const ch = s[i];
    if (esc) { esc = false; continue; }
    if (inStr) { if (ch === '\\') esc = true; else if (ch === q) inStr = false; continue; }
    if (ch === "'" || ch === '"' || ch === '`') { inStr = true; q = ch; continue; }
    if (ch === '{') depth++;
    else if (ch === '}') { depth--; if (depth === 0) return s.slice(start, i + 1); }
  }
  throw new Error('Unbalanced braces for ' + sig);
}
const moveSrc = extractFn(src, 'async function moveProject(');
const projSel = { innerHTML: '' };
const calls = [];
const toasts = [];
let mockFail = false;

const ctx = {
  STATS: [
    { id: 'RO-0',  name: 'R0', code: 'R0', workspace_id: 'OM',  can_admin: false },
    { id: 'ADM-1', name: 'A1', code: 'A1', workspace_id: 'ENG', can_admin: true },
    { id: 'RO-2',  name: 'R2', code: 'R2', workspace_id: 'OM',  can_admin: false },
    { id: 'ADM-3', name: 'A3', code: 'A3', workspace_id: 'ENG', can_admin: true },
  ],
  _currentPid: '',
  _currentDisc: '',
  document: { getElementById: id => id === 'proj-sel' ? projSel : null },
  renderAll() {},
  toast(m, k) { toasts.push({ m, k }); },
  async apiFetch(url, opts) {
    const body = JSON.parse(opts.body);
    calls.push({ url, body });
    if (mockFail) return { ok: false, error: 'Rejected' };
    return { ok: true, order: ['RO-0', 'ADM-3', 'RO-2', 'ADM-1'] };
  }
};
vm.runInNewContext(moveSrc, ctx);
(async () => {
  await ctx.moveProject('ADM-1', 1);
  if (JSON.stringify(ctx.STATS.map(x => x.id)) !== JSON.stringify(['RO-0', 'ADM-3', 'RO-2', 'ADM-1'])) {
    throw new Error('Order mismatch: ' + JSON.stringify(ctx.STATS.map(x => x.id)));
  }
  if (JSON.stringify(calls[0].body.order) !== JSON.stringify(['ADM-3', 'ADM-1'])) {
    throw new Error('Submitted order leaked RO project: ' + JSON.stringify(calls[0].body));
  }
  if (calls[0].body.workspace_id !== 'ALL') {
    throw new Error('Expected ALL workspace_id for mixed list: ' + calls[0].body.workspace_id);
  }
  if (!projSel.innerHTML.includes('ADM-3') || projSel.innerHTML.indexOf('ADM-3') > projSel.innerHTML.indexOf('ADM-1')) {
    throw new Error('Select dropdown out of sync with STATS');
  }
  mockFail = true;
  await ctx.moveProject('ADM-3', 1);
  if (JSON.stringify(ctx.STATS.map(x => x.id)) !== JSON.stringify(['RO-0', 'ADM-3', 'RO-2', 'ADM-1'])) {
    throw new Error('Failed reorder mutated STATS');
  }
  console.log('FIVE_NODE_UI_OK');
})().catch(e => { console.error(e); process.exit(1); });
"""
        proc = subprocess.run(
            [_NODE_EXE, "-e", harness_js],
            input=json.dumps({"script": dash_script}),
            text=True,
            capture_output=True,
            timeout=15,
        )
        self.assertEqual(proc.returncode, 0, f"STDOUT:{proc.stdout}\nSTDERR:{proc.stderr}")
        self.assertIn("FIVE_NODE_UI_OK", proc.stdout)


if __name__ == "__main__":
    unittest.main()
