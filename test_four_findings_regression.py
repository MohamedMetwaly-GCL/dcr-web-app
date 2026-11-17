"""Regression suite for the four remaining Codex security findings.

Run through security_test_support.py with a fresh disposable PostgreSQL cluster.
"""

import base64
import io
import json
import re
import shutil
import subprocess
import unittest
import uuid
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from unittest.mock import patch
from contextlib import ExitStack

from security_test_support import disposable_database


_STATE = {}
_NODE_EXE = shutil.which("node") or (
    r"C:\Program Files\nodejs\node.exe" if Path(r"C:\Program Files\nodejs\node.exe").exists() else None
)



def _make_xlsx_b64(sheets: dict) -> str:
    """Create a base64-encoded XLSX workbook from {sheet_name: [rows]}."""
    import openpyxl

    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for name, rows in sheets.items():
        ws = wb.create_sheet(title=name[:31])
        for row in rows:
            ws.append(list(row))
    buf = io.BytesIO()
    wb.save(buf)
    return base64.b64encode(buf.getvalue()).decode("ascii")


def setUpModule():
    stack = ExitStack()
    unittest.addModuleCleanup(stack.close)
    db = stack.enter_context(disposable_database())
    db.create_workspace("HSE", "HSE", "Health, Safety & Environment", "HSE", "🛡️", "hse", 3, "HSE Workspace")

    # ── Seed Projects for Findings 1, 2, and 4 ────────────────────────────────────────────
    # Colliding IDs & aliases for Finding 1:
    db.create_project("ENG-P1", "Shared Plant", "SHARED-CODE", workspace_id="ENG")
    db.create_project("SHARED-CODE", "Trap Project", "TRAP-CODE", workspace_id="ENG")  # ID == ENG-P1.code!
    db.create_project("HSE-P1", "Shared Plant", "SHARED-CODE", workspace_id="HSE")    # shared code & name across ws!
    db.create_project("OM-UNIQ", "Unique OM Facility", "OM-LEGACY-CODE", workspace_id="OM")

    # Dedicated rich project for Finding 2:
    db.create_project("F2-PROJ", "DocType Isolation Project", "F2-CODE", workspace_id="ENG")
    for dt_code, dt_name in (("DS", "DS"), ("RFI", "RFI"), ("PR", "Purchase Request"), ("LTR", "Letters"), ("CUSTOMX", "Custom Deletable")):
        if not db.get_doc_type("F2-PROJ", dt_code):
            db.exe(
                "INSERT INTO doc_types(id, project_id, name, code, sort_order) VALUES(%s,%s,%s,%s,50) ON CONFLICT DO NOTHING",
                (dt_code, "F2-PROJ", dt_name, dt_code),
            )
    # Ensure DS and RFI have a column we can target in column tests
    for dt_c in ("DS", "RFI"):
        if not db.get_columns("F2-PROJ", dt_c):
            db.add_column("F2-PROJ", dt_c, "docNo", "Document No.", "text")
            db.add_column("F2-PROJ", dt_c, "title", "Description / Title", "text")
            db.add_column("F2-PROJ", dt_c, "status", "Status", "status")
    # Mark one custom column so deletion tests have an exact column id
    db.add_column("F2-PROJ", "RFI", "custom_secret_col", "Secret RFI Col", "text")
    db.add_column("F2-PROJ", "DS", "custom_ds_col", "Allowed DS Col", "text")

    # Synthetic records for Finding 1
    db.save_record("ENG-P1", "DS", "F1-REC-ENG1", {"docNo": "F1-ENG1-001", "title": "TOK_F1_ENG1 Secret", "status": "Under Review"})
    db.save_record("SHARED-CODE", "DS", "F1-REC-TRAP", {"docNo": "F1-TRAP-001", "title": "TOK_F1_TRAP Secret", "status": "Under Review"})
    db.save_record("HSE-P1", "DS", "F1-REC-HSE1", {"docNo": "F1-HSE1-001", "title": "TOK_F1_HSE1 Secret", "status": "Under Review"})
    db.save_record("OM-UNIQ", "DS", "F1-REC-OM1", {"docNo": "F1-OM1-001", "title": "TOK_F1_OM1 Secret", "status": "Under Review"})

    # Synthetic records for Finding 2 (DS, RFI, PR + pr_items, LTR thread)
    db.save_record("F2-PROJ", "DS", "F2-DS-1", {
        "docNo": "F2-DS-001", "title": "TOK_F2_DS Visible Design", "status": "Under Review",
        "issuedDate": "2026-01-05", "expectedReplyDate": "2026-01-12", "discipline": "Mechanical",
    })
    db.save_record("F2-PROJ", "RFI", "F2-RFI-1", {
        "docNo": "F2-RFI-001", "title": "TOK_F2_RFI Hidden Query", "status": "Under Review",
        "issuedDate": "2026-01-05", "expectedReplyDate": "2026-01-12", "discipline": "Electrical",
    })
    db.save_record("F2-PROJ", "PR", "F2-PR-1", {
        "docNo": "F2-PR-001", "title": "TOK_F2_PR Hidden Procurement", "status": "Under Review",
        "issuedDate": "2026-01-07", "expectedReplyDate": "2026-01-14",
    })
    db.exe(
        "INSERT INTO pr_items(id, record_id, row_type, item_name, unit, quantity, sort_order) "
        "VALUES (%s,%s,%s,%s,%s,10,0) ON CONFLICT DO NOTHING",
        ("F2-PRI-1", "F2-PR-1", "item", "TOK_F2_PRITEM Secret Pump", "PCS"),
    )
    db.save_record("F2-PROJ", "LTR", "F2-LTR-1", {
        "docNo": "F2-LTR-001", "subject": "TOK_F2_LTR Confidential Parent",
        "title": "TOK_F2_LTR Confidential Parent", "status": "Open",
        "issuedDate": "2026-01-08", "direction": "OUT", "company": "ClientCo",
    })
    db.save_record("F2-PROJ", "LTR", "F2-LTR-2", {
        "docNo": "F2-LTR-002", "subject": "TOK_F2_LTR Confidential Reply",
        "title": "TOK_F2_LTR Confidential Reply", "status": "Replied",
        "issuedDate": "2026-01-09", "direction": "IN", "company": "ClientCo",
        "parentRecordId": "F2-LTR-1",
    })

    # ── Seed Users ────────────────────────────────────────────────────────────────────────
    pw = "Synthetic#" + uuid.uuid4().hex[:8]
    users = {
        # (role, primary_ws, ws_roles, assigned_projects)
        "f_super":         ("superadmin", "ENG", {"ENG": "admin", "HSE": "admin", "OM": "admin"}, ()),
        "f1_p1_only":      ("editor",     "ENG", {"ENG": "editor"}, ("ENG-P1",)),
        "f1_code_trap":    ("editor",     "ENG", {"ENG": "editor"}, ("SHARED-CODE",)),
        "f1_legacy_uniq":  ("editor",     "OM",  {"OM": "editor"},  ("OM-LEGACY-CODE",)),
        "f1_legacy_ambig": ("editor",     "ENG", {"ENG": "editor"}, ("Shared Plant",)),
        "f2_ed_ds":        ("editor",     "ENG", {"ENG": "editor"}, ("F2-PROJ",)),
        "f2_ed_zero":      ("editor",     "ENG", {"ENG": "editor"}, ("F2-PROJ",)),
        "f2_ed_unrestr":   ("editor",     "ENG", {"ENG": "editor"}, ("F2-PROJ",)),
        "f2_admin":        ("admin",      "ENG", {"ENG": "admin"},  ()),
        "f4_mixed_ed":     ("editor",     "ENG", {"ENG": "admin", "HSE": "viewer"}, ("HSE-P1",)),
        "f4_eng_adm_om_vw":("admin",      "ENG", {"ENG": "admin", "OM": "viewer"}, ("OM-UNIQ",)),
        "f4_viewer":       ("viewer",     "ENG", {"ENG": "viewer"}, ("ENG-P1", "F2-PROJ")),
    }
    tokens = {}
    for uname, (role, pws, ws_roles, projs) in users.items():
        db.add_user(uname, pw, role, f"{uname}@synthetic.local", pws, False)
        for ws_id, ws_role in ws_roles.items():
            db.assign_user_workspace(uname, ws_id, ws_role)
        for pid_assign in projs:
            db.assign_project(uname, pid_assign)
        tokens[uname] = db.create_session(uname, role)

    db.set_user_project_doc_types("f2_ed_ds", "F2-PROJ", ["DS"])
    db.set_user_project_doc_types("f2_ed_zero", "F2-PROJ", [])

    import auth
    from app import app
    app.config["TESTING"] = True
    _STATE["tokens"] = tokens
    _STATE["app"] = app
    _STATE["db"] = db
    _STATE["auth"] = auth



class FourFindingsBase(unittest.TestCase):
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


class TestFinding1CanonicalAliases(FourFindingsBase):
    def test_exact_id_permission_never_expands_to_matching_code_or_name(self):
        # f1_code_trap is assigned canonical project SHARED-CODE, which happens to equal ENG-P1.code and HSE-P1.code.
        u = self.db.get_user("f1_code_trap")
        allowed = self.auth.get_allowed_project_ids(u)
        self.assertEqual(allowed, ["SHARED-CODE"])
        self.assertTrue(self.auth.can_view_project("SHARED-CODE", u))
        self.assertTrue(self.auth.has_project_access(u["username"], "SHARED-CODE", u["role"]))
        self.assertFalse(self.auth.can_view_project("ENG-P1", u))
        self.assertFalse(self.auth.can_view_project("HSE-P1", u))

        # db.get_projects(['SHARED-CODE']) returns ONLY SHARED-CODE
        projs = self.db.get_projects(project_ids=["SHARED-CODE"])
        self.assertEqual([p["id"] for p in projs], ["SHARED-CODE"])

        # API routes and search never leak ENG-P1 or HSE-P1 to f1_code_trap
        c = self.client_for("f1_code_trap")
        r_proj = c.get("/api/projects")
        self.assertEqual(r_proj.status_code, 200)
        self.assertEqual([p["id"] for p in r_proj.get_json()], ["SHARED-CODE"])
        self.assertEqual(c.get("/api/records/ENG-P1/DS").status_code, 403)
        self.assertEqual(c.get("/api/records/HSE-P1/DS").status_code, 403)

        r_search = c.get("/api/records/search/global?q=TOK_F1_")
        self.assertEqual(r_search.status_code, 200)
        pids = {row["project_id"] for row in r_search.get_json()["results"]}
        self.assertEqual(pids, {"SHARED-CODE"})

    def test_unambiguous_legacy_alias_works_and_ambiguous_alias_fails_closed(self):
        # Legacy unambiguous code OM-LEGACY-CODE resolves to OM-UNIQ
        u_uniq = self.db.get_user("f1_legacy_uniq")
        self.assertEqual(self.auth.get_allowed_project_ids(u_uniq), ["OM-UNIQ"])
        c_uniq = self.client_for("f1_legacy_uniq")
        self.assertEqual(c_uniq.get("/api/records/OM-UNIQ/DS").status_code, 200)
        r_s = c_uniq.get("/api/records/search/global?q=TOK_F1_OM1&project_id=OM-LEGACY-CODE")
        self.assertEqual(r_s.status_code, 200)
        self.assertEqual([r["project_id"] for r in r_s.get_json()["results"]], ["OM-UNIQ"])

        # Legacy ambiguous name 'Shared Plant' (matches both ENG-P1 and HSE-P1) grants zero access
        u_ambig = self.db.get_user("f1_legacy_ambig")
        self.assertEqual(self.auth.get_allowed_project_ids(u_ambig), [])
        self.assertFalse(self.auth.can_view_project("ENG-P1", u_ambig))
        self.assertFalse(self.auth.can_view_project("HSE-P1", u_ambig))
        c_ambig = self.client_for("f1_legacy_ambig")
        self.assertEqual(c_ambig.get("/api/projects").get_json(), [])
        self.assertEqual(c_ambig.get("/api/records/ENG-P1/DS").status_code, 403)

        # User with ENG-P1 searching by unauthorized exact ID 'SHARED-CODE' or ambiguous 'Shared Plant' is denied
        c_p1 = self.client_for("f1_p1_only")
        r_trap = c_p1.get("/api/records/search/global?q=TOK_F1_&project_id=SHARED-CODE")
        self.assertIn(r_trap.status_code, (200, 403))
        if r_trap.status_code == 200:
            self.assertEqual(r_trap.get_json()["results"], [])
        r_amb = c_p1.get("/api/records/search/global?q=TOK_F1_&project_id=Shared%20Plant")
        self.assertIn(r_amb.status_code, (200, 403))
        if r_amb.status_code == 200:
            self.assertEqual(r_amb.get_json()["results"], [])

    def test_ai_tools_resolution_and_superadmin_cross_workspace_access(self):
        from blueprints import ai_tools

        # Exact-ID permission SHARED-CODE resolves ONLY to SHARED-CODE, never ENG-P1 or HSE-P1
        self.assertEqual(ai_tools._resolve_project_aliases("all", ["SHARED-CODE"]), ["SHARED-CODE"])
        self.assertEqual(ai_tools._resolve_project_aliases("SHARED-CODE", ["ENG-P1"]), [])
        self.assertEqual(ai_tools._resolve_project_aliases("Shared Plant", ["ENG-P1"]), [])
        self.assertEqual(ai_tools._resolve_project_aliases("OM-LEGACY-CODE", ["OM-UNIQ"]), ["OM-UNIQ"])

        projs_trap = ai_tools._resolve_target_projects(["SHARED-CODE"], None)
        self.assertEqual([p["id"] for p in projs_trap], ["SHARED-CODE"])
        self.assertEqual(ai_tools._resolve_target_projects(["ENG-P1"], "SHARED-CODE"), [])
        self.assertEqual(ai_tools._resolve_target_projects(["ENG-P1"], "Shared Plant"), [])

        # Global superadmin sees all projects across ENG/HSE/OM
        c_sa = self.client_for("f_super")
        sa_pids = {p["id"] for p in c_sa.get("/api/projects").get_json()}
        self.assertTrue({"ENG-P1", "SHARED-CODE", "HSE-P1", "OM-UNIQ", "F2-PROJ"}.issubset(sa_pids))


class TestFinding2DocumentTypeConsistency(FourFindingsBase):
    def test_global_search_filters_denied_doc_types_and_counts_accurately(self):
        c_ds = self.client_for("f2_ed_ds")
        r = c_ds.get("/api/records/search/global?q=TOK_F2_")
        self.assertEqual(r.status_code, 200)
        body = r.get_json()
        self.assertEqual(body["total_found"], 1)
        self.assertEqual([(x["project_id"], x["dt_id"]) for x in body["results"]], [("F2-PROJ", "DS")])

        # Explicitly asking for denied dt_id yields 0 results
        r_rfi = c_ds.get("/api/records/search/global?q=TOK_F2_&dt_id=RFI")
        self.assertEqual(r_rfi.status_code, 200)
        self.assertEqual(r_rfi.get_json()["total_found"], 0)
        self.assertEqual(r_rfi.get_json()["results"], [])

        # Unrestricted editor sees all 5 records (DS, RFI, PR, LTR x2)
        c_unr = self.client_for("f2_ed_unrestr")
        body_u = c_unr.get("/api/records/search/global?q=TOK_F2_").get_json()
        self.assertEqual(body_u["total_found"], 5)

    def test_exports_and_counts_respect_allowed_doc_types(self):
        import openpyxl

        c_ds = self.client_for("f2_ed_ds")
        c_zero = self.client_for("f2_ed_zero")

        # Single-tab Excel & PDF
        self.assertEqual(c_ds.get("/api/export/F2-PROJ/RFI").status_code, 403)
        self.assertEqual(c_ds.get("/api/export/F2-PROJ/PR").status_code, 403)
        self.assertEqual(c_ds.get("/api/export_pdf/F2-PROJ/RFI").status_code, 403)
        self.assertEqual(c_ds.get("/api/export/F2-PROJ/DS").status_code, 200)
        r_pdf_ds = c_ds.get("/api/export_pdf/F2-PROJ/DS")
        self.assertEqual(r_pdf_ds.status_code, 200)
        self.assertTrue(r_pdf_ds.data.startswith(b"%PDF"))

        # export_all returns ONLY allowed DS sheets (no RFI / PR / LTR sheets or tokens)
        r_all = c_ds.get("/api/export_all/F2-PROJ")
        self.assertEqual(r_all.status_code, 200)
        wb = openpyxl.load_workbook(io.BytesIO(r_all.data), data_only=True)
        self.assertNotIn("RFI", wb.sheetnames)
        self.assertNotIn("PR", wb.sheetnames)
        self.assertNotIn("LTR", wb.sheetnames)
        self.assertIn("DS", wb.sheetnames)
        all_cells = " ".join(str(cell) for s in wb.sheetnames for row in wb[s].iter_rows(values_only=True) for cell in row if cell is not None)
        self.assertIn("TOK_F2_DS", all_cells)
        self.assertNotIn("TOK_F2_RFI", all_cells)
        self.assertNotIn("TOK_F2_PR", all_cells)
        self.assertNotIn("TOK_F2_LTR", all_cells)

        # User with empty allowed_doc_types [] cannot export_all or export_pdf_all
        self.assertEqual(c_zero.get("/api/export_all/F2-PROJ").status_code, 403)
        self.assertEqual(c_zero.get("/api/export_pdf_all/F2-PROJ").status_code, 403)

        # export_pdf_all filters dts inside _build_executive_summary_pdf
        self.assertEqual(c_ds.get("/api/export_pdf_all/F2-PROJ").status_code, 200)
        from blueprints import exporting
        with _STATE["app"].test_request_context("/api/export_pdf_all/F2-PROJ"):
            with patch.object(exporting.db, "get_records", wraps=exporting.db.get_records) as spy_rec:
                u_ds = self.db.get_user("f2_ed_ds")
                exporting._build_executive_summary_pdf("F2-PROJ", "all", user=u_ds)
                queried_dts = [call.args[1] for call in spy_rec.call_args_list]
                self.assertEqual(queried_dts, ["DS"])

        # export_count
        self.assertEqual(c_ds.get("/api/export_count/F2-PROJ?scope=all").get_json()["count"], 1)
        self.assertEqual(c_ds.get("/api/export_count/F2-PROJ?scope=RFI").status_code, 403)
        self.assertEqual(c_ds.get("/api/export_count/F2-PROJ?scope=current&tab=RFI").status_code, 403)
        self.assertEqual(c_ds.get("/api/export_count/F2-PROJ?scope=current&tab=DS").get_json()["count"], 1)

    def test_imports_block_unauthorized_tabs_pr_raw_redirect_and_bulk_project(self):
        c_ds = self.client_for("f2_ed_ds")

        # 1. Direct import to denied RFI
        csv_b64 = base64.b64encode(b"Document No.,Title\nF2-RFI-NEW,Injected\n").decode("ascii")
        self.assertEqual(c_ds.post("/api/import/F2-PROJ/RFI", json={"file_b64": csv_b64, "ext": "csv"}).status_code, 403)

        # 2. PR Items Raw bypass: uploaded against DS endpoint with PR Raw template columns
        pr_raw_b64 = _make_xlsx_b64({
            "PR Items Raw": [
                ["Row Type", "PR No.", "Description", "Qty", "Unit"],
                ["PARENT", "F2-PR-BYPASS", "Bypass Parent PR", 1, "LOT"],
                ["ITEM", "F2-PR-BYPASS", "Bypass Secret Line", 5, "EA"],
            ]
        })
        r_bypass = c_ds.post("/api/import/F2-PROJ/DS", json={"file_b64": pr_raw_b64, "ext": "xlsx"})
        self.assertEqual(r_bypass.status_code, 403)
        self.assertFalse(any(r.get("docNo") == "F2-PR-BYPASS" for r in self.db.get_records("F2-PROJ", "PR")))

        # 3. Bulk import_project containing DS + RFI sheets must fail atomic before writing DS
        bulk_b64 = _make_xlsx_b64({
            "DS": [["Document No.", "Title"], ["F2-DS-BULK-ATOMIC", "Should Rollback Entirely"]],
            "RFI": [["Document No.", "Title"], ["F2-RFI-BULK-BAD", "Denied Tab"]],
        })
        r_bulk = c_ds.post("/api/import_project/F2-PROJ", json={"file_b64": bulk_b64, "ext": "xlsx"})
        self.assertEqual(r_bulk.status_code, 403)
        self.assertFalse(any(r.get("docNo") == "F2-DS-BULK-ATOMIC" for r in self.db.get_records("F2-PROJ", "DS")))

        # 4. Legitimate DS-only import works
        ok_b64 = base64.b64encode(b"Document No.,Description / Title\nF2-DS-LEGIT,Valid DS Import\n").decode("ascii")
        r_ok = c_ds.post("/api/import/F2-PROJ/DS", json={"file_b64": ok_b64, "ext": "csv"})
        self.assertEqual(r_ok.status_code, 200)
        self.assertTrue(r_ok.get_json()["ok"])
        # Cleanup imported record so counts stay predictable for other tests
        rec_legit = next(r for r in self.db.get_records("F2-PROJ", "DS") if r.get("docNo") == "F2-DS-LEGIT")
        self.db.delete_record(rec_legit["_id"])

    def test_doc_type_and_column_mutations_enforce_restrictions_and_preserve_unrestricted_editors(self):
        c_ds = self.client_for("f2_ed_ds")
        c_unr = self.client_for("f2_ed_unrestr")

        # Restricted editor listing & counts return DS only
        self.assertEqual([d["id"] for d in c_ds.get("/api/doc_types/F2-PROJ").get_json()], ["DS"])
        counts_ds = c_ds.get("/api/counts/F2-PROJ").get_json()
        self.assertEqual(set(counts_ds.keys()), {"DS"})

        # Restricted editor blocked from mutating/adding/reordering/deleting doc types
        self.assertEqual(c_ds.post("/api/doc_types/F2-PROJ", json={"code": "ZZ", "name": "Blocked"}).status_code, 403)
        self.assertEqual(c_ds.patch("/api/doc_types/F2-PROJ/RFI", json={"code": "RFI", "name": "Hacked"}).status_code, 403)
        self.assertEqual(c_ds.post("/api/doc_types/F2-PROJ/reorder", json={"order": ["RFI", "DS"]}).status_code, 403)
        self.assertEqual(c_ds.delete("/api/doc_types/F2-PROJ/RFI").status_code, 403)

        # Unrestricted editor preserves doc-type management rights (create/rename/reorder/delete custom)
        self.assertEqual(c_unr.post("/api/doc_types/F2-PROJ", json={"code": "EDNEW", "name": "Editor Created"}).status_code, 200)
        self.assertEqual(c_unr.patch("/api/doc_types/F2-PROJ/EDNEW", json={"code": "EDNEW", "name": "Editor Renamed"}).status_code, 200)
        self.assertEqual(c_unr.post("/api/doc_types/F2-PROJ/reorder", json={"order": ["DS", "RFI", "PR", "LTR", "CUSTOMX", "EDNEW"]}).status_code, 200)
        self.assertEqual(c_unr.delete("/api/doc_types/F2-PROJ/EDNEW").status_code, 200)

        # Columns enforcement on denied RFI vs allowed DS
        rfi_col = next(c for c in self.db.get_columns("F2-PROJ", "RFI") if c["col_key"] == "custom_secret_col")
        ds_col = next(c for c in self.db.get_columns("F2-PROJ", "DS") if c["col_key"] == "custom_ds_col")
        self.assertEqual(c_ds.get("/api/columns/F2-PROJ/RFI").status_code, 403)
        self.assertEqual(c_ds.post("/api/columns/F2-PROJ/RFI", json={"label": "Blocked"}).status_code, 403)
        self.assertEqual(c_ds.post(f"/api/columns/visibility/{rfi_col['id']}", json={"visible": False}).status_code, 403)
        self.assertEqual(c_ds.post(f"/api/columns/rename/{rfi_col['id']}", json={"label": "Hack"}).status_code, 403)
        self.assertEqual(c_ds.delete(f"/api/columns/{rfi_col['id']}").status_code, 403)
        self.assertEqual(c_ds.post("/api/columns/reorder/F2-PROJ/RFI", json={"order": [rfi_col["id"]]}).status_code, 403)
        self.assertEqual(c_ds.post("/api/columns/reorder", json={"pid": "F2-PROJ", "dt_id": "RFI", "col_id": rfi_col["id"], "new_order": 0}).status_code, 403)
        self.assertEqual(c_ds.get("/api/col_width/F2-PROJ/RFI").status_code, 403)
        # Allowed DS column operations work
        self.assertEqual(c_ds.get("/api/columns/F2-PROJ/DS").status_code, 200)
        self.assertEqual(c_ds.post(f"/api/columns/rename/{ds_col['id']}", json={"label": "Allowed DS Col Renamed"}).status_code, 200)

    def test_letters_daily_digest_dashboard_cache_and_aggregates(self):
        c_admin = self.client_for("f2_admin")
        c_ds = self.client_for("f2_ed_ds")

        # Warm dashboard_cache and PR analytics cache as unrestricted admin first
        # (workflow KPI counts DS+RFI+PR=3; LTR correspondence is reported under p['ltr']['total']==2)
        admin_stats = next(p for p in c_admin.get("/api/dashboard_stats?workspace_id=ENG").get_json() if p["id"] == "F2-PROJ")
        self.assertEqual(admin_stats["total"], 3)
        self.assertEqual(admin_stats["ltr"]["total"], 2)
        c_admin.get("/api/pr_analytics_summary?project_id=F2-PROJ")

        # Restricted user must bypass dashboard_cache -> sees total=1 (DS only) and ltr.total=0
        ds_stats = next(p for p in c_ds.get("/api/dashboard_stats?workspace_id=ENG").get_json() if p["id"] == "F2-PROJ")
        self.assertEqual(ds_stats["total"], 1)
        self.assertEqual(ds_stats["ltr"]["total"], 0)

        # Unrestricted admin still gets full cached counts (restricted call didn't poison cache)
        admin_stats_after = next(p for p in c_admin.get("/api/dashboard_stats?workspace_id=ENG").get_json() if p["id"] == "F2-PROJ")
        self.assertEqual(admin_stats_after["total"], 3)
        self.assertEqual(admin_stats_after["ltr"]["total"], 2)

        # Daily digest excludes RFI / PR / LTR (both DS and RFI are dated 2026-01-05)
        digest = c_ds.get("/api/daily_digest/F2-PROJ?date=2026-01-05").get_json()
        digest_json = json.dumps(digest)
        self.assertIn("F2-DS-001", digest_json)
        self.assertNotIn("F2-RFI-001", digest_json)
        self.assertNotIn("F2-PR-001", digest_json)
        self.assertNotIn("F2-LTR-001", digest_json)

        # Letters endpoints denied when LTR is excluded
        self.assertEqual(c_ds.get("/api/letters/parent-options/F2-PROJ").status_code, 403)
        self.assertEqual(c_ds.get("/api/letters/thread/F2-PROJ/F2-LTR-2").status_code, 403)
        self.assertEqual(c_ds.get("/api/letters/timeline/F2-PROJ/F2-LTR-2").status_code, 403)
        self.assertEqual(c_admin.get("/api/letters/parent-options/F2-PROJ").status_code, 200)
        self.assertEqual(c_admin.get("/api/letters/thread/F2-PROJ/F2-LTR-2").status_code, 200)
        self.assertEqual(c_admin.get("/api/letters/timeline/F2-PROJ/F2-LTR-2").status_code, 200)

        # Analytics & Summary endpoints include only DS for f2_ed_ds (even after unrestricted cache warm)
        self.assertEqual(c_ds.get("/api/records_range/F2-PROJ/RFI").status_code, 403)
        trend_ds = c_ds.get("/api/analytics/trend?pid=F2-PROJ").get_json()
        self.assertEqual(next(m for m in trend_ds if m["month"] == "2026-01")["submitted"], 1)
        self.assertEqual(sum(m["submitted"] for m in trend_ds), 1)
        trend_adm = c_admin.get("/api/analytics/trend?pid=F2-PROJ").get_json()
        self.assertGreaterEqual(sum(m["submitted"] for m in trend_adm), 4)

        aging_ds = c_ds.get("/api/analytics/aging?pid=F2-PROJ").get_json()
        self.assertEqual(sum(x["count"] for x in aging_ds), 1)
        aging_adm = c_admin.get("/api/analytics/aging?pid=F2-PROJ").get_json()
        self.assertEqual(sum(x["count"] for x in aging_adm), 5)

        qual_ds = c_ds.get("/api/analytics/quality?pid=F2-PROJ").get_json()
        self.assertEqual(sum(x["count"] for x in qual_ds), 1)
        qual_adm = c_admin.get("/api/analytics/quality?pid=F2-PROJ").get_json()
        self.assertGreaterEqual(sum(x["count"] for x in qual_adm), 4)

        ov_docs = {r["docNo"] for r in c_ds.get("/api/analytics/overdue?pid=F2-PROJ").get_json()}
        self.assertEqual(ov_docs, {"F2-DS-001"})
        dig_ov = {r["docNo"] for r in c_ds.get("/api/overdue_digest").get_json()["records"]}
        self.assertNotIn("F2-RFI-001", dig_ov)
        self.assertNotIn("F2-PR-001", dig_ov)
        exec_sum = c_ds.get("/api/executive_summary?workspace_id=ENG").get_json()
        f2_exec = next(p for p in exec_sum["projects"] if p["code"] == "F2-CODE")
        self.assertEqual(f2_exec["total"], 1)
        dq_sum = c_ds.get("/api/data_quality_summary?project_id=F2-PROJ").get_json()
        self.assertEqual(dq_sum["total_records"], 1)
        self.assertEqual(dq_sum["pr_records"], 0)
        act_sum = c_ds.get("/api/action_required_summary?project_id=F2-PROJ").get_json()
        act_docs = {r["docNo"] for r in (act_sum["top_overdue"] + act_sum["pending_longest"] + act_sum["recent_rejected"])}
        self.assertEqual(act_docs, {"F2-DS-001"})
        pr_sum = c_ds.get("/api/pr_analytics_summary?project_id=F2-PROJ").get_json()
        self.assertEqual(pr_sum["total_pr_records"], 0)
        self.assertEqual(c_admin.get("/api/pr_analytics_summary?project_id=F2-PROJ").get_json()["total_pr_records"], 1)


class TestFinding3AtomicProjectCreation(FourFindingsBase):
    def test_failed_seed_or_assignment_rolls_back_and_retries_cleanly(self):
        pid = "F3-ATOMIC-1"
        orig_assign = self.db.assign_project

        def boom(*args, **kwargs):
            raise RuntimeError("injected post-seed failure")

        # Inject failure after project + doc_types/columns seeding inside create_project
        with patch.object(self.db, "assign_project", side_effect=boom):
            with self.assertRaises(RuntimeError):
                self.db.create_project(pid, "Atomic Fail Test", "F3A", workspace_id="ENG", creator="f2_admin")

        # Assert complete rollback across projects, doc_types, columns_config, user_projects
        self.assertIsNone(self.db.q("SELECT id FROM projects WHERE id=%s", (pid,), one=True))
        self.assertEqual(self.db.q("SELECT COUNT(*) AS n FROM doc_types WHERE project_id=%s", (pid,), one=True)["n"], 0)
        self.assertEqual(self.db.q("SELECT COUNT(*) AS n FROM columns_config WHERE project_id=%s", (pid,), one=True)["n"], 0)
        self.assertEqual(self.db.q("SELECT COUNT(*) AS n FROM user_projects WHERE project_id=%s", (pid,), one=True)["n"], 0)

        # Same pid retries cleanly without conflict or corruption
        self.assertIsNotNone(orig_assign)
        self.db.create_project(pid, "Atomic Retry OK", "F3A", workspace_id="ENG", creator="f2_admin")
        self.assertIsNotNone(self.db.q("SELECT id FROM projects WHERE id=%s", (pid,), one=True))
        self.assertGreater(self.db.q("SELECT COUNT(*) AS n FROM doc_types WHERE project_id=%s", (pid,), one=True)["n"], 0)
        self.assertEqual(self.db.q("SELECT COUNT(*) AS n FROM user_projects WHERE username='f2_admin' AND project_id=%s", (pid,), one=True)["n"], 1)

    def test_concurrent_create_project_atomicity(self):
        pid = "F3-CONCURRENT-1"

        def worker(idx):
            try:
                self.db.create_project(pid, f"Conc {idx}", f"F3C{idx}", workspace_id="ENG", creator="f2_admin")
                return True
            except self.db.ProjectOwnershipError:
                return False

        with ThreadPoolExecutor(max_workers=4) as pool:
            results = list(pool.map(worker, range(4)))
        self.assertEqual(results.count(True), 1)
        self.assertEqual(results.count(False), 3)
        self.assertEqual(self.db.q("SELECT COUNT(*) AS n FROM projects WHERE id=%s", (pid,), one=True)["n"], 1)


class TestFinding4MixedRoleReorderAndUI(FourFindingsBase):
    def test_can_admin_flags_and_server_reorder_enforcement(self):
        # f4_eng_adm_om_vw: ENG admin + OM viewer on OM-UNIQ
        c_mixed_adm = self.client_for("f4_eng_adm_om_vw")
        eng_stats = {p["id"]: p["can_admin"] for p in c_mixed_adm.get("/api/dashboard_stats?ws=ENG").get_json()}
        om_stats = {p["id"]: p["can_admin"] for p in c_mixed_adm.get("/api/dashboard_stats?ws=OM").get_json()}
        self.assertTrue(eng_stats["ENG-P1"])
        self.assertTrue(eng_stats["F2-PROJ"])
        self.assertFalse(om_stats["OM-UNIQ"])

        projs_api = {p["id"]: p["can_admin"] for p in c_mixed_adm.get("/api/projects").get_json()}
        self.assertTrue(projs_api["ENG-P1"])
        self.assertTrue(projs_api["F2-PROJ"])
        self.assertFalse(projs_api["OM-UNIQ"])

        # f4_mixed_ed: global role 'editor', ENG 'admin', HSE 'viewer' assigned HSE-P1
        c_promoted = self.client_for("f4_mixed_ed")
        eng_prom = {p["id"]: p["can_admin"] for p in c_promoted.get("/api/dashboard_stats?ws=ENG").get_json()}
        hse_prom = {p["id"]: p["can_admin"] for p in c_promoted.get("/api/dashboard_stats?ws=HSE").get_json()}
        self.assertTrue(eng_prom["ENG-P1"])
        self.assertFalse(hse_prom["HSE-P1"])

        # Authorized reorder of admin projects works
        r_ok = c_promoted.post("/api/projects/reorder", json={"order": ["F2-PROJ", "SHARED-CODE", "ENG-P1"]})
        self.assertEqual(r_ok.status_code, 200)
        f2_data = self.db.get_project("F2-PROJ")
        self.assertEqual(f2_data.get("dashboard_order"), 0)

        # Including read-only HSE-P1 or unauthorized OM-UNIQ rejects with 403 and preserves DB order
        self.assertEqual(c_promoted.post("/api/projects/reorder", json={"order": ["F2-PROJ", "HSE-P1"]}).status_code, 403)
        self.assertEqual(c_mixed_adm.post("/api/projects/reorder", json={"order": ["F2-PROJ", "OM-UNIQ"]}).status_code, 403)
        self.assertEqual(self.db.get_project("F2-PROJ").get("dashboard_order"), 0)
        self.assertEqual(self.db.get_project("ENG-P1").get("dashboard_order"), 2)

        # Invalid payloads / viewer / anonymous rejected cleanly
        self.assertEqual(c_promoted.post("/api/projects/reorder", json={"order": "not-a-list"}).status_code, 400)
        self.assertEqual(c_promoted.post("/api/projects/reorder", json={"order": ["Shared Plant"]}).status_code, 404)
        self.assertEqual(self.client_for("f4_viewer").post("/api/projects/reorder", json={"order": ["ENG-P1"]}).status_code, 403)
        self.assertEqual(self.client_for(None).post("/api/projects/reorder", json={"order": ["ENG-P1"]}).status_code, 403)

    def test_rendered_dashboard_html_and_move_project_js_in_node(self):
        if not _NODE_EXE:
            self.skipTest("Node.js runtime not available")

        # Render actual Flask dashboard page for ENG (?ws=ENG bypasses multi-workspace launchpad)
        c_promoted = self.client_for("f4_mixed_ed")
        resp = c_promoted.get("/?ws=ENG")
        self.assertEqual(resp.status_code, 200)
        html = resp.data.decode("utf-8")
        stats_json = c_promoted.get("/api/dashboard_stats?ws=ENG").get_json() + c_promoted.get("/api/dashboard_stats?ws=HSE").get_json()

        # Extract the script containing renderCards & moveProject from rendered HTML
        scripts = re.findall(r"<script[^>]*>(.*?)</script>", html, re.S)
        dash_script = next((s for s in scripts if "function renderCards" in s and "function moveProject" in s), None)
        self.assertIsNotNone(dash_script, "Dashboard HTML did not contain renderCards/moveProject script")

        # Execute the real renderCards & moveProject code in Node with minimal DOM shims.
        harness_js = r"""
const fs = require('fs');
const vm = require('vm');
const payload = JSON.parse(fs.readFileSync(0, 'utf8'));
const scriptSrc = payload.script;
const statsData = payload.stats;

function extractFn(src, sig) {
  const start = src.indexOf(sig);
  if (start < 0) throw new Error('Missing ' + sig);
  let i = src.indexOf('{', start), depth = 0, inStr = false, q = '', esc = false;
  for (; i < src.length; i++) {
    const ch = src[i];
    if (esc) { esc = false; continue; }
    if (inStr) {
      if (ch === '\\') esc = true;
      else if (ch === q) inStr = false;
      continue;
    }
    if (ch === "'" || ch === '"' || ch === '`') { inStr = true; q = ch; continue; }
    if (ch === '{') depth++;
    else if (ch === '}') { depth--; if (depth === 0) return src.slice(start, i + 1); }
  }
  throw new Error('Unbalanced braces for ' + sig);
}

const renderCardsSrc = extractFn(scriptSrc, 'function renderCards(');
const moveProjectSrc = extractFn(scriptSrc, 'async function moveProject(');

const cards = [];
const grid = {
  innerHTML: '',
  style: { setProperty() {} },
  appendChild(el) { cards.push(el); }
};
const projSel = { innerHTML: '' };
const apiCalls = [];
const toasts = [];
let mockFail = false;

const ctx = {
  ROLE: 'editor',
  STATS: statsData,
  _currentPid: '',
  _currentDisc: '',
  document: {
    getElementById(id) {
      if (id === 'pgrid') return grid;
      if (id === 'proj-sel') return projSel;
      return { style: { setProperty() {} }, innerHTML: '', classList: { contains: () => false } };
    },
    createElement(tag) {
      return { tagName: tag, className: '', innerHTML: '', href: '', onclick: null, style: {} };
    }
  },
  renderAll() {},
  openM() {},
  toast(msg, kind) { toasts.push({ msg, kind }); },
  async apiFetch(url, opts) {
    apiCalls.push({ url, body: JSON.parse(opts.body) });
    if (mockFail) return { ok: false, error: 'Forbidden' };
    return { ok: true };
  }
};

vm.runInNewContext(renderCardsSrc + '\n' + moveProjectSrc, ctx);

(async () => {
  const mixedInput = statsData.concat([
    { id: 'READONLY-OM', name: 'Readonly OM', code: 'RO-OM', total: 1, approved: 0, pending: 1, pct: 0, can_admin: false }
  ]);
  ctx.renderCards(mixedInput, '');
  const adminRendered = [];
  const readonlyRendered = [];
  for (const el of cards) {
    const hasMove = el.innerHTML.includes('pcard-move');
    if (el.href && (el.href.includes('READONLY-OM') || el.href.includes('HSE-P1'))) readonlyRendered.push(hasMove);
    else adminRendered.push(hasMove);
  }

  // Test 1: Only can_admin cards get .pcard-move buttons even when ROLE==='editor'
  if (!adminRendered.length || adminRendered.some(x => !x)) {
    throw new Error('Expected can_admin cards to render .pcard-move: ' + JSON.stringify(adminRendered));
  }
  if (!readonlyRendered.length || readonlyRendered.some(x => x)) {
    throw new Error('Expected readonly card to hide .pcard-move: ' + JSON.stringify(readonlyRendered));
  }

  // Interleave read-only card between two admin cards and test moveProject skips readonly and submits ONLY can_admin IDs
  ctx.STATS = [
    { id: 'ADM-1', name: 'A1', code: 'A1', total: 1, approved: 0, pending: 1, pct: 0, can_admin: true },
    { id: 'RO-1',  name: 'R1', code: 'R1', total: 1, approved: 0, pending: 1, pct: 0, can_admin: false },
    { id: 'ADM-2', name: 'A2', code: 'A2', total: 1, approved: 0, pending: 1, pct: 0, can_admin: true },
  ];
  await ctx.moveProject('ADM-1', 1);
  if (JSON.stringify(ctx.STATS.map(p => p.id)) !== JSON.stringify(['ADM-2', 'RO-1', 'ADM-1'])) {
    throw new Error('Unexpected STATS after moveProject: ' + JSON.stringify(ctx.STATS.map(p => p.id)));
  }
  if (JSON.stringify(apiCalls[0].body.order) !== JSON.stringify(['ADM-2', 'ADM-1'])) {
    throw new Error('Order payload leaked non-admin project: ' + JSON.stringify(apiCalls[0].body));
  }

  // Test 3: Server rejection keeps original STATS order and emits error toast
  mockFail = true;
  await ctx.moveProject('ADM-2', 1);
  if (JSON.stringify(ctx.STATS.map(p => p.id)) !== JSON.stringify(['ADM-2', 'RO-1', 'ADM-1'])) {
    throw new Error('STATS mutated on failed reorder!');
  }
  if (toasts[toasts.length - 1].kind !== 'err') {
    throw new Error('Expected error toast on failed reorder');
  }
  console.log('NODE_UI_OK');
})().catch(e => { console.error(e); process.exit(1); });
"""
        proc = subprocess.run(
            [_NODE_EXE, "-e", harness_js],
            input=json.dumps({"script": dash_script, "stats": stats_json}),
            text=True,
            capture_output=True,
            timeout=15,
        )
        self.assertEqual(proc.returncode, 0, f"Node UI check failed:\nSTDOUT:{proc.stdout}\nSTDERR:{proc.stderr}")
        self.assertIn("NODE_UI_OK", proc.stdout)


if __name__ == "__main__":
    unittest.main()
