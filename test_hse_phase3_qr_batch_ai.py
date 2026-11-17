"""Phase 3 tests for HSE QR badges, public /qr/<token>, batch FEI inspection, Chart/Layout fixes, and HSE AI tools."""

from contextlib import ExitStack
from datetime import date, timedelta
import os
import re
import unittest
from unittest.mock import patch


class HsePhase3QrBatchAiTests(unittest.TestCase):
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

        from app import app
        cls.app = app

    def test_01_pure_python_qr_svg_generator(self):
        import hse_qr

        url_short = "http://127.0.0.1:5055/qr/abcDEF_123456-7890"
        matrix = hse_qr.build_qr_matrix(url_short)
        size = len(matrix)
        self.assertGreaterEqual(size, 21)
        self.assertEqual((size - 17) % 4, 0)
        self.assertTrue(all(len(row) == size for row in matrix))

        # Verify 7x7 finder patterns at top-left, top-right, bottom-left
        for r0, c0 in ((0, 0), (0, size - 7), (size - 7, 0)):
            for dr in range(7):
                self.assertEqual(matrix[r0 + dr][c0], 1)
                self.assertEqual(matrix[r0 + dr][c0 + 6], 1)
            for dc in range(7):
                self.assertEqual(matrix[r0][c0 + dc], 1)
                self.assertEqual(matrix[r0 + 6][c0 + dc], 1)
            # Center 3x3 black square
            for dr in (2, 3, 4):
                for dc in (2, 3, 4):
                    self.assertEqual(matrix[r0 + dr][c0 + dc], 1)

        svg = hse_qr.generate_qr_svg(url_short)
        self.assertIn("<svg", svg)
        self.assertIn('viewBox="', svg)
        self.assertIn('role="img"', svg)
        self.assertIn("<rect ", svg)

        # Longer URL + UTF-8 should promote cleanly to higher QR version without failing
        long_url = "https://dcr-hse.example.local/qr/" + ("X" * 80) + "?note=سلامة"
        svg_long = hse_qr.generate_qr_svg(long_url)
        self.assertIn("<svg", svg_long)
        self.assertGreater(len(hse_qr.build_qr_matrix(long_url)), size)

    def test_02_register_breakdown_domain_sub_counters(self):
        import hse_analytics as ha

        ref = date(2026, 11, 16)
        rows = [
            {"id": "ptw1", "project_id": "P1", "dt_code": "PTW", "issuedDate": ref.isoformat(), "status": "Open"},
            {"id": "ptw2", "project_id": "P1", "dt_code": "PTW", "issuedDate": ref.isoformat(), "status": "Closed"},
            {"id": "obs1", "project_id": "P1", "dt_code": "OBS", "issuedDate": ref.isoformat(), "status": "Open"},
            {"id": "svp1", "project_id": "P1", "dt_code": "SVP", "issuedDate": ref.isoformat(), "status": "Closed"},
            {"id": "inc1", "project_id": "P1", "dt_code": "INC", "issuedDate": ref.isoformat(), "injuryClass": "LTI", "lostDays": 3},
            {"id": "inc2", "project_id": "P1", "dt_code": "INC", "issuedDate": ref.isoformat(), "injuryClass": "MTC"},
            {"id": "inc3", "project_id": "P1", "dt_code": "INC", "issuedDate": ref.isoformat(), "injuryClass": "FAC"},
            {"id": "fei1", "project_id": "P1", "dt_code": "FEI", "expiryDate": (ref + timedelta(days=90)).isoformat()},
            {"id": "fei2", "project_id": "P1", "dt_code": "FEI", "expiryDate": (ref + timedelta(days=10)).isoformat()},
            {"id": "fei3", "project_id": "P1", "dt_code": "FEI", "expiryDate": (ref - timedelta(days=5)).isoformat()},
            {"id": "mhr1", "project_id": "P1", "dt_code": "MHR", "year": 2026, "month": 10, "hours": 15000},
            {"id": "mhr2", "project_id": "P1", "dt_code": "MHR", "year": 2026, "month": 11, "hours": 20000},
            {"id": "wkr1", "project_id": "P1", "dt_code": "WKR", "workerName": "Worker Alpha"},
            {"id": "jsa1", "project_id": "P1", "dt_code": "JSA", "status": "Approved"},
            {"id": "jsa2", "project_id": "P1", "dt_code": "JSA", "status": "Pending"},
            {"id": "jsa3", "project_id": "P1", "dt_code": "JSA", "status": "Draft"},
            {"id": "tbt1", "project_id": "P1", "dt_code": "TBT", "issuedDate": ref.isoformat()},
            {"id": "cev1", "project_id": "P1", "dt_code": "CEV", "issuedDate": ref.isoformat()},
            {"id": "evac1", "project_id": "P1", "dt_code": "EVAC", "issuedDate": ref.isoformat()},
        ]
        trends = ha.compute_hse_trends(rows, today=ref, year=2026, warn_days=30)
        by_code = {item["code"]: item for item in trends["register_breakdown"]}
        self.assertEqual(len(by_code), 11)

        # No non-workflow register should show meaningless "Open: 0 • Closed: 0"
        for code in ("INC", "FEI", "MHR", "WKR", "JSA", "TBT", "CEV", "EVAC"):
            self.assertNotIn("Open: 0 • Closed: 0", by_code[code]["subtitle"])

        self.assertEqual(by_code["PTW"]["subtitle"], "Open: 1 • Closed: 1")
        self.assertEqual(by_code["OBS"]["subtitle"], "Open: 1 • Closed: 0")
        self.assertEqual(by_code["SVP"]["subtitle"], "Open: 0 • Closed: 1")
        self.assertEqual(by_code["INC"]["subtitle"], "LTI: 1 • MTC: 1 • FAC: 1")
        self.assertEqual(by_code["FEI"]["subtitle"], "Valid: 1 • Expiring: 1 • Expired: 1")
        self.assertEqual(by_code["MHR"]["subtitle"], "Logged Months: 2 • 35,000 hrs")
        self.assertEqual(by_code["WKR"]["subtitle"], "Registered Workers: 1")
        self.assertEqual(by_code["JSA"]["subtitle"], "Approved: 1 • Pending/Draft: 2")
        self.assertEqual(by_code["TBT"]["subtitle"], "Logged Talks: 1")
        self.assertEqual(by_code["CEV"]["subtitle"], "Evaluations: 1")
        self.assertEqual(by_code["EVAC"]["subtitle"], "Logged Drills: 1")

    def test_03_dashboard_reordering_charts_and_register_controls_ui(self):
        db = self.db
        hse_pid = "SYN-HSE-P3-UI"
        db.create_project(hse_pid, "HSE Phase 3 Site", "HSE-P3", workspace_id="HSE")
        db.add_user("syn_p3_sup", "SynPass!123", role="superadmin", primary_workspace="HSE")
        client = self.app.test_client()
        client.set_cookie("dcr_token", db.create_session("syn_p3_sup", "superadmin"))

        r_hse = client.get("/?ws=HSE")
        self.assertEqual(r_hse.status_code, 200)
        html = r_hse.get_data(as_text=True)

        # Layout ordering: Header/Filters -> Hero KPIs -> 4-card strip -> Alerts ABOVE charts -> Charts -> Registers
        idx_hero = html.find('<div class="hse-hero-grid">')
        idx_strip = html.find('<div class="om-kpi-grid"', idx_hero)
        idx_alerts = html.find('id="hse-alerts-panel"')
        idx_charts = html.find('class="hse-charts-grid"')
        idx_trend_canvas = html.find('id="hseTrendChart"')
        idx_status_canvas = html.find('id="hseStatusChart"')
        idx_regs = html.find('id="hse-reg-grid"')

        self.assertTrue(all(i > 0 for i in (idx_hero, idx_strip, idx_alerts, idx_charts, idx_trend_canvas, idx_status_canvas, idx_regs)))
        self.assertLess(idx_hero, idx_strip)
        self.assertLess(idx_strip, idx_alerts)
        self.assertLess(idx_alerts, idx_charts)
        self.assertLess(idx_charts, idx_trend_canvas)
        self.assertLess(idx_charts, idx_status_canvas)
        self.assertLess(idx_trend_canvas, idx_regs)

        # Tier 2 operational strip has 4 focused cards (not 6) and preserves Completed / Closed class hook
        strip_slice = html[idx_strip:idx_alerts]
        self.assertEqual(strip_slice.count('class="om-kpi-card"'), 4)
        self.assertNotIn("hse-card-sum-sites", strip_slice)
        self.assertIsNotNone(
            re.search(r'Completed / Closed</span>\s*<span class="om-kpi-val success"[^>]*>\d+</span>', html)
        )

        # Chart fixes: min-width: 0 CSS, /api/hse/trends binding, Canvas2D fallback + MutationObserver
        self.assertIn(".hse-charts-grid > .hse-panel", html)
        self.assertIn("min-width: 0", html)
        self.assertIn("/api/hse/trends?", html)
        self.assertIn("_drawHseCanvasChart", html)
        self.assertIn("MutationObserver", html)
        self.assertIn("DCR HSE Safety Assistant", html)

        # Register page includes QR button support + FEI Batch Inspect modal & toolbar button
        r_reg = client.get(f"/app?p={hse_pid}")
        self.assertEqual(r_reg.status_code, 200)
        reg_html = r_reg.get_data(as_text=True)
        self.assertIn('id="btn-batch-inspect"', reg_html)
        self.assertIn('id="hse-qr-modal"', reg_html)
        self.assertIn('id="fei-batch-inspect-modal"', reg_html)
        self.assertIn("openQrBadgeModal", reg_html)
        self.assertIn("isHseQrTab", reg_html)
        self.assertIn("isHseFeiTab", reg_html)
        self.assertIn("syncFeiBatchInspectBtn", reg_html)
        self.assertIn("openFeiBatchInspectModal", reg_html)

        # ENG and OM remain isolated from Phase 3 HSE UI hooks
        r_eng = client.get("/?ws=ENG")
        self.assertEqual(r_eng.status_code, 200)
        self.assertNotIn("hseTrendChart", r_eng.get_data(as_text=True))
        r_om = client.get("/?ws=OM")
        self.assertEqual(r_om.status_code, 200)
        self.assertNotIn("hseTrendChart", r_om.get_data(as_text=True))

    def test_04_qr_generation_rbac_and_public_read_only_verify_card(self):
        db = self.db
        ref = date.today()
        hse_a = "SYN-HSE-P3-QR-A"
        hse_b = "SYN-HSE-P3-QR-B"
        eng_a = "SYN-ENG-P3-QR-A"

        db.create_project(hse_a, "HSE QR Plant A", "QR-A", workspace_id="HSE")
        db.create_project(hse_b, "HSE QR Plant B", "QR-B", workspace_id="HSE")
        db.create_project(eng_a, "ENG QR Site A", "ENG-QR", workspace_id="ENG")

        fei_id = "p3-fei-valid-1"
        fei_exp_id = "p3-fei-exp-1"
        ptw_ovd_id = "p3-ptw-ovd-1"
        obs_id = "p3-obs-1"
        eng_id = "p3-eng-1"

        db.save_record(
            hse_a,
            "FEI",
            fei_id,
            {
                "docNo": "FEI-QR-001",
                "title": "CO2 Extinguisher Bay 3",
                "legacySubtype": "CO2 6kg",
                "location": "Substation B",
                "lastInspection": (ref - timedelta(days=10)).isoformat(),
                "expiryDate": (ref + timedelta(days=120)).isoformat(),
                "inspector": "SecretInspectorName",
                "responsible": "SecretResponsiblePerson",
                "phone": "01099988877",
                "nationalId": "29901011234567",
                "fileLocation": "https://drive.google.com/file/d/secret-drive-id/view",
                "attachments": [{"url": "https://drive.google.com/file/d/secret-drive-id/view", "name": "cert.pdf"}],
            },
        )
        db.save_record(
            hse_a,
            "FEI",
            fei_exp_id,
            {
                "docNo": "FEI-QR-EXP",
                "legacySubtype": "Foam 9L",
                "location": "Fuel Yard",
                "lastInspection": (ref - timedelta(days=380)).isoformat(),
                "expiryDate": (ref - timedelta(days=5)).isoformat(),
            },
        )
        db.save_record(
            hse_a,
            "PTW",
            ptw_ovd_id,
            {
                "docNo": "PTW-QR-007",
                "legacySubtype": "Hot Work",
                "location": "Boiler Deck 2",
                "issuedDate": (ref - timedelta(days=4)).isoformat(),
                "validUntil": (ref - timedelta(days=1)).isoformat(),
                "status": "Open",
                "permitApplicant": "SecretPermitApplicant",
                "fileLocation": "https://drive.google.com/file/d/ptw-secret-file/view",
            },
        )
        db.save_record(hse_a, "OBS", obs_id, {"docNo": "OBS-QR-1", "status": "Open"})
        db.save_record(eng_a, "LTR", eng_id, {"docNo": "LTR-QR-1"})

        db.add_user("syn_p3_qr_a", "SynPass!123", role="editor", primary_workspace="HSE")
        db.assign_project("syn_p3_qr_a", hse_a)
        db.add_user("syn_p3_qr_b", "SynPass!123", role="editor", primary_workspace="HSE")
        db.assign_project("syn_p3_qr_b", hse_b)
        db.add_user("syn_p3_qr_ptw_only", "SynPass!123", role="viewer", primary_workspace="HSE")
        db.assign_project("syn_p3_qr_ptw_only", hse_a)
        db.set_user_project_doc_types("syn_p3_qr_ptw_only", hse_a, ["PTW"])

        anon = self.app.test_client()
        self.assertEqual(anon.get(f"/api/hse/qr/{fei_id}").status_code, 403)

        c_b = self.app.test_client()
        c_b.set_cookie("dcr_token", db.create_session("syn_p3_qr_b", "editor"))
        self.assertEqual(c_b.get(f"/api/hse/qr/{fei_id}").status_code, 403)

        c_ptw = self.app.test_client()
        c_ptw.set_cookie("dcr_token", db.create_session("syn_p3_qr_ptw_only", "viewer"))
        self.assertEqual(c_ptw.get(f"/api/hse/qr/{fei_id}").status_code, 403)
        self.assertEqual(c_ptw.get(f"/api/hse/qr/{ptw_ovd_id}").status_code, 200)

        c_a = self.app.test_client()
        c_a.set_cookie("dcr_token", db.create_session("syn_p3_qr_a", "editor"))
        self.assertEqual(c_a.get(f"/api/hse/qr/{obs_id}").status_code, 400)
        self.assertEqual(c_a.get(f"/api/hse/qr/{eng_id}").status_code, 403)  # Not assigned to eng_a
        self.assertEqual(c_a.get("/api/hse/qr/does-not-exist").status_code, 404)

        # Superadmin calling ENG doc_type returns 400 (not HSE FEI/PTW)
        db.add_user("syn_p3_qr_sup", "SynPass!123", role="superadmin", primary_workspace="HSE")
        c_sup = self.app.test_client()
        c_sup.set_cookie("dcr_token", db.create_session("syn_p3_qr_sup", "superadmin"))
        self.assertEqual(c_sup.get(f"/api/hse/qr/{eng_id}").status_code, 400)

        # Fetch QR badge twice -> immutable qr_token + valid SVG
        res1 = c_a.get(f"/api/hse/qr/{fei_id}")
        self.assertEqual(res1.status_code, 200)
        p1 = res1.get_json()
        self.assertTrue(p1["ok"])
        self.assertEqual(p1["dt_code"], "FEI")
        self.assertEqual(p1["doc_no"], "FEI-QR-001")
        self.assertEqual(p1["location"], "Substation B")
        self.assertIn("<svg", p1["svg"])
        token1 = p1["qr_token"]
        self.assertTrue(len(token1) >= 16)

        res2 = c_a.get(f"/api/hse/qr/{fei_id}")
        self.assertEqual(res2.status_code, 200)
        self.assertEqual(res2.get_json()["qr_token"], token1)

        res_svg = c_a.get(f"/api/hse/qr/{fei_id}?format=svg")
        self.assertEqual(res_svg.status_code, 200)
        self.assertIn("image/svg+xml", res_svg.content_type)
        self.assertIn("<svg", res_svg.get_data(as_text=True))

        # Public unauthenticated /qr/<token> verification card: read-only, 8 safe fields only, NO PII / Drive URLs
        pub = self.app.test_client()
        card_resp = pub.get(f"/qr/{token1}")
        self.assertEqual(card_resp.status_code, 200)
        self.assertEqual(card_resp.headers.get("Cache-Control"), "no-store")
        self.assertEqual(card_resp.headers.get("X-Robots-Tag"), "noindex, nofollow")
        card_html = card_resp.get_data(as_text=True)

        for expected in (
            "QR-A",
            "HSE QR Plant A",
            "Fire &amp; Equipment Inspection (FEI)",
            "FEI-QR-001",
            "CO2 6kg",
            "Substation B",
            "Valid",
            (ref + timedelta(days=120)).isoformat(),
            (ref - timedelta(days=10)).isoformat(),
        ):
            self.assertIn(expected, card_html)

        for forbidden in (
            "SecretInspectorName",
            "SecretResponsiblePerson",
            "01099988877",
            "29901011234567",
            "drive.google.com",
            "secret-drive-id",
            "cert.pdf",
        ):
            self.assertNotIn(forbidden, card_html)

        # Expired FEI badge + Overdue PTW badge status calculations
        exp_token = c_a.get(f"/api/hse/qr/{fei_exp_id}").get_json()["qr_token"]
        exp_html = pub.get(f"/qr/{exp_token}").get_data(as_text=True)
        self.assertIn("Expired", exp_html)

        ptw_token = c_a.get(f"/api/hse/qr/{ptw_ovd_id}").get_json()["qr_token"]
        ptw_html = pub.get(f"/qr/{ptw_token}").get_data(as_text=True)
        self.assertIn("Overdue", ptw_html)
        self.assertIn("Permit to Work (PTW)", ptw_html)
        self.assertNotIn("SecretPermitApplicant", ptw_html)
        self.assertNotIn("ptw-secret-file", ptw_html)

        # Invalid/unknown tokens return 404
        self.assertEqual(pub.get("/qr/unknown_token_123456").status_code, 404)
        self.assertEqual(pub.get("/qr/bad$chars!").status_code, 404)

    def test_05_fei_batch_inspect_rbac_validation_and_audit(self):
        db = self.db
        hse_a = "SYN-HSE-P3-BAT-A"
        hse_b = "SYN-HSE-P3-BAT-B"
        db.create_project(hse_a, "HSE Batch Plant A", "BAT-A", workspace_id="HSE")
        db.create_project(hse_b, "HSE Batch Plant B", "BAT-B", workspace_id="HSE")

        fei1, fei2, fei_b, obs1 = "p3-bat-fei1", "p3-bat-fei2", "p3-bat-feib", "p3-bat-obs1"
        db.save_record(
            hse_a,
            "FEI",
            fei1,
            {"docNo": "FEI-B-01", "location": "Gate 1", "qr_token": "fixed_token_fei_01", "expiryDate": "2026-01-01"},
        )
        db.save_record(
            hse_a,
            "FEI",
            fei2,
            {"docNo": "FEI-B-02", "location": "Gate 2", "expiryDate": "2026-01-02"},
        )
        db.save_record(hse_b, "FEI", fei_b, {"docNo": "FEI-B-99", "location": "Remote Site"})
        db.save_record(hse_a, "OBS", obs1, {"docNo": "OBS-B-01", "status": "Open"})

        db.add_user("syn_p3_bat_view", "SynPass!123", role="viewer", primary_workspace="HSE")
        db.assign_project("syn_p3_bat_view", hse_a)
        db.add_user("syn_p3_bat_ed_a", "SynPass!123", role="editor", primary_workspace="HSE")
        db.assign_project("syn_p3_bat_ed_a", hse_a)
        db.add_user("syn_p3_bat_ed_ptw", "SynPass!123", role="editor", primary_workspace="HSE")
        db.assign_project("syn_p3_bat_ed_ptw", hse_a)
        db.set_user_project_doc_types("syn_p3_bat_ed_ptw", hse_a, ["PTW"])

        payload = {
            "record_ids": [fei1, fei2],
            "last_inspection": "2026-11-16",
            "expiry_date": "2027-05-16",
            "status": "Valid",
            "notes": "Pressure gauge green, tamper seal intact",
        }

        anon = self.app.test_client()
        self.assertEqual(anon.post("/api/hse/fei/batch-inspect", json=payload).status_code, 403)

        c_view = self.app.test_client()
        c_view.set_cookie("dcr_token", db.create_session("syn_p3_bat_view", "viewer"))
        self.assertEqual(c_view.post("/api/hse/fei/batch-inspect", json=payload).status_code, 403)

        c_ptw_ed = self.app.test_client()
        c_ptw_ed.set_cookie("dcr_token", db.create_session("syn_p3_bat_ed_ptw", "editor"))
        self.assertEqual(c_ptw_ed.post("/api/hse/fei/batch-inspect", json=payload).status_code, 403)

        c_ed = self.app.test_client()
        c_ed.set_cookie("dcr_token", db.create_session("syn_p3_bat_ed_a", "editor"))
        # Cross-project FEI rejected before any mutation
        self.assertEqual(
            c_ed.post("/api/hse/fei/batch-inspect", json={**payload, "record_ids": [fei1, fei_b]}).status_code,
            403,
        )
        # Non-FEI record rejected atomically
        self.assertEqual(
            c_ed.post("/api/hse/fei/batch-inspect", json={**payload, "record_ids": [fei1, obs1]}).status_code,
            400,
        )
        # Missing / invalid date rejected
        self.assertEqual(
            c_ed.post("/api/hse/fei/batch-inspect", json={**payload, "record_ids": [fei1, "missing-fei"]}).status_code,
            404,
        )
        self.assertEqual(
            c_ed.post("/api/hse/fei/batch-inspect", json={**payload, "expiry_date": "not-a-valid-date"}).status_code,
            400,
        )
        # Ensure failed calls did not mutate fei1
        self.assertEqual(db.get_record_by_id(fei1)["expiryDate"], "2026-01-01")

        # Valid batch inspection
        ok_res = c_ed.post("/api/hse/fei/batch-inspect", json=payload)
        self.assertEqual(ok_res.status_code, 200)
        ok_body = ok_res.get_json()
        self.assertTrue(ok_body["ok"])
        self.assertEqual(ok_body["updated"], 2)

        for rid in (fei1, fei2):
            rec = db.get_record_by_id(rid)
            self.assertEqual(rec["lastInspection"], "2026-11-16")
            self.assertEqual(rec["last_inspection"], "2026-11-16")
            self.assertEqual(rec["expiryDate"], "2027-05-16")
            self.assertEqual(rec["expiry_date"], "2027-05-16")
            self.assertEqual(rec["status"], "Valid")
            self.assertEqual(rec["extinguisherStatus"], "Valid")
            self.assertEqual(rec["notes"], "Pressure gauge green, tamper seal intact")
        self.assertEqual(db.get_record_by_id(fei1)["qr_token"], "fixed_token_fei_01")

        logs = db.q(
            "SELECT record_id, action, doc_no, username FROM audit_log WHERE action = 'BATCH_INSPECT' AND project_id = %s",
            (hse_a,),
        )
        self.assertEqual({r["record_id"] for r in logs}, {fei1, fei2})
        self.assertEqual({r["username"] for r in logs}, {"syn_p3_bat_ed_a"})

    def test_06_hse_ai_tools_rbac_and_safety_assistant_endpoint(self):
        from blueprints import ai as ai_bp
        from blueprints import ai_tools

        db = self.db
        ref = date.today()
        hse_a = "SYN-HSE-P3-AI-A"
        hse_b = "SYN-HSE-P3-AI-B"
        eng_a = "SYN-ENG-P3-AI-A"

        db.create_project(hse_a, "HSE AI Refinery A", "AIA", workspace_id="HSE")
        db.create_project(hse_b, "HSE AI Terminal B", "AIB", workspace_id="HSE")
        db.create_project(eng_a, "ENG AI Site A", "ENA", workspace_id="ENG")

        db.save_record(hse_a, "MHR", "p3-ai-mhr-a", {"docNo": "MHR-AIA-1", "year": ref.year, "month": ref.month, "hours": 200000})
        db.save_record(
            hse_a,
            "INC",
            "p3-ai-inc-a",
            {"docNo": "INC-AIA-1", "title": "Scaffold slip", "issuedDate": (ref - timedelta(days=15)).isoformat(), "injuryClass": "LTI", "lostDays": 4},
        )
        db.save_record(
            hse_a,
            "PTW",
            "p3-ai-ptw-a",
            {"docNo": "PTW-AIA-10", "title": "Hot Work Tank 2", "issuedDate": ref.isoformat(), "validUntil": (ref + timedelta(days=5)).isoformat(), "status": "Open", "location": "Tank Farm"},
        )
        db.save_record(
            hse_a,
            "FEI",
            "p3-ai-fei-a",
            {"docNo": "FEI-AIA-20", "title": "DCP 9kg", "location": "Pump Room", "expiryDate": (ref - timedelta(days=3)).isoformat()},
        )

        # Plant B data must NEVER leak to user scoped only to HSE-A
        db.save_record(hse_b, "MHR", "p3-ai-mhr-b", {"docNo": "MHR-AIB-1", "year": ref.year, "month": ref.month, "hours": 500000})
        db.save_record(
            hse_b,
            "PTW",
            "p3-ai-ptw-b",
            {"docNo": "PTW-AIB-SECRET", "title": "Secret Site B Permit", "issuedDate": ref.isoformat(), "validUntil": (ref + timedelta(days=2)).isoformat(), "status": "Open"},
        )

        # Direct ai_tools checks with RBAC boundaries
        kpi_a = ai_tools.get_hse_kpis([hse_a, eng_a], year=ref.year)
        self.assertEqual([p["id"] for p in kpi_a["projects"]], [hse_a])
        self.assertEqual(kpi_a["man_hours"], 200000.0)
        self.assertEqual(kpi_a["ltif"], 5.0)
        self.assertIsNone(kpi_a.get("secret_b"))

        # Doc-type restricted scope (PTW only) sees zero LTI/man-hours and cannot search FEI
        kpi_ptw = ai_tools.get_hse_kpis([hse_a], doc_type_scope={hse_a: ["PTW"]}, year=ref.year)
        self.assertEqual(kpi_ptw["man_hours"], 0.0)
        self.assertEqual(kpi_ptw["incidents"]["lti"], 0)

        srch_a = ai_tools.search_hse_records([hse_a], doc_type="PTW", keyword="Tank")
        self.assertEqual(srch_a["count"], 1)
        self.assertEqual(srch_a["records"][0]["doc_no"], "PTW-AIA-10")
        self.assertEqual(ai_tools.search_hse_records([hse_a], doc_type="FEI", doc_type_scope={hse_a: ["PTW"]})["count"], 0)

        exp_a = ai_tools.get_expiring_permits_and_equipment([hse_a], warn_days=30)
        self.assertGreaterEqual(exp_a["total_alerts"], 2)
        exp_doc_nos = {item["doc_no"] for item in exp_a["alerts"]}
        self.assertIn("PTW-AIA-10", exp_doc_nos)
        self.assertIn("FEI-AIA-20", exp_doc_nos)
        self.assertNotIn("PTW-AIB-SECRET", exp_doc_nos)

        # HTTP POST /api/ai/query (offline deterministic fallback + RBAC)
        db.add_user("syn_p3_ai_user", "SynPass!123", role="viewer", primary_workspace="HSE")
        db.assign_project("syn_p3_ai_user", hse_a)
        c_ai = self.app.test_client()
        c_ai.set_cookie("dcr_token", db.create_session("syn_p3_ai_user", "viewer"))

        with patch.dict(os.environ, {"GEMINI_API_KEY": "synthetic-test-key"}):
            from types import SimpleNamespace

            class _FakePartType(SimpleNamespace):
                @classmethod
                def from_text(cls, text=""):
                    return cls(text=text)

                @classmethod
                def from_function_response(cls, name="", response=None):
                    return cls(name=name, response=response)

            # 1) Offline / Gemini-unavailable path triggers _execute_hse_deterministic_fallback with strict RBAC
            class OfflineClient:
                def __init__(self, api_key=None):
                    self.models = self

                def generate_content(self, **kwargs):
                    raise RuntimeError("offline synthetic test")

            with patch.object(ai_bp.types, "Part", _FakePartType), patch.object(ai_bp.genai, "Client", OfflineClient):
                r_kpi = c_ai.post("/api/ai/query", json={"prompt": "What are our LTIF and man-hours?", "workspace_id": "HSE"})
                self.assertEqual(r_kpi.status_code, 200)
                b_kpi = r_kpi.get_json()
                self.assertIn("200,000", b_kpi["reply"])
                self.assertIn("5.0", b_kpi["reply"])
                self.assertNotIn("500,000", b_kpi["reply"])

                r_exp = c_ai.post("/api/ai/query", json={"prompt": "Which permits or extinguishers are expiring or overdue?", "workspace_id": "HSE"})
                self.assertEqual(r_exp.status_code, 200)
                b_exp = r_exp.get_json()
                self.assertIn("FEI-AIA-20", b_exp["reply"])
                self.assertIn("PTW-AIA-10", b_exp["reply"])
                self.assertNotIn("PTW-AIB-SECRET", b_exp["reply"])

                r_search = c_ai.post("/api/ai/query", json={"prompt": "Show permits for Tank 2", "workspace_id": "HSE"})
                self.assertEqual(r_search.status_code, 200)
                self.assertIn("PTW-AIA-10", r_search.get_json()["reply"])
                self.assertNotIn("PTW-AIB-SECRET", r_search.get_json()["reply"])

                r_std = c_ai.post("/api/ai/query", json={"prompt": "What are the OSHA hot work requirements?", "workspace_id": "HSE"})
                self.assertEqual(r_std.status_code, 200)
                self.assertIn("OSHA 1910.252", r_std.get_json()["reply"])
                self.assertIn("Decision 211", r_std.get_json()["reply"])

            # 2) Mocked Gemini call confirms HSE compliance system prompt (OSHA + NFPA + Decision 211/134 + NCR) + RBAC tool execution
            captured_calls = []

            class FakePart:
                def __init__(self, text=None, function_call=None):
                    self.text = text
                    self.function_call = function_call

            class FakeFnCall:
                def __init__(self, name, args):
                    self.name = name
                    self.args = args

            class FakeContent:
                def __init__(self, parts):
                    self.parts = parts

            class FakeCandidate:
                def __init__(self, parts):
                    self.content = FakeContent(parts)

            class FakeResp:
                def __init__(self, parts):
                    self.candidates = [FakeCandidate(parts)]
                    self.function_calls = [p.function_call for p in parts if p.function_call is not None]
                    self.text = "".join(p.text for p in parts if p.text)

            class FakeGeminiClient:
                def __init__(self, api_key=None):
                    self.models = self

                def generate_content(self, model=None, contents=None, config=None):
                    captured_calls.append({"config": config, "contents": contents})
                    if len(captured_calls) == 1:
                        return FakeResp([FakePart(function_call=FakeFnCall("get_hse_kpis", {}))])
                    return FakeResp([FakePart(text="LTIF is 5.0 per OSHA / Decision 211.")])

            with patch.object(ai_bp.genai, "Client", FakeGeminiClient):
                r_gem = c_ai.post(
                    "/api/ai/query",
                    json={"prompt": "Draft safety summary and check NCR standards", "workspace_id": "HSE"},
                )
                self.assertEqual(r_gem.status_code, 200)
                self.assertIn("LTIF is 5.0", r_gem.get_json()["reply"])
                self.assertEqual(len(captured_calls), 2)
                cfg = captured_calls[0]["config"]
                sys_prompt = getattr(cfg, "system_instruction", "") or ""
                self.assertIn("OSHA 1926.451", sys_prompt)
                self.assertIn("Decision 211", sys_prompt)
                self.assertIn("Probable Root Cause", sys_prompt)
                self.assertIn("AIA", sys_prompt)
                self.assertNotIn("AIB", sys_prompt)
                decl_names = {d.name for t in (cfg.tools or []) for d in (t.function_declarations or [])}
                self.assertEqual(decl_names, {"get_hse_kpis", "search_hse_records", "get_expiring_permits_and_equipment"})

        # Anonymous request to /api/ai/query still blocked
        anon = self.app.test_client()
        self.assertEqual(anon.post("/api/ai/query", json={"prompt": "LTIF?", "workspace_id": "HSE"}).status_code, 403)


if __name__ == "__main__":
    unittest.main()
