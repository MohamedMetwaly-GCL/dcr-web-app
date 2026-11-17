"""Phase 2 tests for pure HSE analytics engine, RBAC-scoped HSE APIs, and Executive Command UI."""

from contextlib import ExitStack
from datetime import date, timedelta
import os
import re
import unittest


class HsePhase2DashboardTests(unittest.TestCase):
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

    def test_01_pure_rates_and_zero_manhours_safety(self):
        import hse_analytics as ha

        # Safe zero-division handling
        zero = ha.compute_frequency_rates(lti_count=2, mtc_count=3, lost_days=15, man_hours=0)
        self.assertFalse(zero["has_hours"])
        self.assertEqual(zero["man_hours"], 0.0)
        self.assertIsNone(zero["ltif"])
        self.assertIsNone(zero["trir"])
        self.assertIsNone(zero["severity_rate"])

        neg = ha.compute_frequency_rates(lti_count=1, mtc_count=0, lost_days=5, man_hours=-100)
        self.assertFalse(neg["has_hours"])
        self.assertIsNone(neg["ltif"])

        # Per 1,000,000 worked hours formulas
        rates = ha.compute_frequency_rates(lti_count=2, mtc_count=3, lost_days=45, man_hours=500_000)
        self.assertTrue(rates["has_hours"])
        self.assertEqual(rates["man_hours"], 500000.0)
        self.assertEqual(rates["ltif"], 4.0)
        self.assertEqual(rates["trir"], 10.0)
        self.assertEqual(rates["severity_rate"], 90.0)

        # Mixed DCR camelCase + snake_case record normalization
        ref = date(2026, 11, 16)
        rows = [
            {"id": "m1", "project_id": "P1", "dt_code": "MHR", "year": 2026, "month": 10, "hours": 250000},
            {"id": "m2", "project_id": "P1", "dt_code": "MHR", "data": {"year": 2026, "month": 11, "man_hours": 250000}},
            {"id": "i1", "project_id": "P1", "dt_id": "INC", "issuedDate": "2026-10-15", "injuryClass": "LTI", "lostDays": 10},
            {"id": "i2", "project_id": "P1", "dt_id": "INC", "doc_date": "2026-11-01", "injury_type": "MTC", "lost_days": 5},
            {"id": "i3", "project_id": "P1", "dt_id": "INC", "issuedDate": "2026-11-05", "injuryClass": "FAC"},
        ]
        kpis = ha.compute_hse_kpis(rows, today=ref, year=2026)
        self.assertEqual(kpis["man_hours"], 500000.0)
        self.assertEqual(kpis["incidents"]["fac"], 1)
        self.assertEqual(kpis["incidents"]["mtc"], 1)
        self.assertEqual(kpis["incidents"]["lti"], 1)
        self.assertEqual(kpis["incidents"]["lost_days"], 10)  # Only LTI contributes lost days
        self.assertEqual(kpis["rates"]["ltif"], 2.0)
        self.assertEqual(kpis["rates"]["trir"], 4.0)
        self.assertEqual(kpis["rates"]["severity_rate"], 20.0)

    def test_02_streaks_observations_ptw_and_fei_alerts(self):
        import hse_analytics as ha

        ref = date(2026, 11, 16)
        # Empty records -> safe None streaks
        empty = ha.compute_streaks([], today=ref)
        self.assertIsNone(empty["lti_free"]["days"])
        self.assertEqual(empty["lti_free"]["basis"], "none")
        self.assertIsNone(empty["any_injury_free"]["days"])

        # Records starting 80d ago, LTI 40d ago, MTC 12d ago, FAC 5d ago, Near Miss 1d ago (non-injury!)
        rows = [
            {"id": "p0", "dt_code": "PTW", "issuedDate": (ref - timedelta(days=80)).isoformat(), "validUntil": (ref + timedelta(days=90)).isoformat(), "status": "Open"},
            {"id": "p_due", "dt_code": "PTW", "issuedDate": (ref - timedelta(days=10)).isoformat(), "validUntil": (ref + timedelta(days=10)).isoformat(), "status": "Open"},
            {"id": "p_ovd", "dt_code": "PTW", "issuedDate": (ref - timedelta(days=20)).isoformat(), "validUntil": (ref - timedelta(days=3)).isoformat(), "status": "Open"},
            {"id": "p_cls", "dt_code": "PTW", "issuedDate": (ref - timedelta(days=25)).isoformat(), "validUntil": (ref - timedelta(days=5)).isoformat(), "status": "Closed"},
            {"id": "o_stale", "dt_code": "OBS", "issuedDate": (ref - timedelta(days=31)).isoformat(), "status": "Open"},
            {"id": "o_30", "dt_code": "OBS", "issuedDate": (ref - timedelta(days=30)).isoformat(), "status": "Open"},
            {"id": "o_closed", "dt_code": "OBS", "issuedDate": (ref - timedelta(days=45)).isoformat(), "status": "Closed"},
            {"id": "i_lti", "dt_code": "INC", "issuedDate": (ref - timedelta(days=40)).isoformat(), "injuryClass": "LTI", "lostDays": 7},
            {"id": "i_mtc", "dt_code": "INC", "issuedDate": (ref - timedelta(days=12)).isoformat(), "injuryClass": "MTC"},
            {"id": "i_fac", "dt_code": "INC", "issuedDate": (ref - timedelta(days=5)).isoformat(), "injuryClass": "FAC"},
            {"id": "i_nm", "dt_code": "INC", "issuedDate": (ref - timedelta(days=1)).isoformat(), "injuryClass": "Near Miss"},
            {"id": "i_blank", "dt_code": "INC", "issuedDate": ref.isoformat(), "injuryClass": ""},
            {"id": "f_exp", "dt_code": "FEI", "expiryDate": (ref - timedelta(days=2)).isoformat()},
            {"id": "f_20", "dt_code": "FEI", "expiryDate": (ref + timedelta(days=20)).isoformat()},
            {"id": "f_45", "dt_code": "FEI", "expiryDate": (ref + timedelta(days=45)).isoformat()},
        ]

        kpis30 = ha.compute_hse_kpis(rows, today=ref, warn_days=30)
        self.assertEqual(kpis30["streaks"]["lti_free"]["days"], 40)
        self.assertEqual(kpis30["streaks"]["lti_free"]["best_days"], 40)
        self.assertEqual(kpis30["streaks"]["lti_free"]["basis"], "event")
        self.assertEqual(kpis30["streaks"]["any_injury_free"]["days"], 5)
        self.assertEqual(kpis30["streaks"]["any_injury_free"]["basis"], "event")

        self.assertEqual(kpis30["observations"]["open"], 2)
        self.assertEqual(kpis30["observations"]["closed"], 1)
        self.assertEqual(kpis30["observations"]["stale"], 1)

        self.assertEqual(kpis30["permits"]["open"], 3)
        self.assertEqual(kpis30["permits"]["closed"], 1)
        self.assertEqual(kpis30["permits"]["overdue"], 1)
        self.assertEqual(kpis30["permits"]["expiring"], 1)
        self.assertEqual(kpis30["permits"]["valid_active"], 1)

        self.assertEqual(kpis30["equipment"]["expired"], 1)
        self.assertEqual(kpis30["equipment"]["expiring"], 1)
        self.assertEqual(kpis30["equipment"]["valid"], 1)

        # Custom warn_days=60 catches the +45d extinguisher too
        kpis60 = ha.compute_hse_kpis(rows, today=ref, warn_days=60)
        self.assertEqual(kpis60["equipment"]["expiring"], 2)
        alerts60 = ha.compute_hse_alerts(rows, today=ref, warn_days=60)
        self.assertEqual(alerts60["counts"]["ptw_overdue"], 1)
        self.assertEqual(alerts60["counts"]["ptw_expiring"], 1)
        self.assertEqual(alerts60["counts"]["fei_expired"], 1)
        self.assertEqual(alerts60["counts"]["fei_expiring"], 2)
        self.assertEqual(alerts60["counts"]["obs_stale"], 1)
        self.assertEqual(alerts60["counts"]["total"], 6)

    def test_03_hse_api_rbac_project_and_doc_type_isolation(self):
        db = self.db
        ref = date.today()
        hse_a = "SYN-HSE-D2-A"
        hse_b = "SYN-HSE-D2-B"
        eng_p = "SYN-ENG-D2-A"
        om_p = "SYN-OM-D2-A"

        db.create_project(hse_a, "HSE Site Alpha", "HSE-A", workspace_id="HSE")
        db.create_project(hse_b, "HSE Site Beta", "HSE-B", workspace_id="HSE")
        db.create_project(eng_p, "ENG Site Alpha", "ENG-A", workspace_id="ENG")
        db.create_project(om_p, "OM Site Alpha", "OM-A", workspace_id="OM")

        # Seed Site A: 100k hrs, 1 LTI (12 lost days), 1 Open stale OBS, 1 Expiring PTW, 1 Expired FEI
        db.save_record(hse_a, "MHR", "d2-a-mhr", {"docNo": "MHR-A-1", "year": ref.year, "month": ref.month, "hours": 100000})
        db.save_record(hse_a, "INC", "d2-a-inc", {
            "docNo": "INC-A-1", "title": "Slip LTI", "issuedDate": (ref - timedelta(days=10)).isoformat(),
            "injuryClass": "LTI", "lostDays": 12,
        })
        db.save_record(hse_a, "OBS", "d2-a-obs", {
            "docNo": "OBS-A-1", "title": "Open guard", "issuedDate": (ref - timedelta(days=40)).isoformat(), "status": "Open",
        })
        db.save_record(hse_a, "PTW", "d2-a-ptw", {
            "docNo": "PTW-A-1", "title": "Hot work A", "issuedDate": ref.isoformat(),
            "validUntil": (ref + timedelta(days=5)).isoformat(), "status": "Open",
        })
        db.save_record(hse_a, "FEI", "d2-a-fei", {
            "docNo": "FEI-A-1", "title": "Ext A-01", "expiryDate": (ref - timedelta(days=4)).isoformat(),
        })

        # Seed Site B: 400k hrs, 1 MTC (must stay isolated from Site-A-only users!)
        db.save_record(hse_b, "MHR", "d2-b-mhr", {"docNo": "MHR-B-1", "year": ref.year, "month": ref.month, "hours": 400000})
        db.save_record(hse_b, "INC", "d2-b-inc", {
            "docNo": "INC-B-1", "title": "Cut MTC", "issuedDate": (ref - timedelta(days=2)).isoformat(),
            "injuryClass": "MTC",
        })
        db.save_record(hse_b, "OBS", "d2-b-obs", {
            "docNo": "OBS-B-1", "title": "Fixed cable", "issuedDate": ref.isoformat(), "status": "Closed",
        })

        db.add_user("syn_d2_hse_mgr", "SynPass!123", role="admin", primary_workspace="HSE")
        db.add_user("syn_d2_hse_a", "SynPass!123", role="editor", primary_workspace="HSE")
        db.assign_project("syn_d2_hse_a", hse_a)
        db.add_user("syn_d2_hse_obs", "SynPass!123", role="viewer", primary_workspace="HSE")
        db.assign_project("syn_d2_hse_obs", hse_a)
        db.set_user_project_doc_types("syn_d2_hse_obs", hse_a, ["OBS", "PTW"])
        db.add_user("syn_d2_eng_mgr", "SynPass!123", role="admin", primary_workspace="ENG")

        # Anonymous requests blocked
        anon = self.app.test_client()
        for path in ("/api/hse/kpis", "/api/hse/trends", "/api/hse/alerts", "/api/hse/overview"):
            self.assertEqual(anon.get(path).status_code, 403)

        # HSE Admin sees both HSE-A and HSE-B (restricted to HSE only)
        c_mgr = self.app.test_client()
        c_mgr.set_cookie("dcr_token", db.create_session("syn_d2_hse_mgr", "admin"))
        ov_all = c_mgr.get("/api/hse/overview").get_json()
        self.assertTrue(ov_all["ok"])
        self.assertGreaterEqual(ov_all["kpis"]["man_hours"], 500000.0)
        # Filter to HSE-A only
        k_a = c_mgr.get(f"/api/hse/kpis?pid={hse_a}").get_json()
        self.assertEqual(k_a["man_hours"], 100000.0)
        self.assertEqual(k_a["incidents"]["lti"], 1)
        self.assertEqual(k_a["incidents"]["mtc"], 0)
        self.assertEqual(k_a["rates"]["ltif"], 10.0)
        self.assertEqual(k_a["rates"]["severity_rate"], 120.0)
        # Cross-workspace queries rejected
        self.assertEqual(c_mgr.get(f"/api/hse/kpis?pid={eng_p}").status_code, 403)
        self.assertEqual(c_mgr.get(f"/api/hse/alerts?pid={om_p}").status_code, 403)

        # HSE Editor assigned only HSE-A cannot see HSE-B metrics
        c_a = self.app.test_client()
        c_a.set_cookie("dcr_token", db.create_session("syn_d2_hse_a", "editor"))
        r_a = c_a.get("/api/hse/kpis").get_json()
        self.assertEqual([p["id"] for p in r_a["projects"]], [hse_a])
        self.assertEqual(r_a["man_hours"], 100000.0)
        self.assertEqual(r_a["incidents"]["mtc"], 0)
        self.assertEqual(c_a.get(f"/api/hse/kpis?pid={hse_b}").status_code, 403)
        self.assertEqual(c_a.get(f"/api/hse/trends?pid={hse_b}").status_code, 403)

        # Doc-type restricted viewer (OBS + PTW only) gets zero leaked INC, MHR, or FEI data
        c_obs = self.app.test_client()
        c_obs.set_cookie("dcr_token", db.create_session("syn_d2_hse_obs", "viewer"))
        ov_obs = c_obs.get("/api/hse/overview").get_json()
        self.assertEqual(ov_obs["kpis"]["man_hours"], 0.0)
        self.assertIsNone(ov_obs["kpis"]["rates"]["ltif"])
        self.assertEqual(ov_obs["kpis"]["incidents"]["lti"], 0)
        self.assertEqual(ov_obs["kpis"]["equipment"]["total"], 0)
        self.assertEqual(ov_obs["kpis"]["observations"]["open"], 1)
        self.assertEqual(ov_obs["kpis"]["permits"]["expiring"], 1)
        alert_codes = {a["dt_code"] for a in ov_obs["alerts"]}
        self.assertEqual(alert_codes, {"OBS", "PTW"})

        # ENG Admin gets no HSE records and 403 when targeting HSE project
        c_eng = self.app.test_client()
        c_eng.set_cookie("dcr_token", db.create_session("syn_d2_eng_mgr", "admin"))
        eng_kpis = c_eng.get("/api/hse/kpis").get_json()
        self.assertEqual(eng_kpis["summary"]["total_records"], 0)
        self.assertEqual(eng_kpis["projects"], [])
        self.assertEqual(c_eng.get(f"/api/hse/overview?pid={hse_a}").status_code, 403)

    def test_04_hse_command_dashboard_ui_and_eng_om_non_regression(self):
        db = self.db
        db.add_user("syn_d2_sup", "SynPass!123", role="superadmin", primary_workspace="ENG")
        client = self.app.test_client()
        client.set_cookie("dcr_token", db.create_session("syn_d2_sup", "superadmin"))

        # HSE workspace renders Executive Command Dashboard + keeps Completed / Closed KPI compatibility
        r_hse = client.get("/?ws=HSE")
        self.assertEqual(r_hse.status_code, 200)
        hse_html = r_hse.get_data(as_text=True)
        self.assertIn("Executive Command Center", hse_html)
        self.assertIn("LTI-Free Days Streak", hse_html)
        self.assertIn("Any-Injury-Free Days", hse_html)
        self.assertIn("LTIF (per 1M Man-Hours)", hse_html)
        self.assertIn("TRIR (per 1M Man-Hours)", hse_html)
        self.assertIn("Severity Rate (per 1M hrs)", hse_html)
        self.assertIn("hseTrendChart", hse_html)
        self.assertIn("hseStatusChart", hse_html)
        self.assertIn("MutationObserver", hse_html)
        self.assertIn('id="admin-modal"', hse_html)
        self.assertIn('id="permissions-modal"', hse_html)
        self.assertNotIn("Facility / Contract Status Breakdown", hse_html)
        self.assertIsNotNone(re.search(r'Completed / Closed</span>\s*<span class="om-kpi-val success"[^>]*>\d+</span>', hse_html))

        # ENG and OM dashboards remain intact
        r_eng = client.get("/?ws=ENG")
        self.assertEqual(r_eng.status_code, 200)
        eng_html = r_eng.get_data(as_text=True)
        self.assertIn("<title>DCR — Dashboard</title>", eng_html)
        self.assertNotIn("hseTrendChart", eng_html)

        r_om = client.get("/?ws=OM")
        self.assertEqual(r_om.status_code, 200)
        om_html = r_om.get_data(as_text=True)
        self.assertIn("Total Facilities & Contracts", om_html)
        self.assertIn("Facility / Contract Status Breakdown", om_html)
        self.assertNotIn("hseTrendChart", om_html)


if __name__ == "__main__":
    unittest.main()
