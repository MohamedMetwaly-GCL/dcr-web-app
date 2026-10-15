"""test_dashboard_kpi_hierarchy.py - Validation test suite for:
1. Status-bearing submittals KPI aggregation (Total == Approved + Pending + Rejected).
2. Elimination of untracked non-status records from Submittals KPI.
3. Dashboard HTML visual hierarchy reordering:
   - 1. Top Control & Navigation Bar
   - 2. PROJECTS Grid
   - 3. SUBMITTALS KPI SUMMARY
   - 4. DOCUMENTS BY PROJECT & STATUS DISTRIBUTION
   - 5. OPERATIONAL BREAKDOWN
   - 6. PR ANALYTICS
   - 7. LETTERS OVERVIEW
4. Donut Chart overdue double-counting prevention (Pending (On Time) + Overdue == Pending).
5. Elimination of 'Unspecified' in PR Analytics (graceful fallback to trade, subTrade, discipline, 'General / Multi-Discipline').
"""
import unittest
from unittest.mock import MagicMock, patch
import html_render
import db

class TestDashboardKPIHierarchy(unittest.TestCase):

    def test_dashboard_stats_submittal_kpi_strict_math(self):
        """Verify that get_dashboard_stats strictly counts only status-bearing submittals,
        and Total == Approved + Pending + Rejected."""
        mock_projects = [{"id": "P01", "name": "Project Alpha", "code": "PA-01", "data": {}}]
        mock_dts = [
            {"id": 1, "project_id": "P01", "code": "MS", "name": "Method Statement", "sort_order": 1},
            {"id": 2, "project_id": "P01", "code": "GEN", "name": "General Letter", "sort_order": 2}
        ]
        # dt 1 has status, dt 2 has NO status
        mock_exp_cols = [
            {"col_key": "status", "dt_id": 1},
            {"col_key": "expectedReplyDate", "dt_id": 1},
            {"col_key": "discipline", "dt_id": 1}
        ]
        # records for dt 1:
        # doc 1: Approved
        # doc 2: Pending (Under Review)
        # doc 3: Rejected
        # doc 4: Cancelled (should be excluded)
        # doc 5: empty status (should be excluded)
        # records for dt 2 (no status column):
        # doc 6, doc 7 (should be completely excluded from submittals total)
        mock_records = [
            {"project_id": "P01", "dt_id": 1, "data": {"docNo": "MS-001", "status": "Approved", "discipline": "Civil"}},
            {"project_id": "P01", "dt_id": 1, "data": {"docNo": "MS-002", "status": "Under Review", "discipline": "Civil", "issuedDate": "2026-01-01"}},
            {"project_id": "P01", "dt_id": 1, "data": {"docNo": "MS-003", "status": "Revise and Resubmit", "discipline": "Mechanical"}},
            {"project_id": "P01", "dt_id": 1, "data": {"docNo": "MS-004", "status": "Cancelled", "discipline": "Civil"}},
            {"project_id": "P01", "dt_id": 1, "data": {"docNo": "MS-005", "status": "", "discipline": "Civil"}},
            {"project_id": "P01", "dt_id": 2, "data": {"docNo": "GEN-001", "subject": "Notice"}},
            {"project_id": "P01", "dt_id": 2, "data": {"docNo": "GEN-002", "subject": "Letter"}},
        ]

        with patch("db.get_projects", return_value=mock_projects), \
             patch("db.q") as mock_q, \
             patch("db.exe"):
            
            def q_side_effect(sql, params=None, one=False):
                sql_s = str(sql).strip().upper()
                if "FROM DASHBOARD_CACHE" in sql_s:
                    return []
                if "FROM DROPDOWN_LISTS" in sql_s:
                    return [
                        {"project_id": "P01", "item_value": "Approved", "meta": "approved"},
                        {"project_id": "P01", "item_value": "Under Review", "meta": "pending"},
                        {"project_id": "P01", "item_value": "Revise and Resubmit", "meta": "rejected"},
                        {"project_id": "P01", "item_value": "Cancelled", "meta": "cancelled"}
                    ]
                if "FROM RECORDS" in sql_s:
                    return mock_records
                if "FROM DOC_TYPES" in sql_s:
                    return mock_dts
                if "FROM COLUMNS_CONFIG" in sql_s:
                    return mock_exp_cols
                return []

            mock_q.side_effect = q_side_effect

            stats = db.get_dashboard_stats(["P01"])
            self.assertEqual(len(stats), 1)
            p = stats[0]

            # Mathematical consistency checks
            self.assertEqual(p["approved"], 1)
            self.assertEqual(p["pending"], 1)
            self.assertEqual(p["rejected"], 1)
            self.assertEqual(p["total"], 3, "Total must strictly equal approved (1) + pending (1) + rejected (1)")
            self.assertEqual(p["total"], p["approved"] + p["pending"] + p["rejected"])
            self.assertEqual(p["pct"], 33, "Completion pct should be round((1 / 3) * 100) = 33%")

            # dt_stats consistency
            dt1_stats = [dt for dt in p["dt_stats"] if dt["id"] == 1][0]
            self.assertEqual(dt1_stats["total"], 3)
            self.assertEqual(dt1_stats["approved"], 1)
            self.assertEqual(dt1_stats["pending"], 1)
            self.assertEqual(dt1_stats["rejected"], 1)

            # dt 2 has no status column, so it is omitted from submittal stats
            dt2_stats = [dt for dt in p["dt_stats"] if dt["id"] == 2]
            self.assertEqual(len(dt2_stats), 0, "Doc types without status column must not be in submittals dt_stats")

    def test_stale_cache_auto_recalculation(self):
        """Verify that if dashboard_cache contains stale data where total != approved + pending + rejected,
        it is automatically treated as missing and recalculated."""
        mock_projects = [{"id": "P01", "name": "Project Alpha", "code": "PA-01", "data": {}}]
        stale_cached_data = {
            "id": "P01", "name": "Project Alpha", "code": "PA-01",
            "total": 100, "approved": 50, "pending": 10, "rejected": 5, # 100 != 50 + 10 + 5 (stale legacy cache)
            "overdue": 2, "pct": 50, "dt_stats": []
        }

        with patch("db.get_projects", return_value=mock_projects), \
             patch("db.q") as mock_q, \
             patch("db.exe"):
            
            def q_side_effect(sql, params=None, one=False):
                sql_s = str(sql).strip().upper()
                if "FROM DASHBOARD_CACHE" in sql_s:
                    return [{"project_id": "P01", "cache_data": stale_cached_data}]
                if "FROM DROPDOWN_LISTS" in sql_s:
                    return []
                if "FROM RECORDS" in sql_s:
                    return []
                if "FROM DOC_TYPES" in sql_s:
                    return []
                if "FROM COLUMNS_CONFIG" in sql_s:
                    return []
                return []

            mock_q.side_effect = q_side_effect

            stats = db.get_dashboard_stats(["P01"])
            self.assertEqual(len(stats), 1)
            p = stats[0]
            # Must have been recalculated fresh from records (0 records -> 0 total) instead of serving the stale 100 total
            self.assertEqual(p["total"], 0)
            self.assertEqual(p["approved"], 0)
            self.assertEqual(p["pending"], 0)
            self.assertEqual(p["rejected"], 0)

    def test_dashboard_html_reordered_hierarchy(self):
        """Verify HTML layout order:
        1. Top Control & Navigation Bar (.psel-bar)
        2. In Overview: PROJECTS Grid (.overview-projects-panel)
        3. In Overview: SUBMITTALS KPI SUMMARY (Submittals KPI Summary + .kpi-grid)
        4. In Overview: OPERATIONAL BREAKDOWN (Operational Breakdown)
        5. In Overview: PR ANALYTICS (.overview-pr-panel)
        6. In Overview: LETTERS OVERVIEW (Letters Overview)
        7. In Analytics: Documents by Project & Status Distribution Charts
        """
        html = html_render.render_dashboard(
            {"id": 1, "username": "admin", "role": "admin"}
        )

        # 1. Check positions strictly in sequence
        psel_pos = html.find('class="psel-bar"')
        tab_overview_pos = html.find('id="tab-overview"')
        projects_grid_pos = html.find('class="panel overview-projects-panel"')
        kpi_summary_pos = html.find('Submittals KPI Summary')
        op_breakdown_pos = html.find('Operational Breakdown')
        pr_panel_pos = html.find('class="panel overview-pr-panel"')
        ltr_panel_pos = html.find('Letters Overview')
        tab_analytics_pos = html.find('id="tab-analytics"')
        cproj_pos = html.find('id="cProj"')
        cstatus_pos = html.find('id="cStatus"')

        self.assertTrue(psel_pos > 0, "Toolbar .psel-bar should be present")
        self.assertTrue(tab_overview_pos > psel_pos, "Toolbar should be FIRST, above the tabs")
        self.assertTrue(projects_grid_pos > tab_overview_pos, "Projects Grid should be inside Overview tab")
        self.assertTrue(kpi_summary_pos > projects_grid_pos, "Submittals KPI Summary should be directly below Projects Grid")
        self.assertTrue(op_breakdown_pos > kpi_summary_pos, "Operational Breakdown must be below Submittals KPI Summary")
        self.assertTrue(pr_panel_pos > op_breakdown_pos, "PR Analytics must be below Operational Breakdown")
        self.assertTrue(ltr_panel_pos > pr_panel_pos, "Letters Overview must be below PR Analytics")
        self.assertTrue(tab_analytics_pos > ltr_panel_pos, "Analytics tab should be after Overview tab")

        # 2. Charts must be located inside the Analytics tab
        self.assertTrue(cproj_pos > tab_analytics_pos, "Documents by Project chart (cProj) must be inside tab-analytics")
        self.assertTrue(cstatus_pos > tab_analytics_pos, "Status Distribution chart (cStatus) must be inside tab-analytics")

        # 3. Check labels & features
        self.assertIn("Total Submittals", html, "First KPI card must be labeled Total Submittals")
        self.assertIn("FROM PENDING", html, "Overdue KPI card must contain FROM PENDING subtitle")
        self.assertIn("Most Active Disciplines / Trades", html, "PR Analytics header must be Most Active Disciplines / Trades")
        self.assertIn("Pending (On Time)", html, "Donut chart must contain Pending (On Time) slice")

    def test_pr_analytics_unspecified_elimination(self):
        """Verify that get_pr_analytics_summary falls back through trade -> subTrade -> discipline -> General / Multi-Discipline
        and strictly eliminates 'Unspecified'."""
        mock_projects = [{"id": "P01", "name": "CFC Project", "code": "PEM-058"}]
        mock_pr_rows = [
            # 1. Has trade
            {"id": "r1", "project_id": "P01", "data": {"trade": "HVAC", "docNo": "PR-01"}},
            # 2. Has subTrade
            {"id": "r2", "project_id": "P01", "data": {"subTrade": "Plumbing", "docNo": "PR-02"}},
            # 3. Has discipline
            {"id": "r3", "project_id": "P01", "data": {"discipline": "Electrical", "docNo": "PR-03"}},
            # 4. Has discipline with whitespace
            {"id": "r4", "project_id": "P01", "data": {"discipline": "  Civil  ", "docNo": "PR-04"}},
            # 5. Has no trade/discipline info at all
            {"id": "r5", "project_id": "P01", "data": {"docNo": "PR-05"}},
            # 6. Has literal 'Unspecified'
            {"id": "r6", "project_id": "P01", "data": {"discipline": "Unspecified", "docNo": "PR-06"}},
        ]

        with patch("db.get_projects", return_value=mock_projects), \
             patch("db.q") as mock_q, \
             patch("db.exe"):
            
            def q_side_effect(sql, params=None, one=False):
                sql_s = str(sql).strip().upper()
                if "FROM DASHBOARD_CACHE" in sql_s:
                    return None
                if "FROM RECORDS R" in sql_s:
                    return mock_pr_rows
                return []

            mock_q.side_effect = q_side_effect

            summary = db.get_pr_analytics_summary(pid="P01")
            self.assertEqual(summary["total_pr_records"], 6)

            top_trades = summary["top_trades"]
            trade_names = [t["trade"] for t in top_trades]

            # Verify 'Unspecified' is 100% eliminated
            self.assertNotIn("Unspecified", trade_names)
            self.assertNotEqual(summary["top_trade_name"], "Unspecified")

            # Verify expected trade/discipline categories
            self.assertIn("HVAC", trade_names)
            self.assertIn("Plumbing", trade_names)
            self.assertIn("Electrical", trade_names)
            self.assertIn("Civil", trade_names)
            self.assertIn("General / Multi-Discipline", trade_names)

            # Rows 5 and 6 should be mapped to General / Multi-Discipline (count = 2)
            gen_count = [t["pr_count"] for t in top_trades if t["trade"] == "General / Multi-Discipline"][0]
            self.assertEqual(gen_count, 2)

    def test_dashboard_tab_switching_and_panes(self):
        """Verify that all tabs exist in the DOM and showTab properly activates selected panes."""
        html = html_render.render_dashboard(
            {"id": 1, "username": "admin", "role": "admin"}
        )

        # 1. Verify all 6 tab panes exist
        expected_panes = [
            'id="tab-overview"',
            'id="tab-analytics"',
            'id="tab-overdue"',
            'id="tab-executive"',
            'id="tab-daily-digest"',
            'id="tab-audit"',
        ]
        for pane in expected_panes:
            self.assertIn(pane, html, f"Dashboard must contain {pane}")

        # 2. Verify Overview is active by default
        self.assertIn('id="tab-overview" class="tab-pane active"', html)

        # 3. Verify showTab JavaScript implementation has pane activation logic
        self.assertIn("pane.classList.add('active')", html, "showTab must add 'active' class to the targeted pane")
        self.assertIn("document.querySelectorAll('.tab-pane').forEach", html)
        self.assertIn("document.querySelectorAll('.psel-bar .tbtn').forEach", html)

        # 4. Verify renderCharts has guards against hidden container
        self.assertIn("tab-analytics", html)
        self.assertIn("pane.classList.contains('active')", html, "renderCharts must verify tab-analytics is active before rendering")

if __name__ == "__main__":
    unittest.main()

