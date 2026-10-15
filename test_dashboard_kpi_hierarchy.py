"""test_dashboard_kpi_hierarchy.py - Validation test suite for:
1. Status-bearing submittals KPI aggregation (Total == Approved + Pending + Rejected).
2. Elimination of untracked non-status records from Submittals KPI.
3. Dashboard HTML visual hierarchy reordering (Toolbar FIRST, Projects Grid SECOND, Submittals KPI Summary THIRD).
4. Submittal completion percentage based on status-bearing items.
5. Invalidation / recalculation of stale dashboard cache rows.
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
        1. Toolbar (.psel-bar) FIRST
        2. Projects Grid (.overview-projects-panel) SECOND
        3. Submittals KPI Summary (Submittals KPI Summary + .kpi-grid) THIRD
        4. First card label is 'Total Submittals'
        5. Overdue card has subtitle 'FROM PENDING'
        """
        html = html_render.render_dashboard(
            {"id": 1, "username": "admin", "role": "admin"}
        )

        # 1. Check positions
        psel_pos = html.find('class="psel-bar"')
        tab_overview_pos = html.find('id="tab-overview"')
        projects_grid_pos = html.find('class="panel overview-projects-panel"')
        kpi_summary_pos = html.find('Submittals KPI Summary')
        pr_panel_pos = html.find('class="panel overview-pr-panel"')

        self.assertTrue(psel_pos > 0, "Toolbar .psel-bar should be present")
        self.assertTrue(tab_overview_pos > psel_pos, "Toolbar should be FIRST, above the tabs")
        self.assertTrue(projects_grid_pos > tab_overview_pos, "Projects Grid should be inside Overview tab")
        self.assertTrue(kpi_summary_pos > projects_grid_pos, "Submittals KPI Summary should be directly below Projects Grid")
        self.assertTrue(pr_panel_pos > kpi_summary_pos, "PR Analytics should be below Submittals KPI Summary")

        # 2. Check labels
        self.assertIn("Total Submittals", html, "First KPI card must be labeled Total Submittals")
        self.assertIn("FROM PENDING", html, "Overdue KPI card must contain FROM PENDING subtitle")

if __name__ == "__main__":
    unittest.main()
