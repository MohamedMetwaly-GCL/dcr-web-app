"""test_ai_tools_milestone1.py - Comprehensive Unit Tests for Milestone 1 Tools."""
import unittest
from unittest.mock import patch
from blueprints.ai_tools import (
    search_submittals,
    get_project_summary,
    get_noc_financials,
    get_overdue_submittals,
    get_document_history,
    _resolve_project_aliases,
)


class TestAiToolsMilestone1(unittest.TestCase):

    def setUp(self):
        self.mock_projects = [
            {"id": "PEM-058", "name": "CFC District Cooling Plant Phase 2", "code": "PEM-058"},
            {"id": "24CP01", "name": "EKH Hospital", "code": "24CP01"},
        ]

    def test_01_search_submittals_rbac_enforcement(self):
        """User cannot access projects not in allowed_project_ids."""
        with patch("db.q") as mock_q:
            mock_q.return_value = [{"id": "24CP01", "name": "EKH Hospital", "code": "24CP01"}]
            # User only has access to 24CP01, but queries PEM-058
            res = search_submittals(
                allowed_project_ids=["24CP01"],
                project_code="PEM-058",
                doc_type="MS"
            )
            self.assertEqual(res["count"], 0)
            self.assertEqual(len(res["records"]), 0)
        print("PASS: search_submittals strictly blocks unauthorized projects.")

    def test_02_search_submittals_valve_query_and_rev_dedup(self):
        """Verify search_submittals returns valve submittals and deduplicates older revisions."""
        raw_valve_records = [
            {
                "id": "r1",
                "project_id": "PEM-058",
                "doc_type": "MS",
                "doc_no": "MS-CY002P608-00019 REV00",
                "title": "Duty Valves for HVAC Piping",
                "status": "C - Revise & Resubmit",
                "discipline": "Mechanical",
                "trade": "HVAC",
            },
            {
                "id": "r2",
                "project_id": "PEM-058",
                "doc_type": "MS",
                "doc_no": "MS-CY002P608-00019 REV01",
                "title": "Duty Valves for HVAC Piping",
                "status": "B - Approved As Noted",
                "discipline": "Mechanical",
                "trade": "HVAC",
            },
            {
                "id": "r3",
                "project_id": "PEM-058",
                "doc_type": "MS",
                "doc_no": "MS-CY002P608-00020 REV00",
                "title": "Fire Fighting Valves (ZCV)",
                "status": "B - Approved As Noted",
                "discipline": "Mechanical",
                "trade": "Fire Fighting",
            },
        ]

        with patch("db.q") as mock_q:
            def q_side_effect(sql, params=()):
                if "FROM records r" in sql:
                    return raw_valve_records
                if "FROM projects" in sql:
                    return self.mock_projects
                return []
            mock_q.side_effect = q_side_effect

            res = search_submittals(
                allowed_project_ids=["PEM-058"],
                project_code="PEM-058",
                doc_type="MS",
                keyword="valve",
                latest_revision_only=True
            )

            # REV00 should be superseded by REV01, so exactly 2 unique submittals
            self.assertEqual(res["count"], 2)
            doc_nos = [r["doc_no"] for r in res["records"]]
            self.assertIn("MS-CY002P608-00019 REV01", doc_nos)
            self.assertNotIn("MS-CY002P608-00019 REV00", doc_nos)
            self.assertIn("MS-CY002P608-00020 REV00", doc_nos)
        print("PASS: search_submittals retrieves valves and deduplicates revisions.")

    def test_03_get_project_summary(self):
        """Verify get_project_summary aggregates KPI totals accurately."""
        mock_stats = [
            {"id": "PEM-058", "code": "PEM-058", "name": "CFC Ph2", "total": 100, "approved": 75, "pending": 20, "overdue": 5, "rejected": 5, "pct": 75},
            {"id": "24CP01", "code": "24CP01", "name": "EKH", "total": 50, "approved": 25, "pending": 20, "overdue": 5, "rejected": 5, "pct": 50},
        ]
        with patch("db.q") as mock_q, patch("db.get_dashboard_stats") as mock_dash:
            mock_q.return_value = self.mock_projects
            mock_dash.return_value = mock_stats

            summary = get_project_summary(allowed_project_ids=["PEM-058", "24CP01"])
            self.assertEqual(summary["total_projects"], 2)
            self.assertEqual(summary["grand_total_documents"], 150)
            self.assertEqual(summary["grand_total_approved"], 100)
            self.assertEqual(summary["overall_approval_pct"], 67)
        print("PASS: get_project_summary aggregates dashboard statistics.")

    def test_04_get_noc_financials(self):
        """Verify get_noc_financials calculates submitted and approved costs."""
        mock_noc_rows = [
            {
                "project_id": "PEM-058",
                "proj_code": "PEM-058",
                "proj_name": "CFC",
                "data": {
                    "docNo": "NOC-CY002P608-001",
                    "title": "Chilled Water Pipe Rerouting",
                    "submittedCost": "150000.00",
                    "finalApprovedCost": "120000.00",
                    "partDStatus": "Accepted & to proceed with Part C",
                    "voNo": "VO-01",
                }
            },
            {
                "project_id": "PEM-058",
                "proj_code": "PEM-058",
                "proj_name": "CFC",
                "data": {
                    "docNo": "NOC-CY002P608-002",
                    "title": "Additional Ductwork at Plant",
                    "submittedCost": "80000.00",
                    "finalApprovedCost": "0.00",
                    "partBStatus": "Under Review",
                }
            },
        ]
        with patch("db.q") as mock_q:
            def q_side_effect(sql, params=()):
                if "FROM records r" in sql:
                    return mock_noc_rows
                if "FROM projects" in sql:
                    return self.mock_projects
                return []
            mock_q.side_effect = q_side_effect

            fin = get_noc_financials(allowed_project_ids=["PEM-058"], project_code="CFC")
            self.assertEqual(fin["total_count"], 2)
            self.assertEqual(fin["total_submitted_cost"], 230000.0)
            self.assertEqual(fin["total_approved_cost"], 120000.0)
            self.assertEqual(fin["approved_count"], 1)
            self.assertEqual(fin["pending_count"], 1)
        print("PASS: get_noc_financials accurately calculates costs and statuses.")

    def test_05_get_overdue_submittals(self):
        """Verify get_overdue_submittals filters and sorts by delay."""
        mock_overdue = [
            {"project_id": "PEM-058", "dt_code": "IR", "docNo": "IR-100", "title": "Duct Inspection", "days_overdue": 14, "discipline": "Mechanical"},
            {"project_id": "PEM-058", "dt_code": "MS", "docNo": "MS-010", "title": "Paint Submittal", "days_overdue": 5, "discipline": "Architecture"},
            {"project_id": "PEM-058", "dt_code": "IR", "docNo": "IR-102", "title": "Cable Pulling", "days_overdue": 20, "discipline": "Electrical"},
        ]
        with patch("db.q") as mock_q, patch("db.get_overdue_records") as mock_ov:
            mock_q.return_value = self.mock_projects
            mock_ov.return_value = mock_overdue

            res = get_overdue_submittals(
                allowed_project_ids=["PEM-058"],
                project_code="PEM-058",
                doc_type="IR"
            )
            self.assertEqual(res["total_overdue"], 2)
            # Longest delay first
            self.assertEqual(res["records"][0]["docNo"], "IR-102")
            self.assertEqual(res["records"][0]["days_overdue"], 20)
            self.assertEqual(res["records"][1]["docNo"], "IR-100")
            self.assertEqual(res["records"][1]["days_overdue"], 14)
        print("PASS: get_overdue_submittals sorts and filters overdue documents.")

    def test_06_get_document_history(self):
        """Verify get_document_history tracks revisions chronologically."""
        mock_history_rows = [
            {
                "id": "r1",
                "project_id": "PEM-058",
                "dt_id": "MS",
                "proj_code": "PEM-058",
                "created_at": "2026-01-01",
                "data": {"docNo": "MS-CY002P608-00019 REV00", "status": "C - Revise & Resubmit", "issuedDate": "2026-01-01", "actualReplyDate": "2026-01-10"}
            },
            {
                "id": "r2",
                "project_id": "PEM-058",
                "dt_id": "MS",
                "proj_code": "PEM-058",
                "created_at": "2026-02-01",
                "data": {"docNo": "MS-CY002P608-00019 REV01", "status": "B - Approved As Noted", "issuedDate": "2026-02-01", "actualReplyDate": "2026-02-08"}
            },
        ]
        with patch("db.q") as mock_q:
            def q_side_effect(sql, params=()):
                if "FROM records r" in sql:
                    return mock_history_rows
                if "FROM projects" in sql:
                    return self.mock_projects
                return []
            mock_q.side_effect = q_side_effect

            hist = get_document_history(
                allowed_project_ids=["PEM-058"],
                doc_no_pattern="MS-CY002P608-00019 REV01"
            )
            self.assertEqual(hist["total_revisions"], 2)
            self.assertEqual(hist["latest_revision"], "REV 01")
            self.assertEqual(hist["latest_status"], "B - Approved As Noted")
            self.assertEqual(hist["revisions"][0]["rev_num"], 0)
            self.assertEqual(hist["revisions"][1]["rev_num"], 1)
        print("PASS: get_document_history compiles chronological revisions.")

    def test_07_cfc_alias_resolution_translates_to_cfc_dcp_ph2a(self):
        """Verify project alias resolver translates 'PEM-058' to include 'CFC DCP PH-2A' and returns financials."""
        aliases = _resolve_project_aliases(None, "PEM-058", ["PEM-058"])
        self.assertIn("CFC DCP PH-2A", aliases)
        self.assertIn("PEM-058", aliases)

        cfc_records = [
            {
                "project_id": "CFC DCP PH-2A",
                "proj_code": "PEM-058",
                "data": {
                    "docNo": "NOC-CY002P608-001",
                    "submittedCost": "150000.00",
                    "finalApprovedCost": "120000.00",
                    "partDStatus": "Approved",
                }
            }
        ]
        with patch("db.q") as mock_q:
            def q_side_effect(sql, params=()):
                if "FROM records r" in sql:
                    params_flat = str(params)
                    self.assertIn("CFC DCP PH-2A", params_flat)
                    return cfc_records
                if "FROM projects" in sql:
                    return self.mock_projects
                return []
            mock_q.side_effect = q_side_effect

            res = get_noc_financials(allowed_project_ids=["PEM-058"], project_code="PEM-058")
            self.assertEqual(res["total_count"], 1)
            self.assertGreater(res["total_submitted_cost"], 0)
            self.assertGreater(res["total_approved_cost"], 0)
            self.assertEqual(res["total_approved_cost"], 120000.0)
        print("PASS: get_noc_financials for PEM-058 correctly queries 'CFC DCP PH-2A' and returns financial values.")

    def test_08_array_binding_and_project_display_name(self):
        """Verify search_submittals uses r.project_id = ANY(%s) with flat Python list and returns clean display name."""
        mock_valves = [
            {
                "id": "v1",
                "project_id": "CFC DCP PH-2A",
                "doc_type": "MS",
                "doc_no": "MS-CY002P608-00019 REV01",
                "title": "Duty Valves for HVAC Piping",
                "status": "B - Approved As Noted",
                "discipline": "Mechanical",
            }
        ]
        with patch("db.q") as mock_q:
            captured_sql = []
            captured_params = []

            def q_side_effect(sql, params=()):
                captured_sql.append(sql)
                captured_params.append(params)
                if "FROM records r" in sql:
                    return mock_valves
                if "FROM projects" in sql:
                    return self.mock_projects
                return []
            mock_q.side_effect = q_side_effect

            # 1. Explicit project requested
            res = search_submittals(
                allowed_project_ids=["PEM-058"],
                project_code="PEM-058",
                doc_type="MS",
                keyword="valve"
            )

            # Find the query to records
            record_queries = [sql for sql in captured_sql if "FROM records r" in sql]
            self.assertTrue(len(record_queries) > 0)
            rec_sql = record_queries[0]
            self.assertIn("r.project_id = ANY(%s)", rec_sql)

            # Find corresponding params
            rec_idx = captured_sql.index(rec_sql)
            rec_params = captured_params[rec_idx]
            array_param = rec_params[0]
            self.assertIsInstance(array_param, list)
            self.assertIn("CFC DCP PH-2A", array_param)
            self.assertIn("PEM-058", array_param)

            self.assertEqual(res["project_display_name"], "PEM-058")
            self.assertEqual(res["count"], 1)
            self.assertEqual(res["records"][0]["doc_no"], "MS-CY002P608-00019 REV01")

            # 2. Global query across all projects
            res_global = search_submittals(
                allowed_project_ids=["PEM-058"],
                project_code=None,
                doc_type="MS"
            )
            self.assertEqual(res_global["project_display_name"], "All Projects")
        print("PASS: search_submittals correctly binds array and returns project_display_name.")


if __name__ == "__main__":
    unittest.main()


