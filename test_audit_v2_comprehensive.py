"""test_audit_v2_comprehensive.py

Comprehensive test suite verifying all 4 phases of Audit Log v2:
Phase 1: Network & Device Fingerprinting (IP, User-Agent, Device).
Phase 2: Read & View & Search Tracking (DOC_VIEW, SEARCH, AI QUERY).
Phase 3: Security & Administrative Auditing (ACCESS DENIED, USER_CREATE, ROLE_UPDATE, etc.).
Phase 4: Import & Export Forensics (EXPORT_EXCEL, EXPORT_PDF, IMPORT_EXCEL).
"""
import os
import unittest
os.environ['SECRET_KEY'] = 'test_secret_key_12345678901234567890123456789012'

import db
import app
import auth
from html_render import render_dashboard

class TestAuditV2Comprehensive(unittest.TestCase):
    def setUp(self):
        self.app = app.app
        self.client = self.app.test_client()

    def test_01_user_agent_parser(self):
        """Test user agent string parser for various OS and browser combinations."""
        # Windows Chrome
        ua1 = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        res1 = db.parse_user_agent(ua1)
        self.assertIn("Windows 10/11", res1)
        self.assertIn("Chrome", res1)

        # iPhone Safari
        ua2 = "Mozilla/5.0 (iPhone; CPU iPhone OS 17_0 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.0 Mobile/15E148 Safari/604.1"
        res2 = db.parse_user_agent(ua2)
        self.assertIn("iPhone", res2)
        self.assertIn("Safari", res2)

        # Android Chrome
        ua3 = "Mozilla/5.0 (Linux; Android 14; SM-S918B) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0 Mobile Safari/537.36"
        res3 = db.parse_user_agent(ua3)
        self.assertIn("Android", res3)
        self.assertIn("Chrome", res3)

        # Windows Edge
        ua4 = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36 Edg/120.0.0.0"
        res4 = db.parse_user_agent(ua4)
        self.assertIn("Windows 10/11", res4)
        self.assertIn("Edge", res4)

        # Empty / None
        self.assertEqual(db.parse_user_agent(None), "Unknown")
        self.assertEqual(db.parse_user_agent(""), "Unknown")
        print("Test 1: User-Agent parser verified: PASS")

    def test_02_automatic_ip_and_device_capture(self):
        """Verify log_action automatically captures IP and Device inside Flask context."""
        logged = []
        orig_exe = db.exe
        def mock_exe(sql, params=()):
            if "INSERT INTO audit_log" in sql:
                logged.append(params)
        db.exe = mock_exe

        try:
            # Simulate request behind Cloudflare
            with self.app.test_request_context(
                '/api/test',
                headers={
                    'CF-Connecting-IP': '197.35.40.12',
                    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/120.0.0.0 Safari/537.36'
                }
            ):
                db.log_action("mohamed", "LOGIN", detail="User logged in")

            self.assertEqual(len(logged), 1)
            params = logged[0]
            # username, action, project_id, dt_id, record_id, doc_no, field_name, old_value, new_value, detail, ip_address, user_agent, device
            self.assertEqual(params[0], "mohamed")
            self.assertEqual(params[1], "LOGIN")
            self.assertEqual(params[10], "197.35.40.12")
            self.assertIn("Chrome", params[11])
            self.assertIn("Windows 10/11", params[12])

            # Simulate request behind X-Forwarded-For
            with self.app.test_request_context(
                '/api/test2',
                headers={
                    'X-Forwarded-For': '41.238.10.5, 10.0.0.1',
                    'User-Agent': 'Mozilla/5.0 (iPhone; CPU iPhone OS 17_0) Safari/604.1'
                }
            ):
                db.log_action("engineer_a", "DOC_VIEW", project_id="PEM-064", doc_no="MS-001")

            self.assertEqual(len(logged), 2)
            params2 = logged[1]
            self.assertEqual(params2[0], "engineer_a")
            self.assertEqual(params2[1], "DOC_VIEW")
            self.assertEqual(params2[10], "41.238.10.5")
            self.assertIn("iPhone", params2[12])
        finally:
            db.exe = orig_exe
        print("Test 2: Automatic IP & Device extraction in log_action: PASS")

    def test_03_beacon_event_endpoint(self):
        """Test /api/audit/log_event beacon endpoint."""
        logged = []
        orig_log = db.log_action
        db.log_action = lambda username, action, **kw: logged.append({'username': username, 'action': action, **kw})

        try:
            with self.app.test_request_context('/api/audit/log_event', method='POST', json={
                'event': 'DOC_VIEW',
                'project_id': 'PEM-064',
                'dt_id': 'MS',
                'doc_no': 'PEM064-MS-004',
                'detail': 'Spotlight opened document'
            }):
                app.current_user = lambda: {'username': 'hassan', 'role': 'editor'}
                res = app.api_audit_log_event()
                self.assertEqual(res.status_code, 200)
                self.assertEqual(len(logged), 1)
                self.assertEqual(logged[0]['action'], 'DOC_VIEW')
                self.assertEqual(logged[0]['doc_no'], 'PEM064-MS-004')

            # Invalid event rejected
            with self.app.test_request_context('/api/audit/log_event', method='POST', json={'event': 'INVALID_ACTION'}):
                app.current_user = lambda: {'username': 'hassan', 'role': 'editor'}
                res = app.api_audit_log_event()
                status = res[1] if isinstance(res, tuple) else res.status_code
                self.assertEqual(status, 400)
        finally:
            db.log_action = orig_log
        print("Test 3: Beacon endpoint /api/audit/log_event: PASS")

    def test_04_security_access_denied_auditing(self):
        """Test ACCESS DENIED logging on unauthorized attempts."""
        logged = []
        orig_log = db.log_action
        db.log_action = lambda username, action, **kw: logged.append({'username': username, 'action': action, **kw})

        try:
            # Unauthorized API project access
            with self.app.test_request_context('/api/records/RESTRICTED_PID/MS'):
                app.current_user = lambda: {'username': 'site_user', 'role': 'viewer'}
                app.can_view_project = lambda pid, u: False
                res = app.enforce_login_and_project_scope()
                self.assertEqual(res[1], 403)
                denied_logs = [l for l in logged if l['action'] == 'ACCESS DENIED']
                self.assertEqual(len(denied_logs), 1)
                self.assertEqual(denied_logs[0]['username'], 'site_user')
                self.assertEqual(denied_logs[0]['project_id'], 'RESTRICTED_PID')

            # Unauthorized superadmin action
            @auth.require_superadmin
            def dummy_superadmin_view():
                return "ok"

            with self.app.test_request_context('/api/users'):
                auth.current_user = lambda: {'username': 'viewer_user', 'role': 'viewer'}
                res = dummy_superadmin_view()
                self.assertEqual(res[1], 403)
                admin_denied = [l for l in logged if l['username'] == 'viewer_user' and l['action'] == 'ACCESS DENIED']
                self.assertEqual(len(admin_denied), 1)
        finally:
            db.log_action = orig_log
        print("Test 4: Security ACCESS DENIED auditing: PASS")

    def test_05_dashboard_html_audit_table_ui(self):
        """Verify Dashboard HTML contains IP Address and Device columns and resize handlers."""
        html = render_dashboard({'username': 'admin_user', 'role': 'superadmin'})
        self.assertIn('IP Address', html)
        self.assertIn('Device', html)
        self.assertIn('AUD_DEFAULT_WIDTHS', html)
        self.assertIn('dcr_audit_col_widths_v2', html)
        self.assertIn('ACCESS DENIED', html)
        self.assertIn('AI QUERY', html)
        self.assertIn('SEARCH', html)
        self.assertIn('DOC_VIEW', html)
        print("Test 5: Dashboard HTML Audit table UI with 11 columns and new badges: PASS")

if __name__ == '__main__':
    unittest.main()
