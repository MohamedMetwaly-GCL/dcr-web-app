"""Tier 1 unit/regression tests for LTR-DATE-01 (LTR Date Range Filter)."""

import ast
import os
import unittest
from unittest.mock import patch

os.environ.setdefault("DATABASE_URL", "")
os.environ.setdefault("SECRET_KEY", "synthetic-test-secret-key-only")
os.environ.setdefault("MAGIC_SECRET", "synthetic-test-magic-secret-only")
os.environ.setdefault("GEMINI_API_KEY", "synthetic-test-gemini-key-only")

import html_render
import i18n


def _sample_user():
    return {
        "id": 1,
        "username": "tester",
        "full_name": "Test User",
        "role": "editor",
        "workspace_role": "editor",
        "is_superadmin": False,
        "can_download": True,
    }


def _sample_project():
    return {
        "id": 101,
        "code": "PRJ-01",
        "name": "Sample Project",
        "contractor": "Contractor",
        "consultant": "Consultant",
        "owner": "Owner",
        "workspace_id": "ENG",
    }


class TestLTRDateFilter(unittest.TestCase):
    """Verify LTR date range filter UI, localization, JS hooks, and Python 3.11 AST safety."""

    def test_i18n_keys_exist_with_en_ar_parity(self):
        expected_keys = [
            "reg.ltr_date_label",
            "reg.ltr_date_issued",
            "reg.ltr_date_received",
            "reg.ltr_date_from",
            "reg.ltr_date_to",
            "reg.ltr_date_clear",
        ]
        en_dict = i18n.TRANSLATIONS["en"]
        ar_dict = i18n.TRANSLATIONS["ar"]
        self.assertEqual(set(en_dict.keys()), set(ar_dict.keys()))
        for key in expected_keys:
            self.assertIn(key, en_dict)
            self.assertIn(key, ar_dict)
            self.assertTrue(en_dict[key].strip())
            self.assertTrue(ar_dict[key].strip())

    def test_register_renders_ltr_date_filter_en_and_ar(self):
        user = _sample_user()
        proj = _sample_project()
        dts = [
            {"id": "NOC", "code": "NOC", "name": "Notice of Change"},
            {"id": "LTR", "code": "LTR", "name": "Letters"},
        ]
        workspaces = [
            {
                "id": "ENG",
                "code": "ENG",
                "name": "Engineering & Projects",
                "name_ar": "المشروعات الهندسية",
                "icon": "🏗️",
            }
        ]
        with patch("html_render.can_edit", return_value=True), \
             patch("html_render.is_project_admin", return_value=True), \
             patch("html_render.get_allowed_doc_type_filter", return_value=None), \
             patch("db.can_user_download", return_value=True), \
             patch("db.get_projects_for_user", return_value=[proj]), \
             patch("db.get_doc_types", return_value=dts), \
             patch("db.get_logo", return_value=None), \
             patch("db.get_user_workspaces", return_value=workspaces):
            for lang, expected_dir, expected_labels in [
                ("en", 'dir="ltr"', ["Date:", "Issue Date", "Received Date", "From", "To", "Clear"]),
                ("ar", 'dir="rtl"', ["التاريخ:", "تاريخ الإصدار", "تاريخ الاستلام", "من", "إلى", "مسح"]),
            ]:
                with self.subTest(lang=lang):
                    out = html_render.render_register(user, proj, active_ws_id="ENG", lang=lang)
                    self.assertIn(expected_dir, out)
                    for marker in [
                        'id="ltr-quickbar"',
                        'id="ltr-date-filter"',
                        'id="ltr-date-field"',
                        'value="issuedDate"',
                        'value="receivedDate"',
                        'id="ltr-date-from"',
                        'id="ltr-date-to"',
                        'id="ltr-date-clear"',
                        'type="date"',
                        "function parseLTRDateKey(",
                        "function setLTRDateFilter(",
                        "function clearLTRDateFilter(",
                        "state.filters._ltrDateFrom",
                        "state.filters._ltrDateTo",
                        "state.filters._ltrDateField",
                    ]:
                        self.assertIn(marker, out)
                    for lbl in expected_labels:
                        self.assertIn(lbl, out)

    def test_no_backslashes_in_fstring_expressions(self):
        for path in ("html_render.py", "i18n.py"):
            with open(path, "r", encoding="utf-8") as fh:
                source = fh.read()
            tree = ast.parse(source, filename=path)
            lines = source.splitlines()
            for node in ast.walk(tree):
                if isinstance(node, ast.FormattedValue):
                    # Fast line-slice check before calling get_source_segment
                    start = max(0, node.lineno - 1)
                    end = getattr(node, "end_lineno", node.lineno)
                    chunk = "\n".join(lines[start:end])
                    if "\\" in chunk:
                        seg = ast.get_source_segment(source, node.value) or ""
                        self.assertNotIn(
                            "\\",
                            seg,
                            f"Backslash inside f-string expression in {path}:{node.lineno}: {seg}",
                        )


if __name__ == "__main__":
    unittest.main()
