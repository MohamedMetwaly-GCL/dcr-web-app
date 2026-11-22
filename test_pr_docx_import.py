"""Offline unit & security tests for PR Word (.docx) Import (Phase 1A).

Run with: python -B -m unittest -v test_pr_docx_import
"""

import base64
import io
import os
import unittest
import zipfile
from contextlib import ExitStack
from copy import deepcopy
from unittest.mock import patch

from pr_docx_parser import normalize_unit, parse_date_to_iso, parse_pr_docx
from security import scan_and_validate_file


_CONTENT_TYPES_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>
"""


def _build_docx_bytes(document_xml: str, extra_files=None, include_content_types=True, include_document=True) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        if include_content_types:
            zf.writestr("[Content_Types].xml", _CONTENT_TYPES_XML)
        if include_document:
            zf.writestr("word/document.xml", document_xml)
        if extra_files:
            for name, content in extra_files.items():
                zf.writestr(name, content)
    return buf.getvalue()


def _gas_chill_document_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>طلب احتياج مواد - Gas Chill</w:t></w:r></w:p>
    <w:tbl>
      <w:tr>
        <w:tc><w:p><w:r><w:t>رقم أشعار الاحتياج</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>PEM064-PR-075</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>تاريخ ورود الأشعار</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>15/10/2026</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
    <w:tbl>
      <w:tr>
        <w:tc><w:p><w:r><w:t>رقم التصنيف</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>رقم القطعة</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>م</w:t></w:r></w:p></w:tc>
        <w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr><w:p><w:r><w:t>بيان المطلوب</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>الكمية</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>الوحدة</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>ملاحظات</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:tcPr><w:gridSpan w:val="8"/></w:tcPr><w:p><w:r><w:t>Air Duct &amp; Insulation Accessories</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p><w:r><w:t></w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t></w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>1</w:t></w:r></w:p></w:tc>
        <w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr><w:p><w:r><w:t>صاج مجلفن - Galvanized Steel Sheet 0.8mm</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>250</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>كيلوجرام</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>Urgent</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p/></w:tc>
        <w:tc><w:p/></w:tc>
        <w:tc><w:p/></w:tc>
        <w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr><w:p/></w:tc>
        <w:tc><w:p/></w:tc>
        <w:tc><w:p/></w:tc>
        <w:tc><w:p/></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p><w:r><w:t>CAT-02</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>PN-99</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>2</w:t></w:r></w:p></w:tc>
        <w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr><w:p><w:r><w:t>Duct Damper 400x400</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>12</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>عدد</w:t></w:r></w:p></w:tc>
        <w:tc><w:p/></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:tcPr><w:gridSpan w:val="8"/></w:tcPr><w:p><w:r><w:t>Approved Material Submittal: PEM064-MS-016 REV01</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:tcPr><w:gridSpan w:val="4"/></w:tcPr><w:p><w:r><w:t>الميزانية تسمح</w:t></w:r></w:p></w:tc>
        <w:tc><w:tcPr><w:gridSpan w:val="4"/></w:tcPr><w:p><w:r><w:t>سلطات الاعتماد</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:tcPr><w:gridSpan w:val="8"/></w:tcPr><w:p><w:r><w:t>يعتمد</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
    <w:p><w:r><w:t>Approved Material Submittal: PEM064-MS-017 REV00</w:t></w:r></w:p>
  </w:body>
</w:document>
"""


class DocxSecurityValidationTests(unittest.TestCase):
    def test_valid_docx_passes(self):
        raw = _build_docx_bytes(_gas_chill_document_xml())
        ok, err = scan_and_validate_file(raw, "docx", ["docx"], max_size_mb=10)
        self.assertTrue(ok, err)

    def test_rejects_non_zip_fake_docx(self):
        ok, err = scan_and_validate_file(b"Not a zip file at all", "docx", ["docx"])
        self.assertFalse(ok)
        self.assertIn("OpenXML", err)

    def test_rejects_vba_macro_payload(self):
        raw = _build_docx_bytes(
            _gas_chill_document_xml(),
            extra_files={"word/vbaProject.bin": b"fake-vba-bytes"},
        )
        ok, err = scan_and_validate_file(raw, "docx", ["docx"])
        self.assertFalse(ok)
        self.assertTrue("binary" in err.lower() or "vba" in err.lower() or "macro" in err.lower(), err)

    def test_rejects_arbitrary_bin_and_exe_members(self):
        raw_bin = _build_docx_bytes(_gas_chill_document_xml(), extra_files={"word/custom.bin": b"123"})
        ok_bin, _ = scan_and_validate_file(raw_bin, "docx", ["docx"])
        self.assertFalse(ok_bin)

        raw_exe = _build_docx_bytes(_gas_chill_document_xml(), extra_files={"word/dropper.exe": b"123"})
        ok_exe, _ = scan_and_validate_file(raw_exe, "docx", ["docx"])
        self.assertFalse(ok_exe)

    def test_rejects_path_traversal_member(self):
        raw = _build_docx_bytes(_gas_chill_document_xml(), extra_files={"../../evil.xml": b"<x/>"})
        ok, err = scan_and_validate_file(raw, "docx", ["docx"])
        self.assertFalse(ok)
        self.assertIn("unsafe archive member path", err)

    def test_rejects_missing_document_xml(self):
        raw = _build_docx_bytes(_gas_chill_document_xml(), include_document=False)
        ok, err = scan_and_validate_file(raw, "docx", ["docx"])
        self.assertFalse(ok)
        self.assertIn("word/document.xml", err)

    def test_rejects_missing_content_types(self):
        raw = _build_docx_bytes(_gas_chill_document_xml(), include_content_types=False)
        ok, err = scan_and_validate_file(raw, "docx", ["docx"])
        self.assertFalse(ok)
        self.assertIn("[Content_Types].xml", err)

    def test_rejects_oversized_uncompressed_zip_bomb(self):
        # 51 MB of zeros compresses to ~50 KB
        huge_xml = b" " * (51 * 1024 * 1024)
        raw = _build_docx_bytes(_gas_chill_document_xml(), extra_files={"word/huge.xml": huge_xml})
        ok, err = scan_and_validate_file(raw, "docx", ["docx"])
        self.assertFalse(ok)
        self.assertIn("50MB", err)


class DocxParserTests(unittest.TestCase):
    def test_gas_chill_fixture_parsing(self):
        raw = _build_docx_bytes(_gas_chill_document_xml())
        result = parse_pr_docx(raw)

        self.assertEqual(result["header"]["docNo"], "PEM064-PR-075")
        self.assertEqual(result["header"]["issuedDate"], "2026-10-15")
        self.assertEqual(
            result["header"]["msRef"],
            "PEM064-MS-016 REV01\nPEM064-MS-017 REV00",
        )
        self.assertEqual(result["warnings"], [])

        items = result["items"]
        self.assertEqual(len(items), 3)
        self.assertEqual(items[0], {"row_type": "header", "item_name": "Air Duct & Insulation Accessories"})
        self.assertEqual(items[1], {
            "row_type": "item",
            "item_name": "صاج مجلفن - Galvanized Steel Sheet 0.8mm",
            "unit": "Kg",
            "quantity": 250,
            "remarks": "Urgent",
        })
        self.assertEqual(items[2], {
            "row_type": "item",
            "item_name": "Duct Damper 400x400",
            "unit": "No",
            "quantity": 12,
            "remarks": "",
        })

    def test_unit_and_date_normalization(self):
        self.assertEqual(normalize_unit("  عدد "), "No")
        self.assertEqual(normalize_unit("قطعة"), "No")
        self.assertEqual(normalize_unit("حبة"), "No")
        self.assertEqual(normalize_unit("كيلو جرام"), "Kg")
        self.assertEqual(normalize_unit("كجم"), "Kg")
        self.assertEqual(normalize_unit("متر طولي"), "m")
        self.assertEqual(normalize_unit("م.ط"), "m")
        self.assertEqual(normalize_unit("متر مربع"), "m2")
        self.assertEqual(normalize_unit("م2"), "m2")
        self.assertEqual(normalize_unit("متر مكعب"), "m3")
        self.assertEqual(normalize_unit("لتر"), "L")
        self.assertEqual(normalize_unit("طقم"), "Set")
        self.assertEqual(normalize_unit("مجموعة"), "Set")
        self.assertEqual(normalize_unit("  CustomBox  "), "CustomBox")

        self.assertEqual(parse_date_to_iso("05/11/2026"), "2026-11-05")
        self.assertEqual(parse_date_to_iso("2026-11-05"), "2026-11-05")

    def test_warnings_and_vmerge(self):
        xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:tbl>
      <w:tr>
        <w:tc><w:p><w:r><w:t>Sr</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>Description</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>Qty</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>Unit</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p><w:r><w:t>1</w:t></w:r></w:p></w:tc>
        <w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr><w:p><w:r><w:t>Valve DN50</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>TBD</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>No</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p/></w:tc>
        <w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p/></w:tc>
        <w:tc><w:p><w:r><w:t>5</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>m</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
  </w:body>
</w:document>
"""
        raw = _build_docx_bytes(xml)
        result = parse_pr_docx(raw)
        self.assertEqual(len(result["items"]), 1)
        self.assertEqual(result["items"][0]["quantity"], "TBD")
        self.assertEqual(len(result["warnings"]), 2)


class ParsePrDocxEndpointTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        isolation = ExitStack()
        cls.addClassCleanup(isolation.close)
        isolation.enter_context(patch.dict(os.environ, {
            "SECRET_KEY": "synthetic-pr-docx-session-key",
            "MAGIC_SECRET": "synthetic-pr-docx-digest-key",
            "DATABASE_URL": "dbname=offline_pr_docx",
        }, clear=True))

        def blocked(*args, **kwargs):
            raise AssertionError("External I/O attempted in offline PR DOCX suite")

        for target in ("socket.socket.connect", "socket.socket.connect_ex",
                       "socket.socket.sendto", "socket.getaddrinfo",
                       "psycopg2.connect", "psycopg2._connect"):
            isolation.enter_context(patch(target, side_effect=blocked))

        import db
        isolation.enter_context(patch.object(db, "get_pool", side_effect=blocked))
        isolation.enter_context(patch.object(db, "init", side_effect=blocked))
        from app import app

        cls.db = db
        cls.app = app
        isolation.enter_context(patch.dict(app.config, {"TESTING": True}))

    def setUp(self):
        self.client = self.app.test_client()
        self.projects = {
            "ENG-P1": {"id": "ENG-P1", "workspace_id": "ENG", "name": "ENG-P1", "code": "ENG-P1", "data": {}},
        }
        self.allowed_dts = ["PR", "DS"]
        self.login_as("editor", {"ENG": "editor"}, ("ENG-P1",))

        mocks = ExitStack()
        self.addCleanup(mocks.close)
        mocks.enter_context(patch.object(self.db, "q", side_effect=self.read_fixture))
        self.write = mocks.enter_context(patch.object(self.db, "exe"))
        self.audit = mocks.enter_context(patch.object(self.db, "log_action"))

    def login_as(self, role, workspace_roles, project_ids=()):
        self.user = {"username": "synthetic-user", "role": role, "primary_workspace": "ENG"}
        self.workspace_roles = workspace_roles
        self.project_ids = project_ids
        self.client.set_cookie("dcr_token", "synthetic-session")

    def read_fixture(self, sql, params=(), one=False):
        query = " ".join(sql.lower().split())
        if "from sessions " in query:
            return deepcopy(self.user) if params == ("synthetic-session",) else None
        if "from users " in query:
            return deepcopy(self.user) if params == (self.user["username"],) else None
        if "from user_workspaces " in query:
            if one:
                role = self.workspace_roles.get(params[1])
                return {"role": role} if role else None
            return [{"workspace_id": ws, "role": role} for ws, role in self.workspace_roles.items()]
        if "from user_projects " in query:
            if query.startswith("select allowed_doc_types from "):
                return {"allowed_doc_types": self.allowed_dts} if params[1] in self.project_ids else None
            return [{"project_id": pid} for pid in self.project_ids]
        if "from projects " in query:
            if one and params:
                return deepcopy(self.projects.get(params[0]))
            return list(deepcopy(self.projects).values())
        if "from doc_types " in query:
            dts = [
                {"id": "PR", "project_id": "ENG-P1", "code": "PR", "name": "Purchase Requisition"},
                {"id": "DS", "project_id": "ENG-P1", "code": "DS", "name": "Document Submittal"},
            ]
            if one:
                dt_id = params[0] if len(params) == 1 else params[1]
                return next((d for d in dts if d["id"] == dt_id), None)
            return dts
        raise AssertionError(f"Unexpected SQL in fixture: {query}")

    def test_requires_edit_permission_and_rbac(self):
        raw = _build_docx_bytes(_gas_chill_document_xml())
        b64 = base64.b64encode(raw).decode("ascii")

        # Viewer cannot edit
        self.login_as("viewer", {"ENG": "viewer"}, ("ENG-P1",))
        resp = self.client.post("/api/pr_items/parse_docx/ENG-P1/PR", json={"file_b64": b64})
        self.assertEqual(resp.status_code, 403)

        # Editor without PR doc-type access
        self.login_as("editor", {"ENG": "editor"}, ("ENG-P1",))
        self.allowed_dts = ["DS"]
        resp = self.client.post("/api/pr_items/parse_docx/ENG-P1/PR", json={"file_b64": b64})
        self.assertEqual(resp.status_code, 403)

        # Non-PR doc type rejected with 400
        resp = self.client.post("/api/pr_items/parse_docx/ENG-P1/DS", json={"file_b64": b64})
        self.assertEqual(resp.status_code, 400)
        self.write.assert_not_called()

    def test_parse_base64_and_multipart_with_zero_db_writes(self):
        raw = _build_docx_bytes(_gas_chill_document_xml())
        b64 = base64.b64encode(raw).decode("ascii")

        # Base64 JSON
        resp = self.client.post(
            "/api/pr_items/parse_docx/ENG-P1/PR",
            json={"file_b64": b64, "ext": "docx"},
        )
        self.assertEqual(resp.status_code, 200)
        body = resp.get_json()
        self.assertTrue(body["ok"])
        self.assertEqual(body["header"]["docNo"], "PEM064-PR-075")
        self.assertEqual(len(body["items"]), 3)
        self.write.assert_not_called()

        # Multipart upload
        resp2 = self.client.post(
            "/api/pr_items/parse_docx/ENG-P1/PR",
            data={"file": (io.BytesIO(raw), "requisition.docx")},
            content_type="multipart/form-data",
        )
        self.assertEqual(resp2.status_code, 200)
        body2 = resp2.get_json()
        self.assertTrue(body2["ok"])
        self.assertEqual(len(body2["items"]), 3)
        self.write.assert_not_called()

    def test_endpoint_rejects_malicious_or_fake_docx(self):
        fake_b64 = base64.b64encode(b"not-a-docx").decode("ascii")
        resp = self.client.post("/api/pr_items/parse_docx/ENG-P1/PR", json={"file_b64": fake_b64})
        self.assertEqual(resp.status_code, 400)
        self.assertFalse(resp.get_json()["ok"])

        macro_raw = _build_docx_bytes(
            _gas_chill_document_xml(),
            extra_files={"word/vbaProject.bin": b"vba"},
        )
        macro_b64 = base64.b64encode(macro_raw).decode("ascii")
        resp2 = self.client.post("/api/pr_items/parse_docx/ENG-P1/PR", json={"file_b64": macro_b64})
        self.assertEqual(resp2.status_code, 400)
        self.assertFalse(resp2.get_json()["ok"])
        self.write.assert_not_called()
