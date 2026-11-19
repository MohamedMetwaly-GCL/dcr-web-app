"""Comprehensive regression test suite for the 5 DCR-FIXES-P1 improvements.

Designed to run inside the isolated PostgreSQL 17.12 sandbox harness
(see pg17_harness.py / DCR_TEST_DB_PREINJECTED=1) with zero live Drive or network calls.
"""

import io
import json
import os
from pathlib import Path
import re
import sys
import unittest
from unittest.mock import patch
import zipfile

from PIL import Image


REPO_ROOT = Path(__file__).resolve().parent


def _build_xlsx_bytes():
    import openpyxl

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Breakdown"
    ws.append(["Code", "Item", "Qty"])
    ws.append(["ST-01", "Steel Truss", 5])
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


class FixesP1IntegrationTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        if str(REPO_ROOT) not in sys.path:
            sys.path.insert(0, str(REPO_ROOT))
        import db
        import drive_routing
        import app as dcr_app
        from pg17_harness import SimDriveState

        drive_routing.invalidate_subfolder_caches()
        cls.db = db
        cls.app = dcr_app.app
        cls.app.config["TESTING"] = True
        cls.sim = SimDriveState()

        # Users
        for uname, role, can_dl, ws in (
            ("p1_super", "superadmin", True, "ENG"),
            ("p1_editor", "editor", True, "ENG"),
            ("p1_viewer_nodl", "viewer", False, "ENG"),
            ("p1_viewer_dl", "viewer", True, "ENG"),
            ("p1_outsider", "viewer", False, "OM"),
        ):
            if not db.get_user(uname):
                db.add_user(uname, "StrongPass!123", role=role, primary_workspace=ws, can_download=can_dl)
            db.assign_user_workspace(uname, ws, role=role)

        # Projects
        for pid, code, name, ws, dfid in (
            ("PEM-058", "PEM-058", "CFC Phase 2", "ENG", "SIM_ROOT_ENG"),
            ("EKH-001", "EKH-001", "EKH Project", "ENG", "SIM_ROOT_ENG"),
            ("OM-01", "OM-01", "OM Project", "OM", "SIM_ROOT_OM"),
            ("HSE-01", "HSE-01", "HSE Project", "HSE", "SIM_ROOT_HSE"),
        ):
            if not db.get_project(pid):
                db.create_project(pid, name, code, workspace_id=ws)
            db.update_project_settings(pid, {"drive_folder_id": dfid})

        if db.project_has_doc_type("PEM-058", "PCQ"):
            db.delete_doc_type("PEM-058", "PCQ")
        if not db.project_has_doc_type("EKH-001", "PCQ"):
            db.add_doc_type("EKH-001", "PCQ", "Pre-Qualification")

        for u in ("p1_editor", "p1_viewer_nodl", "p1_viewer_dl"):
            db.assign_project(u, "PEM-058")
            db.assign_project(u, "EKH-001")
        db.assign_project("p1_outsider", "OM-01")

        # Seed Drive root hierarchy for PEM-058 NOC
        cls.sim.add_folder("F_DT_NOC", "NOC", "SIM_ROOT_ENG")
        cls.sim.add_folder("F_PART_C", "Part C", "F_DT_NOC")

    def _client_as(self, username):
        client = self.app.test_client()
        u = self.db.get_user(username)
        tok = self.db.create_session(username, u["role"])
        client.set_cookie("dcr_token", tok, domain="localhost")
        return client

    def _upload_via_api(self, client, pid, dt_id, slot, filename, content_bytes,
                        subfolder_name="", rel_subfolder="", link_as_folder=False,
                        replace_existing=False, rev_mode="overwrite", part_rev=None,
                        record_id=""):
        import blueprints.drive_uploads as du

        with patch.object(du, "_token_provider", return_value="sim-tok"), \
             patch.object(du, "_http_client", self.sim.make_http()):
            s_resp = client.post("/api/drive/upload-session", json={
                "project_id": pid,
                "dt_id": dt_id,
                "record_id": record_id,
                "slot": slot,
                "file_name": filename,
                "file_size": len(content_bytes),
                "mime_type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" if filename.endswith(".xlsx") else "application/pdf",
                "create_folders": True,
                "subfolder_name": subfolder_name,
                "rel_subfolder": rel_subfolder,
                "link_as_folder": link_as_folder,
                "rev_mode": rev_mode,
                "part_rev": part_rev,
                "doc_no": subfolder_name or "DOC-TEST",
            })
            if s_resp.status_code != 200:
                return s_resp, None
            sess = s_resp.get_json()
            fid = sess["session_url"].rsplit("/", 1)[-1]
            # Simulate chunk proxy or direct PUT
            c_resp = client.post(
                "/api/drive/upload-chunk",
                headers={
                    "X-Upload-Ticket": sess["upload_ticket"],
                    "X-Content-Range": f"bytes 0-{len(content_bytes) - 1}/{len(content_bytes)}",
                },
                data=content_bytes,
            )
            self.assertEqual(c_resp.status_code, 200, c_resp.get_data(as_text=True))
            cmp_resp = client.post("/api/drive/upload-complete", json={
                "upload_ticket": sess["upload_ticket"],
                "file_id": fid,
                "record_id": record_id,
                "link_as_folder": link_as_folder,
                "replace_existing": replace_existing,
                "rel_subfolder": rel_subfolder,
                "rev_mode": rev_mode,
                "part_rev": part_rev,
            })
            return cmp_resp, fid

    # ── Feature 1: Favicon, PWA icons, manifest & all <head> tags ───────

    def test_01_favicon_manifest_assets_and_head_tags(self):
        static_dir = REPO_ROOT / "static"
        for name in ("favicon.ico", "icon-192.png", "icon-512.png", "apple-touch-icon.png", "manifest.json"):
            self.assertTrue((static_dir / name).is_file(), f"Missing static/{name}")
            self.assertGreater((static_dir / name).stat().st_size, 0)

        with Image.open(static_dir / "icon-192.png") as im:
            self.assertEqual(im.size, (192, 192))
        with Image.open(static_dir / "icon-512.png") as im:
            self.assertEqual(im.size, (512, 512))
        with Image.open(static_dir / "apple-touch-icon.png") as im:
            self.assertEqual(im.size, (180, 180))

        manifest = json.loads((static_dir / "manifest.json").read_text(encoding="utf-8"))
        self.assertEqual(manifest["start_url"], "/hub")
        self.assertEqual(manifest["scope"], "/")
        self.assertGreaterEqual(len(manifest["icons"]), 2)

        anon = self.app.test_client()
        r_ico = anon.get("/favicon.ico", follow_redirects=False)
        self.assertEqual(r_ico.status_code, 200)
        self.assertIn("icon", r_ico.content_type.lower())

        r_man = anon.get("/manifest.webmanifest", follow_redirects=False)
        self.assertEqual(r_man.status_code, 200)
        self.assertIn("manifest+json", r_man.content_type.lower())

        self.db.save_record("HSE-01", "PTW", "REC-HSE-QR-TEST-1", {
            "docNo": "PTW-TEST-001",
            "title": "Permit to Work QR Test",
            "status": "Open",
            "qr_token": "P1_TEST_QR_2026_OK",
        })
        import app as dcr_app
        magic_tok = dcr_app._generate_master_magic_token("PEM-058")
        r_qr200 = anon.get("/qr/P1_TEST_QR_2026_OK")
        self.assertEqual(r_qr200.status_code, 200)
        for url in ("/login", f"/magic/PEM-058?token={magic_tok}", "/qr/P1_TEST_QR_2026_OK", "/qr/UNKNOWN_QR_TOKEN_XYZ"):
            html = anon.get(url).get_data(as_text=True)
            self.assertIn("/favicon.ico", html, url)
            self.assertIn("/manifest.webmanifest", html, url)
            self.assertIn("apple-touch-icon", html, url)

        auth = self._client_as("p1_super")
        for url in ("/hub", "/?ws=ENG", "/?ws=OM", "/?ws=HSE", "/app?p=PEM-058"):
            html = auth.get(url).get_data(as_text=True)
            self.assertIn("/favicon.ico", html, url)
            self.assertIn("/manifest.webmanifest", html, url)

        tpl_html = (REPO_ROOT / "templates" / "executive_summary.html").read_text(encoding="utf-8")
        self.assertIn("favicon.ico", tpl_html)
        self.assertIn("manifest.webmanifest", tpl_html)

    # ── Feature 2: Project-scoped Related PCQ on NOC ────────────────────

    def test_02_related_pcq_scoped_to_projects_with_pcq(self):
        self.assertFalse(self.db.project_has_doc_type("PEM-058", "PCQ"))
        self.assertTrue(self.db.project_has_doc_type("EKH-001", "PCQ"))

        cfc_cols = [c["col_key"] for c in self.db.get_columns("PEM-058", "NOC")]
        ekh_cols = [c["col_key"] for c in self.db.get_columns("EKH-001", "NOC")]
        self.assertNotIn("related_pcq", cfc_cols)
        self.assertIn("related_pcq", ekh_cols)

        # Seed matching doc number in both EKH PCQ+NOC and CFC NOC (cross-project guard!)
        self.db.save_record("EKH-001", "PCQ", "REC-PCQ-EKH-1", {
            "docNo": "PCQ-MATCH-001 REV00",
            "nocRef": "NOC-SHARED-999 REV00",
            "title": "EKH Cost Quote",
        })
        self.db.save_record("EKH-001", "NOC", "REC-NOC-EKH-1", {
            "docNo": "NOC-SHARED-999 REV00",
            "title": "EKH NOC",
        })
        self.db.save_record("PEM-058", "NOC", "REC-NOC-CFC-1", {
            "docNo": "NOC-SHARED-999 REV00",
            "title": "CFC NOC",
        })

        ekh_recs = self.db.get_records("EKH-001", "NOC")
        cfc_recs = self.db.get_records("PEM-058", "NOC")
        ekh_match = next(r for r in ekh_recs if r["_id"] == "REC-NOC-EKH-1")
        cfc_match = next(r for r in cfc_recs if r["_id"] == "REC-NOC-CFC-1")
        self.assertEqual(ekh_match.get("related_pcq"), "PCQ-MATCH-001 REV00")
        self.assertNotIn("related_pcq", cfc_match)

        # Cross-check HTTP endpoints
        c = self._client_as("p1_super")
        api_cfc_cols = [x["col_key"] for x in c.get("/api/columns/PEM-058/NOC").get_json()]
        api_ekh_cols = [x["col_key"] for x in c.get("/api/columns/EKH-001/NOC").get_json()]
        self.assertNotIn("related_pcq", api_cfc_cols)
        self.assertIn("related_pcq", api_ekh_cols)

        # Rendered JS maps relatedpcq in getNocFieldProfile
        html = c.get("/app?p=EKH-001").get_data(as_text=True)
        self.assertIn("relatedpcq:profile('Calculated / System Fields'", html)

    # ── Feature 3: Nested Drive Subfolders, Recursive Viewer & Junk Filter ──

    def test_03_junk_filter_nested_subfolders_and_recursive_viewer(self):
        import drive_routing
        import blueprints.drive_uploads as du

        for junk in ("desktop.ini", "Thumbs.db", ".DS_Store", "~$CostSheet.xlsx", "scan.tmp", "old.bak", "shortcut.lnk"):
            self.assertTrue(du.is_ignored_junk_filename(junk), junk)
            clean, reason = du.sanitize_upload_filename(junk)
            self.assertEqual(clean, "")
            self.assertEqual(reason, "ignored_system_file")

        self.assertEqual(
            drive_routing.normalize_rel_subfolder(r"Take-off Documents\Steel Structure/"),
            ("Take-off Documents", "Steel Structure"),
        )
        with self.assertRaises(drive_routing.RoutingError):
            drive_routing.normalize_rel_subfolder("../escape/secret")

        rec_id = "REC-NESTED-001"
        self.db.save_record("PEM-058", "NOC", rec_id, {
            "docNo": "NOC-CY002P608-00047 REV01",
            "title": "Nested Subfolder Claim",
        })
        c = self._client_as("p1_editor")

        # Junk rejection via API
        r_junk, _ = self._upload_via_api(
            c, "PEM-058", "NOC", "part_c", "desktop.ini", b"[.ShellClassInfo]",
            subfolder_name="NOC-CY002P608-00047 - Part C REV00",
            rel_subfolder="Take-off Documents/Steel Structure",
            link_as_folder=True, replace_existing=True, record_id=rec_id,
        )
        self.assertEqual(r_junk.status_code, 400)
        self.assertEqual(r_junk.get_json()["error"], "ignored_system_file")

        # Upload 3 valid files across root, 1-level subfolder, and 2-level nested subfolder
        xlsx_data = _build_xlsx_bytes()
        r1, fid1 = self._upload_via_api(
            c, "PEM-058", "NOC", "part_c", "NOC-00047.pdf", b"%PDF-1.4\nroot",
            subfolder_name="NOC-CY002P608-00047 - Part C REV00", rel_subfolder="",
            link_as_folder=True, replace_existing=True, rev_mode="overwrite", part_rev=0, record_id=rec_id,
        )
        self.assertEqual(r1.status_code, 200, r1.get_data(as_text=True))

        r2, fid2 = self._upload_via_api(
            c, "PEM-058", "NOC", "part_c", "Cost_Breakdown.xlsx", xlsx_data,
            subfolder_name="NOC-CY002P608-00047 - Part C REV00", rel_subfolder="Final 03-01-2026",
            link_as_folder=True, replace_existing=False, rev_mode="overwrite", part_rev=0, record_id=rec_id,
        )
        self.assertEqual(r2.status_code, 200)

        r3, fid3 = self._upload_via_api(
            c, "PEM-058", "NOC", "part_c", "Steel_Qty.txt", b"Tons=42.5",
            subfolder_name="NOC-CY002P608-00047 - Part C REV00", rel_subfolder="Take-off Documents/Steel Structure",
            link_as_folder=True, replace_existing=False, rev_mode="overwrite", part_rev=0, record_id=rec_id,
        )
        self.assertEqual(r3.status_code, 200)
        att_c = r3.get_json()["attachments"]["part_c"]
        self.assertTrue(att_c["is_folder"])
        pkg_folder_id = att_c["folder_id"]

        # Verify files live inside distinct subfolder nodes on Drive, NOT flattened in pkg_folder_id
        self.assertEqual(self.sim.files[fid1]["parents"], [pkg_folder_id])
        self.assertNotEqual(self.sim.files[fid2]["parents"], [pkg_folder_id])
        self.assertNotEqual(self.sim.files[fid3]["parents"], [pkg_folder_id])
        leaf_parent = self.sim.files[fid3]["parents"][0]
        mid_parent = self.sim.folders[leaf_parent]["parents"][0]
        self.assertEqual(self.sim.folders[leaf_parent]["name"], "Steel Structure")
        self.assertEqual(self.sim.folders[mid_parent]["name"], "Take-off Documents")
        self.assertEqual(self.sim.folders[mid_parent]["parents"], [pkg_folder_id])

        # Verify viewer recursively discovers all 3 files & subfolder badges
        with patch.object(du, "_token_provider", return_value="sim-tok"), \
             patch.object(du, "_http_client", self.sim.make_http()):
            v_resp = c.get(f"/api/drive/view-info/{rec_id}/part_c")
            self.assertEqual(v_resp.status_code, 200)
            vj = v_resp.get_json()
            self.assertEqual(vj["kind"], "folder")
            by_id = {f["id"]: f for f in vj["files"]}
            self.assertEqual(set(by_id.keys()), {fid1, fid2, fid3})
            self.assertEqual(by_id[fid1]["subfolder"], "")
            self.assertEqual(by_id[fid2]["subfolder"], "Final 03-01-2026")
            self.assertEqual(by_id[fid3]["subfolder"], "Take-off Documents/Steel Structure")

            # Nested file streaming + Office preview + Folder ZIP structure
            s_nested = c.get(f"/api/drive/stream/{rec_id}/part_c/{fid3}")
            self.assertEqual(s_nested.status_code, 200)
            self.assertEqual(s_nested.get_data(), b"Tons=42.5")

            off_resp = c.get(f"/api/drive/office-preview/{rec_id}/part_c/{fid2}")
            self.assertEqual(off_resp.status_code, 200)
            self.assertEqual(off_resp.get_json()["sheets"][0]["name"], "Breakdown")

            zip_resp = c.get(f"/api/drive/folder-zip/{rec_id}/part_c")
            self.assertEqual(zip_resp.status_code, 200)
            with zipfile.ZipFile(io.BytesIO(zip_resp.get_data())) as zf:
                names = set(zf.namelist())
                self.assertIn("NOC-00047.pdf", names)
                self.assertIn("Final 03-01-2026/Cost_Breakdown.xlsx", names)
                self.assertIn("Take-off Documents/Steel Structure/Steel_Qty.txt", names)

            # Security: unrelated Drive file is rejected; viewer without download gets 403 on ZIP & ?dl=1
            other_fid = self.sim.add_file("SIM_UNRELATED_FILE", "secret.pdf", "SIM_ROOT_OM", b"secret")
            self.assertEqual(c.get(f"/api/drive/stream/{rec_id}/part_c/{other_fid}").status_code, 403)

            v_nodl = self._client_as("p1_viewer_nodl")
            self.assertEqual(v_nodl.get(f"/api/drive/stream/{rec_id}/part_c/{fid3}").status_code, 200)
            self.assertEqual(v_nodl.get(f"/api/drive/stream/{rec_id}/part_c/{fid3}?dl=1").status_code, 403)
            self.assertEqual(v_nodl.get(f"/api/drive/folder-zip/{rec_id}/part_c").status_code, 403)

            outsider = self._client_as("p1_outsider")
            self.assertEqual(outsider.get(f"/api/drive/view-info/{rec_id}/part_c").status_code, 403)

    # ── Feature 4: DS Item Ref structured rendering & [object Object] guard ─

    def test_04_ds_item_ref_normalization_and_corruption_guard(self):
        c = self._client_as("p1_editor")

        bad_str = c.post("/api/records/PEM-058/DS", json={
            "docNo": "DS-CY002P608-00901 REV00",
            "title": "Bad Serialization Test",
            "itemRef": "[object Object],[object Object]",
        })
        self.assertEqual(bad_str.status_code, 400)
        self.assertIn("[object Object]", bad_str.get_json()["error"])

        bad_nested = c.post("/api/records/PEM-058/DS", json={
            "docNo": "DS-CY002P608-00902 REV00",
            "title": "Bad Nested Object String",
            "itemRef": [{"item_ref": "[object Object]", "item_status": "Approved"}],
        })
        self.assertEqual(bad_nested.status_code, 400)

        good = c.post("/api/records/PEM-058/DS", json={
            "docNo": "DS-CY002P608-00903 REV00",
            "title": "Good Structured DS ItemRef",
            "itemRef": [
                {"item_ref": "CFC-GCL-EE-001", "item_status": "Approved"},
                {"ref": "CFC-GCL-EE-002", "status": "Revise & Resubmit"},
            ],
        })
        self.assertEqual(good.status_code, 200, good.get_data(as_text=True))
        rid = good.get_json()["id"]

        recs = c.get("/api/records/PEM-058/DS").get_json()["records"]
        saved = next(r for r in recs if r["_id"] == rid)
        self.assertEqual(saved["itemRef"], [
            {"item_ref": "CFC-GCL-EE-001", "item_status": "Approved"},
            {"item_ref": "CFC-GCL-EE-002", "item_status": "Revise & Resubmit"},
        ])

        html = self._client_as("p1_super").get("/app?p=PEM-058").get_data(as_text=True)
        self.assertIn("function parseItemRefItems(", html)
        self.assertIn("function formatItemRefPlainText(", html)
        self.assertNotIn("if(isSDOrABDTab())", html)

    # ── Feature 5: Per-Part NOC Revisions (REV00 preserved, viewer & delete) ─

    def test_05_noc_per_part_revisions_lifecycle_and_sync(self):
        import attachments
        import drive_service
        import blueprints.drive_uploads as du

        rec_id = "REC-NOC-REV-TEST"
        self.db.save_record("PEM-058", "NOC", rec_id, {
            "docNo": "NOC-CY002P608-00088 REV01",
            "title": "Per-Part Rev Test",
        })
        c = self._client_as("p1_editor")

        # 1. Upload Part C REV00 folder
        r_r0, fid_r0 = self._upload_via_api(
            c, "PEM-058", "NOC", "part_c", "Cost_Rev0.pdf", b"%PDF-1.4\nrev0",
            subfolder_name="NOC-CY002P608-00088 - Part C REV00",
            link_as_folder=True, replace_existing=True, rev_mode="overwrite", part_rev=0, record_id=rec_id,
        )
        self.assertEqual(r_r0.status_code, 200, r_r0.get_data(as_text=True))
        att0 = r_r0.get_json()["attachments"]["part_c"]
        self.assertEqual(att0["rev"], 0)
        self.assertEqual(att0["rev_label"], "REV00")
        folder_r0 = att0["folder_id"]

        # 2. Upload Part C REV01 as new_revision -> REV00 archived & NOT trashed!
        r_r1, fid_r1 = self._upload_via_api(
            c, "PEM-058", "NOC", "part_c", "Cost_Rev1.pdf", b"%PDF-1.4\nrev1",
            subfolder_name="NOC-CY002P608-00088 - Part C REV01",
            link_as_folder=True, replace_existing=True, rev_mode="new_revision", part_rev=1, record_id=rec_id,
        )
        self.assertEqual(r_r1.status_code, 200, r_r1.get_data(as_text=True))
        att1 = r_r1.get_json()["attachments"]["part_c"]
        self.assertEqual(att1["rev"], 1)
        self.assertEqual(att1["rev_label"], "REV01")
        folder_r1 = att1["folder_id"]
        self.assertNotEqual(folder_r0, folder_r1)
        self.assertEqual(len(att1["history"]), 1)
        self.assertEqual(att1["history"][0]["rev"], 0)
        self.assertNotIn(folder_r0, self.sim.trashed_ids)
        self.assertNotIn(fid_r0, self.sim.trashed_ids)

        # 3. View-info & stream across REV01 (default) and ?rev=0 (archived)
        with patch.object(du, "_token_provider", return_value="sim-tok"), \
             patch.object(du, "_http_client", self.sim.make_http()):
            vi1 = c.get(f"/api/drive/view-info/{rec_id}/part_c").get_json()
            self.assertEqual(vi1["rev"], 1)
            self.assertEqual(vi1["rev_label"], "REV01")
            self.assertTrue(vi1["is_current_rev"])
            self.assertEqual([x["rev"] for x in vi1["revisions"]], [1, 0])
            self.assertEqual([f["id"] for f in vi1["files"]], [fid_r1])

            vi0 = c.get(f"/api/drive/view-info/{rec_id}/part_c?rev=0").get_json()
            self.assertEqual(vi0["rev"], 0)
            self.assertEqual(vi0["rev_label"], "REV00")
            self.assertFalse(vi0["is_current_rev"])
            self.assertEqual([f["id"] for f in vi0["files"]], [fid_r0])

            self.assertEqual(c.get(f"/api/drive/stream/{rec_id}/part_c/{fid_r0}?rev=0").get_data(), b"%PDF-1.4\nrev0")
            # Cross-revision file leak prevention
            self.assertEqual(c.get(f"/api/drive/stream/{rec_id}/part_c/{fid_r1}?rev=0").status_code, 403)
            self.assertEqual(c.get(f"/api/drive/folder-zip/{rec_id}/part_c?rev=0").status_code, 200)
            self.assertEqual(self._client_as("p1_viewer_nodl").get(f"/api/drive/folder-zip/{rec_id}/part_c?rev=0").status_code, 403)

            # 4. Deleting active REV01 trashes REV01 & promotes REV00 back to active without trashing REV00
            d_active = c.delete(f"/api/drive/attachment/{rec_id}/part_c")
            self.assertEqual(d_active.status_code, 200)
            promoted = d_active.get_json()["attachments"]["part_c"]
            self.assertEqual(promoted["rev"], 0)
            self.assertEqual(promoted["folder_id"], folder_r0)
            self.assertIn(folder_r1, self.sim.trashed_ids)
            self.assertNotIn(folder_r0, self.sim.trashed_ids)

            # 5. Re-upload REV01 and delete only archived ?rev=0 -> trashes REV00 while keeping REV01
            r_r1b, _ = self._upload_via_api(
                c, "PEM-058", "NOC", "part_c", "Cost_Rev1_v2.pdf", b"%PDF-1.4\nrev1v2",
                subfolder_name="NOC-CY002P608-00088 - Part C REV01",
                link_as_folder=True, replace_existing=True, rev_mode="new_revision", part_rev=1, record_id=rec_id,
            )
            self.assertEqual(r_r1b.status_code, 200)
            d_arch = c.delete(f"/api/drive/attachment/{rec_id}/part_c?rev=0")
            self.assertEqual(d_arch.status_code, 200)
            rem_c = d_arch.get_json()["attachments"]["part_c"]
            self.assertEqual(rem_c["rev"], 1)
            self.assertNotIn("history", rem_c)
            self.assertIn(folder_r0, self.sim.trashed_ids)

        # 6. drive_service.process_drive_folder revision-aware NOC sync test
        sync_root = self.sim.add_folder("F_SYNC_ROOT", "SYNC ROOT", "")
        sync_noc = self.sim.add_folder("F_SYNC_NOC", "NOC", sync_root)
        sync_pa = self.sim.add_folder("F_SYNC_PA", "Part A", sync_noc)
        self.sim.add_file("SF_R0", "NOC-CY002P608-00555 Part A REV00.pdf", sync_pa, b"r0")
        self.sim.add_file("SF_R1", "NOC-CY002P608-00555 Part A REV01.pdf", sync_pa, b"r1")
        if not self.db.get_project("SYNC-P1"):
            self.db.create_project("SYNC-P1", "Sync Test", "SYNC-P1", workspace_id="ENG")
        self.db.update_project_settings("SYNC-P1", {"drive_folder_id": sync_root})
        self.db.save_record("SYNC-P1", "NOC", "REC-SYNC-555", {
            "docNo": "NOC-CY002P608-00555",
            "title": "Sync Multi-Rev NOC",
        })
        with patch.object(drive_service, "get_drive_service", return_value=self.sim.make_drive_v3_service()):
            drive_service.process_drive_folder(sync_root)
        synced = self.db.get_record_by_id("REC-SYNC-555")
        pa = (synced.get("attachments") or {}).get("part_a") or {}
        self.assertEqual(pa.get("rev"), 1)
        self.assertIn("SF_R1", pa.get("url", ""))
        self.assertEqual(len(pa.get("history") or []), 1)
        self.assertEqual(pa["history"][0]["rev"], 0)
        self.assertIn("SF_R0", pa["history"][0]["url"])


if __name__ == "__main__":
    unittest.main()
