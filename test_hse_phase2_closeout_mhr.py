"""HSE-02 tests: Close-Out attachments (SLOT_APPROVAL), MHR duplicate month guard, and ENG/OM non-regression."""

from contextlib import ExitStack
import os
from pathlib import Path
import sys
import unittest
from unittest.mock import patch


def _fake_http_and_token():
    import blueprints.drive_uploads as du

    files_db = {}
    folders_db = {
        "ROOT_HSE": {"id": "ROOT_HSE", "name": "HSE Root", "mimeType": "application/vnd.google-apps.folder", "trashed": False},
        "ROOT_OM": {"id": "ROOT_OM", "name": "OM Root", "mimeType": "application/vnd.google-apps.folder", "trashed": False},
        "ROOT_ENG": {"id": "ROOT_ENG", "name": "ENG Root", "mimeType": "application/vnd.google-apps.folder", "trashed": False},
    }
    seq = {"n": 0}

    class _Resp:
        def __init__(self, status_code, payload=None, headers=None):
            self.status_code = status_code
            self._payload = payload or {}
            self.headers = headers or {}
            self.text = ""

        def json(self):
            return self._payload

    class _FakeHTTP:
        def get(self, url, params=None, headers=None, timeout=None):
            import re
            if "/drive/v3/files/" in url and "?alt=media" not in url:
                fid = url.rsplit("/", 1)[-1].split("?", 1)[0]
                if fid in folders_db:
                    return _Resp(200, folders_db[fid])
                if fid in files_db:
                    return _Resp(200, files_db[fid])
                return _Resp(404, {"error": {"message": "Not found"}})
            q = (params or {}).get("q", "")
            m = re.search(r"'([^']+)'\s+in\s+parents", q)
            parent_id = m.group(1) if m else ""
            children = [
                f for f in folders_db.values()
                if parent_id in (f.get("parents") or []) and not f.get("trashed")
            ]
            return _Resp(200, {"files": children})

        def post(self, url, params=None, json=None, headers=None, data=None, timeout=None):
            seq["n"] += 1
            if "upload/drive/v3/files" in url:
                meta = json or {}
                parent = (meta.get("parents") or [""])[0]
                fid = f"FAKE_FILE_{seq['n']}"
                files_db[fid] = {
                    "id": fid,
                    "name": meta.get("name") or "file.pdf",
                    "size": str(headers.get("X-Upload-Content-Length") or "128"),
                    "mimeType": meta.get("mimeType") or "application/pdf",
                    "parents": [parent],
                    "trashed": False,
                    "webViewLink": f"https://drive.google.com/file/d/{fid}/view",
                }
                return _Resp(200, {}, headers={"Location": f"https://sim.local/upload/{fid}"})
            meta = json or {}
            fid = f"FAKE_FLD_{seq['n']}"
            folders_db[fid] = {
                "id": fid,
                "name": meta.get("name") or "Folder",
                "mimeType": meta.get("mimeType") or "application/vnd.google-apps.folder",
                "parents": meta.get("parents") or [],
                "trashed": False,
            }
            return _Resp(200, folders_db[fid])

        def patch(self, url, params=None, json=None, headers=None, timeout=None):
            return _Resp(200, {})

    return du, _FakeHTTP(), files_db


class HsePhase2CloseoutAndMhrTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        if os.environ.get("DCR_TEST_DB_PREINJECTED") == "1":
            import db
            cls.db = db
            return
        from security_test_support import disposable_database

        stack = ExitStack()
        cls.addClassCleanup(stack.close)
        cls.db = stack.enter_context(disposable_database())

    def test_01_hse_approval_slot_matrix_and_no_engineering_codes(self):
        import attachments
        from app import app

        db = self.db
        du, fake_http, files_db = _fake_http_and_token()

        hse_pid = "SYN-HSE-02-ATT"
        om_pid = "SYN-OM-02-ATT"
        eng_pid = "SYN-ENG-02-ATT"
        db.create_project(hse_pid, hse_pid, hse_pid, workspace_id="HSE")
        db.update_project_settings(hse_pid, {"drive_folder_id": "ROOT_HSE"})
        db.create_project(om_pid, om_pid, om_pid, workspace_id="OM")
        db.update_project_settings(om_pid, {"drive_folder_id": "ROOT_OM"})
        db.create_project(eng_pid, eng_pid, eng_pid, workspace_id="ENG")
        db.update_project_settings(eng_pid, {"drive_folder_id": "ROOT_ENG"})

        db.add_user("syn_hse02_sa", "synthetic-hse-02", role="superadmin", primary_workspace="HSE")
        client = app.test_client()
        client.set_cookie("dcr_token", db.create_session("syn_hse02_sa", "superadmin"))

        with patch.object(du, "_token_provider", return_value="FAKE_TOKEN"), patch.object(du, "_http_client", fake_http):
            # 1. Close-out HSE doc types permit SLOT_APPROVAL; others reject it
            for code in ("PTW", "OBS", "SVP", "INC", "JSA"):
                r = client.post("/api/drive/upload-session", json={
                    "project_id": hse_pid, "dt_id": code, "slot": "approval",
                    "file_name": f"{code}_closeout.pdf", "file_size": 256,
                })
                self.assertEqual(r.status_code, 200, (code, r.get_json()))
                self.assertTrue(r.get_json()["ok"])

            for code in ("TBT", "FEI", "CEV", "EVAC", "MHR", "WKR"):
                r = client.post("/api/drive/upload-session", json={
                    "project_id": hse_pid, "dt_id": code, "slot": "approval",
                    "file_name": f"{code}_closeout.pdf", "file_size": 256,
                })
                self.assertEqual(r.status_code, 400, (code, r.get_json()))

            # 2. NOC_SLOTS remain strictly restricted to ENG
            for slot in attachments.NOC_SLOTS:
                r = client.post("/api/drive/upload-session", json={
                    "project_id": hse_pid, "dt_id": "PTW", "slot": slot,
                    "file_name": "noc.pdf", "file_size": 128,
                })
                self.assertEqual(r.status_code, 400, (slot, r.get_json()))

            # 3. OM INC still rejects SLOT_APPROVAL (no regression to OM restriction)
            r_om = client.post("/api/drive/upload-session", json={
                "project_id": om_pid, "dt_id": "INC", "slot": "approval",
                "file_name": "om_inc.pdf", "file_size": 128,
            })
            self.assertEqual(r_om.status_code, 400)

            # 4. Upload completion for HSE PTW (Open -> Closed) and OBS (Open -> Closed)
            for dt_id, start_status, expected_end in (("PTW", "Open", "Closed"), ("OBS", "Open", "Closed")):
                created = client.post(f"/api/records/{hse_pid}/{dt_id}", json={
                    "docNo": f"HSE-{dt_id}-001", "title": "Close-out test", "status": start_status,
                })
                self.assertEqual(created.status_code, 200)
                rid = created.get_json()["id"]

                sess = client.post("/api/drive/upload-session", json={
                    "project_id": hse_pid, "dt_id": dt_id, "record_id": rid,
                    "slot": "approval", "file_name": f"{dt_id}_closed.pdf", "file_size": 512,
                })
                self.assertEqual(sess.status_code, 200)
                sdata = sess.get_json()
                fid = sdata["session_url"].rsplit("/", 1)[-1]
                files_db[fid] = {
                    "id": fid, "name": f"{dt_id}_closed.pdf", "size": "512",
                    "mimeType": "application/pdf", "parents": [sdata["folder_id"]],
                    "trashed": False, "webViewLink": f"https://drive.google.com/file/d/{fid}/view",
                }
                comp = client.post("/api/drive/upload-complete", json={
                    "upload_ticket": sdata["upload_ticket"], "file_id": fid, "record_id": rid,
                })
                self.assertEqual(comp.status_code, 200, comp.get_json())
                cdata = comp.get_json()
                att = cdata["attachment"]
                self.assertNotIn("consultant_status", att)
                self.assertEqual(att.get("closeout_status"), expected_end)
                self.assertEqual(cdata.get("status"), expected_end)
                stored = db.get_record_by_id(rid)
                self.assertEqual(stored.get("status"), expected_end)
                self.assertFalse(any(str(stored.get("status", "")).startswith(p) for p in ("A -", "B -", "C -", "D -")))

            # Legacy 'مغلق' / 'مفتوح' values also normalize cleanly to English 'Closed'
            legacy_out = attachments.resolve_hse_closeout_status("مفتوح")
            self.assertEqual(legacy_out, ("Closed", "Closed"))
            self.assertEqual(attachments.resolve_hse_closeout_status("مغلق"), ("Closed", "Closed"))

            # 5. ENG DS SLOT_APPROVAL preserves consultant_status and NOC_SLOTS still work in ENG
            eng_rec = client.post(f"/api/records/{eng_pid}/DS", json={
                "docNo": "ENG-DS-001", "title": "ENG submittal", "status": "A - Approved",
            })
            self.assertEqual(eng_rec.status_code, 200)
            eng_rid = eng_rec.get_json()["id"]
            eng_sess = client.post("/api/drive/upload-session", json={
                "project_id": eng_pid, "dt_id": "DS", "record_id": eng_rid,
                "slot": "approval", "file_name": "ds_reply.pdf", "file_size": 256,
            })
            self.assertEqual(eng_sess.status_code, 200)
            esdata = eng_sess.get_json()
            efid = esdata["session_url"].rsplit("/", 1)[-1]
            files_db[efid] = {
                "id": efid, "name": "ds_reply.pdf", "size": "256",
                "mimeType": "application/pdf", "parents": [esdata["folder_id"]],
                "trashed": False, "webViewLink": f"https://drive.google.com/file/d/{efid}/view",
            }
            eng_comp = client.post("/api/drive/upload-complete", json={
                "upload_ticket": esdata["upload_ticket"], "file_id": efid, "record_id": eng_rid,
            })
            self.assertEqual(eng_comp.status_code, 200)
            self.assertEqual(eng_comp.get_json()["attachment"].get("consultant_status"), "A - Approved")

    def test_02_mhr_duplicate_month_guard_and_updates(self):
        from app import app

        db = self.db
        p1 = "SYN-HSE-02-MHR-1"
        p2 = "SYN-HSE-02-MHR-2"
        db.create_project(p1, p1, p1, workspace_id="HSE")
        db.create_project(p2, p2, p2, workspace_id="HSE")
        db.add_user("syn_hse02_ed", "synthetic-hse-02", role="editor", primary_workspace="HSE")
        db.assign_user_workspace("syn_hse02_ed", "HSE", "editor")
        db.assign_project("syn_hse02_ed", p1)
        db.assign_project("syn_hse02_ed", p2)

        client = app.test_client()
        client.set_cookie("dcr_token", db.create_session("syn_hse02_ed", "editor"))

        # First MHR log for 2026-11 succeeds
        r1 = client.post(f"/api/records/{p1}/MHR", json={
            "docNo": "MHR-2026-11", "year": 2026, "month": 11, "hours": 1400,
        })
        self.assertEqual(r1.status_code, 200, r1.get_json())
        rec1_id = r1.get_json()["id"]

        # Duplicate in same project & month (including padded string '03' vs 3 after insert) is blocked
        r_dup = client.post(f"/api/records/{p1}/MHR", json={
            "docNo": "MHR-2026-11-DUP", "year": "2026", "month": "11", "hours": 500,
        })
        self.assertEqual(r_dup.status_code, 409, r_dup.get_json())
        self.assertEqual(r_dup.get_json().get("code"), "DUPLICATE_MHR_PERIOD")
        self.assertEqual(db.count_records(p1, "MHR"), 1)

        # Zero-padded month collision check ('03' then 3)
        r_mar = client.post(f"/api/records/{p1}/MHR", json={
            "docNo": "MHR-2026-03", "year": "2026", "month": "03", "hours": 900,
        })
        self.assertEqual(r_mar.status_code, 200, r_mar.get_json())
        rec_mar_id = r_mar.get_json()["id"]

        r_mar_dup = client.post(f"/api/records/{p1}/MHR", json={
            "docNo": "MHR-2026-3-DUP", "year": 2026, "month": 3, "hours": 950,
        })
        self.assertEqual(r_mar_dup.status_code, 409, r_mar_dup.get_json())

        # Same month in another HSE project succeeds
        r_p2 = client.post(f"/api/records/{p2}/MHR", json={
            "docNo": "MHR-P2-2026-11", "year": 2026, "month": 11, "hours": 2100,
        })
        self.assertEqual(r_p2.status_code, 200, r_p2.get_json())

        # Editing the existing MHR record for the same month succeeds
        r_edit = client.post(f"/api/records/{p1}/MHR", json={
            "_id": rec1_id, "docNo": "MHR-2026-11", "year": 2026, "month": 11, "hours": 1550,
        })
        self.assertEqual(r_edit.status_code, 200, r_edit.get_json())
        self.assertEqual(db.get_record_by_id(rec1_id).get("hours"), 1550)

        # Editing an existing MHR record to collide with another existing month fails (409)
        r_edit_dup = client.post(f"/api/records/{p1}/MHR", json={
            "_id": rec_mar_id, "docNo": "MHR-2026-03", "year": 2026, "month": 11, "hours": 900,
        })
        self.assertEqual(r_edit_dup.status_code, 409, r_edit_dup.get_json())

        # Invalid/missing period checks
        for bad in ({"docNo": "MHR-BAD-1", "year": 2026}, {"docNo": "MHR-BAD-2", "year": 2026, "month": 13}):
            self.assertEqual(client.post(f"/api/records/{p1}/MHR", json=bad).status_code, 400)

        # Direct db.save_record enforcement
        with self.assertRaises(db.DuplicateMhrPeriodError):
            db.save_record(p1, "MHR", "DIRECT-DUP-MHR", {"docNo": "MHR-DIR", "year": 2026, "month": 11, "hours": 100})

    def test_03_status_meta_green_badge_kpi_and_hse_admin_scope(self):
        import re
        from app import app
        from utils import STATUS_COLORS

        db = self.db
        self.assertEqual(db.resolve_status_meta("A - Approved"), "approved")
        self.assertEqual(db.resolve_status_meta("Open"), "pending")
        self.assertEqual(db.resolve_status_meta("Closed"), "approved")
        self.assertEqual(STATUS_COLORS["Closed"], ("bbf7d0", "166534"))

        hse_p = "SYN-HSE-02-KPI"
        eng_p = "SYN-ENG-02-SCOPE"
        om_p = "SYN-OM-02-SCOPE"
        db.create_project(hse_p, "HSE KPI Site", "HSE-KPI", workspace_id="HSE")
        db.create_project(eng_p, "ENG Scope Site", "ENG-SCP", workspace_id="ENG")
        db.create_project(om_p, "OM Scope Site", "OM-SCP", workspace_id="OM")

        # All 11 HSE doc types include fileLocation Drive Link column
        hse_dts = db.get_doc_types(hse_p)
        self.assertEqual(len(hse_dts), 11)
        for dt in hse_dts:
            cols = {c["col_key"]: c for c in db.get_columns(hse_p, dt["id"], visible_only=True)}
            self.assertIn("fileLocation", cols, dt["id"])
            self.assertEqual(cols["fileLocation"]["col_type"], "url")

        db.save_record(hse_p, "PTW", "kpi-open-1", {"docNo": "PTW-KPI-1", "title": "Open permit", "status": "Open"})
        db.save_record(hse_p, "OBS", "kpi-closed-1", {"docNo": "OBS-KPI-1", "title": "Closed obs", "status": "Closed"})

        # Scoped hse_admin (role='admin', HSE only) vs super_admin
        db.add_user("syn_hse02_mgr", "synthetic-hse-02", role="admin", primary_workspace="HSE")
        db.assign_project("syn_hse02_mgr", hse_p)
        db.add_user("syn_hse02_sup", "synthetic-hse-02", role="superadmin", primary_workspace="ENG")

        c_mgr = app.test_client()
        c_mgr.set_cookie("dcr_token", db.create_session("syn_hse02_mgr", "admin"))
        self.assertEqual(c_mgr.get(f"/app?p={hse_p}").status_code, 200)
        self.assertEqual(c_mgr.get(f"/app?p={eng_p}").status_code, 403)
        self.assertEqual(c_mgr.get(f"/app?p={om_p}").status_code, 403)
        self.assertEqual(c_mgr.get("/api/dashboard_stats?ws=ENG").get_json(), [])
        self.assertEqual(c_mgr.get("/api/dashboard_stats?ws=OM").get_json(), [])

        hse_dash = c_mgr.get("/?ws=HSE")
        self.assertEqual(hse_dash.status_code, 200)
        html = hse_dash.get_data(as_text=True)
        m_comp = re.search(r'Completed / Closed</span>\s*<span class="om-kpi-val success">(\d+)</span>', html)
        self.assertIsNotNone(m_comp)
        self.assertGreaterEqual(int(m_comp.group(1)), 1)

        reg_html = c_mgr.get(f"/app?p={hse_p}").get_data(as_text=True)
        self.assertIn("1. Submission", reg_html)
        self.assertIn("2. Close-Out", reg_html)
        self.assertIn("'Closed':['bbf7d0','166534']", reg_html)

        c_sup = app.test_client()
        c_sup.set_cookie("dcr_token", db.create_session("syn_hse02_sup", "superadmin"))
        for p in (hse_p, eng_p, om_p):
            self.assertEqual(c_sup.get(f"/app?p={p}").status_code, 200)


if __name__ == "__main__":
    sys.path.insert(0, str(Path(__file__).resolve().parent))
    import security_test_support as runner

    module = Path(__file__).stem
    runner.SUITES = (module,)
    sys.argv.insert(1, module)
    raise SystemExit(runner.main())
