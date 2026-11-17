"""Phase 4 tests: 100% English standardization, Interactive INC Body Map, and read-only legacy migration."""

from contextlib import ExitStack
from datetime import date, timedelta
import io
import json
import os
from pathlib import Path
import re
import sqlite3
import tempfile
import unittest
from unittest.mock import patch


_ARABIC_RE = re.compile(r"[\u0600-\u06FF]")


def _collect_strings(obj):
    if isinstance(obj, str):
        yield obj
    elif isinstance(obj, dict):
        for k, v in obj.items():
            yield from _collect_strings(k)
            yield from _collect_strings(v)
    elif isinstance(obj, (list, tuple, set)):
        for item in obj:
            yield from _collect_strings(item)


class HsePhase4EnglishBodyMapMigrationTests(unittest.TestCase):
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

    def test_01_zero_arabic_in_hse_defaults_modals_and_qr_verify(self):
        import hse_defaults

        # 1. HSE defaults structures contain ZERO Arabic characters anywhere (keys or values)
        for name, obj in (
            ("HSE_REGISTERS", hse_defaults.HSE_REGISTERS),
            ("HSE_SUBTYPES", hse_defaults.HSE_SUBTYPES),
            ("HSE_LISTS", hse_defaults.HSE_LISTS),
            ("HSE_COLUMNS", hse_defaults.HSE_COLUMNS),
        ):
            hits = [s for s in _collect_strings(obj) if _ARABIC_RE.search(s)]
            self.assertEqual(hits, [], f"Arabic text remained in {name}: {hits}")

        for code, _nm, _src in hse_defaults.HSE_REGISTERS:
            cols = hse_defaults.hse_columns(code)
            hits = [s for s in _collect_strings(cols) if _ARABIC_RE.search(s)]
            self.assertEqual(hits, [], f"Arabic text in hse_columns({code}): {hits}")

        self.assertEqual(
            hse_defaults.HSE_LISTS["hse_extinguisher_status"],
            ("Valid", "Needs Maintenance", "Expired", "Under Refill", "Out of Service"),
        )
        self.assertEqual(
            hse_defaults.HSE_LISTS["hse_extinguisher_type"],
            ("CO2", "Dry Powder", "Foam", "Water", "Wet Chemical", "Fire Hose Reel"),
        )
        inc_keys = [c[0] for c in hse_defaults.HSE_COLUMNS["INC"]]
        self.assertIn("bodyPart", inc_keys)

        # 2. Rendered HSE Dashboard & Batch Inspect / QR modals contain strictly English UI options
        db = self.db
        hse_pid = "SYN-HSE-P4-EN"
        db.create_project(hse_pid, "HSE English Site", "HSE-P4", workspace_id="HSE")
        db.add_user("syn_p4_sa", "SynPass!123", role="superadmin", primary_workspace="HSE")
        client = self.app.test_client()
        client.set_cookie("dcr_token", db.create_session("syn_p4_sa", "superadmin"))

        dash_html = client.get("/?ws=HSE").get_data(as_text=True)
        ov_start = dash_html.find('id="tab-hse-overview"')
        ov_end = dash_html.find('id="hse-ai-fab"', ov_start)
        if ov_start > 0 and ov_end > ov_start:
            self.assertIsNone(_ARABIC_RE.search(dash_html[ov_start:ov_end]))

        reg_html = client.get(f"/app?p={hse_pid}").get_data(as_text=True)
        bi_start = reg_html.find('id="fei-batch-inspect-modal"')
        bi_end = reg_html.find('id="doc-viewer-modal"', bi_start)
        self.assertTrue(0 < bi_start < bi_end)
        batch_modal_html = reg_html[bi_start:bi_end]
        self.assertIsNone(_ARABIC_RE.search(batch_modal_html))
        for st_opt in ("Valid", "Needs Maintenance", "Expired", "Under Refill", "Out of Service"):
            self.assertIn(f'<option value="{st_opt}">{st_opt}</option>', batch_modal_html)

        # 3. Batch inspect endpoint enforces English options + silently converts legacy values on /qr/<token>
        fei_legacy = "p4-fei-legacy"
        fei_ok = "p4-fei-english"
        ref = date.today()
        db.save_record(
            hse_pid,
            "FEI",
            fei_legacy,
            {
                "docNo": "FEI-P4-LEG",
                "extinguisherType": "بودرة جافة",
                "extinguisherStatus": "صيانة",
                "location": "Workshop 1",
                "expiryDate": (ref + timedelta(days=90)).isoformat(),
            },
        )
        db.save_record(
            hse_pid,
            "FEI",
            fei_ok,
            {
                "docNo": "FEI-P4-EN",
                "extinguisherType": "CO2",
                "extinguisherStatus": "Valid",
                "location": "Main Panel",
                "expiryDate": (ref + timedelta(days=90)).isoformat(),
            },
        )

        qr_leg_tok = client.get(f"/api/hse/qr/{fei_legacy}").get_json()["qr_token"]
        qr_leg_html = self.app.test_client().get(f"/qr/{qr_leg_tok}").get_data(as_text=True)
        self.assertIsNone(_ARABIC_RE.search(qr_leg_html))
        self.assertIn("Needs Maintenance", qr_leg_html)
        self.assertIn("Dry Powder", qr_leg_html)

        # Reject invalid batch status cleanly
        r_bad = client.post(
            "/api/hse/fei/batch-inspect",
            json={"record_ids": [fei_ok], "last_inspection": ref.isoformat(), "status": "NotAValidStatus"},
        )
        self.assertEqual(r_bad.status_code, 400)

        # Omit status -> defaults strictly to English "Valid"
        r_def = client.post(
            "/api/hse/fei/batch-inspect",
            json={"record_ids": [fei_ok], "last_inspection": ref.isoformat()},
        )
        self.assertEqual(r_def.status_code, 200)
        self.assertEqual(db.get_record_by_id(fei_ok)["status"], "Valid")

        # Update to English "Needs Maintenance" -> reflected cleanly on QR verification page
        r_maint = client.post(
            "/api/hse/fei/batch-inspect",
            json={
                "record_ids": [fei_ok],
                "last_inspection": ref.isoformat(),
                "expiry_date": (ref + timedelta(days=60)).isoformat(),
                "status": "Needs Maintenance",
                "notes": "Pressure gauge inspection scheduled",
            },
        )
        self.assertEqual(r_maint.status_code, 200)
        self.assertEqual(db.get_record_by_id(fei_ok)["extinguisherStatus"], "Needs Maintenance")
        qr_ok_tok = client.get(f"/api/hse/qr/{fei_ok}").get_json()["qr_token"]
        qr_ok_html = self.app.test_client().get(f"/qr/{qr_ok_tok}").get_data(as_text=True)
        self.assertIn("Needs Maintenance", qr_ok_html)

    def test_02_inc_interactive_body_map_ui_and_persistence(self):
        db = self.db
        hse_pid = "SYN-HSE-P4-INC"
        eng_pid = "SYN-ENG-P4-NR"
        db.create_project(hse_pid, "HSE Body Map Plant", "HSE-BM", workspace_id="HSE")
        db.create_project(eng_pid, "ENG Non-Regression", "ENG-BM", workspace_id="ENG")
        db.add_user("syn_p4_ed", "SynPass!123", role="editor", primary_workspace="HSE")
        db.assign_project("syn_p4_ed", hse_pid)
        db.add_user("syn_p4_eng", "SynPass!123", role="superadmin", primary_workspace="ENG")

        c_ed = self.app.test_client()
        c_ed.set_cookie("dcr_token", db.create_session("syn_p4_ed", "editor"))
        reg_html = c_ed.get(f"/app?p={hse_pid}").get_data(as_text=True)

        # Body map SVG, Front & Back views, and all required anatomical regions exist cleanly in INC script
        for marker in (
            "isHseIncTab",
            "INC_BODY_REGIONS",
            "mountIncBodyMapSection",
            "selectIncBodyRegion",
            "syncIncBodyMapFromField",
            "inc-body-map-svg",
            "FRONT VIEW",
            "BACK VIEW",
            "Head / Face",
            "Neck",
            "Chest",
            "Abdomen",
            "Upper Back",
            "Lower Back",
            "Left Shoulder",
            "Right Shoulder",
            "Left Arm",
            "Right Arm",
            "Left Hand",
            "Right Hand",
            "Pelvis / Hip",
            "Left Thigh",
            "Right Thigh",
            "Left Knee",
            "Right Knee",
            "Left Leg",
            "Right Leg",
            "Left Foot",
            "Right Foot",
        ):
            self.assertIn(marker, reg_html)

        # Save INC record with bodyPart, bodyRegion, and bodyCoords -> persists in records.data JSONB
        created = c_ed.post(
            f"/api/records/{hse_pid}/INC",
            json={
                "docNo": "INC-P4-001",
                "title": "Minor laceration during valve overhaul",
                "issuedDate": date.today().isoformat(),
                "injuryClass": "FAC",
                "bodyPart": "Left Knee",
                "bodyRegion": "left_knee",
                "bodyCoords": "front:128,246",
                "lostDays": 0,
            },
        )
        self.assertEqual(created.status_code, 200, created.get_json())
        rid = created.get_json()["id"]
        stored = db.get_record_by_id(rid)
        self.assertEqual(stored["bodyPart"], "Left Knee")
        self.assertEqual(stored["bodyRegion"], "left_knee")
        self.assertEqual(stored["bodyCoords"], "front:128,246")

        listed = c_ed.get(f"/api/records/{hse_pid}/INC").get_json()["records"]
        rec = next(r for r in listed if r["_id"] == rid)
        self.assertEqual(rec["bodyPart"], "Left Knee")

        # ENG workspace continues rendering normally with no HSE dashboard leakage
        c_sup = self.app.test_client()
        c_sup.set_cookie("dcr_token", db.create_session("syn_p4_eng", "superadmin"))
        self.assertNotIn("hseTrendChart", c_sup.get("/?ws=ENG").get_data(as_text=True))
        self.assertNotIn("hseTrendChart", c_sup.get("/?ws=OM").get_data(as_text=True))

    def test_03_hse_migration_localhost_guard_pii_filter_and_foram_junk(self):
        import hse_migration as hm

        # 1. Strict target DB localhost-only guards
        for ok_url in (
            "postgresql://u:p@127.0.0.1:5432/dcr",
            "postgresql://u:p@localhost:5432/dcr_local",
            "postgres://u:p@[::1]:5432/dcr",
        ):
            self.assertEqual(hm.validate_local_postgres_url(ok_url), ok_url)

        for bad_url in (
            "",
            "sqlite:///tmp/local.db",
            "postgresql://u:p@postgres.railway.internal:5432/railway",
            "postgresql://u:p@roundhouse.proxy.rlwy.net:12345/railway",
            "postgresql://u:p@db.prod.example.com:5432/dcr",
            "postgresql://u:p@127.0.0.1:5432,remote.example.com:5432/dcr",
        ):
            with self.assertRaises(ValueError, msg=bad_url):
                hm.validate_local_postgres_url(bad_url)

        # 2. Isolated synthetic SQLite + FORAM directory dry-run & idempotent local DB import
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp_path = Path(tmp_dir)
            syn_db = tmp_path / "syn_safetyqr.db"
            foram_dir = tmp_path / "FORAM"
            foram_dir.mkdir()
            (foram_dir / "Hot Work Permit.docx").write_bytes(b"docx-stub")
            (foram_dir / "~$Hot Work Permit.docx").write_bytes(b"lock-stub")
            (foram_dir / "Working at Height.LNK").write_bytes(b"lnk-stub")

            conn = sqlite3.connect(syn_db)
            conn.executescript(
                """
                CREATE TABLE projects (id INTEGER PRIMARY KEY, name TEXT, location TEXT, start_date TEXT, project_type TEXT);
                CREATE TABLE docs (id INTEGER PRIMARY KEY, category TEXT, number TEXT, title TEXT, doc_date TEXT,
                                   valid_until TEXT, location TEXT, responsible TEXT, status TEXT, notes TEXT,
                                   injury_type TEXT, lost_days INTEGER, closed_at TEXT, project_id INTEGER);
                CREATE TABLE injuries (id INTEGER PRIMARY KEY, doc_id INTEGER, project_id INTEGER, incident_date TEXT,
                                       incident_time TEXT, injury_category TEXT, body_part TEXT, lost_days INTEGER,
                                       location TEXT, incident_description TEXT, root_cause TEXT,
                                       corrective_action TEXT, status TEXT);
                CREATE TABLE items (id INTEGER PRIMARY KEY, type TEXT, code TEXT, title TEXT, data TEXT, project_id INTEGER);
                CREATE TABLE records (id INTEGER PRIMARY KEY, item_id INTEGER, kind TEXT, title TEXT, record_date TEXT, expiry_date TEXT, notes TEXT);
                CREATE TABLE risk_assessments (id INTEGER PRIMARY KEY, project_id INTEGER, code TEXT, title TEXT, activity_name TEXT,
                                                 department TEXT, location TEXT, assessment_date TEXT, review_date TEXT,
                                                 assessor_name TEXT, approver_name TEXT, status TEXT, notes TEXT);
                CREATE TABLE evacuation_drills (id INTEGER PRIMARY KEY, project_id INTEGER, title TEXT, drill_date TEXT, drill_time TEXT,
                                                  scenario TEXT, location TEXT, assembly_point TEXT, leader_name TEXT,
                                                  target_duration_sec INTEGER, actual_duration_sec INTEGER, evaluation_grade TEXT,
                                                  alarm_working TEXT, exit_routes_clear TEXT, assembly_adequate TEXT,
                                                  corrective_actions TEXT, notes TEXT);
                CREATE TABLE man_hours (project_id INTEGER, year INTEGER, month INTEGER, hours REAL);
                """
            )
            conn.execute("INSERT INTO projects VALUES (1, 'Plant Alpha', 'Cairo', '2026-01-01', 'general')")
            conn.execute(
                "INSERT INTO docs VALUES (10, 'hot_permit', 'PTW-S-01', 'Welding header', '2026-11-01', '2026-11-30', 'Bay 1', 'HSE Eng', 'مفتوح', '', NULL, 0, NULL, 1)"
            )
            conn.execute(
                "INSERT INTO docs VALUES (11, 'weekly_report', 'WR-01', 'Weekly Digest', '2026-11-01', NULL, '', '', 'مغلق', '', NULL, 0, NULL, 1)"
            )
            conn.execute(
                "INSERT INTO injuries VALUES (1, NULL, 1, '2026-11-05', '09:30', 'FAC', 'left_hand', 0, 'Workshop', 'Cut on ID 29901011234567 hand', 'No gloves', 'Wear cut gloves', 'مغلق')"
            )
            conn.execute(
                "INSERT INTO items VALUES (1, 'extinguisher', 'EXT-S-01', '01', ?, 1)",
                (json.dumps({"number": "01", "ext_type": "بودرة جافة", "status": "صيانة", "location": "Control Room", "expiry_date": "2027-01-01"}, ensure_ascii=False),),
            )
            conn.execute(
                "INSERT INTO items VALUES (2, 'person', 'WKR-S-01', 'Synthetic Worker', ?, 1)",
                (json.dumps({"name": "Worker One", "national_id": "29805051234567", "id_number": "29805051234567", "phone": "01000000000", "person_type": "عامل", "notes": "مقاول"}, ensure_ascii=False),),
            )
            conn.execute(
                "INSERT INTO risk_assessments VALUES (1, 1, 'RA-S-01', 'Lifting JSA', 'Crane Lift', 'Mech', 'Yard', '2026-11-01', '2027-11-01', 'Assessor', 'Manager', 'معتمد', 'OK')"
            )
            conn.execute(
                "INSERT INTO evacuation_drills VALUES (1, 1, 'Fire Drill 1', '2026-11-02', '10:00', 'حريق مفاجئ بموقع العمل (Fire Drill)', 'Plant', 'Gate A', 'Leader', 180, 165, 'ممتاز', 'نعم', 'نعم', 'نعم', 'None', 'Smooth')"
            )
            conn.execute("INSERT INTO man_hours VALUES (1, 2026, 11, 50000)")
            conn.commit()
            conn.close()

            before_dir_files = sorted(p.name for p in tmp_path.iterdir())
            plan = hm.build_migration_plan(legacy_db=syn_db, foram_dir=foram_dir)
            after_dir_files = sorted(p.name for p in tmp_path.iterdir())
            self.assertEqual(before_dir_files, after_dir_files)  # Read-only: zero -wal/-shm files created!

            self.assertEqual(plan["counts_by_doc_type"], {"EVAC": 1, "FEI": 1, "INC": 1, "JSA": 1, "MHR": 1, "PTW": 1, "WKR": 1})
            self.assertEqual(plan["skipped_outside_catalog"], {"weekly_report": 1})
            self.assertGreaterEqual(plan["filtered_national_ids"], 3)
            self.assertGreaterEqual(plan["filtered_pii_fields"], 1)
            self.assertEqual(plan["foram"]["valid_count"], 1)
            self.assertEqual(plan["foram"]["skipped_junk_count"], 2)

            serialized = json.dumps(plan["records"], ensure_ascii=False)
            self.assertNotIn("29805051234567", serialized)
            self.assertNotIn("29901011234567", serialized)
            self.assertNotIn("01000000000", serialized)

            by_id = {r["id"]: r["data"] for r in plan["records"]}
            self.assertEqual(by_id["leg-fei-item-1"]["extinguisherType"], "Dry Powder")
            self.assertEqual(by_id["leg-fei-item-1"]["extinguisherStatus"], "Needs Maintenance")
            self.assertEqual(by_id["leg-wkr-item-2"]["personType"], "Worker")
            self.assertEqual(by_id["leg-wkr-item-2"]["affiliation"], "Contractor")
            self.assertEqual(by_id["leg-jsa-1"]["status"], "Approved")
            self.assertEqual(by_id["leg-evac-1"]["evaluationGrade"], "Excellent")
            self.assertEqual(by_id["leg-inc-inj-1"]["bodyPart"], "Left Hand")

            # Verify idempotent execute against local PostgreSQL HSE project (and non-HSE rejection)
            target_hse = "SYN-HSE-P4-MIG"
            target_eng = "SYN-ENG-P4-MIG"
            self.db.create_project(target_hse, "Migration Local Target", "HSE-MIG", workspace_id="HSE")
            self.db.create_project(target_eng, "ENG Reject Target", "ENG-MIG", workspace_id="ENG")
            local_url = str(getattr(self.db, "DATABASE_URL", "") or os.environ.get("DATABASE_URL") or "").strip()
            psy_mod = None
            if not local_url:
                pool = self.db.get_pool()
                probe_conn = pool.getconn()
                try:
                    dsn = probe_conn.get_dsn_parameters()
                    local_url = f"postgresql://{dsn['user']}@{dsn.get('hostaddr') or dsn.get('host') or '127.0.0.1'}:{dsn['port']}/{dsn['dbname']}"
                finally:
                    pool.putconn(probe_conn)

                class _PooledConn:
                    def __init__(self, inner):
                        self._inner = inner

                    def __enter__(self):
                        return self._inner.__enter__()

                    def __exit__(self, *args):
                        return self._inner.__exit__(*args)

                    def cursor(self, *args, **kwargs):
                        return self._inner.cursor(*args, **kwargs)

                    def close(self):
                        pool.putconn(self._inner)

                class _LocalPoolPsycopg2:
                    @staticmethod
                    def connect(dsn_url):
                        hm.validate_local_postgres_url(dsn_url)
                        return _PooledConn(pool.getconn())

                psy_mod = _LocalPoolPsycopg2

            with self.assertRaises(ValueError):
                hm.execute_local_migration(local_url, target_eng, legacy_db=syn_db, foram_dir=foram_dir, psycopg2_module=psy_mod)

            res1 = hm.execute_local_migration(local_url, target_hse, legacy_db=syn_db, foram_dir=foram_dir, psycopg2_module=psy_mod)
            self.assertEqual(res1["inserted"], 7)
            self.assertEqual(res1["updated"], 0)

            # Second run must update cleanly with 0 new duplicates
            res2 = hm.execute_local_migration(local_url, target_hse, legacy_db=syn_db, foram_dir=foram_dir, psycopg2_module=psy_mod)
            self.assertEqual(res2["inserted"], 0)
            self.assertEqual(res2["updated"], 7)

        # 3. Read-only smoke against real D:\SafetyQR_v9 when present on workstation
        if hm.DEFAULT_LEGACY_DB.exists() and hm.DEFAULT_FORAM_DIR.exists():
            buf = io.StringIO()
            with patch("sys.stdout", buf):
                rc = hm.main(["--dry-run", "--json"])
            self.assertEqual(rc, 0)
            summary = json.loads(buf.getvalue())
            self.assertTrue(summary["read_only"])
            self.assertTrue(summary["source_exists"])
            self.assertGreater(summary["total_mapped_records"], 0)
            self.assertGreater(summary["filtered_national_ids"], 0)
            self.assertGreaterEqual(summary["foram"]["skipped_junk_count"], 1)


if __name__ == "__main__":
    unittest.main()
