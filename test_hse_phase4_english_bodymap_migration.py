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
            self.assertEqual(plan["filtered_national_ids"], 0)
            self.assertEqual(plan["filtered_pii_fields"], 0)
            self.assertGreaterEqual(plan["preserved_national_ids"], 3)
            self.assertGreaterEqual(plan["preserved_pii_fields"], 1)
            self.assertEqual(plan["foram"]["valid_count"], 1)
            self.assertEqual(plan["foram"]["skipped_junk_count"], 2)

            serialized = json.dumps(plan["records"], ensure_ascii=False)
            self.assertIn("29805051234567", serialized)
            self.assertIn("29901011234567", serialized)
            self.assertIn("01000000000", serialized)

            by_id = {r["id"]: r["data"] for r in plan["records"]}
            self.assertEqual(by_id["leg-fei-item-1"]["extinguisherType"], "Dry Powder")
            self.assertEqual(by_id["leg-fei-item-1"]["extinguisherStatus"], "Needs Maintenance")
            self.assertEqual(by_id["leg-wkr-item-2"]["personType"], "Worker")
            self.assertEqual(by_id["leg-wkr-item-2"]["affiliation"], "Contractor")
            self.assertEqual(by_id["leg-wkr-item-2"]["nationalId"], "29805051234567")
            self.assertEqual(by_id["leg-wkr-item-2"]["phone"], "01000000000")
            self.assertEqual(by_id["leg-jsa-1"]["status"], "Approved")
            self.assertEqual(by_id["leg-evac-1"]["evaluationGrade"], "Excellent")
            self.assertEqual(by_id["leg-inc-inj-1"]["bodyPart"], "Left Hand")
            self.assertIn("29901011234567", by_id["leg-inc-inj-1"]["description"])

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
            self.assertEqual(summary["filtered_national_ids"], 0)
            self.assertGreater(summary["preserved_national_ids"], 0)
            self.assertEqual(summary["attachment_inventory"]["mapped_files_count"], 219)
            self.assertEqual(summary["attachment_inventory"]["worker_photos_count"], 26)
            self.assertGreaterEqual(summary["foram"]["skipped_junk_count"], 1)

    def test_04_multi_project_date_path_sanitization_attachments_and_drive_mock(self):
        from types import SimpleNamespace
        import hse_migration as hm

        # 1. Double-explicit production DB guard (--allow-remote-db + --confirm-production)
        remote_url = "postgresql://user:SecretPass99@db.prod.example.com:5432/dcr"
        with self.assertRaises(ValueError):
            hm.validate_target_postgres_url(remote_url, allow_remote=False, confirm_production=False)
        with self.assertRaises(ValueError):
            hm.validate_target_postgres_url(remote_url, allow_remote=True, confirm_production=False)
        with self.assertRaises(ValueError):
            hm.validate_target_postgres_url(remote_url, allow_remote=False, confirm_production=True)
        self.assertEqual(
            hm.validate_target_postgres_url(remote_url, allow_remote=True, confirm_production=True),
            remote_url,
        )
        self.assertNotIn("SecretPass99", hm._redact_sensitive_text(remote_url))
        assembled = hm.resolve_target_db_url_from_env(
            env={
                "RAILWAY_TCP_PROXY_DOMAIN": "roundhouse.proxy.rlwy.net",
                "RAILWAY_TCP_PROXY_PORT": "19452",
                "POSTGRES_USER": "postgres",
                "POSTGRES_PASSWORD": "p@ss:word/1",
                "POSTGRES_DB": "railway",
            }
        )
        self.assertEqual(
            assembled,
            "postgresql://postgres:p%40ss%3Aword%2F1@roundhouse.proxy.rlwy.net:19452/railway",
        )
        self.assertEqual(hm.validate_target_postgres_url(assembled, allow_remote=True, confirm_production=True), assembled)

        # CLI also blocks remote --target-db unless both confirmation flags are provided
        with self.assertRaises(ValueError):
            hm.main(["--target-db", remote_url, "--multi-project"])
        with self.assertRaises(ValueError):
            hm.main(["--target-db", remote_url, "--multi-project", "--allow-remote-db"])

        # 2. Read-only probe verifies BEGIN READ ONLY + workspace filtering
        recorded_sql = []

        class _MockReadOnlyCur:
            def execute(self, sql, params=None):
                recorded_sql.append(" ".join(str(sql).split()))

            def fetchall(self):
                return [("HSE-EX-1", "HSE-EX-1", "Existing Site", "HSE", True)]

            def __enter__(self):
                return self

            def __exit__(self, *a):
                return False

        class _MockReadOnlyConn:
            def __init__(self):
                self.rolled_back = False
                self.closed = False

            def cursor(self):
                return _MockReadOnlyCur()

            def rollback(self):
                self.rolled_back = True

            def close(self):
                self.closed = True

        mock_ro_conn = _MockReadOnlyConn()

        class _MockPsyRo:
            @staticmethod
            def connect(_url):
                return mock_ro_conn

        ro_projects = hm.fetch_existing_hse_projects_readonly(
            db_url="postgresql://ro:pw@localhost:5432/dcr",
            psycopg2_module=_MockPsyRo,
        )
        self.assertEqual(len(ro_projects), 1)
        self.assertTrue(mock_ro_conn.rolled_back and mock_ro_conn.closed)
        self.assertEqual(recorded_sql[0], "BEGIN READ ONLY")
        self.assertIn("WHERE UPPER(COALESCE(workspace_id, 'ENG')) = 'HSE'", recorded_sql[1])

        # 3. Synthetic multi-project SQLite + local files tree: date/path sanitization & attachment inventory
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp_path = Path(tmp_dir)
            syn_db = tmp_path / "syn_multi.db"
            foram_dir = tmp_path / "FORAM"
            foram_dir.mkdir()
            (foram_dir / "Permit.docx").write_bytes(b"docx")

            files_dir = tmp_path / "files"
            (files_dir / "HSE" / "unsafe_condition").mkdir(parents=True)
            (files_dir / "HSE" / "weekly_report").mkdir(parents=True)
            (files_dir / "PER-0010").mkdir(parents=True)
            (files_dir / "RISK_ASSESSMENTS").mkdir(parents=True)
            (files_dir / "EVACUATIONS").mkdir(parents=True)
            (files_dir / "photos").mkdir(parents=True)

            (files_dir / "HSE" / "unsafe_condition" / "obs_a.pdf").write_bytes(b"obs-pdf-1")
            (files_dir / "HSE" / "unsafe_condition" / "obs_b.pdf").write_bytes(b"obs-pdf-2-longer")
            (files_dir / "HSE" / "weekly_report" / "wr1.pdf").write_bytes(b"weekly-ignored")
            (files_dir / "PER-0010" / "cert.pdf").write_bytes(b"worker-cert")
            (files_dir / "EVACUATIONS" / "evac.pdf").write_bytes(b"evac-pdf")
            (files_dir / "photos" / "per_secret.jpg").write_bytes(b"pii-per-bytes")
            (files_dir / "photos" / "nid_secret.jpg").write_bytes(b"pii-nid-bytes")
            (files_dir / "HSE" / "orphan_disk_file.pdf").write_bytes(b"unref")
            # Intentionally omit RISK_ASSESSMENTS/missing_jsa.pdf from disk to test missing-file reporting

            conn = sqlite3.connect(syn_db)
            conn.executescript(
                """
                CREATE TABLE projects (id INTEGER PRIMARY KEY, name TEXT, location TEXT, start_date TEXT, project_type TEXT);
                CREATE TABLE docs (id INTEGER PRIMARY KEY, category TEXT, number TEXT, title TEXT, doc_date TEXT,
                                   valid_until TEXT, location TEXT, responsible TEXT, status TEXT, notes TEXT,
                                   injury_type TEXT, lost_days INTEGER, closed_at TEXT, project_id INTEGER);
                CREATE TABLE doc_files (id INTEGER PRIMARY KEY, doc_id INTEGER, file_name TEXT, file_path TEXT);
                CREATE TABLE items (id INTEGER PRIMARY KEY, type TEXT, code TEXT, title TEXT, data TEXT, project_id INTEGER);
                CREATE TABLE records (id INTEGER PRIMARY KEY, item_id INTEGER, kind TEXT, title TEXT, record_date TEXT, expiry_date TEXT, notes TEXT);
                CREATE TABLE record_files (id INTEGER PRIMARY KEY, record_id INTEGER, file_name TEXT, file_path TEXT);
                CREATE TABLE risk_assessments (id INTEGER PRIMARY KEY, project_id INTEGER, code TEXT, title TEXT, activity_name TEXT,
                                                 department TEXT, location TEXT, assessment_date TEXT, review_date TEXT,
                                                 assessor_name TEXT, approver_name TEXT, status TEXT, source_file TEXT, notes TEXT);
                CREATE TABLE evacuation_drills (id INTEGER PRIMARY KEY, project_id INTEGER, title TEXT, drill_date TEXT, drill_time TEXT,
                                                  scenario TEXT, location TEXT, assembly_point TEXT, leader_name TEXT,
                                                  target_duration_sec INTEGER, actual_duration_sec INTEGER, evaluation_grade TEXT,
                                                  alarm_working TEXT, exit_routes_clear TEXT, assembly_adequate TEXT,
                                                  corrective_actions TEXT, source_file TEXT, notes TEXT);
                CREATE TABLE man_hours (project_id INTEGER, year INTEGER, month INTEGER, hours REAL);
                """
            )
            conn.execute("INSERT INTO projects VALUES (1, 'CFC PHAES 2', 'Cairo', '2026-04-08', 'general')")
            conn.execute("INSERT INTO projects VALUES (6, 'جامعه مصر', 'Giza', '2025-01-01', 'maintenance')")
            # Multi-file OBS on project 1 + remark containing a raw Windows local path + valid URL
            conn.execute(
                "INSERT INTO docs VALUES (1, 'unsafe_condition', 'OBS-1/A', 'Scaffold guard', '2026-10-15', NULL, 'Block A', 'Eng', 'مفتوح', 'See D:\\SafetyQR_v9\\data\\files\\temp.pdf and https://example.com/ref', NULL, 0, NULL, 1)"
            )
            conn.execute(
                "INSERT INTO docs VALUES (2, 'weekly_report', 'WR-1', 'Weekly report', '2026-10-16', NULL, 'Site', 'Eng', 'مغلق', '', NULL, 0, NULL, 1)"
            )
            conn.execute("INSERT INTO doc_files VALUES (10, 1, 'obs_a.pdf', 'HSE/unsafe_condition/obs_a.pdf')")
            conn.execute("INSERT INTO doc_files VALUES (11, 1, 'obs_b.pdf', 'HSE/unsafe_condition/obs_b.pdf')")
            conn.execute("INSERT INTO doc_files VALUES (12, 2, 'wr1.pdf', 'HSE/weekly_report/wr1.pdf')")

            # Worker with 0026-10-27 typo, Arabic affiliation, full National ID / phone / blood type & photos + record_file
            conn.execute(
                "INSERT INTO items VALUES (10, 'person', 'PER-0010', 'Samy', ?, 6)",
                (
                    json.dumps(
                        {
                            "name": "Samy Test",
                            "affiliation": "مقاول",
                            "person_type": "عامل",
                            "national_id": "29801011234567",
                            "phone": "01012345678",
                            "hire_date": "0026-10-27",
                            "blood_type": "O+",
                            "photo": "photos/per_secret.jpg",
                            "national_id_card_img": "photos/nid_secret.jpg",
                            "emergency_contact": "01223242577",
                            "notes": "",
                        },
                        ensure_ascii=False,
                    ),
                ),
            )
            conn.execute("INSERT INTO records VALUES (50, 10, 'تصريح', 'Permit', '2026-10-27', '2027-10-27', '')")
            conn.execute("INSERT INTO record_files VALUES (70, 50, 'cert.pdf', 'PER-0010/cert.pdf')")

            # JSA with missing source_file on disk + local server path in notes
            conn.execute(
                "INSERT INTO risk_assessments VALUES (3, 6, 'RA-03', 'Roof JSA', 'Roofing', 'Civil', 'Hall', '2026-11-01', '2027-11-01', 'Assessor', 'Manager', 'معتمد', 'RISK_ASSESSMENTS/missing_jsa.pdf', 'File data/files/RISK_ASSESSMENTS/missing_jsa.pdf')"
            )
            # EVAC with valid PDF on disk
            conn.execute(
                "INSERT INTO evacuation_drills VALUES (5, 1, 'Fire Drill', '2026-06-09', '10:00', 'حريق مفاجئ بموقع العمل (fire drill)', 'Plant', 'Gate', 'Leader', 180, 150, 'ممتاز', 'نعم', 'نعم', 'نعم', 'None', 'EVACUATIONS/evac.pdf', '')"
            )
            conn.execute("INSERT INTO man_hours VALUES (1, 2026, 10, 42000)")
            conn.commit()
            conn.close()

            # Pre-seed existing DB state: matched MUST project + superadmin + HSE admin + ENG-only admin
            db = self.db
            existing_must_pid = "SYN-MUST-PRE"
            db.create_project(existing_must_pid, "جامعة مصر", "MUST-H", workspace_id="HSE")
            db.add_user("syn_p4_hse_adm", "SynPass!123", role="admin", primary_workspace="HSE")
            db.set_user_workspaces("syn_p4_hse_adm", ["HSE"], "admin")
            db.add_user("syn_p4_eng_adm_only", "SynPass!123", role="admin", primary_workspace="ENG")
            db.set_user_workspaces("syn_p4_eng_adm_only", ["ENG"], "admin")

            existing_hse_meta = [
                {"id": p["id"], "code": p["code"], "name": p["name"], "workspace_id": p["workspace_id"], "has_drive_folder": False}
                for p in db.get_projects(workspace_id="HSE")
            ]
            plan = hm.build_migration_plan(
                legacy_db=syn_db,
                foram_dir=foram_dir,
                files_dir=files_dir,
                existing_hse_projects=existing_hse_meta,
            )

            # Date & local path sanitization assertions (with full worker PII preserved)
            by_id = {r["id"]: r["data"] for r in plan["records"]}
            self.assertEqual(by_id["leg-wkr-item-10"]["hireDate"], "2026-10-27")
            self.assertEqual(by_id["leg-wkr-item-10"]["affiliation"], "Contractor")
            self.assertEqual(by_id["leg-wkr-item-10"]["nationalId"], "29801011234567")
            self.assertEqual(by_id["leg-wkr-item-10"]["phone"], "01012345678")
            self.assertEqual(by_id["leg-wkr-item-10"]["emergencyContact"], "01223242577")
            self.assertEqual(by_id["leg-wkr-item-10"]["bloodType"], "O+")
            self.assertGreaterEqual(plan["sanitized_dates"], 1)
            self.assertGreaterEqual(plan["sanitized_local_paths"], 2)
            self.assertNotIn("D:\\SafetyQR_v9", by_id["leg-obs-doc-1"]["remarks"])
            self.assertIn("https://example.com/ref", by_id["leg-obs-doc-1"]["remarks"])
            self.assertNotIn("RISK_ASSESSMENTS/missing_jsa.pdf", by_id["leg-jsa-3"]["remarks"])
            self.assertEqual(plan["filtered_pii_fields"], 0)
            self.assertGreaterEqual(plan["preserved_pii_fields"], 4)

            # Attachment inventory assertions (including worker portrait + NID card scan)
            inv = plan["attachment_inventory"]
            self.assertEqual(inv["mapped_files_count"], 7)  # 2 OBS + 1 WKR cert + 2 WKR photos + 1 JSA(missing) + 1 EVAC
            self.assertEqual(inv["worker_photos_count"], 2)
            self.assertEqual(inv["records_with_attachments_count"], 4)
            self.assertEqual(inv["multi_file_records_count"], 2)
            self.assertEqual(inv["missing_on_disk_count"], 1)
            self.assertEqual(inv["missing_on_disk_files"][0]["record_id"], "leg-jsa-3")
            self.assertEqual(inv["skipped_outside_catalog_count"], 1)
            self.assertIn("weekly_report", inv["skipped_outside_catalog_by_category"])
            self.assertEqual(inv["pii_photos_excluded_count"], 0)
            self.assertEqual(inv["unreferenced_disk_files_count"], 1)

            # Project matching plan assertions (1 new CFC, 1 existing MUST matched despite ه vs ة)
            pmap = {m["legacy_id"]: m for m in plan["project_mapping_plan"]}
            self.assertEqual(pmap[1]["action"], "create_new")
            self.assertEqual(pmap[1]["target_project_id"], "HSE-CFC-P2")
            self.assertEqual(pmap[6]["action"], "match_existing")
            self.assertEqual(pmap[6]["target_project_id"], existing_must_pid)

            # Resolve local psycopg2 connection wrapper cleanly
            local_url = str(getattr(db, "DATABASE_URL", "") or os.environ.get("DATABASE_URL") or "").strip()
            psy_mod = None
            if not local_url:
                pool = db.get_pool()
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

                    def commit(self):
                        return self._inner.commit()

                    def rollback(self):
                        return self._inner.rollback()

                    def close(self):
                        pool.putconn(self._inner)

                class _LocalPoolPsycopg2:
                    @staticmethod
                    def connect(dsn_url):
                        hm.validate_local_postgres_url(dsn_url)
                        return _PooledConn(pool.getconn())

                psy_mod = _LocalPoolPsycopg2

            # Mock Drive HTTP client (zero network calls, zero local server copies)
            class _MockResp:
                def __init__(self, status_code, payload=None, headers=None):
                    self.status_code = status_code
                    self._payload = payload or {}
                    self.headers = headers or {}
                    self.text = json.dumps(self._payload)

                def json(self):
                    return self._payload

            class _MockDriveHttp:
                def __init__(self):
                    self.folders = {"drv-root": {"id": "drv-root", "name": "HSE Root"}}
                    self.children = {"drv-root": []}
                    self.sessions = {}
                    self.uploaded = {}
                    self._seq = 0
                    self.upload_puts = 0

                def get(self, url, params=None, headers=None, timeout=30):
                    if "/files/" in url:
                        fid = url.rsplit("/", 1)[-1]
                        meta = self.folders.get(fid)
                        if not meta:
                            return _MockResp(404, {"error": "not found"})
                        return _MockResp(200, {"id": fid, "mimeType": "application/vnd.google-apps.folder", "trashed": False})
                    q = (params or {}).get("q", "")
                    m_par = re.search(r"'([^']+)'\s+in\s+parents", q)
                    pid_drive = m_par.group(1) if m_par else ""
                    files = list(self.children.get(pid_drive, []))
                    m_nm = re.search(r"name\s*=\s*'([^']+)'", q)
                    if m_nm:
                        files = [f for f in files if f["name"] == m_nm.group(1)]
                    return _MockResp(200, {"files": files})

                def post(self, url, params=None, headers=None, data=None, timeout=30):
                    self._seq += 1
                    body = json.loads(data or "{}")
                    if "upload/drive/v3/files" in url:
                        sess = f"https://mock-upload/session/{self._seq}"
                        self.sessions[sess] = {
                            "name": body["name"],
                            "parents": body.get("parents") or [],
                        }
                        return _MockResp(200, {}, headers={"Location": sess})
                    fid = f"drv-f-{self._seq}"
                    parent = (body.get("parents") or [""])[0]
                    item = {"id": fid, "name": body["name"], "mimeType": body.get("mimeType")}
                    self.folders[fid] = item
                    self.children.setdefault(parent, []).append(item)
                    return _MockResp(200, item)

                def put(self, url, data=None, headers=None, timeout=300):
                    self.upload_puts += 1
                    sess = self.sessions[url]
                    self._seq += 1
                    fid = f"drv-file-{self._seq}"
                    payload_bytes = data or b""
                    meta = {
                        "id": fid,
                        "name": sess["name"],
                        "size": str(len(payload_bytes)),
                        "mimeType": "application/pdf",
                        "webViewLink": f"https://drive.google.com/file/d/{fid}/view",
                    }
                    self.uploaded[fid] = payload_bytes
                    return _MockResp(200, meta)

            mock_http = _MockDriveHttp()
            mock_creds = SimpleNamespace(valid=True, token="syn-drive-token")

            res1 = hm.execute_multi_project_migration(
                local_url,
                legacy_db=syn_db,
                foram_dir=foram_dir,
                files_dir=files_dir,
                psycopg2_module=psy_mod,
                upload_drive=True,
                drive_root_id="drv-root",
                http_module=mock_http,
                drive_creds=mock_creds,
            )
            self.assertEqual(res1["projects_created"], 1)  # Only CFC PHAES 2 created; MUST already existed
            self.assertEqual(res1["inserted"], 5)  # 1 OBS, 1 WKR, 1 JSA, 1 EVAC, 1 MHR
            self.assertEqual(res1["drive_upload"]["uploaded_files"], 6)  # 2 OBS + 1 WKR cert + 2 WKR photos + 1 EVAC (JSA missing)
            self.assertEqual(res1["drive_upload"]["missing_files"], 1)
            self.assertEqual(mock_http.upload_puts, 6)

            # Verify created CFC project seeded all 11 English registers, Drive folder routed in drv-root, and HSE/SA assignment ONLY
            self.assertEqual(len(db.get_doc_types("HSE-CFC-P2")), 11)
            cfc_proj = db.get_project("HSE-CFC-P2")
            must_proj = db.get_project(existing_must_pid)
            self.assertTrue(cfc_proj.get("drive_folder_id"))
            self.assertTrue(must_proj.get("drive_folder_id"))
            root_child_ids = {f["id"] for f in mock_http.children["drv-root"]}
            self.assertIn(cfc_proj["drive_folder_id"], root_child_ids)
            self.assertIn(must_proj["drive_folder_id"], root_child_ids)
            sa_pids = {p["project_id"] for p in db.get_user_projects("syn_p4_sa")}
            hse_adm_pids = {p["project_id"] for p in db.get_user_projects("syn_p4_hse_adm")}
            eng_adm_pids = {p["project_id"] for p in db.get_user_projects("syn_p4_eng_adm_only")}
            self.assertIn("HSE-CFC-P2", sa_pids)
            self.assertIn("HSE-CFC-P2", hse_adm_pids)
            self.assertNotIn("HSE-CFC-P2", eng_adm_pids)

            # Verify Drive attachment metadata stored cleanly in PostgreSQL JSONB (single & multi-file folder + worker photos)
            obs_rec = db.get_record_by_id("leg-obs-doc-1")
            obs_sub = obs_rec["attachments"]["submission"]
            self.assertTrue(obs_sub["is_folder"])
            self.assertEqual(obs_sub["file_count"], 2)
            self.assertEqual(obs_rec["fileLocation"], obs_sub["folder_url"])

            wkr_rec = db.get_record_by_id("leg-wkr-item-10")
            self.assertEqual(wkr_rec["hireDate"], "2026-10-27")
            self.assertEqual(wkr_rec["nationalId"], "29801011234567")
            self.assertEqual(wkr_rec["phone"], "01012345678")
            self.assertEqual(wkr_rec["bloodType"], "O+")
            wkr_sub = wkr_rec["attachments"]["submission"]
            self.assertTrue(wkr_sub["is_folder"])
            self.assertEqual(wkr_sub["file_count"], 3)
            self.assertEqual({f["name"] for f in wkr_sub["files"]}, {"cert.pdf", "per_secret.jpg", "nid_secret.jpg"})
            self.assertEqual(wkr_rec["fileLocation"], wkr_sub["folder_url"])

            # Verify uploaded bytes equal exact legacy bytes including worker photos & NID scan
            uploaded_contents = set(mock_http.uploaded.values())
            self.assertEqual(
                uploaded_contents,
                {b"obs-pdf-1", b"obs-pdf-2-longer", b"worker-cert", b"pii-per-bytes", b"pii-nid-bytes", b"evac-pdf"},
            )

            # Second run must preserve attachments cleanly and skip re-uploading identical files (0 new PUTs!)
            res2 = hm.execute_multi_project_migration(
                local_url,
                legacy_db=syn_db,
                foram_dir=foram_dir,
                files_dir=files_dir,
                psycopg2_module=psy_mod,
                upload_drive=True,
                drive_root_id="drv-root",
                http_module=mock_http,
                drive_creds=mock_creds,
            )
            self.assertEqual(res2["projects_created"], 0)
            self.assertEqual(res2["inserted"], 0)
            self.assertEqual(res2["updated"], 5)
            self.assertEqual(res2["drive_upload"]["uploaded_files"], 0)
            self.assertEqual(res2["drive_upload"]["skipped_existing"], 6)
            self.assertEqual(mock_http.upload_puts, 6)

            # Duplicate docNo records must still receive distinct multi-file Drive subfolder names
            self.assertNotEqual(
                hm._safe_drive_subfolder_name("PEM-058-005", "leg-obs-doc-89"),
                hm._safe_drive_subfolder_name("PEM-058-005", "leg-obs-doc-96"),
            )


if __name__ == "__main__":
    unittest.main()
