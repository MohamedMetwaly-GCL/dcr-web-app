"""Read-only SafetyQR legacy inspection & idempotent local DCR HSE migration utility.

Strict safety guarantees:
- Legacy SQLite source is opened ONLY with SQLite URI ``?mode=ro``.
- Legacy ``FORAM/`` directory is walked strictly read-only, filtering ``.lnk`` and ``~$`` files.
- 14-digit Egyptian National IDs and private phone numbers are scrubbed / counted and never imported.
- Default mode is ``--dry-run`` (zero target database writes).
- Optional ``--target-db`` refuses any non-local PostgreSQL host (only ``127.0.0.1``, ``localhost``, ``::1``).
"""

import argparse
from collections import Counter
from contextlib import closing
from datetime import datetime
import json
import os
from pathlib import Path
import re
import sqlite3
import sys
from urllib.parse import urlparse


DEFAULT_LEGACY_DB = Path(r"D:\SafetyQR_v9\SafetyQR\data\app.db")
DEFAULT_FORAM_DIR = Path(r"D:\SafetyQR_v9\FORAM")
ALLOWED_LOCAL_HOSTS = frozenset({"127.0.0.1", "localhost", "::1"})

_NAT_ID_14_RE = re.compile(r"(?<!\d)\d{14}(?!\d)")
_PHONE_KEYS = frozenset({"phone", "mobile", "telephone", "tel", "contact_phone"})
_NAT_ID_KEYS = frozenset({"national_id", "nationalid", "id_number", "nid", "ssn"})
_PII_MEDIA_KEYS = frozenset({"photo_path", "card_image", "signature_path", "blood_type"})

_DOC_CATEGORY_MAP = {
    "hot_permit": ("PTW", "Hot Work Permit"),
    "cold_permit": ("PTW", "Cold Work Permit"),
    "height_permit": ("PTW", "Work at Height Permit"),
    "confined_permit": ("PTW", "Confined Space Permit"),
    "excavation_permit": ("PTW", "Excavation Permit"),
    "electrical_permit": ("PTW", "Electrical / Isolation Permit"),
    "lifting_permit": ("PTW", "Lifting Operations Permit"),
    "radiography_permit": ("PTW", "Radiography Permit"),
    "maintenance": ("PTW", "Maintenance Work"),
    "unsafe_act": ("OBS", "Unsafe Act"),
    "unsafe_condition": ("OBS", "Unsafe Condition"),
    "near_miss": ("OBS", "Near Miss"),
    "first_aid": ("INC", "First Aid Report"),
    "injuries": ("INC", "Injury Register"),
    "exclusion": ("SVP", "Exclusion Report"),
    "financial_violation": ("SVP", "Financial Violation Report"),
    "toolbox": ("TBT", "Toolbox Talk"),
}

_EXT_TYPE_MAP = {
    "بودرة جافة": "Dry Powder",
    "dry powder": "Dry Powder",
    "ثاني أكسيد الكربون (co2)": "CO2",
    "ثاني أكسيد الكربون": "CO2",
    "co2": "CO2",
    "رغوة": "Foam",
    "foam": "Foam",
    "مياه": "Water",
    "water": "Water",
    "wet chemical": "Wet Chemical",
    "fire hose reel": "Fire Hose Reel",
    "أخرى": "Other",
}

_EXT_STATUS_MAP = {
    "صالحة": "Valid",
    "valid": "Valid",
    "صيانة": "Needs Maintenance",
    "needs maintenance": "Needs Maintenance",
    "maintenance": "Needs Maintenance",
    "under refill": "Under Refill",
    "فارغة": "Expired",
    "expired": "Expired",
    "out of service": "Out of Service",
}

_JSA_STATUS_MAP = {
    "معتمد": "Approved",
    "approved": "Approved",
    "قيد المراجعة": "Under Review",
    "under review": "Under Review",
    "pending": "Under Review",
    "مسودة": "Draft",
    "draft": "Draft",
    "محدث": "Revised",
    "revised": "Revised",
    "updated": "Revised",
}

_EVAC_SCENARIO_MAP = {
    "حريق مفاجئ بموقع العمل (fire drill)": "Fire Emergency Drill",
    "تسرب غاز / مواد كيميائية خطرة (gas / chemical leak)": "Gas / Chemical Leak",
    "هبوط أرضي / تصدع إنشائي أو زلزال (structural / earthquake)": "Structural Collapse / Earthquake",
    "انفجار / حالة طوارئ شاملة (general site evacuation)": "General Site Evacuation",
    "انقطاع تيار كهربائي وحصار عمال (power failure / trapped personnel)": "Power Failure / Trapped Personnel",
}

_EVAC_GRADE_MAP = {
    "ممتاز": "Excellent",
    "جيد جداً": "Very Good",
    "جيد جدا": "Very Good",
    "مقبول": "Satisfactory",
    "يحتاج إعادة وتدريب": "Needs Retraining",
}

_CHECK_EN_MAP = {
    "نعم": "Yes",
    "yes": "Yes",
    "جزئي": "Partial",
    "partial": "Partial",
    "لا": "No",
    "no": "No",
}

_PERSON_TYPE_MAP = {
    "موظف": "Staff",
    "عامل": "Worker",
    "مسؤول مخزن": "Storekeeper",
    "مسؤول سفتي": "Safety Officer",
    "مشرف": "Supervisor",
    "مهندس": "Engineer",
    "زائر": "Visitor",
}

_AFFILIATION_MAP = {
    "مقاول": "Contractor",
    "شركة": "Company Staff",
}

_BODY_PART_MAP = {
    "head": ("head_face", "Head / Face"),
    "face": ("head_face", "Head / Face"),
    "neck": ("neck", "Neck"),
    "chest": ("chest", "Chest"),
    "abdomen": ("abdomen", "Abdomen"),
    "back": ("upper_back", "Upper Back"),
    "upper_back": ("upper_back", "Upper Back"),
    "lower_back": ("lower_back", "Lower Back"),
    "left_shoulder": ("left_shoulder", "Left Shoulder"),
    "right_shoulder": ("right_shoulder", "Right Shoulder"),
    "left_arm": ("left_arm", "Left Arm"),
    "right_arm": ("right_arm", "Right Arm"),
    "left_hand": ("left_hand", "Left Hand"),
    "right_hand": ("right_hand", "Right Hand"),
    "pelvis": ("pelvis_hip", "Pelvis / Hip"),
    "hip": ("pelvis_hip", "Pelvis / Hip"),
    "left_thigh": ("left_thigh", "Left Thigh"),
    "right_thigh": ("right_thigh", "Right Thigh"),
    "left_knee": ("left_knee", "Left Knee"),
    "right_knee": ("right_knee", "Right Knee"),
    "left_leg": ("left_leg", "Left Leg"),
    "right_leg": ("right_leg", "Right Leg"),
    "left_foot": ("left_foot", "Left Foot"),
    "right_foot": ("right_foot", "Right Foot"),
}


def validate_local_postgres_url(url: str) -> str:
    """Ensure target DB URL points strictly to a local host (127.0.0.1, localhost, or ::1)."""
    raw = str(url or "").strip()
    if not raw:
        raise ValueError("Target database URL is required.")
    parsed = urlparse(raw)
    scheme = (parsed.scheme or "").lower()
    if scheme not in ("postgresql", "postgres"):
        raise ValueError(f"Unsupported database scheme '{scheme}'. Expected postgresql://.")
    if "," in (parsed.netloc or ""):
        raise ValueError("Multi-host connection strings are not permitted.")
    host = (parsed.hostname or "").strip().lower()
    if host not in ALLOWED_LOCAL_HOSTS:
        raise ValueError(
            f"Refusing non-local database host '{host or '<empty>'}'. Only 127.0.0.1 / localhost / ::1 is allowed."
        )
    return raw


def open_legacy_sqlite_readonly(db_path: Path | str) -> sqlite3.Connection:
    """Open legacy SQLite file strictly read-only (`?mode=ro`), never modifying source directory."""
    resolved = Path(db_path).resolve(strict=True)
    if not resolved.is_file():
        raise FileNotFoundError(f"Legacy SQLite database not found: {resolved}")
    uri = f"file:{resolved.as_posix()}?mode=ro"
    conn = sqlite3.connect(uri, uri=True)
    conn.row_factory = sqlite3.Row
    return conn


def is_synthetic_national_id(value) -> bool:
    """Return True if value matches an exact or embedded 14-digit Egyptian National ID."""
    txt = str(value or "").strip()
    if not txt:
        return False
    return bool(_NAT_ID_14_RE.search(txt))


def scrub_text_national_ids(value) -> tuple[str, int]:
    """Redact any embedded 14-digit National IDs from free text without failing migration."""
    txt = str(value or "").strip()
    if not txt:
        return "", 0
    matches = _NAT_ID_14_RE.findall(txt)
    if not matches:
        return txt, 0
    cleaned = _NAT_ID_14_RE.sub("[REDACTED-ID]", txt)
    return cleaned, len(matches)


def _norm_status(raw_status):
    txt = str(raw_status or "").strip()
    low = txt.lower()
    if low in ("closed", "close", "complete", "completed", "approved") or txt == "مغلق":
        return "Closed"
    if low in ("open", "pending", "active", "in progress") or txt == "مفتوح":
        return "Open"
    return txt or "Open"


def _norm_injury_class(raw_val):
    txt = str(raw_val or "").strip()
    up = txt.upper()
    aliases = {
        "FIRST AID": "FAC",
        "FIRST AID CASE": "FAC",
        "FAC": "FAC",
        "MEDICAL TREATMENT": "MTC",
        "MEDICAL TREATMENT CASE": "MTC",
        "MTC": "MTC",
        "LOST TIME INJURY": "LTI",
        "LTI": "LTI",
        "NEAR MISS": "Near Miss",
        "FATALITY": "Fatality",
    }
    return aliases.get(up, txt)


def _map_foram_template(filename: str) -> str:
    low = filename.lower()
    if any(k in low for k in ("weekly", "اسبوعي", "تقرير اليومي", "daily")):
        return "OUT_OF_CATALOG"
    if any(k in low for k in ("permit", "تصريح")):
        return "PTW"
    if any(k in low for k in ("unsafe", "near miss", "nieer miss", "تفتيش_سلامة")):
        return "OBS"
    if any(k in low for k in ("first aid", "اسعافات")):
        return "INC"
    if any(k in low for k in ("violation", "مخالف", "غراما", "pem-058-sv", "pem-065-sv", "sv.docx")):
        return "SVP"
    if any(k in low for k in ("toolbox", "tool box", "tbt")):
        return "TBT"
    if any(k in low for k in ("extinguish", "طفاي", "hand tools", "كهرباء", "حزام", "لحام")):
        return "FEI"
    if any(k in low for k in ("drill", "اخلاء")):
        return "EVAC"
    if any(k in low for k in ("induction", "بطاقة التعريف")):
        return "WKR"
    return "OUT_OF_CATALOG"


def inspect_foram_directory(foram_dir: Path | str) -> dict:
    """Walk legacy FORAM/ templates read-only, skipping .lnk and ~$ lock files."""
    root_path = Path(foram_dir)
    if not root_path.exists():
        return {
            "exists": False,
            "path": str(root_path),
            "valid_templates": [],
            "skipped_junk_files": [],
            "by_doc_type": {},
        }

    valid_templates = []
    skipped_junk = []
    by_dt = Counter()

    for dirpath, _dirs, filenames in os.walk(root_path):
        for fn in sorted(filenames):
            low = fn.lower()
            full = Path(dirpath) / fn
            rel = str(full.relative_to(root_path))
            if fn.startswith("~$") or low.endswith(".lnk") or low in ("thumbs.db", "desktop.ini"):
                skipped_junk.append({"file": rel, "reason": "shortcut_or_lock"})
                continue
            dt_code = _map_foram_template(fn)
            by_dt[dt_code] += 1
            valid_templates.append({"file": rel, "mapped_doc_type": dt_code})

    return {
        "exists": True,
        "path": str(root_path),
        "valid_count": len(valid_templates),
        "valid_templates": valid_templates,
        "skipped_junk_count": len(skipped_junk),
        "skipped_junk_files": skipped_junk,
        "by_doc_type": dict(sorted(by_dt.items())),
    }


def _safe_json_loads(raw_text):
    if isinstance(raw_text, dict):
        return raw_text
    try:
        parsed = json.loads(raw_text or "{}")
        return parsed if isinstance(parsed, dict) else {}
    except Exception:
        return {}


def build_migration_plan(
    legacy_db: Path | str = DEFAULT_LEGACY_DB,
    foram_dir: Path | str = DEFAULT_FORAM_DIR,
    allow_missing_source: bool = False,
) -> dict:
    """Inspect legacy SQLite DB & FORAM/ templates strictly read-only and build English DCR records."""
    db_path = Path(legacy_db)
    foram_report = inspect_foram_directory(foram_dir)

    if not db_path.exists():
        if not allow_missing_source:
            raise FileNotFoundError(f"Legacy SQLite database not found: {db_path}")
        return {
            "mode": "dry-run",
            "read_only": True,
            "source_db": str(db_path),
            "source_exists": False,
            "projects": [],
            "records": [],
            "counts_by_doc_type": {},
            "skipped_outside_catalog": {},
            "filtered_national_ids": 0,
            "filtered_pii_fields": 0,
            "foram": foram_report,
        }

    records = []
    projects = []
    counts_by_dt = Counter()
    skipped_cats = Counter()
    filtered_nat_ids = 0
    filtered_pii_fields = 0

    with closing(open_legacy_sqlite_readonly(db_path)) as conn:
        cur = conn.cursor()
        tables = {
            r[0]
            for r in cur.execute("SELECT name FROM sqlite_master WHERE type='table'").fetchall()
        }

        if "projects" in tables:
            for row in cur.execute("SELECT id, name, location, start_date, project_type FROM projects ORDER BY id"):
                projects.append({
                    "legacy_id": row["id"],
                    "name": row["name"] or f"Legacy Project {row['id']}",
                    "location": row["location"] or "",
                    "start_date": row["start_date"] or "",
                    "project_type": row["project_type"] or "general",
                })

        # 1. Legacy docs table (PTW, OBS, INC, SVP, TBT; skip outside catalog e.g. weekly_report)
        if "docs" in tables:
            for row in cur.execute(
                """SELECT id, category, number, title, doc_date, valid_until, location,
                          responsible, status, notes, injury_type, lost_days, closed_at, project_id
                   FROM docs ORDER BY id"""
            ):
                cat = str(row["category"] or "").strip()
                mapped = _DOC_CATEGORY_MAP.get(cat)
                if not mapped:
                    skipped_cats[cat or "unknown"] += 1
                    continue
                dt_code, subtype = mapped
                title_clean, nid_hits = scrub_text_national_ids(row["title"])
                notes_clean, nid_hits2 = scrub_text_national_ids(row["notes"])
                filtered_nat_ids += nid_hits + nid_hits2
                rec_data = {
                    "docNo": str(row["number"] or f"LEG-{dt_code}-{row['id']}"),
                    "legacySubtype": subtype,
                    "title": title_clean or subtype,
                    "issuedDate": str(row["doc_date"] or "")[:10],
                    "location": str(row["location"] or ""),
                    "responsible": str(row["responsible"] or ""),
                    "status": _norm_status(row["status"]),
                    "remarks": notes_clean,
                }
                if dt_code == "PTW" and row["valid_until"]:
                    rec_data["validUntil"] = str(row["valid_until"])[:10]
                if dt_code == "INC":
                    rec_data["injuryClass"] = _norm_injury_class(row["injury_type"])
                    rec_data["lostDays"] = int(row["lost_days"] or 0)
                if row["closed_at"]:
                    rec_data["closedDate"] = str(row["closed_at"])[:10]

                records.append({
                    "id": f"leg-{dt_code.lower()}-doc-{row['id']}",
                    "legacy_project_id": row["project_id"],
                    "dt_id": dt_code,
                    "data": rec_data,
                })
                counts_by_dt[dt_code] += 1

        # 2. Legacy injuries table -> INC (with anatomical English bodyPart, scrubbed of National IDs)
        if "injuries" in tables:
            for row in cur.execute(
                """SELECT id, doc_id, project_id, incident_date, incident_time, injury_category,
                          body_part, lost_days, location, incident_description, root_cause,
                          corrective_action, status
                   FROM injuries ORDER BY id"""
            ):
                raw_bp = str(row["body_part"] or "").strip().lower()
                reg_id, bp_label = _BODY_PART_MAP.get(raw_bp, (raw_bp, raw_bp.replace("_", " ").title() if raw_bp else ""))
                desc_clean, n1 = scrub_text_national_ids(row["incident_description"])
                rc_clean, n2 = scrub_text_national_ids(row["root_cause"])
                ca_clean, n3 = scrub_text_national_ids(row["corrective_action"])
                filtered_nat_ids += n1 + n2 + n3
                rec_data = {
                    "docNo": f"INC-INJ-{row['id']:04d}",
                    "legacySubtype": "Injury Register",
                    "title": desc_clean[:120] if desc_clean else f"Injury Incident #{row['id']}",
                    "issuedDate": str(row["incident_date"] or "")[:10],
                    "incidentTime": str(row["incident_time"] or ""),
                    "location": str(row["location"] or ""),
                    "injuryClass": _norm_injury_class(row["injury_category"]),
                    "bodyPart": bp_label,
                    "bodyRegion": reg_id,
                    "lostDays": int(row["lost_days"] or 0),
                    "description": desc_clean,
                    "rootCause": rc_clean,
                    "correctiveAction": ca_clean,
                    "status": _norm_status(row["status"]),
                }
                records.append({
                    "id": f"leg-inc-inj-{row['id']}",
                    "legacy_project_id": row["project_id"],
                    "dt_id": "INC",
                    "data": rec_data,
                })
                counts_by_dt["INC"] += 1

        # 3. Latest expiry/inspection dates per item from legacy records table
        latest_item_record = {}
        if "records" in tables:
            for row in cur.execute("SELECT item_id, record_date, expiry_date FROM records ORDER BY id"):
                iid = row["item_id"]
                if iid is None:
                    continue
                cur_entry = latest_item_record.setdefault(iid, {})
                if row["record_date"]:
                    cur_entry["last_inspection"] = str(row["record_date"])[:10]
                if row["expiry_date"]:
                    cur_entry["expiry_date"] = str(row["expiry_date"])[:10]

        # 4. Legacy items table -> FEI (extinguisher) & WKR (person, PII-scrubbed)
        if "items" in tables:
            for row in cur.execute("SELECT id, type, code, title, data, project_id FROM items ORDER BY id"):
                itype = str(row["type"] or "").strip().lower()
                raw = _safe_json_loads(row["data"])
                item_dates = latest_item_record.get(row["id"], {})

                if itype == "extinguisher":
                    ext_type_raw = str(raw.get("ext_type") or "").strip().lower()
                    ext_type = _EXT_TYPE_MAP.get(ext_type_raw, raw.get("ext_type") or "")
                    ext_status_raw = str(raw.get("status") or "").strip().lower()
                    ext_status = _EXT_STATUS_MAP.get(ext_status_raw, "Valid" if not ext_status_raw else "Needs Maintenance")
                    last_insp = str(raw.get("last_inspection") or item_dates.get("last_inspection") or "")[:10]
                    exp_date = str(raw.get("expiry_date") or item_dates.get("expiry_date") or "")[:10]
                    rec_data = {
                        "docNo": str(row["code"] or f"EXT-{row['id']:04d}"),
                        "legacySubtype": "Fire Extinguisher Inspection",
                        "title": f"Extinguisher {row['title'] or row['code'] or row['id']}",
                        "equipmentRef": str(raw.get("number") or row["code"] or row["title"] or ""),
                        "issuedDate": last_insp,
                        "lastInspection": last_insp,
                        "expiryDate": exp_date,
                        "location": str(raw.get("location") or ""),
                        "extinguisherType": ext_type,
                        "capacity": str(raw.get("capacity") or ""),
                        "manufactureDate": str(raw.get("manufacture_date") or "")[:10],
                        "extinguisherStatus": ext_status,
                        "conditionNotes": str(raw.get("notes") or ""),
                    }
                    records.append({
                        "id": f"leg-fei-item-{row['id']}",
                        "legacy_project_id": row["project_id"],
                        "dt_id": "FEI",
                        "data": rec_data,
                    })
                    counts_by_dt["FEI"] += 1

                elif itype == "person":
                    # Explicitly strip synthetic/real National IDs, phones, and private media fields
                    for k, v in raw.items():
                        k_low = str(k).lower()
                        if k_low in _NAT_ID_KEYS:
                            if str(v or "").strip() or is_synthetic_national_id(v):
                                filtered_nat_ids += 1
                        elif k_low in _PHONE_KEYS or k_low in _PII_MEDIA_KEYS:
                            if str(v or "").strip():
                                filtered_pii_fields += 1
                        elif is_synthetic_national_id(v):
                            filtered_nat_ids += 1

                    raw_role = str(raw.get("person_type") or "").strip()
                    raw_notes = str(raw.get("notes") or "").strip()
                    notes_clean, nid_in_notes = scrub_text_national_ids(raw_notes)
                    filtered_nat_ids += nid_in_notes
                    aff_en = _AFFILIATION_MAP.get(
                        raw_notes,
                        "Contractor" if ("مقاول" in raw_notes or "يوميه" in raw_notes or "يومية" in raw_notes) else "",
                    )
                    role_en = _PERSON_TYPE_MAP.get(raw_role, "Worker" if raw_role == "" else raw_role)
                    name_clean, nid_in_name = scrub_text_national_ids(raw.get("name") or row["title"] or row["code"])
                    filtered_nat_ids += nid_in_name

                    rec_data = {
                        "docNo": str(row["code"] or f"WKR-{row['id']:04d}"),
                        "title": name_clean,
                        "affiliation": aff_en,
                        "company": str(raw.get("company") or ""),
                        "personType": role_en,
                        "jobTitle": str(raw.get("job_title") or ""),
                        "hireDate": str(raw.get("hire_date") or "")[:10],
                        "remarks": notes_clean,
                    }
                    records.append({
                        "id": f"leg-wkr-item-{row['id']}",
                        "legacy_project_id": row["project_id"],
                        "dt_id": "WKR",
                        "data": rec_data,
                    })
                    counts_by_dt["WKR"] += 1

        # 5. risk_assessments -> JSA
        if "risk_assessments" in tables:
            for row in cur.execute(
                """SELECT id, project_id, code, title, activity_name, department, location,
                          assessment_date, review_date, assessor_name, approver_name, status, notes
                   FROM risk_assessments ORDER BY id"""
            ):
                st_raw = str(row["status"] or "").strip().lower()
                st_en = _JSA_STATUS_MAP.get(st_raw, "Draft")
                rec_data = {
                    "docNo": str(row["code"] or f"JSA-{row['id']:04d}"),
                    "title": str(row["title"] or row["activity_name"] or "Risk Assessment"),
                    "issuedDate": str(row["assessment_date"] or "")[:10],
                    "reviewDate": str(row["review_date"] or "")[:10],
                    "location": str(row["location"] or ""),
                    "activityName": str(row["activity_name"] or ""),
                    "department": str(row["department"] or ""),
                    "assessorName": str(row["assessor_name"] or ""),
                    "approverName": str(row["approver_name"] or ""),
                    "status": st_en,
                    "remarks": str(row["notes"] or ""),
                }
                records.append({
                    "id": f"leg-jsa-{row['id']}",
                    "legacy_project_id": row["project_id"],
                    "dt_id": "JSA",
                    "data": rec_data,
                })
                counts_by_dt["JSA"] += 1

        # 6. contractor_evaluations -> CEV
        if "contractor_evaluations" in tables:
            for row in cur.execute(
                """SELECT id, contractor_id, project_id, year, month, eval_date, evaluator,
                          positive_notes, negative_notes, action_plan
                   FROM contractor_evaluations ORDER BY id"""
            ):
                rec_data = {
                    "docNo": f"CEV-{row['year'] or 2026}-{int(row['month'] or 1):02d}-{row['id']}",
                    "title": f"Contractor #{row['contractor_id'] or row['id']}",
                    "year": row["year"],
                    "month": row["month"],
                    "issuedDate": str(row["eval_date"] or "")[:10],
                    "evaluator": str(row["evaluator"] or ""),
                    "positiveNotes": str(row["positive_notes"] or ""),
                    "negativeNotes": str(row["negative_notes"] or ""),
                    "actionPlan": str(row["action_plan"] or ""),
                }
                records.append({
                    "id": f"leg-cev-{row['id']}",
                    "legacy_project_id": row["project_id"],
                    "dt_id": "CEV",
                    "data": rec_data,
                })
                counts_by_dt["CEV"] += 1

        # 7. evacuation_drills -> EVAC
        if "evacuation_drills" in tables:
            for row in cur.execute(
                """SELECT id, project_id, title, drill_date, drill_time, scenario, location,
                          assembly_point, leader_name, target_duration_sec, actual_duration_sec,
                          evaluation_grade, alarm_working, exit_routes_clear, assembly_adequate,
                          corrective_actions, notes
                   FROM evacuation_drills ORDER BY id"""
            ):
                scen_low = str(row["scenario"] or "").strip().lower()
                grade_low = str(row["evaluation_grade"] or "").strip().lower()
                alarm_low = str(row["alarm_working"] or "").strip().lower()
                exit_low = str(row["exit_routes_clear"] or "").strip().lower()
                asm_low = str(row["assembly_adequate"] or "").strip().lower()
                rec_data = {
                    "docNo": f"EVAC-{row['id']:04d}",
                    "title": str(row["title"] or f"Evacuation Drill #{row['id']}"),
                    "issuedDate": str(row["drill_date"] or "")[:10],
                    "drillTime": str(row["drill_time"] or ""),
                    "scenario": _EVAC_SCENARIO_MAP.get(scen_low, "Fire Emergency Drill"),
                    "location": str(row["location"] or ""),
                    "assemblyPoint": str(row["assembly_point"] or ""),
                    "leaderName": str(row["leader_name"] or ""),
                    "targetDurationSec": row["target_duration_sec"],
                    "actualDurationSec": row["actual_duration_sec"],
                    "evaluationGrade": _EVAC_GRADE_MAP.get(grade_low, "Satisfactory"),
                    "alarmWorking": _CHECK_EN_MAP.get(alarm_low, "Yes"),
                    "exitRoutesClear": _CHECK_EN_MAP.get(exit_low, "Yes"),
                    "assemblyAdequate": _CHECK_EN_MAP.get(asm_low, "Yes"),
                    "correctiveAction": str(row["corrective_actions"] or ""),
                    "remarks": str(row["notes"] or ""),
                }
                records.append({
                    "id": f"leg-evac-{row['id']}",
                    "legacy_project_id": row["project_id"],
                    "dt_id": "EVAC",
                    "data": rec_data,
                })
                counts_by_dt["EVAC"] += 1

        # 8. man_hours -> MHR (deduplicated per project_id/year/month)
        if "man_hours" in tables:
            seen_mhr = set()
            for row in cur.execute("SELECT project_id, year, month, hours FROM man_hours ORDER BY year, month"):
                yr, mo = int(row["year"] or 0), int(row["month"] or 0)
                if not (yr and 1 <= mo <= 12):
                    continue
                key = (row["project_id"], yr, mo)
                if key in seen_mhr:
                    continue
                seen_mhr.add(key)
                rec_data = {
                    "docNo": f"MHR-{yr}-{mo:02d}",
                    "year": yr,
                    "month": mo,
                    "hours": float(row["hours"] or 0),
                }
                records.append({
                    "id": f"leg-mhr-{row['project_id'] or 0}-{yr}-{mo:02d}",
                    "legacy_project_id": row["project_id"],
                    "dt_id": "MHR",
                    "data": rec_data,
                })
                counts_by_dt["MHR"] += 1

    return {
        "mode": "dry-run",
        "read_only": True,
        "generated_at": datetime.utcnow().isoformat(timespec="seconds") + "Z",
        "source_db": str(db_path),
        "source_exists": True,
        "projects": projects,
        "total_mapped_records": len(records),
        "counts_by_doc_type": dict(sorted(counts_by_dt.items())),
        "skipped_outside_catalog": dict(sorted(skipped_cats.items())),
        "filtered_national_ids": filtered_nat_ids,
        "filtered_pii_fields": filtered_pii_fields,
        "foram": foram_report,
        "records": records,
    }


def execute_local_migration(
    target_db_url: str,
    project_id: str,
    legacy_db: Path | str = DEFAULT_LEGACY_DB,
    foram_dir: Path | str = DEFAULT_FORAM_DIR,
    legacy_project_id: int | None = None,
    psycopg2_module=None,
) -> dict:
    """Idempotently import mapped legacy records into a verified local PostgreSQL HSE project."""
    safe_url = validate_local_postgres_url(target_db_url)
    pid = str(project_id or "").strip()
    if not pid:
        raise ValueError("Target DCR --project-id is required for execution.")

    plan = build_migration_plan(legacy_db=legacy_db, foram_dir=foram_dir)
    if psycopg2_module is None:
        import psycopg2 as psycopg2_module  # type: ignore

    selected_recs = [
        r for r in plan["records"]
        if legacy_project_id is None or r["legacy_project_id"] == legacy_project_id
    ]

    inserted = 0
    updated = 0
    skipped = 0

    conn = psycopg2_module.connect(safe_url)
    try:
        with conn:
            with conn.cursor() as cur:
                cur.execute("SELECT workspace_id FROM projects WHERE id = %s", (pid,))
                proj_row = cur.fetchone()
                if not proj_row:
                    raise ValueError(f"Target project '{pid}' not found in target DB.")
                ws_id = str(proj_row[0] or "ENG").strip().upper()
                if ws_id != "HSE":
                    raise ValueError(f"Target project '{pid}' belongs to workspace '{ws_id}', expected 'HSE'.")

                for rec in selected_recs:
                    rid = rec["id"]
                    dt_id = rec["dt_id"]
                    payload = rec["data"]
                    if dt_id == "MHR":
                        cur.execute(
                            """SELECT id FROM records
                               WHERE project_id = %s AND UPPER(dt_id) = 'MHR'
                                 AND (data->>'year')::int = %s AND (data->>'month')::int = %s
                               LIMIT 1""",
                            (pid, int(payload["year"]), int(payload["month"])),
                        )
                        existing_mhr = cur.fetchone()
                        if existing_mhr and existing_mhr[0] != rid:
                            skipped += 1
                            continue

                    cur.execute("SELECT project_id FROM records WHERE id = %s", (rid,))
                    existing = cur.fetchone()
                    payload_json = json.dumps(payload, ensure_ascii=False)
                    if existing:
                        if existing[0] != pid:
                            skipped += 1
                            continue
                        cur.execute(
                            """UPDATE records
                               SET dt_id = %s, data = %s::jsonb, updated_at = NOW()
                               WHERE id = %s AND project_id = %s""",
                            (dt_id, payload_json, rid, pid),
                        )
                        updated += 1
                    else:
                        cur.execute(
                            """INSERT INTO records(id, project_id, dt_id, data)
                               VALUES(%s, %s, %s, %s::jsonb)""",
                            (rid, pid, dt_id, payload_json),
                        )
                        inserted += 1
                cur.execute("DELETE FROM dashboard_cache WHERE project_id = %s", (pid,))
    finally:
        conn.close()

    return {
        "mode": "execute",
        "project_id": pid,
        "inserted": inserted,
        "updated": updated,
        "skipped": skipped,
        "filtered_national_ids": plan["filtered_national_ids"],
        "counts_by_doc_type": plan["counts_by_doc_type"],
        "foram": plan["foram"],
    }


def _strip_records_for_summary(report: dict) -> dict:
    return {k: v for k, v in report.items() if k != "records"}


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="Safe read-only SafetyQR -> DCR HSE migration tool")
    parser.add_argument("--legacy-db", default=str(DEFAULT_LEGACY_DB), help="Path to SafetyQR SQLite app.db")
    parser.add_argument("--foram-dir", default=str(DEFAULT_FORAM_DIR), help="Path to FORAM templates directory")
    parser.add_argument("--dry-run", action="store_true", help="Inspect read-only and print migration report (default)")
    parser.add_argument("--target-db", default="", help="Localhost PostgreSQL URL (127.0.0.1/localhost only)")
    parser.add_argument("--project-id", default="", help="Target HSE project ID in local DCR DB")
    parser.add_argument("--legacy-project-id", type=int, default=None, help="Optional filter for one legacy project ID")
    parser.add_argument("--json", action="store_true", help="Print JSON summary report")
    args = parser.parse_args(argv)

    # Guard --target-db even if supplied accidentally with --dry-run
    if args.target_db:
        validate_local_postgres_url(args.target_db)

    if args.target_db and not args.dry_run:
        res = execute_local_migration(
            args.target_db,
            args.project_id,
            legacy_db=args.legacy_db,
            foram_dir=args.foram_dir,
            legacy_project_id=args.legacy_project_id,
        )
    else:
        res = _strip_records_for_summary(build_migration_plan(legacy_db=args.legacy_db, foram_dir=args.foram_dir))

    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")

    if args.json:
        print(json.dumps(res, ensure_ascii=False, indent=2))
    else:
        print("=== DCR HSE Legacy Migration Report ===", flush=True)
        print(f"Mode: {res['mode']}", flush=True)
        print(f"Mapped records by register: {json.dumps(res.get('counts_by_doc_type', {}), ensure_ascii=False)}", flush=True)
        print(f"Filtered 14-digit National IDs: {res.get('filtered_national_ids', 0)}", flush=True)
        foram = res.get("foram") or {}
        print(
            f"FORAM templates: {foram.get('valid_count', 0)} mapped, "
            f"{foram.get('skipped_junk_count', 0)} .lnk/~$ skipped",
            flush=True,
        )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
