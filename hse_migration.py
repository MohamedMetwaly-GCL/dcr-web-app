"""Read-only SafetyQR legacy inspection & idempotent DCR HSE migration utility.

Strict safety guarantees:
- Legacy SQLite source is opened ONLY with SQLite URI ``?mode=ro``.
- Legacy ``FORAM/`` and ``data/files/`` directories are walked strictly read-only, filtering ``.lnk`` and ``~$`` files.
- All worker identity data, phone numbers, and ID card/portrait photos are preserved intact inside RBAC-protected internal registers (``WKR``, etc.); public ``/qr/<token>`` remains strictly non-PII.
- Raw local Windows/server filesystem paths are stripped from migrated ``records.data`` and malformed dates (e.g. ``0026-10-27``) are sanitized to ISO ``YYYY-MM-DD``.
- Default mode is ``--dry-run`` (zero target database writes, zero Drive uploads).
- Optional ``--target-db`` / remote execution refuses any non-local PostgreSQL host unless BOTH ``--allow-remote-db`` and ``--confirm-production`` are passed.
- Direct Google Drive migration (``--upload-drive``) streams directly from legacy files on disk using resumable uploads (zero server disk storage).
"""

import argparse
from collections import Counter
from contextlib import closing
from datetime import datetime
import json
import mimetypes
import os
from pathlib import Path
import re
import sqlite3
import sys
from urllib.parse import quote, urlparse


DEFAULT_LEGACY_DB = Path(r"D:\SafetyQR_v9\SafetyQR\data\app.db")
DEFAULT_FORAM_DIR = Path(r"D:\SafetyQR_v9\FORAM")
DEFAULT_FILES_DIR = Path(r"D:\SafetyQR_v9\SafetyQR\data\files")
ALLOWED_LOCAL_HOSTS = frozenset({"127.0.0.1", "localhost", "::1"})

_NAT_ID_14_RE = re.compile(r"(?<!\d)\d{14}(?!\d)")
_PHONE_KEYS = frozenset({
    "phone",
    "mobile",
    "telephone",
    "tel",
    "contact_phone",
    "safety_officer_phone",
    "emergency_contact",
})
_NAT_ID_KEYS = frozenset({"national_id", "nationalid", "id_number", "nid", "ssn"})
_PII_MEDIA_KEYS = frozenset({
    "photo",
    "photo_path",
    "card_image",
    "national_id_card_img",
    "signature_path",
    "blood_type",
})

_DATE_TYPO_00_RE = re.compile(r"^00(\d{2}-\d{2}-\d{2})$")
_ISO_DATE_RE = re.compile(r"^\d{4}-\d{2}-\d{2}$")
_WIN_PATH_RE = re.compile(r"(?<![A-Za-z0-9_])[A-Za-z]:(?:\\|/(?!/))[^\s\"'<>|]+")
_LOCAL_SERVER_PATH_RE = re.compile(
    r"(?:data[\\/]files[\\/][^\s\"'<>|]+|(?:photos|HSE|EVACUATIONS|RISK_ASSESSMENTS|PER-\d+)[\\/][^\s\"'<>|]*\.(?:pdf|jpe?g|png|docx?|xlsx?|webp|txt|csv))",
    re.IGNORECASE,
)
_LOCAL_URL_SECRET_RE = re.compile(r"(postgres(?:ql)?://)([^@\s]+@)", re.IGNORECASE)
_BEARER_SECRET_RE = re.compile(r"(Bearer\s+)[A-Za-z0-9._\-]+", re.IGNORECASE)
_FOLDER_SAFE_CHARS_RE = re.compile(r'[<>:"/\\|?*\x00-\x1f]+')

LEGACY_PROJECT_DEFS = {
    1: {
        "suggested_id": "HSE-CFC-P2",
        "suggested_code": "HSE-CFC-P2",
        "aliases": ("cfc phaes 2", "cfc phase 2", "cfc"),
    },
    2: {
        "suggested_id": "HSE-EASTMAIN",
        "suggested_code": "HSE-EASTMAIN",
        "aliases": ("eastmain", "east main"),
    },
    6: {
        "suggested_id": "HSE-MUST",
        "suggested_code": "HSE-MUST",
        "aliases": ("جامعه مصر", "جامعة مصر", "must"),
    },
    7: {
        "suggested_id": "HSE-ANOPC",
        "suggested_code": "HSE-ANOPC",
        "aliases": ("anopc assuit", "anopc assiut", "anopc"),
    },
}

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


def _redact_sensitive_text(text) -> str:
    """Scrub embedded DB passwords or bearer tokens from diagnostic or exception strings."""
    cleaned = _LOCAL_URL_SECRET_RE.sub(r"\1***@", str(text or ""))
    return _BEARER_SECRET_RE.sub(r"\1[REDACTED]", cleaned)


def validate_local_postgres_url(url: str) -> str:
    """Ensure target DB URL points strictly to a local host (127.0.0.1, localhost, or ::1)."""
    return validate_target_postgres_url(url, allow_remote=False, confirm_production=False)


def validate_target_postgres_url(
    url: str,
    allow_remote: bool = False,
    confirm_production: bool = False,
) -> str:
    """Validate target PostgreSQL URL; refuse remote hosts without double-explicit approval."""
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
    if not host:
        raise ValueError("Refusing database URL with empty host.")
    if host not in ALLOWED_LOCAL_HOSTS:
        if not (allow_remote and confirm_production):
            raise ValueError(
                f"Refusing non-local database host '{host}'. "
                "Only 127.0.0.1 / localhost / ::1 is allowed unless BOTH --allow-remote-db and --confirm-production are set."
            )
    return raw


def resolve_target_db_url_from_env(explicit_url: str = "", env=None) -> str:
    """Resolve target PostgreSQL URL from explicit arg or in-process Railway/Postgres env vars."""
    raw = str(explicit_url or "").strip()
    if raw:
        return raw
    environ = env if env is not None else os.environ
    for key in ("DCR_MIGRATION_TARGET_DB", "DATABASE_PUBLIC_URL"):
        cand = str(environ.get(key) or "").strip()
        if cand:
            return cand

    tcp_domain = str(environ.get("RAILWAY_TCP_PROXY_DOMAIN") or "").strip()
    tcp_port = str(environ.get("RAILWAY_TCP_PROXY_PORT") or "").strip()
    pg_user = str(environ.get("PGUSER") or environ.get("POSTGRES_USER") or "").strip()
    pg_pass = str(environ.get("PGPASSWORD") or environ.get("POSTGRES_PASSWORD") or "").strip()
    pg_db = str(environ.get("PGDATABASE") or environ.get("POSTGRES_DB") or "").strip()
    if tcp_domain and tcp_port and pg_user and pg_db:
        auth_part = f"{quote(pg_user, safe='')}:{quote(pg_pass, safe='')}" if pg_pass else quote(pg_user, safe="")
        return f"postgresql://{auth_part}@{tcp_domain}:{tcp_port}/{quote(pg_db, safe='')}"

    fallback = str(environ.get("DATABASE_URL") or "").strip()
    return fallback


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


def sanitize_iso_date(value) -> tuple[str, int]:
    """Normalize ISO dates and repair legacy year-00xx typos such as ``0026-10-27``."""
    raw = str(value or "").strip()
    if not raw:
        return "", 0
    cand = raw[:10]
    fixed = 0
    typo_match = _DATE_TYPO_00_RE.match(cand)
    if typo_match:
        cand = f"20{typo_match.group(1)}"
        fixed = 1
    if not _ISO_DATE_RE.match(cand):
        return "", 1
    try:
        datetime.strptime(cand, "%Y-%m-%d").date()
    except ValueError:
        return "", 1
    return cand, fixed


def _strip_local_paths_from_text(value: str) -> tuple[str, int]:
    """Remove local Windows or SafetyQR server upload paths from text values."""
    txt = str(value or "").strip()
    if not txt:
        return "", 0
    low = txt.lower()
    if low.startswith("http://") or low.startswith("https://"):
        return txt, 0
    if _WIN_PATH_RE.fullmatch(txt) or _LOCAL_SERVER_PATH_RE.fullmatch(txt):
        return "", 1
    hits = 0
    cleaned = txt
    for pattern in (_WIN_PATH_RE, _LOCAL_SERVER_PATH_RE):
        found = pattern.findall(cleaned)
        if found:
            hits += len(found)
            cleaned = pattern.sub("", cleaned)
    if hits:
        cleaned = re.sub(r"\s{2,}", " ", cleaned).strip(" ,;-")
    return cleaned, hits


def _strip_local_paths_in_dict(data: dict) -> tuple[dict, int]:
    """Strip any local filesystem paths across all string fields in a record JSONB dict."""
    total_hits = 0
    cleaned = {}
    local_only_keys = frozenset({"filelocation", "source_file", "sourcefile", "file_path", "photo", "national_id_card_img"})
    for key, val in data.items():
        if str(key).lower() in local_only_keys:
            if isinstance(val, str) and val.strip():
                low = val.strip().lower()
                if not (low.startswith("http://") or low.startswith("https://")):
                    if _WIN_PATH_RE.search(val) or _LOCAL_SERVER_PATH_RE.search(val) or "\\" in val or "/" in val:
                        total_hits += 1
                        continue
        if isinstance(val, str):
            new_val, hits = _strip_local_paths_from_text(val)
            total_hits += hits
            cleaned[key] = new_val
        else:
            cleaned[key] = val
    return cleaned, total_hits


def _normalize_rel_attachment_path(raw_path: str) -> str:
    """Normalize a legacy attachment relative path, rejecting traversal outside ``files_dir``."""
    raw = str(raw_path or "").strip().replace("\\", "/")
    if not raw:
        return ""
    marker = "data/files/"
    idx = raw.lower().find(marker)
    if idx >= 0:
        raw = raw[idx + len(marker):]
    raw = raw.lstrip("/")
    if not raw or ":" in raw.split("/", 1)[0]:
        return ""
    parts = [p for p in raw.split("/") if p and p != "."]
    if not parts or any(p == ".." for p in parts):
        return ""
    return "/".join(parts)


def _normalize_project_match_key(value: str) -> str:
    """Normalize Arabic/Latin project names and codes for deterministic matching."""
    txt = str(value or "").strip().lower()
    if not txt:
        return ""
    txt = (
        txt.replace("ة", "ه")
        .replace("أ", "ا")
        .replace("إ", "ا")
        .replace("آ", "ا")
        .replace("ى", "ي")
    )
    txt = re.sub(r"[\s_\-]+", " ", txt)
    txt = re.sub(r"[^\w\u0600-\u06ff ]+", "", txt)
    return re.sub(r"\s+", " ", txt).strip()


def _slugify_project_code(name: str, legacy_id: int | str) -> str:
    """Generate a deterministic uppercase HSE project ID/code for unmapped projects."""
    ascii_slug = re.sub(r"[^A-Za-z0-9]+", "-", str(name or "").upper()).strip("-")
    if len(ascii_slug) >= 3:
        return f"HSE-{ascii_slug[:20].strip('-')}"
    return f"HSE-PRJ-{legacy_id}"


def build_project_mapping_plan(
    legacy_projects: list[dict],
    existing_hse_projects: list[dict] | None = None,
    records: list[dict] | None = None,
    attachments_by_record: dict[str, list[dict]] | None = None,
) -> list[dict]:
    """Match legacy SafetyQR projects against existing DCR HSE projects or propose new ones."""
    existing_list = list(existing_hse_projects or [])
    by_norm_key = {}
    for ex in existing_list:
        for cand in (ex.get("id"), ex.get("code"), ex.get("name")):
            nk = _normalize_project_match_key(cand)
            if nk:
                by_norm_key.setdefault(nk, ex)

    rec_counts_by_proj: dict = {}
    att_stats_by_proj: dict = {}
    for rec in records or []:
        lpid = rec.get("legacy_project_id")
        bucket = rec_counts_by_proj.setdefault(lpid, Counter())
        bucket[rec["dt_id"]] += 1
        att_list = (attachments_by_record or {}).get(rec["id"], [])
        if att_list:
            st = att_stats_by_proj.setdefault(lpid, {"count": 0, "bytes": 0, "missing": 0})
            for item in att_list:
                st["count"] += 1
                st["bytes"] += int(item.get("size_bytes") or 0)
                if not item.get("exists_on_disk", True):
                    st["missing"] += 1

    mapping_plan = []
    for lp in legacy_projects:
        lpid = lp.get("legacy_id")
        lname = str(lp.get("name") or f"Legacy Project {lpid}")
        known_def = LEGACY_PROJECT_DEFS.get(int(lpid)) if str(lpid).isdigit() else None
        suggested_id = known_def["suggested_id"] if known_def else _slugify_project_code(lname, lpid)
        suggested_code = known_def["suggested_code"] if known_def else suggested_id

        candidates = [lname, suggested_id, suggested_code]
        if known_def:
            candidates.extend(known_def.get("aliases", ()))

        matched = None
        for cand in candidates:
            nk = _normalize_project_match_key(cand)
            if nk and nk in by_norm_key:
                matched = by_norm_key[nk]
                break

        dt_counts = dict(sorted((rec_counts_by_proj.get(lpid) or {}).items()))
        att_st = att_stats_by_proj.get(lpid) or {"count": 0, "bytes": 0, "missing": 0}
        if matched:
            mapping_plan.append({
                "legacy_id": lpid,
                "legacy_name": lname,
                "action": "match_existing",
                "target_project_id": matched["id"],
                "target_code": matched.get("code") or matched["id"],
                "target_name": matched.get("name") or lname,
                "workspace_id": "HSE",
                "has_drive_folder": bool(matched.get("has_drive_folder")),
                "total_records": sum(dt_counts.values()),
                "counts_by_doc_type": dt_counts,
                "attachment_count": att_st["count"],
                "attachment_bytes": att_st["bytes"],
                "attachment_mb": round(att_st["bytes"] / (1024 * 1024), 2),
                "missing_attachment_count": att_st["missing"],
            })
        else:
            mapping_plan.append({
                "legacy_id": lpid,
                "legacy_name": lname,
                "action": "create_new",
                "target_project_id": suggested_id,
                "target_code": suggested_code,
                "target_name": lname,
                "workspace_id": "HSE",
                "has_drive_folder": False,
                "total_records": sum(dt_counts.values()),
                "counts_by_doc_type": dt_counts,
                "attachment_count": att_st["count"],
                "attachment_bytes": att_st["bytes"],
                "attachment_mb": round(att_st["bytes"] / (1024 * 1024), 2),
                "missing_attachment_count": att_st["missing"],
            })
    return mapping_plan


def fetch_existing_hse_projects_readonly(
    db_url: str = "",
    psycopg2_module=None,
) -> list[dict]:
    """Read ONLY existing HSE workspace projects using an explicit ``BEGIN READ ONLY`` transaction."""
    raw = resolve_target_db_url_from_env(db_url)
    if not raw:
        raise ValueError("No target DB URL supplied or present in DATABASE_PUBLIC_URL / RAILWAY_TCP_PROXY_* / DATABASE_URL.")
    parsed = urlparse(raw)
    scheme = (parsed.scheme or "").lower()
    if scheme not in ("postgresql", "postgres") or "," in (parsed.netloc or "") or not (parsed.hostname or "").strip():
        raise ValueError("Invalid PostgreSQL URL for read-only probe.")

    if psycopg2_module is None:
        import psycopg2 as psycopg2_module  # type: ignore

    try:
        conn = psycopg2_module.connect(raw)
    except Exception as exc:
        raise RuntimeError(f"Read-only DB probe connection failed: {_redact_sensitive_text(exc)}") from None

    rows_out = []
    try:
        with conn.cursor() as cur:
            cur.execute("BEGIN READ ONLY")
            cur.execute(
                """SELECT id, code, name,
                          COALESCE(workspace_id, 'ENG') AS workspace_id,
                          COALESCE(data->>'drive_folder_id', '') != '' AS has_drive_folder
                   FROM projects
                   WHERE UPPER(COALESCE(workspace_id, 'ENG')) = 'HSE'
                   ORDER BY id"""
            )
            for row in cur.fetchall():
                rows_out.append({
                    "id": row[0],
                    "code": row[1],
                    "name": row[2],
                    "workspace_id": row[3],
                    "has_drive_folder": bool(row[4]),
                })
            conn.rollback()
    except Exception as exc:
        try:
            conn.rollback()
        except Exception:
            pass
        raise RuntimeError(f"Read-only DB probe query failed: {_redact_sensitive_text(exc)}") from None
    finally:
        conn.close()
    return rows_out


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


def _inspect_attachment_on_disk(
    files_root: Path | None,
    files_root_resolved: Path | None,
    raw_path: str,
    file_name: str = "",
    include_photos: bool = True,
) -> dict:
    """Resolve and stat one legacy attachment safely under ``files_root`` (strictly read-only)."""
    rel_path = _normalize_rel_attachment_path(raw_path)
    clean_name = Path(str(file_name or rel_path or "attachment")).name
    mime_type = mimetypes.guess_type(clean_name or rel_path)[0] or "application/octet-stream"
    is_photo = rel_path.lower().startswith("photos/")
    if not rel_path or (is_photo and not include_photos):
        return {
            "rel_path": rel_path,
            "file_name": clean_name,
            "mime_type": mime_type,
            "exists_on_disk": False,
            "size_bytes": 0,
            "abs_path": "",
            "is_worker_photo": is_photo,
        }
    if not files_root or not files_root_resolved or not files_root.exists():
        return {
            "rel_path": rel_path,
            "file_name": clean_name,
            "mime_type": mime_type,
            "exists_on_disk": False,
            "size_bytes": 0,
            "abs_path": "",
            "is_worker_photo": is_photo,
        }
    try:
        candidate = (files_root / rel_path).resolve()
        candidate.relative_to(files_root_resolved)
    except Exception:
        return {
            "rel_path": rel_path,
            "file_name": clean_name,
            "mime_type": mime_type,
            "exists_on_disk": False,
            "size_bytes": 0,
            "abs_path": "",
            "is_worker_photo": is_photo,
        }
    if candidate.is_file():
        return {
            "rel_path": rel_path,
            "file_name": clean_name,
            "mime_type": mime_type,
            "exists_on_disk": True,
            "size_bytes": int(candidate.stat().st_size),
            "abs_path": str(candidate),
            "is_worker_photo": is_photo,
        }
    return {
        "rel_path": rel_path,
        "file_name": clean_name,
        "mime_type": mime_type,
        "exists_on_disk": False,
        "size_bytes": 0,
        "abs_path": "",
        "is_worker_photo": is_photo,
    }


def _inventory_disk_files(
    files_root: Path,
    files_root_resolved: Path | None,
    referenced_rel_paths: set[str],
    include_photos: bool = True,
) -> dict:
    """Walk ``files_root`` read-only to classify referenced photos and unreferenced disk files."""
    if not files_root.exists() or not files_root_resolved:
        return {
            "files_dir_exists": False,
            "disk_total_files": 0,
            "disk_total_bytes": 0,
            "disk_total_mb": 0.0,
            "pii_photos_excluded_count": 0,
            "pii_photos_excluded_bytes": 0,
            "unreferenced_disk_files_count": 0,
            "unreferenced_disk_bytes": 0,
        }
    ref_norm = {p.lower() for p in referenced_rel_paths if p}
    total_files = 0
    total_bytes = 0
    excluded_photo_count = 0
    excluded_photo_bytes = 0
    unref_count = 0
    unref_bytes = 0
    for dirpath, _dirs, filenames in os.walk(files_root):
        for fn in filenames:
            full = Path(dirpath) / fn
            try:
                rel = full.relative_to(files_root).as_posix()
                size = int(full.stat().st_size)
            except Exception:
                continue
            total_files += 1
            total_bytes += size
            rel_low = rel.lower()
            if rel_low.startswith("photos/") and not include_photos:
                excluded_photo_count += 1
                excluded_photo_bytes += size
            elif rel_low not in ref_norm:
                unref_count += 1
                unref_bytes += size
    return {
        "files_dir_exists": True,
        "disk_total_files": total_files,
        "disk_total_bytes": total_bytes,
        "disk_total_mb": round(total_bytes / (1024 * 1024), 2),
        "pii_photos_excluded_count": excluded_photo_count,
        "pii_photos_excluded_bytes": excluded_photo_bytes,
        "unreferenced_disk_files_count": unref_count,
        "unreferenced_disk_bytes": unref_bytes,
    }


def _build_empty_attachment_inventory(files_dir_path: Path) -> dict:
    return {
        "files_dir": str(files_dir_path),
        "files_dir_exists": files_dir_path.exists(),
        "mapped_files_count": 0,
        "mapped_total_bytes": 0,
        "mapped_total_mb": 0.0,
        "document_files_count": 0,
        "document_files_bytes": 0,
        "worker_photos_count": 0,
        "worker_photos_bytes": 0,
        "records_with_attachments_count": 0,
        "multi_file_records_count": 0,
        "missing_on_disk_count": 0,
        "missing_on_disk_files": [],
        "skipped_outside_catalog_count": 0,
        "skipped_outside_catalog_bytes": 0,
        "skipped_outside_catalog_by_category": {},
        "skipped_orphan_count": 0,
        "pii_photos_excluded_count": 0,
        "pii_photos_excluded_bytes": 0,
        "disk_total_files": 0,
        "disk_total_bytes": 0,
        "disk_total_mb": 0.0,
        "unreferenced_disk_files_count": 0,
        "unreferenced_disk_bytes": 0,
        "by_doc_type": {},
        "matched_records": [],
    }


def build_migration_plan(
    legacy_db: Path | str = DEFAULT_LEGACY_DB,
    foram_dir: Path | str = DEFAULT_FORAM_DIR,
    allow_missing_source: bool = False,
    files_dir: Path | str = DEFAULT_FILES_DIR,
    existing_hse_projects: list[dict] | None = None,
    redact_pii: bool = False,
    include_photos: bool = True,
) -> dict:
    """Inspect legacy SQLite DB, FORAM templates, and local attachments strictly read-only."""
    db_path = Path(legacy_db)
    files_root = Path(files_dir)
    files_root_resolved = files_root.resolve() if files_root.exists() else None
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
            "project_mapping_plan": [],
            "total_mapped_records": 0,
            "counts_by_doc_type": {},
            "skipped_outside_catalog": {},
            "filtered_national_ids": 0,
            "filtered_pii_fields": 0,
            "preserved_national_ids": 0,
            "preserved_pii_fields": 0,
            "sanitized_dates": 0,
            "sanitized_local_paths": 0,
            "foram": foram_report,
            "attachment_inventory": _build_empty_attachment_inventory(files_root),
            "record_attachments": {},
            "records": [],
        }

    records = []
    projects = []
    counts_by_dt = Counter()
    skipped_cats = Counter()
    filtered_nat_ids = 0
    filtered_pii_fields = 0
    preserved_nat_ids = 0
    preserved_pii_fields = 0
    sanitized_dates = 0
    sanitized_local_paths = 0
    mapped_doc_targets: dict[int, tuple[str, str, int | None]] = {}
    mapped_records_by_id: dict[str, dict] = {}
    worker_photo_attachments: list[tuple[str, str, int | None, str, int, str, str]] = []

    def _fix_date(val) -> str:
        nonlocal sanitized_dates
        clean, fixed = sanitize_iso_date(val)
        sanitized_dates += fixed
        return clean

    def _maybe_scrub_text(val) -> str:
        nonlocal filtered_nat_ids, preserved_nat_ids
        if redact_pii:
            cleaned, hits = scrub_text_national_ids(val)
            filtered_nat_ids += hits
            return cleaned
        txt = str(val or "").strip()
        if txt:
            preserved_nat_ids += len(_NAT_ID_14_RE.findall(txt))
        return txt

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
                    "start_date": _fix_date(row["start_date"]),
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
                title_clean = _maybe_scrub_text(row["title"])
                notes_clean = _maybe_scrub_text(row["notes"])
                rec_data = {
                    "docNo": str(row["number"] or f"LEG-{dt_code}-{row['id']}"),
                    "legacySubtype": subtype,
                    "title": title_clean or subtype,
                    "issuedDate": _fix_date(row["doc_date"]),
                    "location": str(row["location"] or ""),
                    "responsible": str(row["responsible"] or ""),
                    "status": _norm_status(row["status"]),
                    "remarks": notes_clean,
                }
                if dt_code == "PTW" and row["valid_until"]:
                    vu = _fix_date(row["valid_until"])
                    if vu:
                        rec_data["validUntil"] = vu
                if dt_code == "INC":
                    rec_data["injuryClass"] = _norm_injury_class(row["injury_type"])
                    rec_data["lostDays"] = int(row["lost_days"] or 0)
                if row["closed_at"]:
                    cd = _fix_date(row["closed_at"])
                    if cd:
                        rec_data["closedDate"] = cd

                rec_data, path_hits = _strip_local_paths_in_dict(rec_data)
                sanitized_local_paths += path_hits
                rid = f"leg-{dt_code.lower()}-doc-{row['id']}"
                rec_obj = {
                    "id": rid,
                    "legacy_project_id": row["project_id"],
                    "dt_id": dt_code,
                    "data": rec_data,
                }
                records.append(rec_obj)
                mapped_records_by_id[rid] = rec_obj
                mapped_doc_targets[int(row["id"])] = (rid, dt_code, row["project_id"])
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
                desc_clean = _maybe_scrub_text(row["incident_description"])
                rc_clean = _maybe_scrub_text(row["root_cause"])
                ca_clean = _maybe_scrub_text(row["corrective_action"])
                rec_data = {
                    "docNo": f"INC-INJ-{row['id']:04d}",
                    "legacySubtype": "Injury Register",
                    "title": desc_clean[:120] if desc_clean else f"Injury Incident #{row['id']}",
                    "issuedDate": _fix_date(row["incident_date"]),
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
                rec_data, path_hits = _strip_local_paths_in_dict(rec_data)
                sanitized_local_paths += path_hits
                rid = f"leg-inc-inj-{row['id']}"
                rec_obj = {
                    "id": rid,
                    "legacy_project_id": row["project_id"],
                    "dt_id": "INC",
                    "data": rec_data,
                }
                records.append(rec_obj)
                mapped_records_by_id[rid] = rec_obj
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
                    rd = _fix_date(row["record_date"])
                    if rd:
                        cur_entry["last_inspection"] = rd
                if row["expiry_date"]:
                    ed = _fix_date(row["expiry_date"])
                    if ed:
                        cur_entry["expiry_date"] = ed

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
                    last_insp = _fix_date(raw.get("last_inspection")) or item_dates.get("last_inspection") or ""
                    exp_date = _fix_date(raw.get("expiry_date")) or item_dates.get("expiry_date") or ""
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
                        "manufactureDate": _fix_date(raw.get("manufacture_date")),
                        "extinguisherStatus": ext_status,
                        "conditionNotes": str(raw.get("notes") or ""),
                    }
                    rec_data, path_hits = _strip_local_paths_in_dict(rec_data)
                    sanitized_local_paths += path_hits
                    rid = f"leg-fei-item-{row['id']}"
                    rec_obj = {
                        "id": rid,
                        "legacy_project_id": row["project_id"],
                        "dt_id": "FEI",
                        "data": rec_data,
                    }
                    records.append(rec_obj)
                    mapped_records_by_id[rid] = rec_obj
                    counts_by_dt["FEI"] += 1

                elif itype == "person":
                    for k, v in raw.items():
                        k_low = str(k).lower()
                        has_val = bool(str(v or "").strip())
                        if redact_pii:
                            if k_low in _NAT_ID_KEYS:
                                if has_val or is_synthetic_national_id(v):
                                    filtered_nat_ids += 1
                            elif k_low in _PHONE_KEYS or k_low in _PII_MEDIA_KEYS:
                                if has_val:
                                    filtered_pii_fields += 1
                            elif is_synthetic_national_id(v):
                                filtered_nat_ids += 1
                        else:
                            if k_low in _NAT_ID_KEYS and (has_val or is_synthetic_national_id(v)):
                                preserved_nat_ids += 1
                            elif (k_low in _PHONE_KEYS or k_low in _PII_MEDIA_KEYS) and has_val:
                                preserved_pii_fields += 1

                    raw_role = str(raw.get("person_type") or "").strip()
                    raw_aff = str(raw.get("affiliation") or "").strip()
                    raw_notes = str(raw.get("notes") or "").strip()
                    notes_clean = _maybe_scrub_text(raw_notes)
                    aff_source = raw_aff or raw_notes
                    aff_en = _AFFILIATION_MAP.get(
                        aff_source,
                        "Contractor" if ("مقاول" in aff_source or "يوميه" in aff_source or "يومية" in aff_source) else "",
                    )
                    role_en = _PERSON_TYPE_MAP.get(raw_role, "Worker" if raw_role == "" else raw_role)
                    name_clean = _maybe_scrub_text(raw.get("name") or row["title"] or row["code"])

                    rec_data = {
                        "docNo": str(row["code"] or f"WKR-{row['id']:04d}"),
                        "title": name_clean,
                        "affiliation": aff_en,
                        "company": str(raw.get("company") or ""),
                        "personType": role_en,
                        "jobTitle": str(raw.get("job_title") or ""),
                        "hireDate": _fix_date(raw.get("hire_date")),
                        "remarks": notes_clean,
                    }
                    if not redact_pii:
                        nat_val = str(raw.get("national_id") or raw.get("id_number") or "").strip()
                        if nat_val:
                            rec_data["nationalId"] = nat_val
                        if str(raw.get("phone") or "").strip():
                            rec_data["phone"] = str(raw["phone"]).strip()
                        if str(raw.get("emergency_contact") or "").strip():
                            rec_data["emergencyContact"] = str(raw["emergency_contact"]).strip()
                        if str(raw.get("safety_officer_phone") or "").strip():
                            rec_data["safetyOfficerPhone"] = str(raw["safety_officer_phone"]).strip()
                        if str(raw.get("blood_type") or "").strip():
                            rec_data["bloodType"] = str(raw["blood_type"]).strip()
                        if str(raw.get("contractor_id") or "").strip():
                            rec_data["contractorId"] = str(raw["contractor_id"]).strip()

                    rec_data, path_hits = _strip_local_paths_in_dict(rec_data)
                    sanitized_local_paths += path_hits
                    rid = f"leg-wkr-item-{row['id']}"
                    rec_obj = {
                        "id": rid,
                        "legacy_project_id": row["project_id"],
                        "dt_id": "WKR",
                        "data": rec_data,
                    }
                    records.append(rec_obj)
                    mapped_records_by_id[rid] = rec_obj
                    counts_by_dt["WKR"] += 1

                    if include_photos:
                        for media_key in ("photo", "national_id_card_img"):
                            raw_media = str(raw.get(media_key) or "").strip()
                            if raw_media:
                                worker_photo_attachments.append(
                                    (rid, "WKR", row["project_id"], f"items_{media_key}", row["id"], Path(raw_media).name, raw_media)
                                )

        # 5. risk_assessments -> JSA (keep source_file only in attachment inventory, never in rec_data)
        ra_attachments = []
        if "risk_assessments" in tables:
            ra_cols = {r[1] for r in cur.execute("PRAGMA table_info(risk_assessments)").fetchall()}
            src_expr = "source_file" if "source_file" in ra_cols else "'' AS source_file"
            for row in cur.execute(
                f"""SELECT id, project_id, code, title, activity_name, department, location,
                           assessment_date, review_date, assessor_name, approver_name, status, notes, {src_expr}
                    FROM risk_assessments ORDER BY id"""
            ):
                st_raw = str(row["status"] or "").strip().lower()
                st_en = _JSA_STATUS_MAP.get(st_raw, "Draft")
                rec_data = {
                    "docNo": str(row["code"] or f"JSA-{row['id']:04d}"),
                    "title": str(row["title"] or row["activity_name"] or "Risk Assessment"),
                    "issuedDate": _fix_date(row["assessment_date"]),
                    "reviewDate": _fix_date(row["review_date"]),
                    "location": str(row["location"] or ""),
                    "activityName": str(row["activity_name"] or ""),
                    "department": str(row["department"] or ""),
                    "assessorName": str(row["assessor_name"] or ""),
                    "approverName": str(row["approver_name"] or ""),
                    "status": st_en,
                    "remarks": str(row["notes"] or ""),
                }
                rec_data, path_hits = _strip_local_paths_in_dict(rec_data)
                sanitized_local_paths += path_hits
                rid = f"leg-jsa-{row['id']}"
                rec_obj = {
                    "id": rid,
                    "legacy_project_id": row["project_id"],
                    "dt_id": "JSA",
                    "data": rec_data,
                }
                records.append(rec_obj)
                mapped_records_by_id[rid] = rec_obj
                counts_by_dt["JSA"] += 1
                raw_src = str(row["source_file"] or "").strip()
                if raw_src:
                    ra_attachments.append((rid, "JSA", row["project_id"], "risk_assessments", row["id"], Path(raw_src).name, raw_src))

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
                    "issuedDate": _fix_date(row["eval_date"]),
                    "evaluator": str(row["evaluator"] or ""),
                    "positiveNotes": str(row["positive_notes"] or ""),
                    "negativeNotes": str(row["negative_notes"] or ""),
                    "actionPlan": str(row["action_plan"] or ""),
                }
                rec_data, path_hits = _strip_local_paths_in_dict(rec_data)
                sanitized_local_paths += path_hits
                rid = f"leg-cev-{row['id']}"
                rec_obj = {
                    "id": rid,
                    "legacy_project_id": row["project_id"],
                    "dt_id": "CEV",
                    "data": rec_data,
                }
                records.append(rec_obj)
                mapped_records_by_id[rid] = rec_obj
                counts_by_dt["CEV"] += 1

        # 7. evacuation_drills -> EVAC (keep source_file only in attachment inventory, never in rec_data)
        evac_attachments = []
        if "evacuation_drills" in tables:
            ev_cols = {r[1] for r in cur.execute("PRAGMA table_info(evacuation_drills)").fetchall()}
            src_expr = "source_file" if "source_file" in ev_cols else "'' AS source_file"
            for row in cur.execute(
                f"""SELECT id, project_id, title, drill_date, drill_time, scenario, location,
                           assembly_point, leader_name, target_duration_sec, actual_duration_sec,
                           evaluation_grade, alarm_working, exit_routes_clear, assembly_adequate,
                           corrective_actions, notes, {src_expr}
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
                    "issuedDate": _fix_date(row["drill_date"]),
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
                rec_data, path_hits = _strip_local_paths_in_dict(rec_data)
                sanitized_local_paths += path_hits
                rid = f"leg-evac-{row['id']}"
                rec_obj = {
                    "id": rid,
                    "legacy_project_id": row["project_id"],
                    "dt_id": "EVAC",
                    "data": rec_data,
                }
                records.append(rec_obj)
                mapped_records_by_id[rid] = rec_obj
                counts_by_dt["EVAC"] += 1
                raw_src = str(row["source_file"] or "").strip()
                if raw_src:
                    evac_attachments.append((rid, "EVAC", row["project_id"], "evacuation_drills", row["id"], Path(raw_src).name, raw_src))

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
                rid = f"leg-mhr-{row['project_id'] or 0}-{yr}-{mo:02d}"
                rec_obj = {
                    "id": rid,
                    "legacy_project_id": row["project_id"],
                    "dt_id": "MHR",
                    "data": rec_data,
                }
                records.append(rec_obj)
                mapped_records_by_id[rid] = rec_obj
                counts_by_dt["MHR"] += 1

        # 9. Read-only Legacy Attachment Inventory (documents + worker photos/ID scans)
        record_attachments: dict[str, list[dict]] = {}
        seen_rec_rel_keys: set[tuple[str, str]] = set()
        referenced_rel_paths: set[str] = set()
        skipped_out_by_cat: dict[str, dict] = {}
        skipped_orphan_count = 0
        missing_on_disk: list[dict] = []
        att_by_dt_counts = Counter()
        att_by_dt_bytes = Counter()

        def _register_mapped_attachment(
            record_id: str,
            dt_code: str,
            legacy_proj_id,
            source_table: str,
            source_row_id,
            file_name: str,
            raw_rel_path: str,
        ) -> None:
            info = _inspect_attachment_on_disk(
                files_root,
                files_root_resolved,
                raw_rel_path,
                file_name=file_name,
                include_photos=include_photos,
            )
            if info["rel_path"]:
                referenced_rel_paths.add(info["rel_path"])
            if not include_photos and info.get("is_worker_photo"):
                return
            dedup_key = (record_id, (info["rel_path"] or info["file_name"]).lower())
            if dedup_key in seen_rec_rel_keys:
                return
            seen_rec_rel_keys.add(dedup_key)
            entry = {
                "record_id": record_id,
                "dt_id": dt_code,
                "legacy_project_id": legacy_proj_id,
                "source_table": source_table,
                "source_row_id": source_row_id,
                "file_name": info["file_name"],
                "rel_path": info["rel_path"],
                "abs_path": info["abs_path"],
                "mime_type": info["mime_type"],
                "exists_on_disk": info["exists_on_disk"],
                "size_bytes": info["size_bytes"],
                "is_worker_photo": bool(info.get("is_worker_photo")),
            }
            record_attachments.setdefault(record_id, []).append(entry)
            att_by_dt_counts[dt_code] += 1
            att_by_dt_bytes[dt_code] += info["size_bytes"]
            if not info["exists_on_disk"]:
                missing_on_disk.append({
                    "record_id": record_id,
                    "source_table": source_table,
                    "source_row_id": source_row_id,
                    "rel_path": info["rel_path"],
                    "file_name": info["file_name"],
                })

        if "doc_files" in tables:
            doc_cat_by_id = {}
            if "docs" in tables:
                for drow in cur.execute("SELECT id, category FROM docs"):
                    doc_cat_by_id[int(drow["id"])] = str(drow["category"] or "unknown").strip() or "unknown"
            for row in cur.execute("SELECT id, doc_id, file_name, file_path FROM doc_files ORDER BY id"):
                did = row["doc_id"]
                target = mapped_doc_targets.get(int(did)) if did is not None else None
                if target:
                    rid, dt_code, lpid = target
                    _register_mapped_attachment(
                        rid, dt_code, lpid, "doc_files", row["id"], row["file_name"], row["file_path"]
                    )
                else:
                    info = _inspect_attachment_on_disk(
                        files_root,
                        files_root_resolved,
                        row["file_path"],
                        file_name=row["file_name"],
                        include_photos=include_photos,
                    )
                    if info["rel_path"]:
                        referenced_rel_paths.add(info["rel_path"])
                    cat = doc_cat_by_id.get(int(did), "orphan") if did is not None else "orphan"
                    if cat == "orphan":
                        skipped_orphan_count += 1
                    else:
                        b = skipped_out_by_cat.setdefault(cat, {"count": 0, "bytes": 0})
                        b["count"] += 1
                        b["bytes"] += info["size_bytes"]

        if "record_files" in tables and "records" in tables and "items" in tables:
            for row in cur.execute(
                """SELECT rf.id, rf.record_id, rf.file_name, rf.file_path,
                          r.item_id, i.type AS item_type, i.project_id
                   FROM record_files rf
                   LEFT JOIN records r ON rf.record_id = r.id
                   LEFT JOIN items i ON r.item_id = i.id
                   ORDER BY rf.id"""
            ):
                iid = row["item_id"]
                itype = str(row["item_type"] or "").strip().lower()
                if iid is not None and itype == "person":
                    rid = f"leg-wkr-item-{iid}"
                    dt_code = "WKR"
                elif iid is not None and itype == "extinguisher":
                    rid = f"leg-fei-item-{iid}"
                    dt_code = "FEI"
                else:
                    rid = ""
                    dt_code = ""
                if rid and rid in mapped_records_by_id:
                    _register_mapped_attachment(
                        rid, dt_code, row["project_id"], "record_files", row["id"], row["file_name"], row["file_path"]
                    )
                else:
                    skipped_orphan_count += 1

        for rid, dt_code, lpid, src_tbl, row_id, fname, raw_src in (
            ra_attachments + evac_attachments + worker_photo_attachments
        ):
            _register_mapped_attachment(rid, dt_code, lpid, src_tbl, row_id, fname, raw_src)

    disk_stats = _inventory_disk_files(files_root, files_root_resolved, referenced_rel_paths, include_photos=include_photos)
    all_mapped_items = [item for items in record_attachments.values() for item in items]
    mapped_files_count = len(all_mapped_items)
    mapped_total_bytes = sum(item["size_bytes"] for item in all_mapped_items)
    worker_photos_items = [item for item in all_mapped_items if item.get("is_worker_photo")]
    doc_items = [item for item in all_mapped_items if not item.get("is_worker_photo")]
    multi_file_records_count = sum(1 for items in record_attachments.values() if len(items) > 1)
    by_dt_inv = {
        dt: {
            "files": int(att_by_dt_counts[dt]),
            "bytes": int(att_by_dt_bytes[dt]),
            "mb": round(att_by_dt_bytes[dt] / (1024 * 1024), 2),
        }
        for dt in sorted(att_by_dt_counts)
    }
    matched_record_summaries = [
        {
            "record_id": rid,
            "dt_id": items[0]["dt_id"],
            "legacy_project_id": items[0]["legacy_project_id"],
            "file_count": len(items),
            "total_bytes": sum(i["size_bytes"] for i in items),
            "missing_count": sum(1 for i in items if not i["exists_on_disk"]),
            "source_tables": sorted({i["source_table"] for i in items}),
        }
        for rid, items in sorted(record_attachments.items())
    ]

    attachment_inventory = {
        "files_dir": str(files_root),
        "files_dir_exists": disk_stats["files_dir_exists"],
        "mapped_files_count": mapped_files_count,
        "mapped_total_bytes": mapped_total_bytes,
        "mapped_total_mb": round(mapped_total_bytes / (1024 * 1024), 2),
        "document_files_count": len(doc_items),
        "document_files_bytes": sum(i["size_bytes"] for i in doc_items),
        "worker_photos_count": len(worker_photos_items),
        "worker_photos_bytes": sum(i["size_bytes"] for i in worker_photos_items),
        "records_with_attachments_count": len(record_attachments),
        "multi_file_records_count": multi_file_records_count,
        "missing_on_disk_count": len(missing_on_disk),
        "missing_on_disk_files": missing_on_disk,
        "skipped_outside_catalog_count": sum(v["count"] for v in skipped_out_by_cat.values()),
        "skipped_outside_catalog_bytes": sum(v["bytes"] for v in skipped_out_by_cat.values()),
        "skipped_outside_catalog_by_category": skipped_out_by_cat,
        "skipped_orphan_count": skipped_orphan_count,
        "pii_photos_excluded_count": disk_stats["pii_photos_excluded_count"],
        "pii_photos_excluded_bytes": disk_stats["pii_photos_excluded_bytes"],
        "disk_total_files": disk_stats["disk_total_files"],
        "disk_total_bytes": disk_stats["disk_total_bytes"],
        "disk_total_mb": disk_stats["disk_total_mb"],
        "unreferenced_disk_files_count": disk_stats["unreferenced_disk_files_count"],
        "unreferenced_disk_bytes": disk_stats["unreferenced_disk_bytes"],
        "by_doc_type": by_dt_inv,
        "matched_records": matched_record_summaries,
    }

    project_mapping_plan = build_project_mapping_plan(
        projects,
        existing_hse_projects=existing_hse_projects,
        records=records,
        attachments_by_record=record_attachments,
    )

    return {
        "mode": "dry-run",
        "read_only": True,
        "generated_at": datetime.utcnow().isoformat(timespec="seconds") + "Z",
        "source_db": str(db_path),
        "source_exists": True,
        "projects": projects,
        "project_mapping_plan": project_mapping_plan,
        "total_mapped_records": len(records),
        "counts_by_doc_type": dict(sorted(counts_by_dt.items())),
        "skipped_outside_catalog": dict(sorted(skipped_cats.items())),
        "filtered_national_ids": filtered_nat_ids,
        "filtered_pii_fields": filtered_pii_fields,
        "preserved_national_ids": preserved_nat_ids,
        "preserved_pii_fields": preserved_pii_fields,
        "sanitized_dates": sanitized_dates,
        "sanitized_local_paths": sanitized_local_paths,
        "foram": foram_report,
        "attachment_inventory": attachment_inventory,
        "record_attachments": record_attachments,
        "records": records,
    }


def _seed_hse_project_with_cursor(cur, pid: str) -> None:
    """Seed all 11 English HSE registers, columns, and dropdown lists inside an active transaction."""
    from hse_defaults import HSE_LISTS, HSE_REGISTERS, HSE_SUBTYPES, PTW_SUBTYPE_CODES, hse_columns

    for order, (code, name, source) in enumerate(HSE_REGISTERS):
        dt_data = {"legacy_source": source, "legacy_subtypes": HSE_SUBTYPES.get(code, {})}
        cur.execute(
            """INSERT INTO doc_types(id, project_id, name, code, sort_order, data)
               VALUES(%s, %s, %s, %s, %s, %s::jsonb) ON CONFLICT DO NOTHING""",
            (code, pid, name, code, order, json.dumps(dt_data, ensure_ascii=False)),
        )
        for sort_order, (key, label, col_type, list_name) in enumerate(hse_columns(code)):
            cur.execute(
                """INSERT INTO columns_config(project_id, dt_id, col_key, label, col_type, list_name, visible, sort_order)
                   VALUES(%s, %s, %s, %s, %s, %s, true, %s) ON CONFLICT DO NOTHING""",
                (pid, code, key, label, col_type, list_name, sort_order),
            )
    for list_name, values in HSE_LISTS.items():
        for order, value in enumerate(values):
            meta = None
            if list_name == "hse_doc_status":
                meta = "pending" if value == "Open" else ("approved" if value == "Closed" else None)
            elif list_name == "hse_ptw_subtype":
                meta = PTW_SUBTYPE_CODES.get(value)
            elif list_name == "hse_extinguisher_status":
                meta = (
                    "approved"
                    if value == "Valid"
                    else ("rejected" if value in ("Expired", "Out of Service") else "pending")
                )
            elif list_name == "hse_jsa_status":
                meta = "approved" if value == "Approved" else "pending"
            cur.execute(
                """INSERT INTO dropdown_lists(project_id, list_name, item_value, sort_order, meta)
                   VALUES(%s, %s, %s, %s, %s)
                   ON CONFLICT(project_id, list_name, item_value)
                   DO UPDATE SET meta = COALESCE(dropdown_lists.meta, EXCLUDED.meta)""",
                (pid, list_name, value, order, meta),
            )


def _assign_hse_admins_with_cursor(cur, pid: str) -> int:
    """Grant access to active superadmins and HSE-workspace admins cleanly (no cross-workspace leak)."""
    cur.execute(
        """INSERT INTO user_projects(username, project_id)
           SELECT u.username, %s
           FROM users u
           LEFT JOIN user_workspaces uw
             ON uw.username = u.username AND UPPER(uw.workspace_id) = 'HSE'
           WHERE LOWER(COALESCE(u.role, '')) = 'superadmin'
              OR LOWER(COALESCE(uw.role, '')) = 'admin'
              OR (
                  uw.username IS NULL
                  AND LOWER(COALESCE(u.role, '')) = 'admin'
                  AND UPPER(COALESCE(u.primary_workspace, 'ENG')) = 'HSE'
              )
           ON CONFLICT DO NOTHING""",
        (pid,),
    )
    return int(cur.rowcount or 0)


def _fetch_existing_hse_projects_from_cursor(cur) -> list[dict]:
    cur.execute(
        """SELECT id, code, name,
                  COALESCE(workspace_id, 'ENG') AS workspace_id,
                  COALESCE(data->>'drive_folder_id', '') != '' AS has_drive_folder
           FROM projects
           WHERE UPPER(COALESCE(workspace_id, 'ENG')) = 'HSE'
           ORDER BY id"""
    )
    return [
        {
            "id": r[0],
            "code": r[1],
            "name": r[2],
            "workspace_id": r[3],
            "has_drive_folder": bool(r[4]),
        }
        for r in cur.fetchall()
    ]


def _ensure_target_hse_project(cur, mapping_entry: dict) -> tuple[str, bool, int]:
    """Ensure target HSE project exists, seeding registers & assigning admins when created."""
    if mapping_entry["action"] == "match_existing":
        return mapping_entry["target_project_id"], False, 0

    base_id = mapping_entry["target_project_id"]
    cand_id = base_id
    suffix = 1
    while True:
        cur.execute("SELECT COALESCE(workspace_id, 'ENG') FROM projects WHERE id = %s", (cand_id,))
        existing = cur.fetchone()
        if not existing:
            break
        if str(existing[0] or "ENG").strip().upper() == "HSE":
            return cand_id, False, 0
        suffix += 1
        cand_id = f"{base_id}-HSE" if suffix == 2 else f"{base_id}-HSE-{suffix}"

    proj_data = {
        "legacy_project_id": mapping_entry.get("legacy_id"),
        "migrated_from": "SafetyQR",
    }
    cur.execute(
        """INSERT INTO projects(id, name, code, data, workspace_id)
           VALUES(%s, %s, %s, %s::jsonb, 'HSE')""",
        (cand_id, mapping_entry["target_name"], mapping_entry["target_code"], json.dumps(proj_data, ensure_ascii=False)),
    )
    _seed_hse_project_with_cursor(cur, cand_id)
    assigned = _assign_hse_admins_with_cursor(cur, cand_id)
    return cand_id, True, assigned


def _upsert_records_for_project(cur, pid: str, selected_recs: list[dict]) -> tuple[int, int, int]:
    """Idempotently upsert legacy records into one HSE project while preserving existing Drive attachments."""
    inserted = 0
    updated = 0
    skipped = 0

    for rec in selected_recs:
        rid = rec["id"]
        dt_id = rec["dt_id"]
        payload = dict(rec["data"])
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

        cur.execute("SELECT project_id, data FROM records WHERE id = %s", (rid,))
        existing = cur.fetchone()
        if existing:
            if existing[0] != pid:
                skipped += 1
                continue
            existing_data = existing[1] if isinstance(existing[1], dict) else _safe_json_loads(existing[1])
            merged = dict(payload)
            if isinstance(existing_data.get("attachments"), dict):
                merged["attachments"] = existing_data["attachments"]
            existing_loc = str(existing_data.get("fileLocation") or "").strip()
            if existing_loc.lower().startswith(("http://", "https://")) and not merged.get("fileLocation"):
                merged["fileLocation"] = existing_loc
            cur.execute(
                """UPDATE records
                   SET dt_id = %s, data = %s::jsonb, updated_at = NOW()
                   WHERE id = %s AND project_id = %s""",
                (dt_id, json.dumps(merged, ensure_ascii=False), rid, pid),
            )
            updated += 1
        else:
            cur.execute(
                """INSERT INTO records(id, project_id, dt_id, data)
                   VALUES(%s, %s, %s, %s::jsonb)""",
                (rid, pid, dt_id, json.dumps(payload, ensure_ascii=False)),
            )
            inserted += 1

    cur.execute("DELETE FROM dashboard_cache WHERE project_id = %s", (pid,))
    return inserted, updated, skipped


class _CursorRoutingStore:
    """Thin DB store adapter over a live psycopg2 connection for ``drive_routing``."""

    def __init__(self, conn):
        self._conn = conn

    def get_project(self, pid: str) -> dict | None:
        with self._conn.cursor() as cur:
            cur.execute(
                "SELECT id, name, code, COALESCE(workspace_id, 'ENG'), data FROM projects WHERE id = %s",
                (pid,),
            )
            row = cur.fetchone()
        if not row:
            return None
        data = row[4] if isinstance(row[4], dict) else _safe_json_loads(row[4])
        return {"id": row[0], "name": row[1], "code": row[2], "workspace_id": row[3], **data}

    def get_doc_type(self, pid: str, dt_id: str) -> dict | None:
        with self._conn.cursor() as cur:
            cur.execute(
                "SELECT id, project_id, name, code, sort_order, data FROM doc_types WHERE id = %s AND project_id = %s",
                (dt_id, pid),
            )
            row = cur.fetchone()
        if not row:
            return None
        data = row[5] if isinstance(row[5], dict) else _safe_json_loads(row[5])
        return {"id": row[0], "project_id": row[1], "name": row[2], "code": row[3], "sort_order": row[4], **data}

    def set_doc_type_drive_cache(self, pid: str, dt_id: str, patch: dict) -> bool:
        allowed = {"drive_folder_id", "drive_folder_name", "drive_slot_folders", "folder_synced_at"}
        filtered = {k: v for k, v in (patch or {}).items() if k in allowed}
        if not filtered:
            return False
        with self._conn.cursor() as cur:
            cur.execute(
                """UPDATE doc_types
                   SET data = COALESCE(data, '{}'::jsonb) || %s::jsonb
                   WHERE id = %s AND project_id = %s
                   RETURNING id""",
                (json.dumps(filtered), dt_id, pid),
            )
            updated = cur.fetchone()
        self._conn.commit()
        return bool(updated)


def _extract_drive_folder_id(val: str) -> str:
    raw = str(val or "").strip()
    if not raw:
        return ""
    m = re.search(r"/folders/([a-zA-Z0-9_-]+)", raw)
    return m.group(1) if m else raw


def _safe_drive_subfolder_name(doc_no: str, rid: str) -> str:
    raw_doc = str(doc_no or "").strip()
    clean_doc = _FOLDER_SAFE_CHARS_RE.sub("_", raw_doc).strip(" .")
    clean_rid = _FOLDER_SAFE_CHARS_RE.sub("_", str(rid or "")).strip(" .")
    if clean_doc and clean_rid and clean_rid not in clean_doc:
        combined = f"{clean_doc[:60]} ({clean_rid})"
        return combined[:80]
    return (clean_doc or clean_rid or "Attachments")[:80]


def _ensure_project_drive_folder(
    conn,
    store: _CursorRoutingStore,
    pid: str,
    service,
    drive_routing_mod,
    drive_root_id: str = "",
) -> str:
    proj = store.get_project(pid) or {}
    existing_fid = _extract_drive_folder_id(proj.get("drive_folder_id", ""))
    if existing_fid:
        return existing_fid

    parent_root = _extract_drive_folder_id(drive_root_id)
    if not parent_root:
        raise ValueError(
            f"Project '{pid}' has no drive_folder_id configured; pass --drive-root-id or set project Drive folder first."
        )
    folder_label = f"{proj.get('code') or pid} - {proj.get('name') or pid}".strip(" -")
    children = drive_routing_mod.list_child_folders(service, parent_root)
    wanted_keys = {
        _normalize_project_match_key(folder_label),
        _normalize_project_match_key(proj.get("code") or ""),
        _normalize_project_match_key(proj.get("name") or ""),
    } - {""}
    matched = next((c for c in children if _normalize_project_match_key(c.get("name", "")) in wanted_keys), None)
    if not matched:
        matched = drive_routing_mod.create_folder(service, parent_root, folder_label)
    new_fid = str(matched["id"])
    with conn.cursor() as cur:
        cur.execute(
            """UPDATE projects
               SET data = COALESCE(data, '{}'::jsonb) || %s::jsonb
               WHERE id = %s""",
            (json.dumps({"drive_folder_id": new_fid}), pid),
        )
    conn.commit()
    return new_fid


def _upload_project_attachments_to_drive(
    conn,
    pid: str,
    selected_recs: list[dict],
    record_attachments: dict[str, list[dict]],
    service,
    token: str,
    store: _CursorRoutingStore,
    drive_root_id: str = "",
    http_module=None,
    sleep_fn=None,
) -> dict:
    """Stream-upload local legacy files directly to Google Drive (zero server temp files)."""
    import attachments as att_mod
    import drive_routing
    import drive_upload

    recs_with_att = [r for r in selected_recs if record_attachments.get(r["id"])]
    if not recs_with_att:
        return {"uploaded_files": 0, "skipped_existing": 0, "missing_files": 0, "uploaded_bytes": 0}

    _ensure_project_drive_folder(conn, store, pid, service, drive_routing, drive_root_id=drive_root_id)

    http_client = http_module
    kwargs_http = {"http": http_client} if http_client is not None else {}
    kwargs_stream = dict(kwargs_http)
    if sleep_fn is not None:
        kwargs_stream["sleep"] = sleep_fn

    uploaded_files = 0
    skipped_existing = 0
    missing_files = 0
    uploaded_bytes = 0

    for rec in recs_with_att:
        rid = rec["id"]
        dt_id = rec["dt_id"]
        doc_no = str(rec["data"].get("docNo") or rid)
        raw_items = record_attachments.get(rid, [])
        uploadable = [it for it in raw_items if it.get("exists_on_disk") and it.get("abs_path")]
        missing_files += len(raw_items) - len(uploadable)
        if not uploadable:
            continue

        with conn.cursor() as cur:
            cur.execute("SELECT data FROM records WHERE id = %s AND project_id = %s", (rid, pid))
            row = cur.fetchone()
        cur_data = (row[0] if isinstance(row[0], dict) else _safe_json_loads(row[0])) if row else dict(rec["data"])
        existing_att = cur_data.get("attachments") if isinstance(cur_data.get("attachments"), dict) else {}
        existing_sub = existing_att.get(att_mod.SLOT_SUBMISSION) if isinstance(existing_att.get(att_mod.SLOT_SUBMISSION), dict) else None

        if len(uploadable) == 1:
            item = uploadable[0]
            if existing_sub and existing_sub.get("file_id"):
                same_rel = existing_sub.get("legacy_rel_path") == item["rel_path"]
                same_sig = (
                    existing_sub.get("file_name") == item["file_name"]
                    and int(existing_sub.get("file_size") or 0) == item["size_bytes"]
                )
                if same_rel or same_sig:
                    skipped_existing += 1
                    continue

            dest = drive_routing.resolve_upload_destination(
                pid,
                dt_id,
                att_mod.SLOT_SUBMISSION,
                service,
                store,
                create_if_missing=True,
                doc_no=doc_no,
            )
            session_url = drive_upload.create_resumable_session(
                token,
                dest["folder_id"],
                item["file_name"],
                item["mime_type"],
                item["size_bytes"],
                **kwargs_http,
            )
            with open(item["abs_path"], "rb") as fh:
                remote = drive_upload.upload_stream_resumable(
                    session_url,
                    fh,
                    item["size_bytes"],
                    **kwargs_stream,
                )
            file_id = str(remote["id"])
            view_url = str(remote.get("webViewLink") or att_mod.drive_view_url(file_id))
            meta = att_mod.build_attachment_meta(
                file_id=file_id,
                name=remote.get("name") or item["file_name"],
                size=int(remote.get("size") or item["size_bytes"]),
                mime_type=remote.get("mimeType") or item["mime_type"],
                uploaded_by="hse_migration",
                url=view_url,
                legacy_rel_path=item["rel_path"],
            )
            new_att = dict(existing_att)
            new_att[att_mod.SLOT_SUBMISSION] = meta
            cur_data["attachments"] = new_att
            cur_data["fileLocation"] = view_url
            with conn.cursor() as cur:
                cur.execute(
                    """UPDATE records SET data = %s::jsonb, updated_at = NOW()
                       WHERE id = %s AND project_id = %s""",
                    (json.dumps(cur_data, ensure_ascii=False), rid, pid),
                )
            conn.commit()
            uploaded_files += 1
            uploaded_bytes += item["size_bytes"]
        else:
            existing_files = list(existing_sub.get("files") or []) if (existing_sub and existing_sub.get("is_folder")) else []
            seen_keys = {(f.get("name"), int(f.get("size") or 0)) for f in existing_files}
            pending = [it for it in uploadable if (it["file_name"], it["size_bytes"]) not in seen_keys]
            if not pending and existing_sub and existing_sub.get("folder_id"):
                skipped_existing += len(uploadable)
                continue

            subfolder_name = _safe_drive_subfolder_name(doc_no, rid)
            dest = drive_routing.resolve_upload_destination(
                pid,
                dt_id,
                att_mod.SLOT_SUBMISSION,
                service,
                store,
                create_if_missing=True,
                subfolder_name=subfolder_name,
                doc_no=doc_no,
            )
            folder_id = str(dest.get("subfolder_id") or dest["folder_id"])
            folder_url = f"https://drive.google.com/drive/folders/{folder_id}"
            files_list = (
                list(existing_sub.get("files") or [])
                if (existing_sub and existing_sub.get("is_folder") and existing_sub.get("folder_id") == folder_id)
                else []
            )
            seen_keys = {(f.get("name"), int(f.get("size") or 0)) for f in files_list}
            last_file_id = str((existing_sub or {}).get("file_id") or "")

            for item in uploadable:
                if (item["file_name"], item["size_bytes"]) in seen_keys:
                    skipped_existing += 1
                    continue
                session_url = drive_upload.create_resumable_session(
                    token,
                    folder_id,
                    item["file_name"],
                    item["mime_type"],
                    item["size_bytes"],
                    **kwargs_http,
                )
                with open(item["abs_path"], "rb") as fh:
                    remote = drive_upload.upload_stream_resumable(
                        session_url,
                        fh,
                        item["size_bytes"],
                        **kwargs_stream,
                    )
                last_file_id = str(remote["id"])
                f_name = remote.get("name") or item["file_name"]
                f_size = int(remote.get("size") or item["size_bytes"])
                f_mime = remote.get("mimeType") or item["mime_type"]
                files_list = [f for f in files_list if f.get("id") != last_file_id]
                files_list.append({"id": last_file_id, "name": f_name, "size": f_size, "mimeType": f_mime})
                seen_keys.add((f_name, f_size))
                uploaded_files += 1
                uploaded_bytes += item["size_bytes"]

                meta = att_mod.build_attachment_meta(
                    file_id=last_file_id or folder_id,
                    name=subfolder_name,
                    size=sum(int(f.get("size") or 0) for f in files_list),
                    mime_type="application/vnd.google-apps.folder",
                    uploaded_by="hse_migration",
                    is_folder=True,
                    folder_id=folder_id,
                    folder_name=subfolder_name,
                    folder_url=folder_url,
                    url=folder_url,
                    files=files_list,
                    file_count=len(files_list),
                )
                new_att = dict(existing_att)
                new_att[att_mod.SLOT_SUBMISSION] = meta
                cur_data["attachments"] = new_att
                cur_data["fileLocation"] = folder_url
                existing_att = new_att
                with conn.cursor() as cur:
                    cur.execute(
                        """UPDATE records SET data = %s::jsonb, updated_at = NOW()
                           WHERE id = %s AND project_id = %s""",
                        (json.dumps(cur_data, ensure_ascii=False), rid, pid),
                    )
                conn.commit()

    return {
        "uploaded_files": uploaded_files,
        "skipped_existing": skipped_existing,
        "missing_files": missing_files,
        "uploaded_bytes": uploaded_bytes,
    }


class _ResilientDriveHttp:
    """Keep-alive ``requests.Session`` wrapper retrying transient TLS resets and 429/5xx errors."""

    def __init__(self, max_retries: int = 5):
        import requests as _req
        import time as _time

        self._requests = _req
        self._sleep = _time.sleep
        self._max_retries = max_retries
        self._session = _req.Session()

    def _reset_session(self) -> None:
        try:
            self._session.close()
        except Exception:
            pass
        self._session = self._requests.Session()

    def _request(self, method: str, url: str, **kwargs):
        last_exc = None
        for attempt in range(self._max_retries):
            try:
                resp = getattr(self._session, method)(url, **kwargs)
                if resp.status_code in (429, 500, 502, 503, 504) and attempt + 1 < self._max_retries:
                    if method == "post" and "/upload/" not in url:
                        return resp
                    self._sleep(min(2.0 ** attempt, 16.0))
                    continue
                return resp
            except self._requests.exceptions.RequestException as exc:
                last_exc = exc
                self._reset_session()
                if attempt + 1 >= self._max_retries:
                    raise
                self._sleep(min(2.0 ** attempt, 16.0))
        raise last_exc  # pragma: no cover

    def get(self, url: str, **kwargs):
        return self._request("get", url, **kwargs)

    def post(self, url: str, **kwargs):
        return self._request("post", url, **kwargs)

    def put(self, url: str, **kwargs):
        return self._request("put", url, **kwargs)

    def patch(self, url: str, **kwargs):
        return self._request("patch", url, **kwargs)

    def close(self) -> None:
        try:
            self._session.close()
        except Exception:
            pass


def _init_drive_client(drive_creds=None, http_module=None):
    import drive_upload

    if drive_creds is not None and getattr(drive_creds, "valid", False) and getattr(drive_creds, "token", ""):
        token = str(drive_creds.token)
    else:
        creds = drive_creds or drive_upload.get_upload_credentials()
        token = drive_upload.access_token(creds)
    active_http = http_module if http_module is not None else _ResilientDriveHttp()
    service = drive_upload.RestDriveAdapter(token, http=active_http)
    return service, token, active_http


def execute_local_migration(
    target_db_url: str,
    project_id: str,
    legacy_db: Path | str = DEFAULT_LEGACY_DB,
    foram_dir: Path | str = DEFAULT_FORAM_DIR,
    files_dir: Path | str = DEFAULT_FILES_DIR,
    legacy_project_id: int | None = None,
    psycopg2_module=None,
    allow_remote: bool = False,
    confirm_production: bool = False,
    upload_drive: bool = False,
    drive_root_id: str = "",
    http_module=None,
    sleep_fn=None,
    drive_creds=None,
) -> dict:
    """Idempotently import mapped legacy records into a single verified HSE project."""
    safe_url = validate_target_postgres_url(
        target_db_url, allow_remote=allow_remote, confirm_production=confirm_production
    )
    pid = str(project_id or "").strip()
    if not pid:
        raise ValueError("Target DCR --project-id is required for single-project execution (or use --multi-project).")

    plan = build_migration_plan(legacy_db=legacy_db, foram_dir=foram_dir, files_dir=files_dir)
    if psycopg2_module is None:
        import psycopg2 as psycopg2_module  # type: ignore

    selected_recs = [
        r for r in plan["records"]
        if legacy_project_id is None or r["legacy_project_id"] == legacy_project_id
    ]

    try:
        conn = psycopg2_module.connect(safe_url)
    except Exception as exc:
        raise RuntimeError(f"Target DB connection failed: {_redact_sensitive_text(exc)}") from None

    drive_stats = {"uploaded_files": 0, "skipped_existing": 0, "missing_files": 0, "uploaded_bytes": 0}
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

                inserted, updated, skipped = _upsert_records_for_project(cur, pid, selected_recs)

        if upload_drive:
            service, token, active_http = _init_drive_client(drive_creds=drive_creds, http_module=http_module)
            try:
                store = _CursorRoutingStore(conn)
                drive_stats = _upload_project_attachments_to_drive(
                    conn,
                    pid,
                    selected_recs,
                    plan["record_attachments"],
                    service,
                    token,
                    store,
                    drive_root_id=drive_root_id,
                    http_module=active_http,
                    sleep_fn=sleep_fn,
                )
            finally:
                if http_module is None and hasattr(active_http, "close"):
                    active_http.close()
    finally:
        conn.close()

    return {
        "mode": "execute",
        "project_id": pid,
        "inserted": inserted,
        "updated": updated,
        "skipped": skipped,
        "filtered_national_ids": plan["filtered_national_ids"],
        "sanitized_dates": plan["sanitized_dates"],
        "sanitized_local_paths": plan["sanitized_local_paths"],
        "counts_by_doc_type": plan["counts_by_doc_type"],
        "drive_upload": drive_stats,
        "foram": plan["foram"],
    }


def execute_multi_project_migration(
    target_db_url: str,
    legacy_db: Path | str = DEFAULT_LEGACY_DB,
    foram_dir: Path | str = DEFAULT_FORAM_DIR,
    files_dir: Path | str = DEFAULT_FILES_DIR,
    legacy_project_id: int | None = None,
    psycopg2_module=None,
    allow_remote: bool = False,
    confirm_production: bool = False,
    upload_drive: bool = False,
    drive_root_id: str = "",
    http_module=None,
    sleep_fn=None,
    drive_creds=None,
) -> dict:
    """Match/create all legacy SafetyQR projects in workspace ``HSE`` and import records (+ optional Drive upload)."""
    safe_url = validate_target_postgres_url(
        target_db_url, allow_remote=allow_remote, confirm_production=confirm_production
    )
    plan = build_migration_plan(legacy_db=legacy_db, foram_dir=foram_dir, files_dir=files_dir)
    if psycopg2_module is None:
        import psycopg2 as psycopg2_module  # type: ignore

    try:
        conn = psycopg2_module.connect(safe_url)
    except Exception as exc:
        raise RuntimeError(f"Target DB connection failed: {_redact_sensitive_text(exc)}") from None

    project_results = []
    total_inserted = 0
    total_updated = 0
    total_skipped = 0
    total_created_projects = 0
    drive_totals = {"uploaded_files": 0, "skipped_existing": 0, "missing_files": 0, "uploaded_bytes": 0}
    target_records_map: dict[str, list[dict]] = {}

    try:
        with conn:
            with conn.cursor() as cur:
                existing_hse = _fetch_existing_hse_projects_from_cursor(cur)
                mapping_plan = build_project_mapping_plan(
                    plan["projects"],
                    existing_hse_projects=existing_hse,
                    records=plan["records"],
                    attachments_by_record=plan["record_attachments"],
                )
                for m_entry in mapping_plan:
                    lpid = m_entry["legacy_id"]
                    if legacy_project_id is not None and lpid != legacy_project_id:
                        continue
                    target_pid, created, admins_assigned = _ensure_target_hse_project(cur, m_entry)
                    if created:
                        total_created_projects += 1
                    proj_recs = [r for r in plan["records"] if r["legacy_project_id"] == lpid]
                    ins, upd, skp = _upsert_records_for_project(cur, target_pid, proj_recs)
                    total_inserted += ins
                    total_updated += upd
                    total_skipped += skp
                    target_records_map[target_pid] = proj_recs
                    project_results.append({
                        "legacy_id": lpid,
                        "legacy_name": m_entry["legacy_name"],
                        "target_project_id": target_pid,
                        "created": created,
                        "admins_assigned": admins_assigned,
                        "inserted": ins,
                        "updated": upd,
                        "skipped": skp,
                    })

        if upload_drive:
            import drive_routing as _dr

            service, token, active_http = _init_drive_client(drive_creds=drive_creds, http_module=http_module)
            try:
                store = _CursorRoutingStore(conn)
                for item in project_results:
                    tpid = item["target_project_id"]
                    if drive_root_id:
                        item["drive_folder_id"] = _ensure_project_drive_folder(
                            conn, store, tpid, service, _dr, drive_root_id=drive_root_id
                        )
                    st = _upload_project_attachments_to_drive(
                        conn,
                        tpid,
                        target_records_map.get(tpid, []),
                        plan["record_attachments"],
                        service,
                        token,
                        store,
                        drive_root_id=drive_root_id,
                        http_module=active_http,
                        sleep_fn=sleep_fn,
                    )
                    item["drive_upload"] = st
                    for k in drive_totals:
                        drive_totals[k] += st[k]
            finally:
                if http_module is None and hasattr(active_http, "close"):
                    active_http.close()
    finally:
        conn.close()

    return {
        "mode": "execute-multi-project",
        "projects_created": total_created_projects,
        "inserted": total_inserted,
        "updated": total_updated,
        "skipped": total_skipped,
        "project_results": project_results,
        "filtered_national_ids": plan["filtered_national_ids"],
        "preserved_national_ids": plan.get("preserved_national_ids", 0),
        "preserved_pii_fields": plan.get("preserved_pii_fields", 0),
        "sanitized_dates": plan["sanitized_dates"],
        "sanitized_local_paths": plan["sanitized_local_paths"],
        "counts_by_doc_type": plan["counts_by_doc_type"],
        "drive_upload": drive_totals,
        "foram": plan["foram"],
    }


def _strip_records_for_summary(report: dict) -> dict:
    return {k: v for k, v in report.items() if k not in ("records", "record_attachments")}


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="Safe read-only SafetyQR -> DCR HSE migration tool")
    parser.add_argument("--legacy-db", default=str(DEFAULT_LEGACY_DB), help="Path to SafetyQR SQLite app.db")
    parser.add_argument("--foram-dir", default=str(DEFAULT_FORAM_DIR), help="Path to FORAM templates directory")
    parser.add_argument("--files-dir", default=str(DEFAULT_FILES_DIR), help="Path to legacy SafetyQR data/files directory")
    parser.add_argument("--dry-run", action="store_true", help="Inspect read-only and print migration report (default)")
    parser.add_argument("--read-only-probe", action="store_true", help="Read existing HSE projects via BEGIN READ ONLY")
    parser.add_argument("--target-db", default="", help="Target PostgreSQL URL (local only unless double-confirmed)")
    parser.add_argument("--project-id", default="", help="Target HSE project ID for single-project execution")
    parser.add_argument("--multi-project", action="store_true", help="Auto-match/create all 4 HSE projects during execute")
    parser.add_argument("--legacy-project-id", type=int, default=None, help="Optional filter for one legacy project ID")
    parser.add_argument("--allow-remote-db", action="store_true", help="Required with --confirm-production for non-local DB")
    parser.add_argument("--confirm-production", action="store_true", help="Required with --allow-remote-db for non-local DB")
    parser.add_argument("--upload-drive", action="store_true", help="Stream local attachments directly to Google Drive")
    parser.add_argument("--drive-root-id", default="", help="Parent Google Drive folder ID for auto-created project folders")
    parser.add_argument("--json", action="store_true", help="Print JSON summary report")
    args = parser.parse_args(argv)

    want_env_execute = bool(
        not args.dry_run
        and not args.read_only_probe
        and (args.multi_project or str(args.project_id or "").strip())
        and (args.target_db or (args.allow_remote_db and args.confirm_production))
    )

    if args.read_only_probe:
        if args.upload_drive or args.multi_project or (args.target_db and not args.dry_run):
            raise ValueError("--read-only-probe is strictly read-only and cannot be combined with DB writes or Drive uploads.")
        existing_hse = fetch_existing_hse_projects_readonly(db_url=args.target_db)
        plan = build_migration_plan(
            legacy_db=args.legacy_db,
            foram_dir=args.foram_dir,
            files_dir=args.files_dir,
            existing_hse_projects=existing_hse,
        )
        res = _strip_records_for_summary(plan)
        res["existing_hse_projects"] = existing_hse
    elif (args.target_db or want_env_execute) and not args.dry_run:
        resolved_target_db = resolve_target_db_url_from_env(args.target_db)
        validate_target_postgres_url(
            resolved_target_db,
            allow_remote=bool(args.allow_remote_db),
            confirm_production=bool(args.confirm_production),
        )
        if args.multi_project or not str(args.project_id or "").strip():
            if not args.multi_project:
                raise ValueError("Specify --project-id <id> or --multi-project for execution.")
            res = execute_multi_project_migration(
                resolved_target_db,
                legacy_db=args.legacy_db,
                foram_dir=args.foram_dir,
                files_dir=args.files_dir,
                legacy_project_id=args.legacy_project_id,
                allow_remote=bool(args.allow_remote_db),
                confirm_production=bool(args.confirm_production),
                upload_drive=bool(args.upload_drive),
                drive_root_id=args.drive_root_id,
            )
        else:
            res = execute_local_migration(
                resolved_target_db,
                args.project_id,
                legacy_db=args.legacy_db,
                foram_dir=args.foram_dir,
                files_dir=args.files_dir,
                legacy_project_id=args.legacy_project_id,
                allow_remote=bool(args.allow_remote_db),
                confirm_production=bool(args.confirm_production),
                upload_drive=bool(args.upload_drive),
                drive_root_id=args.drive_root_id,
            )
    else:
        if args.target_db:
            validate_target_postgres_url(
                args.target_db,
                allow_remote=bool(args.allow_remote_db),
                confirm_production=bool(args.confirm_production),
            )
        res = _strip_records_for_summary(
            build_migration_plan(
                legacy_db=args.legacy_db,
                foram_dir=args.foram_dir,
                files_dir=args.files_dir,
            )
        )

    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")

    if args.json:
        print(json.dumps(res, ensure_ascii=False, indent=2))
    else:
        print("=== DCR HSE Legacy Migration Report ===", flush=True)
        print(f"Mode: {res['mode']}", flush=True)
        print(f"Mapped records by register: {json.dumps(res.get('counts_by_doc_type', {}), ensure_ascii=False)}", flush=True)
        print(
            f"Preserved worker National IDs: {res.get('preserved_national_ids', 0)} | "
            f"Preserved PII fields: {res.get('preserved_pii_fields', 0)}",
            flush=True,
        )
        print(f"Sanitized dates: {res.get('sanitized_dates', 0)} | Stripped local paths: {res.get('sanitized_local_paths', 0)}", flush=True)
        inv = res.get("attachment_inventory") or {}
        if inv:
            print(
                f"Attachments: {inv.get('mapped_files_count', 0)} mapped ({inv.get('mapped_total_mb', 0)} MB) "
                f"[docs={inv.get('document_files_count', 0)}, worker_photos={inv.get('worker_photos_count', 0)}], "
                f"{inv.get('missing_on_disk_count', 0)} missing on disk",
                flush=True,
            )
        drive_up = res.get("drive_upload") or {}
        if drive_up:
            print(
                f"Drive Upload: uploaded={drive_up.get('uploaded_files', 0)}, "
                f"skipped_existing={drive_up.get('skipped_existing', 0)}, "
                f"missing={drive_up.get('missing_files', 0)}",
                flush=True,
            )
        foram = res.get("foram") or {}
        print(
            f"FORAM templates: {foram.get('valid_count', 0)} mapped, "
            f"{foram.get('skipped_junk_count', 0)} .lnk/~$ skipped",
            flush=True,
        )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
