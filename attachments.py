"""attachments.py - Dual attachment model (Submission / Approval) for DCR records.

Pure functions only: NO database or network access, so they are safe to unit-test
offline. Data lives inside records.data (JSONB) under the additive key
'attachments'. The legacy 'fileLocation' key is never modified here.
"""
import copy
import datetime
import re

SLOT_SUBMISSION = "submission"
SLOT_APPROVAL = "approval"
SLOT_PART_A = "part_a"
SLOT_PART_B = "part_b"
SLOT_PART_C = "part_c"
SLOT_PART_D = "part_d"
DUAL_SLOTS = (SLOT_SUBMISSION, SLOT_APPROVAL)
NOC_SLOTS = (SLOT_PART_A, SLOT_PART_B, SLOT_PART_C, SLOT_PART_D)
SLOTS = DUAL_SLOTS + NOC_SLOTS

_INVALID_LINKS = {"", "#", "null", "none"}


def _clean_link(value):
    v = str(value or "").strip()
    return "" if v.lower() in _INVALID_LINKS else v


def parse_rev_number(value, default=0):
    """Parse a revision integer or string (e.g., 1, 'REV01', 'R2') into an int >= 0."""
    if value is None or isinstance(value, bool):
        return default
    if isinstance(value, int):
        return max(0, value)
    s = str(value).strip()
    if not s:
        return default
    if s.isdigit():
        return max(0, int(s))
    m = re.search(r"(?:rev|r)\s*[-_.]?\s*(\d+)", s, re.IGNORECASE)
    if m:
        return max(0, int(m.group(1)))
    return default


def format_rev_label(rev):
    """Format revision number into zero-padded label like 'REV00', 'REV01'."""
    return "REV%02d" % max(0, parse_rev_number(rev, 0))


def get_attachment_rev(entry):
    """Return numeric revision (default 0) for an attachment dict."""
    if not isinstance(entry, dict):
        return 0
    if "rev" in entry and entry.get("rev") is not None:
        return parse_rev_number(entry.get("rev"), 0)
    if entry.get("rev_label"):
        return parse_rev_number(entry.get("rev_label"), 0)
    return 0


def strip_entry_history(entry):
    """Return a clean copy of an attachment dict without nested 'history'."""
    if not isinstance(entry, dict):
        return {}
    snap = copy.deepcopy(entry)
    snap.pop("history", None)
    rev = get_attachment_rev(snap)
    snap.setdefault("rev", rev)
    snap.setdefault("rev_label", format_rev_label(rev))
    return snap


def get_attachment_revisions(entry):
    """Return ordered list of revision snapshots for a slot (active first, then archived history desc)."""
    if not isinstance(entry, dict):
        return []
    active = strip_entry_history(entry)
    active["current"] = True
    out = [active]
    seen_revs = {active["rev"]}
    raw_hist = entry.get("history") if isinstance(entry.get("history"), list) else []
    hist_snaps = []
    for item in raw_hist:
        if not isinstance(item, dict):
            continue
        snap = strip_entry_history(item)
        snap["current"] = False
        if snap["rev"] not in seen_revs:
            seen_revs.add(snap["rev"])
            hist_snaps.append(snap)
    hist_snaps.sort(key=lambda h: (-int(h.get("rev") or 0), str(h.get("archived_at") or h.get("uploaded_at") or "")), reverse=False)
    out.extend(hist_snaps)
    return out


def archive_and_promote_revision(existing_entry, new_meta, rev=None, archived_by=""):
    """Archive existing_entry into new_meta['history'] and stamp the promoted revision number."""
    promoted = copy.deepcopy(new_meta) if isinstance(new_meta, dict) else {}
    promoted.pop("history", None)
    history = []
    if isinstance(existing_entry, dict) and (
        existing_entry.get("file_id") or existing_entry.get("folder_id") or existing_entry.get("url")
    ):
        prev_rev = get_attachment_rev(existing_entry)
        prev_snap = strip_entry_history(existing_entry)
        prev_snap["rev"] = prev_rev
        prev_snap["rev_label"] = format_rev_label(prev_rev)
        prev_snap.setdefault(
            "archived_at",
            datetime.datetime.utcnow().replace(microsecond=0).isoformat() + "Z",
        )
        if archived_by:
            prev_snap.setdefault("archived_by", archived_by)
        history.append(prev_snap)
        old_hist = existing_entry.get("history") if isinstance(existing_entry.get("history"), list) else []
        seen_revs = {prev_rev}
        for h in old_hist:
            if not isinstance(h, dict):
                continue
            hs = strip_entry_history(h)
            if hs["rev"] not in seen_revs:
                seen_revs.add(hs["rev"])
                history.append(hs)
        next_rev = parse_rev_number(rev, prev_rev + 1) if rev is not None else (prev_rev + 1)
    else:
        next_rev = parse_rev_number(rev, 0) if rev is not None else 0
    promoted["rev"] = next_rev
    promoted["rev_label"] = format_rev_label(next_rev)
    if history:
        history.sort(key=lambda h: int(h.get("rev") or 0), reverse=True)
        promoted["history"] = history
    return promoted


def pop_active_revision(existing_entry):
    """Remove the active revision; if history exists, promote the latest historical entry to active.

    Returns `(replacement_entry_or_None, removed_snapshot)`.
    """
    if not isinstance(existing_entry, dict):
        return None, None
    removed = strip_entry_history(existing_entry)
    raw_hist = existing_entry.get("history") if isinstance(existing_entry.get("history"), list) else []
    hist_snaps = [strip_entry_history(h) for h in raw_hist if isinstance(h, dict)]
    if not hist_snaps:
        return None, removed
    hist_snaps.sort(key=lambda h: int(h.get("rev") or 0), reverse=True)
    promoted = copy.deepcopy(hist_snaps[0])
    promoted.pop("archived_at", None)
    promoted.pop("archived_by", None)
    remaining = hist_snaps[1:]
    if remaining:
        promoted["history"] = remaining
    return promoted, removed


def drive_view_url(file_id):
    return "https://drive.google.com/file/d/%s/view" % file_id


def drive_download_url(file_id):
    return "https://drive.google.com/uc?export=download&id=%s" % file_id


def extract_file_id(url):
    """Return the Drive file id from a Drive URL, or '' (folders return '')."""
    url = _clean_link(url)
    if not url or "/folders/" in url:
        return ""
    m = re.search(r"/d/([a-zA-Z0-9_-]+)", url) or re.search(r"[?&]id=([a-zA-Z0-9_-]+)", url)
    return m.group(1) if m else ""


def build_attachment_meta(file_id, name, size=0, mime_type="", uploaded_by="", **extra):
    meta = {
        "file_id": file_id,
        "file_name": name,
        "file_size": int(size or 0),
        "mime_type": mime_type or "",
        "uploaded_at": datetime.datetime.utcnow().replace(microsecond=0).isoformat() + "Z",
        "uploaded_by": uploaded_by or "",
    }
    meta.update(extra)
    return meta


def normalize_attachments(data):
    """Return {'submission': dict|None, 'approval': dict|None} for a record's data.

    - If data['attachments'] has a slot, it is used as-is.
    - Includes 'part_a'..'part_d' when present in data['attachments'].
    - If the submission slot is empty and a legacy fileLocation exists, it is
      exposed as a legacy submission entry ({'url': ..., 'legacy': True}).
    Never mutates `data`.
    """
    data = data if isinstance(data, dict) else {}
    att = data.get("attachments")
    att = att if isinstance(att, dict) else {}
    result = {s: (att.get(s) if isinstance(att.get(s), dict) else None) for s in DUAL_SLOTS}
    for ns in NOC_SLOTS:
        if isinstance(att.get(ns), dict):
            result[ns] = att.get(ns)
    has_noc_part = any(ns in result for ns in NOC_SLOTS)
    if result[SLOT_SUBMISSION] is None and not has_noc_part:
        legacy = _clean_link(data.get("fileLocation"))
        app_entry = result.get(SLOT_APPROVAL) or {}
        app_url = _clean_link(app_entry.get("url"))
        app_fid = app_entry.get("file_id") or extract_file_id(app_url)
        leg_fid = extract_file_id(legacy)
        is_same = bool((legacy and legacy == app_url) or (app_fid and leg_fid and app_fid == leg_fid))
        if legacy and not is_same:
            result[SLOT_SUBMISSION] = {"url": legacy, "legacy": True}
    return result


def attachment_urls(entry):
    """Return (view_url, download_url) for an attachment entry; ('','') if none."""
    if not entry:
        return "", ""
    if entry.get("file_id"):
        return drive_view_url(entry["file_id"]), drive_download_url(entry["file_id"])
    url = _clean_link(entry.get("url"))
    fid = extract_file_id(url)
    return url, (drive_download_url(fid) if fid else "")


def apply_attachment(data, slot, meta):
    """Return a NEW data dict with `meta` stored in the given slot.

    fileLocation is left untouched (legacy compatibility).
    """
    if slot not in SLOTS:
        raise ValueError("Invalid attachment slot: %r" % (slot,))
    new = copy.deepcopy(data) if isinstance(data, dict) else {}
    att = new.get("attachments")
    if not isinstance(att, dict):
        att = {}
    att[slot] = meta
    new["attachments"] = att
    return new


def remove_attachment(data, slot):
    """Return a NEW data dict with the slot cleared (fileLocation untouched)."""
    if slot not in SLOTS:
        raise ValueError("Invalid attachment slot: %r" % (slot,))
    new = copy.deepcopy(data) if isinstance(data, dict) else {}
    att = new.get("attachments")
    if isinstance(att, dict):
        att.pop(slot, None)
        if not att:
            new.pop("attachments", None)
    return new


SYNC_SOURCE = "drive_sync"


def classify_path_kind(full_path):
    """'part_a'..'part_d' if path contains 'part a'..'part d',
    'approval' if the Drive path contains 'reply' or 'replied' or 'approval',
    'submission' if it contains 'submit' or 'submission' (strictly excluding 'submittal'),
    else None (general document without a separate cycle folder).

    Same precedence as drive_service.get_item_priority (part_d..part_a / reply wins over submit).
    """
    p = (full_path or "").lower()
    m = re.search(r"\bpart[\s_-]*([abcd])\b", p)
    if m:
        return "part_" + m.group(1)
    if re.search(r"\b(reply|replied|approval|approved)\b", p):
        return SLOT_APPROVAL
    if re.search(r"\b(submit|submission)\b", p):
        return SLOT_SUBMISSION
    return None


def merge_sync_links(data, sync_links):
    """Return a NEW data dict with Drive-sync links merged into attachments.

    sync_links: {'submission': url|dict|None, ...}.
    Only slots that are empty OR previously created by sync (source == drive_sync)
    are written. Slots uploaded through DCR (no 'drive_sync' source) are never
    overwritten. fileLocation is not touched here.
    """
    new = copy.deepcopy(data) if isinstance(data, dict) else {}
    for slot in SLOTS:
        raw_val = (sync_links or {}).get(slot)
        if isinstance(raw_val, dict):
            url = _clean_link(raw_val.get("url"))
            if not url:
                continue
            sync_entry = copy.deepcopy(raw_val)
            sync_entry["url"] = url
            sync_entry["source"] = SYNC_SOURCE
        else:
            url = _clean_link(raw_val)
            if not url:
                continue
            sync_entry = {"url": url, "source": SYNC_SOURCE}
        att = new.get("attachments")
        existing = att.get(slot) if isinstance(att, dict) else None
        if isinstance(existing, dict) and existing.get("source") != SYNC_SOURCE:
            continue
        new = apply_attachment(new, slot, sync_entry)
    return new


# ---------------- Document types with / without a consultant reply cycle ----------------
REPLY_CYCLE_CODES = {"SD", "DS", "MS", "IR", "MIR", "RFI", "NOC", "ABD", "WN", "NCN", "NCR", "EI", "SI", "CVI", "PCQ", "SVP"}
_NO_CYCLE_NAME_HINTS = ("requisition", "purchase", "letter", "correspondence", "minutes", "mom",
                        "daily report", "site log", "boq", "specs", "variation")
_CYCLE_NAME_HINTS = ("shop drawing", "drawing", "material submittal", "submittal", "inspection request",
                     "request for information", "notice of change", "as-built", "as built",
                     "work notice", "work notification", "works notice", "non-conformance", "non conformance", "instruction")


def doc_type_has_reply_cycle(dt):
    """True when documents of this type go to the consultant and receive a reply / approval status
    (Shop Drawings, Material Submittals, IR, RFI, NOC). False for single-party documents
    (PR, Letters, MOM, Daily Reports). An explicit boolean 'has_reply_cycle' on the doc type wins."""
    dt = dt or {}
    flag = dt.get("has_reply_cycle")
    if isinstance(flag, bool):
        return flag
    code = str(dt.get("code") or dt.get("id") or "").strip().upper()
    if code in REPLY_CYCLE_CODES:
        return True
    name = str(dt.get("name") or "").strip().lower()
    if any(h in name for h in _NO_CYCLE_NAME_HINTS):
        return False
    return any(h in name for h in _CYCLE_NAME_HINTS)


HSE_CLOSEOUT_CODES = frozenset({"PTW", "OBS", "SVP", "INC", "JSA"})


def is_hse_closeout_doc_type(dt):
    """True when an HSE document type supports a dual Close-Out attachment in SLOT_APPROVAL."""
    dt = dt or {}
    flag = dt.get("has_closeout_slot")
    if isinstance(flag, bool):
        return flag
    code = str(dt.get("code") or dt.get("id") or "").strip().upper()
    return code in HSE_CLOSEOUT_CODES


def doc_type_allows_approval_slot(dt, workspace_id="ENG"):
    """True when SLOT_APPROVAL is permitted for `(dt, workspace_id)`."""
    ws = str(workspace_id or "ENG").strip().upper()
    if ws == "HSE":
        return is_hse_closeout_doc_type(dt)
    if ws == "ENG":
        return doc_type_has_reply_cycle(dt)
    return False


def resolve_hse_closeout_status(current_status="", requested_status=""):
    """Resolve English HSE close-out lifecycle status (`Open` -> `Closed`) on SLOT_APPROVAL upload.

    Returns `(closeout_status, new_record_status_or_None)` so HSE close-out uploads
    never emit engineering consultant review codes (`Code A/B/C/D`).
    """
    cur = str(current_status or "").strip()
    req = str(requested_status or "").strip()
    cur_low = cur.lower()
    req_low = req.lower()
    if req_low in ("closed", "مغلق"):
        return "Closed", ("Closed" if cur != "Closed" else None)
    if cur_low in ("open", "مفتوح") or cur == "مغلق":
        return "Closed", "Closed"
    if cur_low == "closed":
        return "Closed", ("Closed" if cur != "Closed" else None)
    return "Closed", None
