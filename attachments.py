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
SLOTS = (SLOT_SUBMISSION, SLOT_APPROVAL)

_INVALID_LINKS = {"", "#", "null", "none"}


def _clean_link(value):
    v = str(value or "").strip()
    return "" if v.lower() in _INVALID_LINKS else v


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
    - If the submission slot is empty and a legacy fileLocation exists, it is
      exposed as a legacy submission entry ({'url': ..., 'legacy': True}).
    Never mutates `data`.
    """
    data = data if isinstance(data, dict) else {}
    att = data.get("attachments")
    att = att if isinstance(att, dict) else {}
    result = {s: (att.get(s) if isinstance(att.get(s), dict) else None) for s in SLOTS}
    if result[SLOT_SUBMISSION] is None:
        legacy = _clean_link(data.get("fileLocation"))
        if legacy:
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
    """'approval' if the Drive path contains 'reply', 'submission' if it contains 'submit', else None.

    Same precedence as drive_service.get_item_priority (reply wins over submit).
    """
    p = (full_path or "").lower()
    if "reply" in p:
        return SLOT_APPROVAL
    if "submit" in p:
        return SLOT_SUBMISSION
    return None


def merge_sync_links(data, sync_links):
    """Return a NEW data dict with Drive-sync links merged into attachments.

    sync_links: {'submission': url|None, 'approval': url|None}.
    Only slots that are empty OR previously created by sync (source == drive_sync)
    are written. Slots uploaded through DCR (no 'drive_sync' source) are never
    overwritten. fileLocation is not touched here.
    """
    new = copy.deepcopy(data) if isinstance(data, dict) else {}
    for slot in SLOTS:
        url = _clean_link((sync_links or {}).get(slot))
        if not url:
            continue
        att = new.get("attachments")
        existing = att.get(slot) if isinstance(att, dict) else None
        if isinstance(existing, dict) and existing.get("source") != SYNC_SOURCE:
            continue
        new = apply_attachment(new, slot, {"url": url, "source": SYNC_SOURCE})
    return new


# ---------------- Document types with / without a consultant reply cycle ----------------
REPLY_CYCLE_CODES = {"SD", "DS", "MS", "IR", "MIR", "RFI", "NOC", "ABD"}
_NO_CYCLE_NAME_HINTS = ("requisition", "purchase", "letter", "correspondence", "minutes", "mom",
                        "daily report", "site log", "boq", "specs", "variation")
_CYCLE_NAME_HINTS = ("shop drawing", "drawing", "material submittal", "submittal", "inspection request",
                     "request for information", "notice of change", "as-built", "as built")


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
