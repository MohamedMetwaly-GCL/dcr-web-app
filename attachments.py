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


def extract_drive_item_id(url):
    """Return the Drive file OR folder id from any Drive URL, or ''."""
    url = _clean_link(url)
    if not url:
        return ""
    m = (
        re.search(r"/folders/([a-zA-Z0-9_-]+)", url)
        or re.search(r"/d/([a-zA-Z0-9_-]+)", url)
        or re.search(r"[?&]id=([a-zA-Z0-9_-]+)", url)
    )
    return m.group(1) if m else ""


LEGACY_LINK_KEYS = ("fileLocation", "drive_link", "driveLink")


def extract_legacy_link(data):
    """Return (key, clean_url) for the first non-empty legacy link field on a record dict."""
    if not isinstance(data, dict):
        return "", ""
    for key in LEGACY_LINK_KEYS:
        val = _clean_link(data.get(key))
        if val:
            return key, val
    return "", ""


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
        _, legacy = extract_legacy_link(data)
        app_entry = result.get(SLOT_APPROVAL) or {}
        app_url = _clean_link(app_entry.get("url"))
        app_fid = app_entry.get("file_id") or app_entry.get("folder_id") or extract_drive_item_id(app_url)
        leg_fid = extract_drive_item_id(legacy)
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


def _extract_dir_segments(full_path):
    """Return clean directory path segments for `full_path`, excluding the item's own basename."""
    raw = str(full_path or "").replace("\\", "/").strip()
    if not raw:
        return []
    has_trailing_sep = raw.endswith("/")
    parts = [seg.strip() for seg in raw.split("/")]
    if has_trailing_sep:
        return [seg for seg in parts if seg]
    if len(parts) <= 1:
        return []
    return [seg for seg in parts[:-1] if seg]


_PART_DIR_RE = re.compile(r"(?:^|[^a-z0-9])part[\s_-]*([abcd])(?:$|[^a-z0-9])", re.IGNORECASE)
_APPROVAL_DIR_RE = re.compile(r"(?:^|[^a-z0-9])(reply|replied|approval|approved)(?:$|[^a-z0-9])", re.IGNORECASE)
_SUBMIT_DIR_RE = re.compile(r"(?:^|[^a-z0-9])(submit|submission)(?:$|[^a-z0-9])", re.IGNORECASE)


def classify_path_kind(full_path):
    """Classify a Drive path into `'part_a'..'part_d'`, `'approval'`, `'submission'`, or `None`.

    Matches slot keywords strictly against **directory path segments** (excluding the file/item's
    own basename), so a file such as `DS-001_For_Approval.pdf` inside `Submit/` is classified as
    `'submission'` and never `'approval'`. Strictly excludes `'submittal'` from `'submit'`.
    """
    dir_segments = _extract_dir_segments(full_path)
    if not dir_segments:
        return None

    # 1) NOC 4-part directory segments (deepest segment wins)
    for seg in reversed(dir_segments):
        m = _PART_DIR_RE.search(seg)
        if m:
            return "part_" + m.group(1).lower()

    # 2) Dedicated slot directory segments (e.g. 'Submit', '01 - Replied') take precedence
    #    over document-named child folders (e.g. 'DS-001 For Approval REV00').
    for seg in reversed(dir_segments):
        unindexed = re.sub(r"^\d+[\s._-]+", "", seg).strip().lower()
        if unindexed in ("reply", "replied", "approval", "approved"):
            return SLOT_APPROVAL
        if unindexed in ("submit", "submission"):
            return SLOT_SUBMISSION

    # 3) General keyword matching on directory segments only (deepest segment wins)
    for seg in reversed(dir_segments):
        if _APPROVAL_DIR_RE.search(seg):
            return SLOT_APPROVAL
        if _SUBMIT_DIR_RE.search(seg):
            return SLOT_SUBMISSION
    return None


def _same_sync_target(entry_a, entry_b):
    if not isinstance(entry_a, dict) or not isinstance(entry_b, dict):
        return False
    url_a = _clean_link(entry_a.get("url"))
    url_b = _clean_link(entry_b.get("url"))
    if url_a and url_b and url_a == url_b:
        return True
    fid_a = entry_a.get("file_id") or entry_a.get("folder_id") or extract_drive_item_id(url_a)
    fid_b = entry_b.get("file_id") or entry_b.get("folder_id") or extract_drive_item_id(url_b)
    return bool(fid_a and fid_b and fid_a == fid_b)


def _clear_matching_legacy_keys(data_dict, target_entry):
    """Remove any top-level legacy link key whose URL/Drive ID matches `target_entry`."""
    if not isinstance(data_dict, dict) or not isinstance(target_entry, dict):
        return False
    cleared = False
    for key in LEGACY_LINK_KEYS:
        val = _clean_link(data_dict.get(key))
        if val and _same_sync_target({"url": val}, target_entry):
            data_dict.pop(key, None)
            cleared = True
    return cleared


def _is_verified_submit_attachment(entry, confirmed_not_in_submit=False):
    """Return True when `entry` is a verified DCR Submit upload that must never be replaced.

    An explicit `slot_verified=True` always protects the submission slot.
    Legacy/imported entries that merely have `uploaded_at` (without `slot_verified=True`)
    are protected when Drive Submit-folder presence is unknown (`confirmed_not_in_submit=False`),
    but may be promoted/replaced when a Drive sync has confirmed the document is NOT in
    `Submit/` and IS in `Replied/` (`confirmed_not_in_submit=True`).
    """
    if not isinstance(entry, dict):
        return False
    if entry.get("slot_verified") is True:
        return True
    if confirmed_not_in_submit:
        return False
    if entry.get("uploaded_at") and not entry.get("legacy") and entry.get("source") != SYNC_SOURCE:
        return True
    return False


def promote_legacy_link_to_reply(data, matched_item=None, reply_url="", file_id="", folder_id="",
                                 file_name="", mime_type="", clear_stale_legacy=False,
                                 confirmed_not_in_submit=False):
    """Safely promote a verified Replied-folder legacy/submission link to `attachments['approval']`.

    Pure function (never mutates `data`). Preserves all metadata and never deletes a distinct
    verified submission file or overwrites a distinct DCR-uploaded approval file.
    When `clear_stale_legacy=True`, also clears any un-verified legacy `submission` / top-level
    legacy link keys whose stored Drive ID differed from the verified `Replied` file.
    Returns `(new_data, changed_bool)`.
    """
    if not isinstance(data, dict):
        return {}, False
    new = copy.deepcopy(data)
    att = new.get("attachments") if isinstance(new.get("attachments"), dict) else {}
    if any(isinstance(att.get(ns), dict) for ns in NOC_SLOTS):
        return new, False

    item = matched_item if isinstance(matched_item, dict) else {}
    resolved_mime = str(mime_type or item.get("mimeType") or "").strip()
    is_folder = bool(
        folder_id
        or resolved_mime == "application/vnd.google-apps.folder"
        or "/folders/" in str(reply_url or item.get("webViewLink") or "")
    )
    resolved_id = str(
        folder_id
        or file_id
        or item.get("id")
        or extract_drive_item_id(reply_url)
        or extract_drive_item_id(item.get("webViewLink"))
        or ""
    ).strip()
    resolved_url = _clean_link(reply_url or item.get("webViewLink"))
    if not resolved_url and resolved_id:
        resolved_url = (
            "https://drive.google.com/drive/folders/%s" % resolved_id
            if is_folder
            else drive_view_url(resolved_id)
        )
    if not resolved_url and not resolved_id:
        return new, False

    target_probe = {"url": resolved_url}
    if resolved_id:
        target_probe["folder_id" if is_folder else "file_id"] = resolved_id

    existing_sub = att.get(SLOT_SUBMISSION) if isinstance(att.get(SLOT_SUBMISSION), dict) else None
    existing_app = att.get(SLOT_APPROVAL) if isinstance(att.get(SLOT_APPROVAL), dict) else None

    if clear_stale_legacy and _is_verified_submit_attachment(
        existing_sub, confirmed_not_in_submit=confirmed_not_in_submit
    ):
        # Safety: never replace/clear a verified DCR Submit upload
        return new, False

    sub_matches = bool(existing_sub and _same_sync_target(existing_sub, target_probe))
    app_matches = bool(existing_app and _same_sync_target(existing_app, target_probe))

    if existing_app and not app_matches and existing_app.get("source") != SYNC_SOURCE:
        # Never overwrite a distinct DCR-uploaded approval attachment
        return new, False

    if app_matches:
        promoted = copy.deepcopy(existing_app)
    elif sub_matches:
        promoted = copy.deepcopy(existing_sub)
        promoted.pop("legacy", None)
    else:
        promoted = {}

    if resolved_url and not _clean_link(promoted.get("url")):
        promoted["url"] = resolved_url
    elif resolved_url and promoted.get("source") == SYNC_SOURCE:
        promoted["url"] = resolved_url

    if resolved_id:
        if is_folder:
            promoted.setdefault("folder_id", resolved_id)
            promoted["is_folder"] = True
        else:
            promoted.setdefault("file_id", resolved_id)

    resolved_name = str(file_name or item.get("name") or "").strip()
    if resolved_name and not promoted.get("file_name"):
        promoted["file_name"] = resolved_name
    if resolved_mime and not promoted.get("mime_type"):
        promoted["mime_type"] = resolved_mime

    promoted.setdefault("source", SYNC_SOURCE)
    promoted["role"] = "reply"
    promoted["slot_verified"] = True

    new = apply_attachment(new, SLOT_APPROVAL, promoted)
    if sub_matches or (
        clear_stale_legacy
        and existing_sub
        and not _is_verified_submit_attachment(existing_sub, confirmed_not_in_submit=confirmed_not_in_submit)
    ):
        new = remove_attachment(new, SLOT_SUBMISSION)
    if clear_stale_legacy:
        for key in LEGACY_LINK_KEYS:
            if _clean_link(new.get(key)):
                new.pop(key, None)
    else:
        _clear_matching_legacy_keys(new, promoted)

    return new, (new != data)


def merge_sync_links(data, sync_links):
    """Return a NEW data dict with Drive-sync links merged into attachments.

    sync_links: {'submission': url|dict|None, ...}.
    Only slots that are empty OR previously created by sync (source == drive_sync)
    are written. Slots uploaded through DCR (no 'drive_sync' source) are never
    overwritten. `submission` and `approval` are merged independently without ever
    writing a `Submit` URL into `approval`.
    When `approval` is synced (and no separate `submission` exists for that target),
    any duplicate legacy `fileLocation`/`drive_link`/`driveLink` or un-typed `submission`
    pointing to the exact same Drive item is cleanly promoted/cleared so the record
    renders as `Reply` without duplication.
    """
    new = copy.deepcopy(data) if isinstance(data, dict) else {}
    prepared = {}
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
        if slot == SLOT_APPROVAL:
            sync_entry.setdefault("role", "reply")
        prepared[slot] = sync_entry

    # Guard: never allow a Submit URL to be written into approval
    if SLOT_SUBMISSION in prepared and SLOT_APPROVAL in prepared:
        if _same_sync_target(prepared[SLOT_SUBMISSION], prepared[SLOT_APPROVAL]):
            prepared.pop(SLOT_APPROVAL, None)

    for slot, sync_entry in prepared.items():
        att = new.get("attachments")
        existing = att.get(slot) if isinstance(att, dict) else None
        if isinstance(existing, dict) and existing.get("source") != SYNC_SOURCE:
            continue
        if isinstance(existing, dict) and _same_sync_target(existing, sync_entry):
            merged_entry = copy.deepcopy(existing)
            merged_entry.update(sync_entry)
            sync_entry = merged_entry
        new = apply_attachment(new, slot, sync_entry)

    # Self-heal 1: if approval holds a stale drive_sync copy of the exact same submission URL, clear approval
    att_after = new.get("attachments") if isinstance(new.get("attachments"), dict) else {}
    sub_after = att_after.get(SLOT_SUBMISSION)
    app_after = att_after.get(SLOT_APPROVAL)
    if (
        isinstance(sub_after, dict)
        and isinstance(app_after, dict)
        and app_after.get("source") == SYNC_SOURCE
        and SLOT_APPROVAL not in prepared
        and _same_sync_target(sub_after, app_after)
    ):
        new = remove_attachment(new, SLOT_APPROVAL)

    # Self-heal 2: if approval was verified from Drive (Replied folder) and submission/legacy holds
    # the exact same Drive item (with no Submit file for it), promote cleanly to reply and clear duplicates
    if SLOT_APPROVAL in prepared and SLOT_SUBMISSION not in prepared and isinstance(app_after, dict):
        sub_is_dup = isinstance(sub_after, dict) and _same_sync_target(sub_after, app_after)
        leg_is_dup = any(
            _clean_link(new.get(k)) and _same_sync_target({"url": _clean_link(new.get(k))}, app_after)
            for k in LEGACY_LINK_KEYS
        )
        if sub_is_dup or leg_is_dup:
            new, _ = promote_legacy_link_to_reply(
                new,
                reply_url=app_after.get("url") or "",
                file_id=app_after.get("file_id") or "",
                folder_id=app_after.get("folder_id") or "",
                file_name=app_after.get("file_name") or "",
                mime_type=app_after.get("mime_type") or "",
            )
    return new


# ---------------- Document types with / without a consultant reply cycle ----------------
REPLY_CYCLE_CODES = frozenset({
    "SD", "DS", "MS", "IR", "MIR", "RFI", "NOC", "ABD", "WN", "NCN", "NCR",
    "EI", "SI", "CVI", "PCQ", "SVP", "PTW", "OBS", "INC", "JSA",
})
NO_REPLY_CYCLE_CODES = frozenset({
    "PR", "LTR", "L", "LETTERS", "COR", "CORRESPONDENCE", "MOM", "IOM", "QS",
    "TBT", "FEI", "CEV", "EVAC", "MHR", "WKR",
    "WO", "PPM", "CM", "MR", "SLA",
})
OM_DEFAULT_NO_CYCLE_CODES = frozenset({"WO", "PPM", "CM", "INC", "MR", "SLA", "COR"})
_NO_CYCLE_NAME_HINTS = (
    "requisition", "purchase", "letter", "correspondence", "minutes", "mom",
    "internal office memo", "quantity surveyor", "daily report", "site log",
    "boq", "specs", "variation", "toolbox", "extinguisher", "contractor evaluation",
    "evacuation", "man-hour", "man hour", "worker register",
)
_CYCLE_NAME_HINTS = (
    "shop drawing", "drawing", "material submittal", "submittal", "inspection request",
    "request for information", "notice of change", "as-built", "as built",
    "work notice", "work notification", "works notice", "non-conformance", "non conformance",
    "instruction", "potential change", "safety violation", "permit to work",
    "safety observation", "job safety analysis",
)


def doc_type_has_reply_cycle(dt, workspace_id=None):
    """True when documents of this type use dual Submit / Replied folders and slots.

    An explicit boolean 'has_reply_cycle' on the doc type always wins across all workspaces.
    For legacy rows without the flag, falls back deterministically to built-in register rules.
    """
    dt = dt or {}
    flag = dt.get("has_reply_cycle")
    if isinstance(flag, bool):
        return flag
    co_flag = dt.get("has_closeout_slot")
    if isinstance(co_flag, bool) and co_flag:
        return True
    ws = str(workspace_id or dt.get("workspace_id") or "").strip().upper()
    code = str(dt.get("code") or dt.get("id") or "").strip().upper()
    if ws in ("OM", "O&M") and code in OM_DEFAULT_NO_CYCLE_CODES:
        return False
    if code in REPLY_CYCLE_CODES:
        return True
    if code in NO_REPLY_CYCLE_CODES:
        return False
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
    rc_flag = dt.get("has_reply_cycle")
    if isinstance(rc_flag, bool):
        return rc_flag
    code = str(dt.get("code") or dt.get("id") or "").strip().upper()
    return code in HSE_CLOSEOUT_CODES


def doc_type_allows_approval_slot(dt, workspace_id="ENG"):
    """True when SLOT_APPROVAL is permitted for `(dt, workspace_id)`."""
    dt = dt or {}
    if isinstance(dt.get("has_reply_cycle"), bool):
        return dt["has_reply_cycle"]
    ws = str(workspace_id or dt.get("workspace_id") or "ENG").strip().upper()
    if ws == "HSE":
        return is_hse_closeout_doc_type(dt)
    return doc_type_has_reply_cycle(dt, workspace_id=ws)


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
