"""drive_routing.py - Resolve the destination Google Drive folder for a DCR upload.

(project root folder) -> (doc type folder) -> (Submit | Replied sub-folder)

Design rules
- The Drive service and the DB access layer are injected, so everything is testable with mocks.
- A resolved folder id is cached in doc_types.data (drive_folder_id / drive_slot_folders).
- Ambiguous matches are NEVER guessed: the caller gets an error listing the candidates.
- Folders are only created when create_if_missing=True (default False) so the existing
  Drive structure is never modified implicitly.
"""
import re
import datetime
import logging
import threading

logger = logging.getLogger(__name__)

FOLDER_MIME = "application/vnd.google-apps.folder"
DELETED_ARCHIVE_FOLDER_NAME = "_DCR_Deleted"

# Drive folder names that mean each slot (same keywords drive_service.get_item_priority uses)
SLOT_KEYWORDS = {
    "submission": ("submit",),
    "approval": ("reply", "replied"),
    "part_a": ("part a", "part_a", "part-a"),
    "part_b": ("part b", "part_b", "part-b"),
    "part_c": ("part c", "part_c", "part-c"),
    "part_d": ("part d", "part_d", "part-d"),
}
SLOT_DEFAULT_NAMES = {
    "submission": "Submit",
    "approval": "Replied",
    "part_a": "Part A",
    "part_b": "Part B",
    "part_c": "Part C",
    "part_d": "Part D",
}

SCORE_EXACT, SCORE_PREFIX, SCORE_TOKEN, SCORE_NAME_IN = 100, 90, 80, 60


class RoutingError(Exception):
    def __init__(self, reason, candidates=None):
        super().__init__(reason)
        self.reason = reason
        self.candidates = candidates or []


def _tokens(s):
    return [t for t in re.split(r"[^a-z0-9]+", str(s or "").lower()) if t]


def _norm(s):
    return " ".join(_tokens(s))


def extract_project_party_aliases(project):
    """Return dict of {code_or_token: phrase, phrase: code_or_token} from project metadata."""
    aliases = {}
    if not isinstance(project, dict):
        return aliases
    party_fields = ["client", "consultant", "contractor", "pmo", "landlord", "mainConsultant", "mepConsultant", "mep"]
    for pf in party_fields:
        val = str(project.get(pf) or "").strip()
        if not val:
            continue
        m = re.search(r"^(.*?)\s*\(([A-Za-z0-9_-]+)\)\s*$", val)
        if m:
            full_name = _norm(m.group(1))
            code = _norm(m.group(2))
            if code and full_name:
                aliases[code] = full_name
                aliases[full_name] = code
        else:
            norm_val = _norm(val)
            if norm_val:
                tokens = _tokens(val)
                initials = "".join(t[0] for t in tokens if t)
                if len(initials) >= 2:
                    aliases[initials] = norm_val
                    aliases[norm_val] = initials
    return aliases


def resolve_hierarchical_subfolder(service, dt_folder_id, project, direction="",
                                   party="", from_party="", to_party="", doc_no=""):
    """Drill down through company and direction subfolders (e.g. Letters > Style Design > Sent)."""
    children = list_child_folders(service, dt_folder_id)
    if not children:
        return dt_folder_id

    # If this is a standard submission/approval layout, let standard slot routing handle it
    if any(any(k in str(c.get("name", "")).lower() for k in SLOT_KEYWORDS["submission"] + SLOT_KEYWORDS["approval"]) for c in children):
        return dt_folder_id

    d_norm = str(direction or "").lower()
    is_inc = any(k in d_norm for k in ("receiv", "incom", "rec", "وارد")) or d_norm in ("in", "i")
    is_out = any(k in d_norm for k in ("issu", "outgo", "sent", "iss", "صادر")) or d_norm in ("out", "o")

    # Check if children are directly Sent / Received / Incoming / Outgoing
    inc_kws = ("in", "incom", "incoming", "receiv", "received", "وارد")
    out_kws = ("out", "outgo", "outgoing", "sent", "issu", "issued", "صادر")
    dir_hits = [c for c in children if any(k in _tokens(c["name"]) for k in inc_kws + out_kws)]
    if dir_hits and len(dir_hits) == len(children):
        if is_inc:
            match = next((c for c in children if any(k in _tokens(c["name"]) for k in inc_kws)), None)
            if match:
                return match["id"]
        elif is_out:
            match = next((c for c in children if any(k in _tokens(c["name"]) for k in out_kws)), None)
            if match:
                return match["id"]
        raise RoutingError("ambiguous", [{"id": c["id"], "name": c["name"]} for c in children])

    # Otherwise children are company / party folders (e.g. 'Style Design', 'SCAS', 'EKH', 'SCAS to SD')
    aliases = extract_project_party_aliases(project)
    p_tokens = set()
    for raw in (party, from_party, to_party, doc_no):
        if raw:
            for t in _tokens(raw):
                p_tokens.add(t)
                if t in aliases:
                    p_tokens.update(_tokens(aliases[t]))

    scored_comps = []
    for c in children:
        c_toks = _tokens(c["name"])
        c_norm = _norm(c["name"])
        score = 0
        if party and (_norm(party) == c_norm or _norm(aliases.get(_norm(party), "")) == c_norm):
            score += 100
        if from_party and (_norm(from_party) == c_norm or _norm(aliases.get(_norm(from_party), "")) == c_norm):
            score += 85
        if to_party and (_norm(to_party) == c_norm or _norm(aliases.get(_norm(to_party), "")) == c_norm):
            score += 85
        common = p_tokens.intersection(set(c_toks))
        score += len(common) * 20
        if "to" in c_toks and len(common) >= 2:
            score += 35
        if score > 0:
            scored_comps.append((score, c))

    scored_comps.sort(key=lambda x: -x[0])
    if scored_comps:
        top_score = scored_comps[0][0]
        chosen_comps = [c for s, c in scored_comps if s == top_score]
    else:
        # None of the child folders match any company / party alias; upload directly to parent folder
        return dt_folder_id

    final_candidates = []
    for comp in chosen_comps:
        comp_children = list_child_folders(service, comp["id"])
        if not comp_children:
            final_candidates.append({"id": comp["id"], "name": comp["name"]})
            continue

        matched_dir = None
        if is_inc:
            matched_dir = next((sc for sc in comp_children if any(k in _tokens(sc["name"]) for k in inc_kws)), None)
        elif is_out:
            matched_dir = next((sc for sc in comp_children if any(k in _tokens(sc["name"]) for k in out_kws)), None)

        if matched_dir:
            final_candidates.append({"id": matched_dir["id"], "name": f"{comp['name']} ❯ {matched_dir['name']}"})
        else:
            for sc in comp_children:
                final_candidates.append({"id": sc["id"], "name": f"{comp['name']} ❯ {sc['name']}"})

    if len(final_candidates) == 1:
        return final_candidates[0]["id"]
    if len(final_candidates) > 1:
        raise RoutingError("ambiguous", final_candidates)
    return dt_folder_id


def score_folder_name(folder_name, code, name, direction=""):
    """Score how well a Drive folder name matches a doc type (code + display name)."""
    raw_name = str(folder_name or "").strip()
    f_tokens, f_norm = _tokens(raw_name), _norm(raw_name)
    c_norm, n_norm = _norm(code), _norm(name)
    c_tokens = _tokens(code)
    n_tokens = _tokens(name)
    if not f_norm:
        return 0

    unindexed = re.sub(r"^\d+[\s._-]+", "", raw_name).strip()
    u_tokens, u_norm = _tokens(unindexed), _norm(unindexed)

    has_code_in_parens = bool(c_norm and re.search(r"[\(\[]\s*" + re.escape(str(code).strip()) + r"\s*[\)\]]", raw_name, re.I))
    has_code_token = bool(c_tokens and any(t in f_tokens for t in c_tokens))
    has_name_token = bool(n_tokens and any(t in f_tokens for t in n_tokens))

    if f_norm in (c_norm, n_norm):
        score = SCORE_EXACT
    elif has_code_in_parens and has_name_token:
        # Full category folder with name & parenthesized code, e.g. '9. Requisition (PR)'
        score = 98
    elif has_code_in_parens:
        # Parenthesized code, e.g. '9. (PR)'
        score = 95
    elif c_tokens and f_tokens[:len(c_tokens)] == c_tokens:
        score = SCORE_PREFIX
    elif has_code_token and has_name_token:
        # Both code and name present, e.g. 'Requisition PR'
        score = 92
    elif has_code_token:
        score = SCORE_TOKEN
    elif n_norm and (" " + n_norm + " ") in (" " + f_norm + " "):
        score = SCORE_NAME_IN
    elif n_tokens and (f_tokens[:len(n_tokens)] == n_tokens or u_tokens[:len(n_tokens)] == n_tokens):
        score = SCORE_NAME_IN
    else:
        score = 0

    # Direction-based tie-breaker (e.g. for Letters where Drive has 'Incoming Letters' vs 'Outgoing Letters')
    if score > 0 and direction:
        d_norm = str(direction).lower()
        is_incoming = any(k in d_norm for k in ("receiv", "incom", "rec", "وارد")) or d_norm in ("in", "i")
        is_outgoing = any(k in d_norm for k in ("issu", "outgo", "sent", "iss", "صادر")) or d_norm in ("out", "o")
        inc_hit = any(t.startswith("incom") or t.startswith("receiv") or t in ("in", "وارد") for t in f_tokens)
        out_hit = any(t.startswith("outgo") or t.startswith("issu") or t.startswith("sent") or t in ("out", "صادر") for t in f_tokens)
        if is_incoming:
            if inc_hit:
                score += 35
            elif out_hit:
                score -= 35
        elif is_outgoing:
            if out_hit:
                score += 35
            elif inc_hit:
                score -= 35

    # Penalize specific document / revision subfolders (e.g. 'PR-373 REV00')
    if score > 0:
        if re.search(r"\b(rev|revision)[._\s-]*\d+\b", raw_name, re.I):
            score -= 40
        if c_norm and re.search(r"\b" + re.escape(str(code).strip()) + r"[-_]\d+\b", raw_name, re.I):
            score -= 30

    return max(0, score)


def match_doc_type_folder(folders, code, name, direction=""):
    """Pick the folder for a doc type. Returns folder dict, None (no match).
    Raises RoutingError('ambiguous') if the top score is shared by several folders."""
    scored = sorted(((score_folder_name(f.get("name"), code, name, direction=direction), f) for f in folders),
                    key=lambda x: -x[0])
    scored = [(s, f) for s, f in scored if s > 0]
    if not scored:
        return None
    top = scored[0][0]
    best = [f for s, f in scored if s == top]
    if len(best) > 1:
        raise RoutingError("ambiguous", [{"id": f["id"], "name": f["name"]} for f in best])
    return best[0]


def match_slot_folder(folders, slot):
    """Find the Submit / Replied / Part A..D sub-folder inside a doc type folder."""
    if slot in ("part_a", "part_b", "part_c", "part_d"):
        letter = slot.split("_")[-1]
        pat = re.compile(r"\bpart[\s_-]*" + re.escape(letter) + r"\b", re.IGNORECASE)
        hits = [f for f in folders if pat.search(str(f.get("name", "")))]
        if len(hits) > 1:
            raise RoutingError("ambiguous", [{"id": f["id"], "name": f["name"]} for f in hits])
        return hits[0] if hits else None
    keys = SLOT_KEYWORDS[slot]
    hits = [f for f in folders if any(k in str(f.get("name", "")).lower() for k in keys)]
    # 'Replied' must not be taken for the Submit slot and vice versa
    if slot == "submission":
        hits = [f for f in hits if "repl" not in f["name"].lower()]
    if len(hits) > 1:
        raise RoutingError("ambiguous", [{"id": f["id"], "name": f["name"]} for f in hits])
    return hits[0] if hits else None


# ---------------- Drive helpers (service is injected) ----------------
import time
_FOLDER_USABLE_CACHE = {}   # folder_id -> (timestamp, bool)
_CHILD_FOLDERS_CACHE = {}   # parent_id -> (timestamp, list)
_SUBFOLDER_LOCKS_GUARD = threading.Lock()
_SUBFOLDER_LOCKS = {}       # (parent_id, seg_lower) -> threading.Lock
_NESTED_SUBFOLDER_CACHE = {}  # (root_folder_id, rel_lower) -> (timestamp, leaf_id, ancestor_ids)
MAX_SUBFOLDER_DEPTH = 10


def _get_subfolder_lock(parent_id, seg_name):
    key = (str(parent_id), str(seg_name).strip().lower())
    with _SUBFOLDER_LOCKS_GUARD:
        lk = _SUBFOLDER_LOCKS.get(key)
        if lk is None:
            lk = threading.Lock()
            _SUBFOLDER_LOCKS[key] = lk
        return lk


def is_real_service(service):
    return type(service).__name__ != "FakeDrive" and not hasattr(service, "items")

def list_child_folders(service, parent_id):
    now = time.time()
    if is_real_service(service):
        cached = _CHILD_FOLDERS_CACHE.get(parent_id)
        if cached and (now - cached[0]) < 60:
            return cached[1]
    q = ("'%s' in parents and mimeType = '%s' and trashed = false" % (parent_id, FOLDER_MIME))
    out, token = [], None
    while True:
        res = service.files().list(
            q=q, fields="nextPageToken, files(id, name)", pageSize=1000, pageToken=token,
            supportsAllDrives=True, includeItemsFromAllDrives=True).execute()
        out.extend(res.get("files", []))
        token = res.get("nextPageToken")
        if not token:
            break
    if is_real_service(service):
        _CHILD_FOLDERS_CACHE[parent_id] = (now, out)
    return out


def folder_is_usable(service, folder_id):
    now = time.time()
    if is_real_service(service):
        cached = _FOLDER_USABLE_CACHE.get(folder_id)
        if cached and (now - cached[0]) < 300:
            return cached[1]
    try:
        f = service.files().get(fileId=folder_id, fields="id, mimeType, trashed",
                                supportsAllDrives=True).execute()
        res = (f.get("mimeType") == FOLDER_MIME and not f.get("trashed", False))
    except Exception:
        res = False
    if is_real_service(service):
        _FOLDER_USABLE_CACHE[folder_id] = (now, res)
    return res


def create_folder(service, parent_id, name):
    clean_name = sanitize_drive_folder_segment(name, max_len=120) or str(name or "").strip() or "Folder"
    body = {"name": clean_name, "mimeType": FOLDER_MIME, "parents": [parent_id]}
    res = service.files().create(body=body, fields="id, name", supportsAllDrives=True).execute()
    if res and res.get("id"):
        _FOLDER_USABLE_CACHE[res["id"]] = (time.time(), True)
        if is_real_service(service) and parent_id in _CHILD_FOLDERS_CACHE:
            ts, children = _CHILD_FOLDERS_CACHE[parent_id]
            updated = list(children)
            if not any(str(c.get("id")) == str(res["id"]) for c in updated):
                updated.append({"id": res["id"], "name": res.get("name") or clean_name})
            _CHILD_FOLDERS_CACHE[parent_id] = (time.time(), updated)
        else:
            _CHILD_FOLDERS_CACHE.pop(parent_id, None)
    else:
        _CHILD_FOLDERS_CACHE.pop(parent_id, None)
    return res


def rename_folder(service, folder_id, new_name, parent_id=None):
    """Rename an existing Drive folder via PATCH files/{id} and refresh caches."""
    clean_name = sanitize_drive_folder_segment(new_name, max_len=120)
    if not folder_id or not clean_name:
        raise RoutingError("invalid_folder_name")
    lock_parent = parent_id or folder_id
    with _get_subfolder_lock(lock_parent, clean_name):
        res = service.files().update(
            fileId=folder_id,
            body={"name": clean_name},
            fields="id, name",
            supportsAllDrives=True,
        ).execute()
        _FOLDER_USABLE_CACHE[folder_id] = (time.time(), True)
        if parent_id:
            invalidate_subfolder_caches(parent_id)
            _FOLDER_USABLE_CACHE[parent_id] = (time.time(), True)
        else:
            invalidate_subfolder_caches()
            _FOLDER_USABLE_CACHE[folder_id] = (time.time(), True)
        return res or {"id": folder_id, "name": clean_name}


def ensure_deleted_archive_folder(service, project_root_id):
    """Find or create `<Project Root>/_DCR_Deleted` with thread-safe locking and supportsAllDrives."""
    root_id = extract_drive_folder_id(project_root_id)
    if not root_id:
        raise RoutingError("project_has_no_drive_folder")
    with _get_subfolder_lock(root_id, DELETED_ARCHIVE_FOLDER_NAME):
        children = list_child_folders(service, root_id)
        match = next(
            (f for f in children if str(f.get("name", "")).strip().lower() == DELETED_ARCHIVE_FOLDER_NAME.lower()),
            None,
        )
        if match is not None:
            return match
        return create_folder(service, root_id, DELETED_ARCHIVE_FOLDER_NAME)


def archive_drive_item(service, project_root_id, item_id, original_name=None):
    """Move a Drive file or folder into `<Project Root>/_DCR_Deleted`, avoiding name collisions."""
    clean_item_id = str(item_id or "").strip()
    if not clean_item_id:
        raise RoutingError("invalid_item_id")
    archive_folder = ensure_deleted_archive_folder(service, project_root_id)
    archive_folder_id = str(archive_folder["id"]).strip()

    meta = service.files().get(
        fileId=clean_item_id,
        fields="id,name,parents,trashed",
        supportsAllDrives=True,
    ).execute() or {}
    current_name = str(original_name or meta.get("name") or "").strip()

    existing_names = set()
    try:
        q = "'%s' in parents and trashed = false" % archive_folder_id
        existing_res = service.files().list(
            q=q,
            fields="nextPageToken, files(id, name)",
            pageSize=1000,
            supportsAllDrives=True,
            includeItemsFromAllDrives=True,
        ).execute() or {}
        for f in (existing_res.get("files") or []):
            if str(f.get("id") or "") != clean_item_id:
                fname = str(f.get("name") or "").strip().lower()
                if fname:
                    existing_names.add(fname)
    except Exception:
        existing_names = set()

    new_name = None
    if current_name and current_name.lower() in existing_names:
        ts = datetime.datetime.now(datetime.timezone.utc).strftime("%Y%m%d-%H%M%S")
        candidate = f"{ts}__{current_name}"
        if candidate.lower() in existing_names:
            ms = datetime.datetime.now(datetime.timezone.utc).strftime("%f")[:3]
            candidate = f"{ts}-{ms}__{current_name}"
        new_name = candidate

    cur_parents = [str(p) for p in (meta.get("parents") or []) if str(p)]
    old_parents = [p for p in cur_parents if p != archive_folder_id]

    update_kwargs = {
        "fileId": clean_item_id,
        "body": ({"name": new_name} if new_name else {}),
        "fields": "id,name,parents,trashed",
        "supportsAllDrives": True,
    }
    if archive_folder_id not in cur_parents:
        update_kwargs["addParents"] = archive_folder_id
    if old_parents:
        update_kwargs["removeParents"] = ",".join(old_parents)

    if update_kwargs.get("addParents") or update_kwargs.get("removeParents") or new_name:
        moved = service.files().update(**update_kwargs).execute() or {}
    else:
        moved = meta

    for pid in old_parents:
        invalidate_subfolder_caches(pid)
    invalidate_subfolder_caches(archive_folder_id)

    return {
        "ok": True,
        "item_id": clean_item_id,
        "archive_folder_id": archive_folder_id,
        "archive_folder_name": archive_folder.get("name") or DELETED_ARCHIVE_FOLDER_NAME,
        "name": moved.get("name") or new_name or current_name,
        "renamed": bool(new_name),
    }


def extract_drive_folder_id(raw_value):
    """Extract a clean Drive folder ID from a Drive URL or raw ID string."""
    val = str(raw_value or "").strip()
    if not val:
        return ""
    m = re.search(r"/folders/([a-zA-Z0-9_-]+)", val)
    return m.group(1) if m else val


def format_doc_type_folder_name(code, name):
    """Format and sanitize a standardized `<CODE> - <Name>` Drive folder segment."""
    c = sanitize_drive_folder_segment(code, max_len=30)
    n = sanitize_drive_folder_segment(name, max_len=80)
    if c and n and c.lower() != n.lower():
        return sanitize_drive_folder_segment(f"{c} - {n}", max_len=110)
    return c or n or "Documents"


def format_workspace_folder_name(ws_id, ws_name=None):
    """Format and sanitize `<WS_ID> - <Workspace Name>` for a Department/Workspace Drive folder."""
    c = sanitize_drive_folder_segment(str(ws_id or "").strip().upper(), max_len=30)
    n = sanitize_drive_folder_segment(ws_name, max_len=80)
    if c and n and c.lower() != n.lower():
        return sanitize_drive_folder_segment(f"{c} - {n}", max_len=115)
    return c or n or "Workspace"


def format_project_folder_name(project):
    """Format and sanitize `<Code> - <Name>` (or `<Name>`) for a Project Drive folder."""
    if not isinstance(project, dict):
        return sanitize_drive_folder_segment(project, max_len=115) or "Project"
    code = sanitize_drive_folder_segment(project.get("code") or "", max_len=30)
    name = sanitize_drive_folder_segment(project.get("name") or project.get("id") or "", max_len=85)
    if code and name and code.lower() != name.lower():
        return sanitize_drive_folder_segment(f"{code} - {name}", max_len=115)
    return name or code or sanitize_drive_folder_segment(project.get("id") or "", max_len=85) or "Project"


def invalidate_subfolder_caches(parent_id=None):
    """Clear child/nested subfolder caches (used when folders are trashed or tests reset state)."""
    if parent_id is None:
        _CHILD_FOLDERS_CACHE.clear()
        _NESTED_SUBFOLDER_CACHE.clear()
        _FOLDER_USABLE_CACHE.clear()
        return
    _CHILD_FOLDERS_CACHE.pop(parent_id, None)
    _FOLDER_USABLE_CACHE.pop(parent_id, None)
    for k in list(_NESTED_SUBFOLDER_CACHE.keys()):
        if k[0] == parent_id:
            _NESTED_SUBFOLDER_CACHE.pop(k, None)


def sanitize_drive_folder_segment(name, max_len=80):
    """Sanitize a single Drive subfolder name by stripping unsafe path/control characters."""
    s = str(name or "")
    s = re.sub(r"[\x00-\x1f\x7f]+", " ", s)
    while ".." in s:
        s = s.replace("..", " ")
    s = re.sub(r"\s*[/\\]+\s*", " - ", s)
    s = re.sub(r"\s+", " ", s).strip(" .-")
    if max_len and len(s) > max_len:
        s = s[:max_len].rstrip(" .-")
    return s


def normalize_rel_subfolder(rel_subfolder):
    """Parse and validate a relative subfolder path into clean directory segments."""
    raw = str(rel_subfolder or "").replace("\\", "/").strip()
    if not raw:
        return ()
    if any(ord(ch) < 32 or ord(ch) == 127 for ch in raw):
        raise RoutingError("invalid_subfolder")
    parts = []
    for seg in raw.split("/"):
        s = seg.strip()
        if not s:
            continue
        if s in (".", ".."):
            raise RoutingError("invalid_subfolder")
        if len(s) > 200:
            raise RoutingError("invalid_subfolder")
        parts.append(s)
    if len(parts) > MAX_SUBFOLDER_DEPTH:
        raise RoutingError("invalid_subfolder")
    return tuple(parts)


def resolve_nested_subfolder_with_ancestors(service, root_folder_id, rel_subfolder, create_if_missing=True):
    """Walk or create nested subfolders inside `root_folder_id`, returning `(leaf_id, ancestor_ids)`.

    `ancestor_ids` lists every intermediate/leaf subfolder id created or traversed under
    `root_folder_id` (excluding `root_folder_id` itself), ordered root-to-leaf.
    Thread-safe per `(parent_id, segment)` so parallel uploads never create duplicate folders.
    """
    if not root_folder_id:
        raise RoutingError("missing_root_folder")
    parts = normalize_rel_subfolder(rel_subfolder)
    if not parts:
        return root_folder_id, []

    cache_key = (str(root_folder_id), "/".join(p.lower() for p in parts))
    now = time.time()
    if is_real_service(service):
        cached = _NESTED_SUBFOLDER_CACHE.get(cache_key)
        if cached and (now - cached[0]) < 60:
            return cached[1], list(cached[2])

    cur_parent = root_folder_id
    ancestors = []
    for seg in parts:
        lk = _get_subfolder_lock(cur_parent, seg)
        with lk:
            children = list_child_folders(service, cur_parent)
            match = next(
                (f for f in children if str(f.get("name", "")).strip().lower() == seg.lower()),
                None,
            )
            if match is None:
                if not create_if_missing:
                    raise RoutingError("subfolder_not_found")
                match = create_folder(service, cur_parent, seg)
            cur_parent = match["id"]
            ancestors.append(cur_parent)

    if is_real_service(service):
        _NESTED_SUBFOLDER_CACHE[cache_key] = (time.time(), cur_parent, tuple(ancestors))
    return cur_parent, ancestors


def resolve_nested_subfolder(service, root_folder_id, rel_subfolder, create_if_missing=True):
    """Resolve/create nested subfolder path under `root_folder_id` and return its Drive folder id."""
    leaf_id, _ = resolve_nested_subfolder_with_ancestors(
        service, root_folder_id, rel_subfolder, create_if_missing=create_if_missing
    )
    return leaf_id


# ---------------- Orchestration ----------------
_service_provider = None


def get_default_write_service():
    """Return an authenticated Drive write adapter, or None if Drive credentials are unconfigured."""
    if _service_provider is not None:
        return _service_provider()
    try:
        import blueprints.drive_uploads as du
        import drive_upload
        if du._token_provider is not None or du._http_client is not None:
            token = du._get_token()
            if not token:
                return None
            return drive_upload.RestDriveAdapter(token, http=du._get_http())
    except Exception:
        pass
    try:
        import drive_upload
        creds = drive_upload.get_upload_credentials()
        token = drive_upload.access_token(creds)
        if not token:
            return None
        return drive_upload.RestDriveAdapter(token)
    except Exception:
        return None


def _is_noc_doc_type(dt):
    code = str((dt or {}).get("code") or (dt or {}).get("id") or "").strip().upper()
    name = str((dt or {}).get("name") or "").strip().lower()
    return code == "NOC" or "notice of change" in name


def _is_ltr_doc_type(dt):
    code = str((dt or {}).get("code") or (dt or {}).get("id") or "").strip().upper()
    name = str((dt or {}).get("name") or "").strip().lower()
    return code in ("LTR", "L", "LETTERS") or "letter" in name


def _is_correspondence_doc_type(dt):
    code = str((dt or {}).get("code") or (dt or {}).get("id") or "").strip().upper()
    name = str((dt or {}).get("name") or "").strip().lower()
    return _is_ltr_doc_type(dt) or code in ("COR", "CORRESPONDENCE") or "correspond" in name


def _is_bad_rev_folder_name(cached_name, code=""):
    low = str(cached_name or "").strip().lower()
    if not low:
        return False
    if re.search(r"\b(rev|revision)[._\s-]*\d+\b", low, re.I):
        return True
    if code and re.search(r"\b" + re.escape(str(code).strip()) + r"[-_]\d+\b", low, re.I):
        return True
    return False


def resolve_upload_destination(pid, dt_id, slot, service, store, create_if_missing=False,
                               subfolder_name=None, flat=False, direction="", destination_folder_id="",
                               party="", from_party="", to_party="", doc_no="", subtype=""):
    """Return {'folder_id', 'doc_type_folder_id', 'source', ...} or raise RoutingError.

    store must provide: get_project(pid), get_doc_type(pid, dt_id),
    set_doc_type_drive_cache(pid, dt_id, patch_dict).
    - Reply-cycle document types (`flat=False`) route into
      `<Project Root>/<DocType Folder>/[<Subtype>/]/Submit` or `Replied` (or `Part A..D` for NOC).
    - Non-reply-cycle document types (`flat=True`) route directly into
      `<Project Root>/<DocType Folder>/[<Subtype>/]`.
    - If `subfolder_name` is provided, creates/finds a sanitized `<DocNo>` subfolder inside the resolved slot.
    """
    if slot not in SLOT_KEYWORDS:
        raise RoutingError("invalid_slot")
    project = store.get_project(pid) or {}
    root_id = extract_drive_folder_id(project.get("drive_folder_id"))
    if not root_id:
        raise RoutingError("project_has_no_drive_folder")
    dt = store.get_doc_type(pid, dt_id)
    if not dt:
        raise RoutingError("doc_type_not_found")

    from hse_defaults import normalize_ptw_subtype_value
    clean_sub = sanitize_drive_folder_segment(normalize_ptw_subtype_value(subtype))

    source = "cache"
    patch = {}
    slot_cache = dict(dt.get("drive_slot_folders") or {})
    dir_cache = dict(dt.get("drive_direction_folders") or {})
    sub_cache = dict(dt.get("drive_subtype_folders") or {})
    dt_folder_id = ""

    d_norm = str(direction or "").lower()
    is_incoming = any(k in d_norm for k in ("receiv", "incom", "rec", "وارد")) or d_norm in ("in", "i")
    is_outgoing = any(k in d_norm for k in ("issu", "outgo", "sent", "iss", "صادر")) or d_norm in ("out", "o")
    dir_key = "incoming" if is_incoming else ("outgoing" if is_outgoing else "")

    # Explicit destination folder chosen by caller (e.g. resolved after ambiguous picker)
    if destination_folder_id:
        if folder_is_usable(service, destination_folder_id):
            dt_folder_id = destination_folder_id
            source = "explicit"
            now_iso = datetime.datetime.now(datetime.timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")
            if dir_key:
                dir_cache[dir_key] = dt_folder_id
                patch.update({"drive_direction_folders": dir_cache, "folder_synced_at": now_iso})
            else:
                patch.update({"drive_folder_id": dt_folder_id, "folder_synced_at": now_iso})
        else:
            raise RoutingError("folder_not_usable")

    # 1) doc type folder
    if not dt_folder_id and dir_key and dir_cache.get(dir_key):
        cand_dir_id = str(dir_cache[dir_key]).strip()
        if folder_is_usable(service, cand_dir_id):
            dt_folder_id = cand_dir_id
            source = "cache"

    if not dt_folder_id:
        cand_id = str(dt.get("drive_folder_id") or "").strip()
        cached_name = str(dt.get("drive_folder_name") or "").lower()
        opposite_inc = is_outgoing and any(k in cached_name for k in ("incom", "receiv", "وارد"))
        opposite_out = is_incoming and any(k in cached_name for k in ("outgo", "issu", "sent", "صادر"))
        is_bad_rev_folder = _is_bad_rev_folder_name(cached_name, dt.get("code"))
        if cand_id and not (opposite_inc or opposite_out or is_bad_rev_folder) and folder_is_usable(service, cand_id):
            dt_folder_id = cand_id
            source = "cache"

    if not dt_folder_id:
        new_dt_folder_name = format_doc_type_folder_name(dt.get("code") or dt_id, dt.get("name"))
        dt_lock_name = f"{new_dt_folder_name}:{dir_key}" if dir_key else new_dt_folder_name
        with _get_subfolder_lock(root_id, dt_lock_name):
            source = "discovered"
            folder = match_doc_type_folder(list_child_folders(service, root_id), dt.get("code"), dt.get("name"), direction=direction)
            if folder is None:
                if not create_if_missing:
                    raise RoutingError("doc_type_folder_not_found")
                folder = create_folder(service, root_id, new_dt_folder_name)
                source = "created"
            dt_folder_id = folder["id"]
            slot_cache = {}
            sub_cache = {}
            now_iso = datetime.datetime.now(datetime.timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")
            if dir_key:
                dir_cache[dir_key] = dt_folder_id
                patch.update({"drive_direction_folders": dir_cache,
                              "folder_synced_at": now_iso})
            else:
                patch.update({"drive_folder_id": dt_folder_id, "drive_folder_name": folder.get("name", ""),
                              "drive_slot_folders": {},
                              "folder_synced_at": now_iso})

    # 2) Generalized optional Subtype subfolder (<DocType Folder>/<Subtype>/)
    sub_folder_id = ""
    sub_key = clean_sub.lower() if clean_sub else ""
    sub_entry = {}
    if destination_folder_id and folder_is_usable(service, destination_folder_id):
        slot_parent_id = destination_folder_id
    elif clean_sub:
        raw_sub_entry = sub_cache.get(sub_key)
        if isinstance(raw_sub_entry, dict):
            sub_entry = dict(raw_sub_entry)
            sub_folder_id = str(sub_entry.get("id") or "").strip()
        elif isinstance(raw_sub_entry, str) and raw_sub_entry.strip():
            sub_folder_id = raw_sub_entry.strip()
            sub_entry = {"id": sub_folder_id, "name": clean_sub, "slots": {}}
        if not sub_folder_id:
            legacy_ptw_id = str(slot_cache.get(f"ptw_subtype:{sub_key}") or "").strip()
            if legacy_ptw_id:
                sub_folder_id = legacy_ptw_id
                sub_entry = {"id": sub_folder_id, "name": clean_sub, "slots": {}}
        if sub_folder_id and not folder_is_usable(service, sub_folder_id):
            sub_folder_id = ""
            sub_entry = {}
        if not sub_folder_id:
            with _get_subfolder_lock(dt_folder_id, clean_sub):
                children = list_child_folders(service, dt_folder_id)
                sub_match = next(
                    (f for f in children if str(f.get("name", "")).strip().lower() == sub_key),
                    None,
                )
                if sub_match is None:
                    if not create_if_missing:
                        raise RoutingError("subfolder_not_found")
                    sub_match = create_folder(service, dt_folder_id, clean_sub)
                    source = "created"
                elif source == "cache":
                    source = "discovered"
                sub_folder_id = sub_match["id"]
                sub_entry = {
                    "id": sub_folder_id,
                    "name": sub_match.get("name") or clean_sub,
                    "slots": dict(sub_entry.get("slots") or {}),
                }
                sub_cache[sub_key] = sub_entry
                patch["drive_subtype_folders"] = sub_cache
        slot_parent_id = sub_folder_id
    else:
        slot_parent_id = dt_folder_id

    # 3) Submit / Replied / Part A..D slot folder (or flat destination)
    slot_id = ""
    if destination_folder_id and folder_is_usable(service, destination_folder_id):
        slot_id = destination_folder_id
    elif flat:
        if clean_sub:
            slot_id = slot_parent_id
        else:
            is_letter = _is_correspondence_doc_type(dt)
            has_party_or_dir = bool(direction or party or from_party or to_party)
            if is_letter or has_party_or_dir:
                slot_id = resolve_hierarchical_subfolder(
                    service, dt_folder_id, project,
                    direction=direction, party=party,
                    from_party=from_party, to_party=to_party, doc_no=doc_no
                )
            else:
                slot_id = dt_folder_id
    else:
        if clean_sub:
            sub_slots = dict(sub_entry.get("slots") or {})
            slot_id = str(sub_slots.get(slot) or "").strip()
            if slot_id and not folder_is_usable(service, slot_id):
                slot_id = ""
            if not slot_id:
                slot_default_name = SLOT_DEFAULT_NAMES[slot]
                with _get_subfolder_lock(slot_parent_id, slot_default_name):
                    sub = match_slot_folder(list_child_folders(service, slot_parent_id), slot)
                    if sub is None:
                        if not create_if_missing:
                            raise RoutingError("slot_folder_not_found")
                        sub = create_folder(service, slot_parent_id, slot_default_name)
                        source = "created"
                    elif source == "cache":
                        source = "discovered"
                    slot_id = sub["id"]
                    sub_slots[slot] = slot_id
                    sub_entry["slots"] = sub_slots
                    sub_cache[sub_key] = sub_entry
                    patch["drive_subtype_folders"] = sub_cache
        else:
            slot_id = str(slot_cache.get(slot) or "").strip()
            if slot_id and not folder_is_usable(service, slot_id):
                slot_id = ""
            if not slot_id:
                slot_default_name = SLOT_DEFAULT_NAMES[slot]
                with _get_subfolder_lock(slot_parent_id, slot_default_name):
                    sub = match_slot_folder(list_child_folders(service, slot_parent_id), slot)
                    if sub is None:
                        if not create_if_missing:
                            raise RoutingError("slot_folder_not_found")
                        sub = create_folder(service, slot_parent_id, slot_default_name)
                        source = "created"
                    elif source == "cache":
                        source = "discovered"
                    slot_id = sub["id"]
                    slot_cache[slot] = slot_id
                    patch["drive_slot_folders"] = slot_cache

    if patch:
        store.set_doc_type_drive_cache(pid, dt_id, patch)

    # 4) Optional subfolder per document (e.g. DS-CY002P608-00001 REV00)
    target_folder_id = slot_id
    subfolder_id = None
    clean_doc_sub = ""
    if subfolder_name:
        clean_doc_sub = sanitize_drive_folder_segment(subfolder_name, max_len=120)
        if clean_doc_sub:
            with _get_subfolder_lock(slot_id, clean_doc_sub):
                children = list_child_folders(service, slot_id)
                match = next((f for f in children if str(f.get("name", "")).strip().lower() == clean_doc_sub.lower()), None)
                if match is None:
                    if not create_if_missing:
                        raise RoutingError("subfolder_not_found")
                    match = create_folder(service, slot_id, clean_doc_sub)
                    source = "created"
                subfolder_id = match["id"]
                target_folder_id = subfolder_id

    result = {
        "folder_id": target_folder_id,
        "doc_type_folder_id": dt_folder_id,
        "slot_folder_id": slot_id,
        "source": source
    }
    if sub_folder_id:
        result["subtype_folder_id"] = sub_folder_id
    if subfolder_id:
        result["subfolder_id"] = subfolder_id
        result["subfolder_name"] = clean_doc_sub
    return result


def _collect_active_subtypes_for_doc_type(project_id, ws_id, dt, store, lists_map=None):
    """Return ordered list of sanitized active subtype names for `(project_id, dt)`."""
    from hse_defaults import normalize_ptw_subtype_value

    lists_dict = lists_map if isinstance(lists_map, dict) else {}
    if not lists_dict and hasattr(store, "get_lists"):
        try:
            lists_dict = store.get_lists(project_id) or {}
        except Exception:
            lists_dict = {}

    code = str(dt.get("code") or dt.get("id") or "").strip().upper()
    code_low = re.sub(r"[^a-z0-9]+", "_", code.lower()).strip("_")
    ws_low = re.sub(r"[^a-z0-9]+", "_", str(ws_id or "ENG").strip().lower()).strip("_") or "eng"

    list_name = ""
    if hasattr(store, "get_columns"):
        try:
            cols = store.get_columns(project_id, dt["id"]) or []
            for c in cols:
                if not isinstance(c, dict) or c.get("col_type") != "dropdown" or not c.get("list_name"):
                    continue
                ck = str(c.get("col_key") or "").strip().lower()
                ln = str(c.get("list_name") or "").strip()
                ln_low = ln.lower()
                if ck in ("legacysubtype", "subtype") or ln_low == "subtype" or ln_low.endswith("_subtype"):
                    list_name = ln
                    break
        except Exception:
            list_name = ""

    if not list_name and code_low:
        candidates = [f"{ws_low}_{code_low}_subtype", f"hse_{code_low}_subtype", f"{code_low}_subtype"]
        for cand in candidates:
            for existing_ln in lists_dict.keys():
                if str(existing_ln).strip().lower() == cand.lower():
                    list_name = existing_ln
                    break
            if list_name:
                break

    raw_items = []
    if list_name and isinstance(lists_dict.get(list_name), (list, tuple)):
        raw_items = list(lists_dict.get(list_name) or [])
    elif not lists_dict and isinstance(dt.get("legacy_subtypes"), dict):
        raw_items = list(dt["legacy_subtypes"].values())

    out = []
    seen = set()
    for item in raw_items:
        clean = sanitize_drive_folder_segment(normalize_ptw_subtype_value(item))
        if clean and clean.lower() not in seen:
            seen.add(clean.lower())
            out.append(clean)
    return out


def _ensure_slot_children(service, parent_id, slots_to_ensure, existing_slot_map=None):
    """Ensure each slot in `slots_to_ensure` exists under `parent_id`; return `(updated_map, created_cnt, existing_cnt)`."""
    slot_map = dict(existing_slot_map or {})
    created = 0
    existing = 0
    for slot in slots_to_ensure:
        cand_id = str(slot_map.get(slot) or "").strip()
        if cand_id and folder_is_usable(service, cand_id):
            existing += 1
            continue
        default_name = SLOT_DEFAULT_NAMES[slot]
        with _get_subfolder_lock(parent_id, default_name):
            children = list_child_folders(service, parent_id)
            matched = match_slot_folder(children, slot)
            if matched is not None:
                slot_map[slot] = matched["id"]
                existing += 1
            else:
                created_f = create_folder(service, parent_id, default_name)
                slot_map[slot] = created_f["id"]
                created += 1
    return slot_map, created, existing


def _ensure_ltr_direction_children(service, dt_folder_id, existing_dir_map=None):
    """Ensure directional `Sent` and `Received` folders exist under `dt_folder_id` when no company folders exist."""
    dir_map = dict(existing_dir_map or {})
    created = 0
    existing = 0
    inc_kws = ("in", "incom", "incoming", "receiv", "received", "وارد")
    out_kws = ("out", "outgo", "outgoing", "sent", "issu", "issued", "صادر")

    with _get_subfolder_lock(dt_folder_id, "ltr_directions"):
        children = list_child_folders(service, dt_folder_id)
        inc_hit = next((c for c in children if any(k in _tokens(c.get("name")) for k in inc_kws)), None)
        out_hit = next((c for c in children if any(k in _tokens(c.get("name")) for k in out_kws)), None)
        non_dir_children = [
            c for c in children
            if not any(k in _tokens(c.get("name")) for k in inc_kws + out_kws)
        ]
        # Preserve existing company-first hierarchy (e.g. Style Design / SCAS / EKH) without adding root Sent/Received
        if non_dir_children and not (inc_hit or out_hit):
            return dir_map, 0, len(non_dir_children)

        if out_hit:
            dir_map["outgoing"] = out_hit["id"]
            existing += 1
        else:
            out_f = create_folder(service, dt_folder_id, "Sent")
            dir_map["outgoing"] = out_f["id"]
            created += 1

        if inc_hit:
            dir_map["incoming"] = inc_hit["id"]
            existing += 1
        else:
            inc_f = create_folder(service, dt_folder_id, "Received")
            dir_map["incoming"] = inc_f["id"]
            created += 1

    return dir_map, created, existing


def _get_master_root_id_from_store(store):
    if hasattr(store, "get_master_drive_root_id"):
        try:
            return extract_drive_folder_id(store.get_master_drive_root_id())
        except Exception:
            return ""
    return ""


def _ensure_workspace_drive_folder_detail(service, ws_id, ws_name=None, store=None):
    """Find or create `<Master Root>/<WS_ID> - <Workspace Name>`, returning `(folder_id, created, folder_name)`."""
    if store is None:
        import db as store
    clean_ws_id = str(ws_id or "").strip().upper()
    if not clean_ws_id:
        return "", False, ""
    master_root_id = _get_master_root_id_from_store(store)
    if not master_root_id:
        return "", False, ""
    svc = service if service is not None else get_default_write_service()
    if svc is None:
        return "", False, ""

    cached_id = ""
    if hasattr(store, "get_workspace_drive_folder_id"):
        try:
            cached_id = extract_drive_folder_id(store.get_workspace_drive_folder_id(clean_ws_id))
        except Exception:
            cached_id = ""
    if cached_id and folder_is_usable(svc, cached_id):
        return cached_id, False, ""

    resolved_name = ws_name
    if not resolved_name and hasattr(store, "get_workspace"):
        try:
            ws_row = store.get_workspace(clean_ws_id) or {}
            resolved_name = ws_row.get("name") or ""
        except Exception:
            resolved_name = ""
    target_name = format_workspace_folder_name(clean_ws_id, resolved_name)
    prefix_low = f"{clean_ws_id.lower()} -"

    with _get_subfolder_lock(master_root_id, clean_ws_id):
        children = list_child_folders(svc, master_root_id)
        match = next(
            (f for f in children if str(f.get("name", "")).strip().lower() == target_name.lower()),
            None,
        )
        if match is None:
            match = next(
                (
                    f for f in children
                    if str(f.get("name", "")).strip().lower() == clean_ws_id.lower()
                    or str(f.get("name", "")).strip().lower().startswith(prefix_low)
                ),
                None,
            )
        if match is not None:
            fid = match["id"]
            actual_name = match.get("name") or target_name
            created = False
        else:
            created_f = create_folder(svc, master_root_id, target_name)
            fid = created_f["id"]
            actual_name = created_f.get("name") or target_name
            created = True
        if hasattr(store, "set_workspace_drive_folder_id"):
            try:
                store.set_workspace_drive_folder_id(clean_ws_id, fid)
            except Exception:
                pass
        return fid, created, actual_name


def ensure_workspace_drive_folder(service, ws_id, ws_name=None, store=None):
    """Find or create `<Master Root>/<WS_ID> - <Workspace Name>` and cache its Drive folder ID."""
    fid, _, _ = _ensure_workspace_drive_folder_detail(service, ws_id, ws_name=ws_name, store=store)
    return fid


def rename_workspace_drive_folder(service, ws_id, new_ws_name, store=None):
    """Rename a cached or discovered `<Master Root>/<WS_ID> - <Name>` folder in-place via PATCH."""
    if store is None:
        import db as store
    clean_ws_id = str(ws_id or "").strip().upper()
    if not clean_ws_id or not sanitize_drive_folder_segment(new_ws_name):
        return {"ok": True, "skipped": True, "reason": "invalid_name", "renamed": False}

    svc = service if service is not None else get_default_write_service()
    if svc is None:
        return {"ok": True, "skipped": True, "reason": "drive_not_configured", "renamed": False}

    master_root_id = _get_master_root_id_from_store(store)
    folder_id = ""
    if hasattr(store, "get_workspace_drive_folder_id"):
        try:
            folder_id = extract_drive_folder_id(store.get_workspace_drive_folder_id(clean_ws_id))
        except Exception:
            folder_id = ""
    if folder_id and not folder_is_usable(svc, folder_id):
        folder_id = ""

    if not folder_id and master_root_id:
        prefix_low = f"{clean_ws_id.lower()} -"
        children = list_child_folders(svc, master_root_id)
        match = next(
            (
                f for f in children
                if str(f.get("name", "")).strip().lower() == clean_ws_id.lower()
                or str(f.get("name", "")).strip().lower().startswith(prefix_low)
            ),
            None,
        )
        if match is not None:
            folder_id = match["id"]

    if not folder_id:
        return {"ok": True, "skipped": False, "renamed": False, "reason": "folder_not_found"}

    new_folder_name = format_workspace_folder_name(clean_ws_id, new_ws_name)
    renamed_res = rename_folder(svc, folder_id, new_folder_name, parent_id=master_root_id or None)
    final_name = (renamed_res or {}).get("name") or new_folder_name
    if hasattr(store, "set_workspace_drive_folder_id"):
        try:
            store.set_workspace_drive_folder_id(clean_ws_id, folder_id)
        except Exception:
            pass
    return {"ok": True, "skipped": False, "renamed": True, "folder_id": folder_id, "folder_name": final_name}


def _ensure_project_drive_folder_detail(service, pid, ws_id=None, store=None):
    """Ensure a project has a Drive folder; returns `(folder_id, created, folder_name)`."""
    if store is None:
        import db as store
    try:
        project = store.get_project(pid) if hasattr(store, "get_project") else None
    except Exception:
        project = None
    if not isinstance(project, dict) or not project:
        return "", False, ""

    existing_fid = extract_drive_folder_id(project.get("drive_folder_id"))
    if existing_fid:
        return existing_fid, False, ""

    master_root_id = _get_master_root_id_from_store(store)
    if not master_root_id:
        return "", False, ""

    svc = service if service is not None else get_default_write_service()
    if svc is None:
        return "", False, ""

    eff_ws_id = str(ws_id or project.get("workspace_id") or "ENG").strip().upper()
    ws_folder_id = ensure_workspace_drive_folder(svc, eff_ws_id, store=store)
    if not ws_folder_id:
        return "", False, ""

    target_name = format_project_folder_name(project)
    alt_name = sanitize_drive_folder_segment(project.get("name") or "", max_len=85)
    with _get_subfolder_lock(ws_folder_id, target_name):
        children = list_child_folders(svc, ws_folder_id)
        match = next(
            (f for f in children if str(f.get("name", "")).strip().lower() == target_name.lower()),
            None,
        )
        if match is None and alt_name and alt_name.lower() != target_name.lower():
            match = next(
                (f for f in children if str(f.get("name", "")).strip().lower() == alt_name.lower()),
                None,
            )
        if match is not None:
            fid = match["id"]
            actual_name = match.get("name") or target_name
            created = False
        else:
            created_f = create_folder(svc, ws_folder_id, target_name)
            fid = created_f["id"]
            actual_name = created_f.get("name") or target_name
            created = True

        if hasattr(store, "update_project_settings"):
            store.update_project_settings(pid, {"drive_folder_id": fid, "auto_provisioned_drive": True})
        return fid, created, actual_name


def ensure_project_drive_folder(service, pid, ws_id=None, store=None):
    """Return existing `drive_folder_id` unchanged, or auto-provision `<Master Root>/<Workspace>/<Project>`."""
    fid, _, _ = _ensure_project_drive_folder_detail(service, pid, ws_id=ws_id, store=store)
    return fid


def rename_project_drive_folder(service, pid, new_project_name, store=None, new_project_code=None):
    """Rename an auto-provisioned project Drive folder in-place via PATCH."""
    if store is None:
        import db as store
    try:
        project = store.get_project(pid) if hasattr(store, "get_project") else None
    except Exception:
        project = None
    if not isinstance(project, dict) or not project:
        return {"ok": True, "skipped": True, "reason": "project_not_found", "renamed": False}
    if not project.get("auto_provisioned_drive"):
        return {"ok": True, "skipped": True, "reason": "not_auto_provisioned", "renamed": False}

    folder_id = extract_drive_folder_id(project.get("drive_folder_id"))
    if not folder_id:
        return {"ok": True, "skipped": True, "reason": "no_drive_folder", "renamed": False}

    svc = service if service is not None else get_default_write_service()
    if svc is None:
        return {"ok": True, "skipped": True, "reason": "drive_not_configured", "renamed": False}
    if not folder_is_usable(svc, folder_id):
        return {"ok": True, "skipped": False, "renamed": False, "reason": "folder_not_found"}

    eff_code = new_project_code if new_project_code is not None else project.get("code")
    new_folder_name = format_project_folder_name({"id": pid, "code": eff_code, "name": new_project_name})
    ws_id = str(project.get("workspace_id") or "ENG").strip().upper()
    ws_folder_id = ""
    if hasattr(store, "get_workspace_drive_folder_id"):
        try:
            ws_folder_id = extract_drive_folder_id(store.get_workspace_drive_folder_id(ws_id))
        except Exception:
            ws_folder_id = ""
    renamed_res = rename_folder(svc, folder_id, new_folder_name, parent_id=ws_folder_id or None)
    final_name = (renamed_res or {}).get("name") or new_folder_name
    return {"ok": True, "skipped": False, "renamed": True, "folder_id": folder_id, "folder_name": final_name}


def provision_project_drive_structure(project_id, only_doc_type_id=None, only_subtype=None, service=None, store=None):
    """Idempotently provision standardized Drive folders for `project_id`.

    If `drive_folder_id` is empty and a global Master Drive Root is configured, automatically
    provisions `<Master Root>/<Workspace>/<Project>` first. Safely no-ops when neither is
    configured or Drive credentials are unavailable.
    """
    import attachments
    from hse_defaults import normalize_ptw_subtype_value

    if store is None:
        import db as store

    try:
        project = store.get_project(project_id) if hasattr(store, "get_project") else None
    except Exception:
        project = None
    if not isinstance(project, dict) or not project:
        return {"ok": True, "skipped": True, "reason": "no_drive_folder", "created": 0, "existing": 0, "errors": []}

    root_id = extract_drive_folder_id(project.get("drive_folder_id"))
    svc = None
    if not root_id:
        master_root_id = _get_master_root_id_from_store(store)
        if not master_root_id:
            return {"ok": True, "skipped": True, "reason": "no_drive_folder", "created": 0, "existing": 0, "errors": []}
        svc = service if service is not None else get_default_write_service()
        if svc is None:
            return {"ok": True, "skipped": True, "reason": "drive_not_configured", "created": 0, "existing": 0, "errors": []}
        root_id = ensure_project_drive_folder(svc, project_id, ws_id=project.get("workspace_id"), store=store)
        if not root_id:
            return {"ok": True, "skipped": True, "reason": "no_drive_folder", "created": 0, "existing": 0, "errors": []}
        project = dict(project)
        project["drive_folder_id"] = root_id

    if svc is None:
        svc = service if service is not None else get_default_write_service()
    if svc is None:
        return {"ok": True, "skipped": True, "reason": "drive_not_configured", "created": 0, "existing": 0, "errors": []}

    ws_id = str(project.get("workspace_id") or "ENG").strip().upper()
    dts = []
    try:
        if only_doc_type_id:
            dt_one = store.get_doc_type(project_id, only_doc_type_id) if hasattr(store, "get_doc_type") else None
            if dt_one:
                dts = [dt_one]
        elif hasattr(store, "get_doc_types"):
            dts = list(store.get_doc_types(project_id) or [])
    except Exception as e:
        return {"ok": False, "skipped": False, "created": 0, "existing": 0, "errors": [str(e)]}

    lists_map = {}
    if hasattr(store, "get_lists"):
        try:
            lists_map = store.get_lists(project_id) or {}
        except Exception:
            lists_map = {}

    clean_only_sub = sanitize_drive_folder_segment(normalize_ptw_subtype_value(only_subtype)) if only_subtype else ""
    created_total = 0
    existing_total = 0
    errors = []

    for dt in dts:
        if not isinstance(dt, dict):
            continue
        dt_id = dt.get("id") or dt.get("code")
        code = str(dt.get("code") or dt_id or "").strip()
        name = str(dt.get("name") or code).strip()
        if not dt_id:
            continue

        try:
            is_noc = _is_noc_doc_type(dt)
            is_ltr = _is_ltr_doc_type(dt) and ws_id == "ENG"
            has_cycle = attachments.doc_type_has_reply_cycle(dt, workspace_id=ws_id)
            now_iso = datetime.datetime.now(datetime.timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")

            patch = {"folder_synced_at": now_iso}
            slot_cache = dict(dt.get("drive_slot_folders") or {})
            dir_cache = dict(dt.get("drive_direction_folders") or {})
            sub_cache = dict(dt.get("drive_subtype_folders") or {})

            # Special check for LTR with pre-existing root-level Incoming/Outgoing Letters folders
            if is_ltr and not clean_only_sub:
                root_children = list_child_folders(svc, root_id)
                try:
                    inc_root = match_doc_type_folder(root_children, code, name, direction="Received")
                    out_root = match_doc_type_folder(root_children, code, name, direction="Sent")
                except RoutingError:
                    inc_root, out_root = None, None
                if inc_root and out_root and inc_root["id"] != out_root["id"]:
                    dir_cache["incoming"] = inc_root["id"]
                    dir_cache["outgoing"] = out_root["id"]
                    patch["drive_direction_folders"] = dir_cache
                    existing_total += 2
                    if hasattr(store, "set_doc_type_drive_cache"):
                        store.set_doc_type_drive_cache(project_id, dt_id, patch)
                    continue

            # 1) Ensure <DocType Folder>
            dt_folder_id = str(dt.get("drive_folder_id") or "").strip()
            dt_folder_name = str(dt.get("drive_folder_name") or "").strip()
            if dt_folder_id and (_is_bad_rev_folder_name(dt_folder_name, code) or not folder_is_usable(svc, dt_folder_id)):
                dt_folder_id = ""
                slot_cache = {}
                sub_cache = {}

            target_dt_name = format_doc_type_folder_name(code, name)
            if dt_folder_id:
                existing_total += 1
            else:
                with _get_subfolder_lock(root_id, target_dt_name):
                    children = list_child_folders(svc, root_id)
                    matched_dt = match_doc_type_folder(children, code, name)
                    if matched_dt is not None:
                        dt_folder_id = matched_dt["id"]
                        dt_folder_name = matched_dt.get("name") or target_dt_name
                        existing_total += 1
                    else:
                        created_dt = create_folder(svc, root_id, target_dt_name)
                        dt_folder_id = created_dt["id"]
                        dt_folder_name = created_dt.get("name") or target_dt_name
                        created_total += 1
                patch["drive_folder_id"] = dt_folder_id
                patch["drive_folder_name"] = dt_folder_name

            # 2) Determine active subtypes
            if clean_only_sub:
                subtypes = [clean_only_sub]
            else:
                subtypes = _collect_active_subtypes_for_doc_type(project_id, ws_id, dt, store, lists_map=lists_map)

            slots_needed = ("part_a", "part_b", "part_c", "part_d") if is_noc else (("submission", "approval") if has_cycle else ())

            if subtypes:
                for sub_name in subtypes:
                    sub_key = sub_name.lower()
                    raw_entry = sub_cache.get(sub_key)
                    sub_entry = dict(raw_entry) if isinstance(raw_entry, dict) else {}
                    sub_id = str(sub_entry.get("id") or "").strip()
                    if sub_id and not folder_is_usable(svc, sub_id):
                        sub_id = ""
                        sub_entry = {}
                    if sub_id:
                        existing_total += 1
                    else:
                        with _get_subfolder_lock(dt_folder_id, sub_name):
                            sub_children = list_child_folders(svc, dt_folder_id)
                            sub_match = next(
                                (f for f in sub_children if str(f.get("name", "")).strip().lower() == sub_key),
                                None,
                            )
                            if sub_match is not None:
                                sub_id = sub_match["id"]
                                sub_actual_name = sub_match.get("name") or sub_name
                                existing_total += 1
                            else:
                                sub_created = create_folder(svc, dt_folder_id, sub_name)
                                sub_id = sub_created["id"]
                                sub_actual_name = sub_created.get("name") or sub_name
                                created_total += 1
                        sub_entry["id"] = sub_id
                        sub_entry["name"] = sub_actual_name

                    if slots_needed:
                        sub_slots, c_cnt, e_cnt = _ensure_slot_children(svc, sub_id, slots_needed, sub_entry.get("slots"))
                        created_total += c_cnt
                        existing_total += e_cnt
                        sub_entry["slots"] = sub_slots
                    else:
                        sub_entry.setdefault("slots", {})
                    sub_cache[sub_key] = sub_entry
                patch["drive_subtype_folders"] = sub_cache
            else:
                if is_ltr:
                    dir_cache, c_cnt, e_cnt = _ensure_ltr_direction_children(svc, dt_folder_id, dir_cache)
                    created_total += c_cnt
                    existing_total += e_cnt
                    if dir_cache:
                        patch["drive_direction_folders"] = dir_cache
                elif slots_needed:
                    slot_cache, c_cnt, e_cnt = _ensure_slot_children(svc, dt_folder_id, slots_needed, slot_cache)
                    created_total += c_cnt
                    existing_total += e_cnt
                    patch["drive_slot_folders"] = slot_cache

            if hasattr(store, "set_doc_type_drive_cache"):
                store.set_doc_type_drive_cache(project_id, dt_id, patch)
        except Exception as err:
            errors.append({"dt_id": str(dt_id), "error": str(err)})

    return {
        "ok": len(errors) == 0,
        "skipped": False,
        "project_id": project_id,
        "drive_folder_id": root_id,
        "created": created_total,
        "existing": existing_total,
        "errors": errors,
    }


_bg_provision_lock = threading.Lock()
_bg_provision_threads = []


def wait_for_background_provisioning(timeout=5.0):
    """Join any active background Drive provisioning threads (primarily for deterministic tests)."""
    with _bg_provision_lock:
        threads = [t for t in _bg_provision_threads if t.is_alive()]
    for t in threads:
        t.join(timeout=timeout)
    with _bg_provision_lock:
        _bg_provision_threads[:] = [t for t in _bg_provision_threads if t.is_alive()]


def provision_new_project_drive_tree(project_id, background=True, service=None, store=None):
    """Fast-path Drive provisioning for newly created projects.

    Synchronously ensures only `<Master Root>/<Workspace>/<Project>` (1-2 API calls)
    so `drive_folder_id` is persisted and returned immediately. When `background=True`,
    all DocType + Submit/Replied subfolder provisioning runs in a daemon thread with
    its own fresh Drive service instance (`get_default_write_service()`).
    """
    if store is None:
        import db as store

    try:
        project = store.get_project(project_id) if hasattr(store, "get_project") else None
    except Exception:
        project = None
    if not isinstance(project, dict) or not project:
        return {
            "ok": True,
            "skipped": True,
            "reason": "no_drive_folder",
            "drive_folder_id": "",
            "drive_provisioning_background": False,
            "created": 0,
            "existing": 0,
            "errors": [],
        }

    root_id = extract_drive_folder_id(project.get("drive_folder_id"))
    folder_created = False
    svc = service

    if not root_id:
        master_root_id = _get_master_root_id_from_store(store)
        if not master_root_id:
            return {
                "ok": True,
                "skipped": True,
                "reason": "no_drive_folder",
                "drive_folder_id": "",
                "drive_provisioning_background": False,
                "created": 0,
                "existing": 0,
                "errors": [],
            }
        if svc is None:
            svc = get_default_write_service()
        if svc is None:
            return {
                "ok": True,
                "skipped": True,
                "reason": "drive_not_configured",
                "drive_folder_id": "",
                "drive_provisioning_background": False,
                "created": 0,
                "existing": 0,
                "errors": [],
            }
        root_id, folder_created, _ = _ensure_project_drive_folder_detail(
            svc, project_id, ws_id=project.get("workspace_id"), store=store
        )
        if not root_id:
            return {
                "ok": True,
                "skipped": True,
                "reason": "no_drive_folder",
                "drive_folder_id": "",
                "drive_provisioning_background": False,
                "created": 0,
                "existing": 0,
                "errors": [],
            }

    if not background:
        res = provision_project_drive_structure(project_id, service=svc, store=store)
        res["drive_provisioning_background"] = False
        res.setdefault("drive_folder_id", root_id)
        return res

    # Thread-safety rule: never share the parent thread's Drive service across threads.
    # Build a fresh Drive service instance inside the background daemon thread.
    def _worker():
        try:
            bg_svc = get_default_write_service()
            if bg_svc is None:
                return
            provision_project_drive_structure(project_id, service=bg_svc, store=store)
        except Exception as exc:
            logger.exception("Background Drive subfolder provisioning failed for project %s: %s", project_id, exc)

    t = threading.Thread(
        target=_worker,
        name=f"drive-prov-{project_id}",
        daemon=True,
    )
    with _bg_provision_lock:
        _bg_provision_threads[:] = [th for th in _bg_provision_threads if th.is_alive()]
        _bg_provision_threads.append(t)
    t.start()

    return {
        "ok": True,
        "skipped": False,
        "project_id": project_id,
        "drive_folder_id": root_id,
        "drive_provisioning_background": True,
        "created": 1 if folder_created else 0,
        "existing": 0 if folder_created else 1,
        "errors": [],
    }


def provision_all_workspaces_and_projects(service=None, store=None):
    """Ensure Drive folders for all workspaces, all projects, and all internal doc-type trees."""
    if store is None:
        import db as store

    master_root_id = _get_master_root_id_from_store(store)
    if not master_root_id:
        return {
            "ok": False,
            "skipped": True,
            "reason": "no_master_root",
            "workspaces_created": 0,
            "workspaces_existing": 0,
            "projects_created": 0,
            "projects_existing": 0,
            "folders_created": 0,
            "folders_existing": 0,
            "errors": ["no_master_root"],
        }

    svc = service if service is not None else get_default_write_service()
    if svc is None:
        return {
            "ok": False,
            "skipped": True,
            "reason": "drive_not_configured",
            "workspaces_created": 0,
            "workspaces_existing": 0,
            "projects_created": 0,
            "projects_existing": 0,
            "folders_created": 0,
            "folders_existing": 0,
            "errors": ["drive_not_configured"],
        }

    workspaces_created = 0
    workspaces_existing = 0
    projects_created = 0
    projects_existing = 0
    folders_created = 0
    folders_existing = 0
    errors = []

    raw_workspaces = []
    if hasattr(store, "get_workspaces"):
        try:
            raw_workspaces = list(store.get_workspaces(active_only=False) or [])
        except TypeError:
            raw_workspaces = list(store.get_workspaces() or [])
        except Exception as err:
            errors.append({"scope": "workspaces", "error": str(err)})

    all_projects = []
    if hasattr(store, "get_projects"):
        try:
            all_projects = list(store.get_projects() or [])
        except Exception as err:
            errors.append({"scope": "projects", "error": str(err)})

    ws_map = {}
    for w in raw_workspaces:
        if not isinstance(w, dict):
            continue
        wid = str(w.get("id") or w.get("code") or "").strip().upper()
        if wid and wid not in ws_map:
            ws_map[wid] = w
    for p in all_projects:
        if not isinstance(p, dict):
            continue
        wid = str(p.get("workspace_id") or "ENG").strip().upper()
        if wid and wid not in ws_map:
            ws_map[wid] = {"id": wid, "name": wid}

    seen_pids = set()
    for wid, ws in ws_map.items():
        ws_name = ws.get("name") or wid
        try:
            ws_fid, ws_created, _ = _ensure_workspace_drive_folder_detail(svc, wid, ws_name=ws_name, store=store)
            if ws_fid:
                if ws_created:
                    workspaces_created += 1
                else:
                    workspaces_existing += 1
        except Exception as err:
            errors.append({"workspace_id": wid, "error": str(err)})

        ws_projects = [
            p for p in all_projects
            if isinstance(p, dict) and str(p.get("workspace_id") or "ENG").strip().upper() == wid
        ]
        for p in ws_projects:
            pid = str(p.get("id") or "").strip()
            if not pid or pid in seen_pids:
                continue
            seen_pids.add(pid)
            try:
                p_fid, p_created, _ = _ensure_project_drive_folder_detail(svc, pid, ws_id=wid, store=store)
                if p_fid:
                    if p_created:
                        projects_created += 1
                    else:
                        projects_existing += 1
                prov = provision_project_drive_structure(pid, service=svc, store=store)
                if not prov.get("skipped"):
                    folders_created += int(prov.get("created") or 0)
                    folders_existing += int(prov.get("existing") or 0)
                    for e in (prov.get("errors") or []):
                        errors.append({"project_id": pid, "error": e})
            except Exception as err:
                errors.append({"project_id": pid, "error": str(err)})

    return {
        "ok": len(errors) == 0,
        "skipped": False,
        "master_root_id": master_root_id,
        "workspaces_created": workspaces_created,
        "workspaces_existing": workspaces_existing,
        "projects_created": projects_created,
        "projects_existing": projects_existing,
        "folders_created": folders_created,
        "folders_existing": folders_existing,
        "errors": errors,
    }


def rename_doc_type_drive_folder(project_id, dt_id, old_code, old_name, new_code, new_name, service=None, store=None):
    """Rename the Drive folder for `(project_id, dt_id)` in-place via PATCH files/{id} without duplicates."""
    if store is None:
        import db as store
    try:
        project = store.get_project(project_id) if hasattr(store, "get_project") else None
    except Exception:
        project = None
    if not isinstance(project, dict) or not project:
        return {"ok": True, "skipped": True, "reason": "no_drive_folder", "renamed": False}

    root_id = extract_drive_folder_id(project.get("drive_folder_id"))
    if not root_id:
        return {"ok": True, "skipped": True, "reason": "no_drive_folder", "renamed": False}

    svc = service if service is not None else get_default_write_service()
    if svc is None:
        return {"ok": True, "skipped": True, "reason": "drive_not_configured", "renamed": False}

    dt = store.get_doc_type(project_id, dt_id) if hasattr(store, "get_doc_type") else {}
    dt = dt or {}
    new_folder_name = format_doc_type_folder_name(new_code, new_name)

    folder_id = str(dt.get("drive_folder_id") or "").strip()
    if folder_id and not folder_is_usable(svc, folder_id):
        folder_id = ""

    if not folder_id:
        try:
            children = list_child_folders(svc, root_id)
            matched = match_doc_type_folder(children, old_code, old_name)
            if matched is not None:
                folder_id = matched["id"]
        except RoutingError:
            folder_id = ""

    if not folder_id:
        return {"ok": True, "skipped": False, "renamed": False, "reason": "folder_not_found"}

    renamed_res = rename_folder(svc, folder_id, new_folder_name, parent_id=root_id)
    final_name = (renamed_res or {}).get("name") or new_folder_name
    now_iso = datetime.datetime.now(datetime.timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")
    if hasattr(store, "set_doc_type_drive_cache"):
        store.set_doc_type_drive_cache(project_id, dt_id, {
            "drive_folder_id": folder_id,
            "drive_folder_name": final_name,
            "folder_synced_at": now_iso,
        })
    return {"ok": True, "skipped": False, "renamed": True, "folder_id": folder_id, "folder_name": final_name}


def rename_subtype_drive_folder(project_id, dt_id, old_subtype, new_subtype, service=None, store=None):
    """Rename a `<Subtype>` subfolder under `<DocType Folder>` in-place via PATCH files/{id} without duplicates."""
    from hse_defaults import normalize_ptw_subtype_value

    if store is None:
        import db as store

    old_clean = sanitize_drive_folder_segment(normalize_ptw_subtype_value(old_subtype))
    new_clean = sanitize_drive_folder_segment(normalize_ptw_subtype_value(new_subtype))
    if not old_clean or not new_clean or old_clean == new_clean:
        return {"ok": True, "skipped": True, "reason": "unchanged", "renamed": False}

    try:
        project = store.get_project(project_id) if hasattr(store, "get_project") else None
    except Exception:
        project = None
    if not isinstance(project, dict) or not project:
        return {"ok": True, "skipped": True, "reason": "no_drive_folder", "renamed": False}

    root_id = extract_drive_folder_id(project.get("drive_folder_id"))
    if not root_id:
        return {"ok": True, "skipped": True, "reason": "no_drive_folder", "renamed": False}

    svc = service if service is not None else get_default_write_service()
    if svc is None:
        return {"ok": True, "skipped": True, "reason": "drive_not_configured", "renamed": False}

    dt = store.get_doc_type(project_id, dt_id) if hasattr(store, "get_doc_type") else None
    if not isinstance(dt, dict) or not dt:
        return {"ok": True, "skipped": True, "reason": "doc_type_not_found", "renamed": False}

    dt_folder_id = str(dt.get("drive_folder_id") or "").strip()
    if dt_folder_id and not folder_is_usable(svc, dt_folder_id):
        dt_folder_id = ""
    if not dt_folder_id:
        try:
            matched_dt = match_doc_type_folder(list_child_folders(svc, root_id), dt.get("code"), dt.get("name"))
            if matched_dt is not None:
                dt_folder_id = matched_dt["id"]
        except RoutingError:
            dt_folder_id = ""
    if not dt_folder_id:
        return {"ok": True, "skipped": False, "renamed": False, "reason": "doc_type_folder_not_found"}

    sub_cache = dict(dt.get("drive_subtype_folders") or {})
    slot_cache = dict(dt.get("drive_slot_folders") or {})
    old_key = old_clean.lower()
    new_key = new_clean.lower()

    old_entry = dict(sub_cache.get(old_key) or {}) if isinstance(sub_cache.get(old_key), dict) else {}
    sub_folder_id = str(old_entry.get("id") or slot_cache.get(f"ptw_subtype:{old_key}") or "").strip()
    if sub_folder_id and not folder_is_usable(svc, sub_folder_id):
        sub_folder_id = ""

    if not sub_folder_id:
        children = list_child_folders(svc, dt_folder_id)
        sub_match = next(
            (f for f in children if str(f.get("name", "")).strip().lower() == old_key),
            None,
        )
        if sub_match is not None:
            sub_folder_id = sub_match["id"]

    if not sub_folder_id:
        return {"ok": True, "skipped": False, "renamed": False, "reason": "subfolder_not_found"}

    renamed_res = rename_folder(svc, sub_folder_id, new_clean, parent_id=dt_folder_id)
    final_name = (renamed_res or {}).get("name") or new_clean

    if old_key != new_key:
        sub_cache.pop(old_key, None)
    sub_cache[new_key] = {
        "id": sub_folder_id,
        "name": final_name,
        "slots": dict(old_entry.get("slots") or {}),
    }
    now_iso = datetime.datetime.now(datetime.timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")
    patch = {
        "drive_folder_id": dt_folder_id,
        "drive_subtype_folders": sub_cache,
        "folder_synced_at": now_iso,
    }
    if f"ptw_subtype:{old_key}" in slot_cache:
        slot_cache.pop(f"ptw_subtype:{old_key}", None)
        patch["drive_slot_folders"] = slot_cache
    if hasattr(store, "set_doc_type_drive_cache"):
        store.set_doc_type_drive_cache(project_id, dt_id, patch)
    return {"ok": True, "skipped": False, "renamed": True, "folder_id": sub_folder_id, "folder_name": final_name}


def folder_has_any_files(service, folder_id, max_depth=5):
    """Return True if `folder_id` (or any subfolder up to `max_depth`) contains a non-folder file."""
    root_id = str(folder_id or "").strip()
    if not root_id:
        return False
    queue = [(root_id, 0)]
    visited = {root_id}
    while queue:
        cur_id, depth = queue.pop(0)
        q = "'%s' in parents and trashed = false" % cur_id
        token = None
        while True:
            res = service.files().list(
                q=q,
                fields="nextPageToken, files(id, name, mimeType)",
                pageSize=1000,
                pageToken=token,
                supportsAllDrives=True,
                includeItemsFromAllDrives=True,
            ).execute() or {}
            for f in (res.get("files") or []):
                if f.get("mimeType") != FOLDER_MIME:
                    return True
                cid = str(f.get("id") or "").strip()
                if cid and depth < max_depth and cid not in visited:
                    visited.add(cid)
                    queue.append((cid, depth + 1))
            token = res.get("nextPageToken")
            if not token:
                break
    return False


def cleanup_or_archive_drive_folder(service, folder_id, archive_parent_id=None):
    """Trash `folder_id` if its subtree has zero files; otherwise archive it into `_DCR_Deleted`."""
    clean_fid = extract_drive_folder_id(folder_id) or str(folder_id or "").strip()
    if not clean_fid:
        return {"ok": True, "skipped": True, "reason": "no_drive_folder"}
    if service is None:
        return {"ok": True, "skipped": True, "reason": "drive_not_configured", "folder_id": clean_fid}
    try:
        if not folder_is_usable(service, clean_fid):
            return {"ok": True, "skipped": True, "reason": "folder_not_found", "folder_id": clean_fid}
        has_files = folder_has_any_files(service, clean_fid)
        clean_archive_parent = extract_drive_folder_id(archive_parent_id) or str(archive_parent_id or "").strip()
        if has_files and clean_archive_parent:
            archived = archive_drive_item(service, clean_archive_parent, clean_fid)
            _FOLDER_USABLE_CACHE.pop(clean_fid, None)
            return {
                "ok": True,
                "action": "archived",
                "folder_id": clean_fid,
                "archive_folder_id": archived.get("archive_folder_id"),
                "name": archived.get("name"),
            }
        service.files().update(
            fileId=clean_fid,
            body={"trashed": True},
            fields="id, name, trashed",
            supportsAllDrives=True,
        ).execute()
        _FOLDER_USABLE_CACHE.pop(clean_fid, None)
        if clean_archive_parent:
            invalidate_subfolder_caches(clean_archive_parent)
        else:
            invalidate_subfolder_caches()
        return {"ok": True, "action": "trashed", "folder_id": clean_fid}
    except Exception as err:
        return {"ok": False, "action": "error", "folder_id": clean_fid, "error": str(err)}


def _resolve_doc_type_folder_for_cleanup(service, root_id, dt):
    """Resolve a DocType folder ID inside `root_id`, returning `(folder_id, skip_reason)`."""
    dt_dict = dt if isinstance(dt, dict) else {}
    dt_folder_id = str(dt_dict.get("drive_folder_id") or "").strip()
    if dt_folder_id and folder_is_usable(service, dt_folder_id):
        return dt_folder_id, None
    try:
        matched = match_doc_type_folder(
            list_child_folders(service, root_id),
            dt_dict.get("code"),
            dt_dict.get("name"),
        )
    except RoutingError:
        return "", "ambiguous"
    if matched is None:
        return "", "folder_not_found"
    return matched["id"], None


def cleanup_project_drive_folder(project, service=None, store=None):
    """Safely trash or archive an auto-provisioned project's Drive folder on deletion."""
    if not isinstance(project, dict) or not project:
        return {"ok": True, "skipped": True, "reason": "project_not_found"}
    if project.get("auto_provisioned_drive") is not True:
        return {"ok": True, "skipped": True, "reason": "not_auto_provisioned"}
    folder_id = extract_drive_folder_id(project.get("drive_folder_id"))
    if not folder_id:
        return {"ok": True, "skipped": True, "reason": "no_drive_folder"}
    svc = service if service is not None else get_default_write_service()
    if svc is None:
        return {"ok": True, "skipped": True, "reason": "drive_not_configured", "folder_id": folder_id}
    if store is None:
        import db as store
    master_root_id = _get_master_root_id_from_store(store)
    return cleanup_or_archive_drive_folder(svc, folder_id, archive_parent_id=master_root_id or None)


def cleanup_doc_type_drive_folder(project, dt, service=None):
    """Safely trash or archive a DocType folder inside `project.drive_folder_id` on tab deletion."""
    if not isinstance(project, dict) or not project:
        return {"ok": True, "skipped": True, "reason": "no_drive_folder"}
    root_id = extract_drive_folder_id(project.get("drive_folder_id"))
    if not root_id:
        return {"ok": True, "skipped": True, "reason": "no_drive_folder"}
    svc = service if service is not None else get_default_write_service()
    if svc is None:
        return {"ok": True, "skipped": True, "reason": "drive_not_configured"}
    dt_folder_id, skip_reason = _resolve_doc_type_folder_for_cleanup(svc, root_id, dt)
    if not dt_folder_id:
        return {"ok": True, "skipped": True, "reason": skip_reason or "folder_not_found"}
    return cleanup_or_archive_drive_folder(svc, dt_folder_id, archive_parent_id=root_id)


def cleanup_subtype_drive_folder(project, dt, subtype_name, service=None, store=None):
    """Safely trash or archive a Sub-Tab folder inside a DocType folder on sub-tab deletion."""
    from hse_defaults import normalize_ptw_subtype_value

    if store is None:
        import db as store
    clean_sub = sanitize_drive_folder_segment(normalize_ptw_subtype_value(subtype_name))
    if not clean_sub:
        return {"ok": True, "skipped": True, "reason": "invalid_subtype"}
    if not isinstance(project, dict) or not project:
        return {"ok": True, "skipped": True, "reason": "no_drive_folder"}
    root_id = extract_drive_folder_id(project.get("drive_folder_id"))
    if not root_id:
        return {"ok": True, "skipped": True, "reason": "no_drive_folder"}
    svc = service if service is not None else get_default_write_service()
    if svc is None:
        return {"ok": True, "skipped": True, "reason": "drive_not_configured"}
    if not isinstance(dt, dict) or not dt:
        return {"ok": True, "skipped": True, "reason": "doc_type_not_found"}

    dt_folder_id, skip_reason = _resolve_doc_type_folder_for_cleanup(svc, root_id, dt)
    if not dt_folder_id:
        return {"ok": True, "skipped": True, "reason": skip_reason or "doc_type_folder_not_found"}

    sub_cache = dict(dt.get("drive_subtype_folders") or {})
    slot_cache = dict(dt.get("drive_slot_folders") or {})
    sub_key = clean_sub.lower()

    sub_entry = dict(sub_cache.get(sub_key) or {}) if isinstance(sub_cache.get(sub_key), dict) else {}
    sub_folder_id = str(sub_entry.get("id") or slot_cache.get(f"ptw_subtype:{sub_key}") or "").strip()
    if sub_folder_id and not folder_is_usable(svc, sub_folder_id):
        sub_folder_id = ""

    if not sub_folder_id:
        children = list_child_folders(svc, dt_folder_id)
        matches = [f for f in children if str(f.get("name", "")).strip().lower() == sub_key]
        if len(matches) > 1:
            return {"ok": True, "skipped": True, "reason": "ambiguous"}
        if len(matches) == 1:
            sub_folder_id = matches[0]["id"]

    if not sub_folder_id:
        return {"ok": True, "skipped": True, "reason": "subfolder_not_found"}

    res = cleanup_or_archive_drive_folder(svc, sub_folder_id, archive_parent_id=root_id)
    if res.get("ok") and not res.get("skipped") and hasattr(store, "set_doc_type_drive_cache"):
        try:
            sub_cache.pop(sub_key, None)
            patch = {"drive_subtype_folders": sub_cache}
            if f"ptw_subtype:{sub_key}" in slot_cache:
                slot_cache.pop(f"ptw_subtype:{sub_key}", None)
                patch["drive_slot_folders"] = slot_cache
            dt_id = dt.get("id") or dt.get("code")
            if project.get("id") and dt_id:
                store.set_doc_type_drive_cache(project["id"], dt_id, patch)
        except Exception:
            pass
    return res
