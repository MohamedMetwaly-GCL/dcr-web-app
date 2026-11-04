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

FOLDER_MIME = "application/vnd.google-apps.folder"

# Drive folder names that mean each slot (same keywords drive_service.get_item_priority uses)
SLOT_KEYWORDS = {"submission": ("submit",), "approval": ("reply", "replied")}
SLOT_DEFAULT_NAMES = {"submission": "Submit", "approval": "Replied"}

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


def score_folder_name(folder_name, code, name):
    """Score how well a Drive folder name matches a doc type (code + display name)."""
    f_tokens, f_norm = _tokens(folder_name), _norm(folder_name)
    c_norm, n_norm = _norm(code), _norm(name)
    if not f_norm:
        return 0
    if f_norm in (c_norm, n_norm):
        return SCORE_EXACT
    # whole-token match only, so code 'IR' never matches 'MIR' or 'SIRIUS'
    c_tokens = _tokens(code)
    if c_tokens and f_tokens[:len(c_tokens)] == c_tokens:
        return SCORE_PREFIX
    if c_tokens and any(f_tokens[i:i + len(c_tokens)] == c_tokens
                        for i in range(len(f_tokens) - len(c_tokens) + 1)):
        return SCORE_TOKEN
    if n_norm and (" " + n_norm + " ") in (" " + f_norm + " "):
        return SCORE_NAME_IN
    return 0


def match_doc_type_folder(folders, code, name):
    """Pick the folder for a doc type. Returns folder dict, None (no match).
    Raises RoutingError('ambiguous') if the top score is shared by several folders."""
    scored = sorted(((score_folder_name(f.get("name"), code, name), f) for f in folders),
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
    """Find the Submit / Replied sub-folder inside a doc type folder."""
    keys = SLOT_KEYWORDS[slot]
    hits = [f for f in folders if any(k in str(f.get("name", "")).lower() for k in keys)]
    # 'Replied' must not be taken for the Submit slot and vice versa
    if slot == "submission":
        hits = [f for f in hits if "repl" not in f["name"].lower()]
    if len(hits) > 1:
        raise RoutingError("ambiguous", [{"id": f["id"], "name": f["name"]} for f in hits])
    return hits[0] if hits else None


# ---------------- Drive helpers (service is injected) ----------------
def list_child_folders(service, parent_id):
    q = ("'%s' in parents and mimeType = '%s' and trashed = false" % (parent_id, FOLDER_MIME))
    out, token = [], None
    while True:
        res = service.files().list(
            q=q, fields="nextPageToken, files(id, name)", pageSize=1000, pageToken=token,
            supportsAllDrives=True, includeItemsFromAllDrives=True).execute()
        out.extend(res.get("files", []))
        token = res.get("nextPageToken")
        if not token:
            return out


def folder_is_usable(service, folder_id):
    try:
        f = service.files().get(fileId=folder_id, fields="id, mimeType, trashed",
                                supportsAllDrives=True).execute()
    except Exception:
        return False
    return f.get("mimeType") == FOLDER_MIME and not f.get("trashed", False)


def create_folder(service, parent_id, name):
    body = {"name": name, "mimeType": FOLDER_MIME, "parents": [parent_id]}
    return service.files().create(body=body, fields="id, name", supportsAllDrives=True).execute()


# ---------------- Orchestration ----------------
def resolve_upload_destination(pid, dt_id, slot, service, store, create_if_missing=False, subfolder_name=None, flat=False):
    """Return {'folder_id', 'doc_type_folder_id', 'source', ...} or raise RoutingError.

    store must provide: get_project(pid), get_doc_type(pid, dt_id),
    set_doc_type_drive_cache(pid, dt_id, patch_dict).
    If subfolder_name is provided, creates/finds a subfolder with that name (e.g. Doc No)
    inside the destination folder and returns its folder_id.
    """
    if slot not in SLOT_KEYWORDS:
        raise RoutingError("invalid_slot")
    project = store.get_project(pid) or {}
    root_id = str(project.get("drive_folder_id") or "").strip()
    if not root_id:
        raise RoutingError("project_has_no_drive_folder")
    dt = store.get_doc_type(pid, dt_id)
    if not dt:
        raise RoutingError("doc_type_not_found")

    source = "cache"
    # 1) doc type folder
    dt_folder_id = str(dt.get("drive_folder_id") or "").strip()
    if dt_folder_id and not folder_is_usable(service, dt_folder_id):
        dt_folder_id = ""          # deleted / trashed on Drive -> rediscover
    patch = {}
    slot_cache = dict(dt.get("drive_slot_folders") or {})
    if not dt_folder_id:
        source = "discovered"
        folder = match_doc_type_folder(list_child_folders(service, root_id), dt.get("code"), dt.get("name"))
        if folder is None:
            if not create_if_missing:
                raise RoutingError("doc_type_folder_not_found")
            folder = create_folder(service, root_id, "%s - %s" % (dt.get("code"), dt.get("name")))
            source = "created"
        dt_folder_id = folder["id"]
        slot_cache = {}            # slot folders belong to the old doc type folder
        patch.update({"drive_folder_id": dt_folder_id, "drive_folder_name": folder.get("name", ""),
                      "drive_slot_folders": {},
                      "folder_synced_at": datetime.datetime.utcnow().replace(microsecond=0).isoformat() + "Z"})

    # 2) Submit / Replied sub-folder (skipped when flat=True: single-party documents such as PR / Letters)
    slot_id = str(slot_cache.get(slot) or "").strip()
    if flat:
        slot_id = dt_folder_id
    elif slot_id and not folder_is_usable(service, slot_id):
        slot_id = ""
    if not slot_id:
        sub = match_slot_folder(list_child_folders(service, dt_folder_id), slot)
        if sub is None:
            if not create_if_missing:
                raise RoutingError("slot_folder_not_found")
            sub = create_folder(service, dt_folder_id, SLOT_DEFAULT_NAMES[slot])
            source = "created"
        slot_id = sub["id"]
        slot_cache[slot] = slot_id
        patch["drive_slot_folders"] = slot_cache
        if source == "cache":
            source = "discovered"

    if patch:
        store.set_doc_type_drive_cache(pid, dt_id, patch)

    # 3) Optional subfolder per document (e.g. DS-CY002P608-00001 REV00)
    target_folder_id = slot_id
    subfolder_id = None
    if subfolder_name:
        clean_sub = str(subfolder_name).strip()
        if clean_sub:
            children = list_child_folders(service, slot_id)
            match = next((f for f in children if str(f.get("name", "")).strip().lower() == clean_sub.lower()), None)
            if match is None:
                if not create_if_missing:
                    raise RoutingError("subfolder_not_found")
                match = create_folder(service, slot_id, clean_sub)
                source = "created"
            subfolder_id = match["id"]
            target_folder_id = subfolder_id

    result = {
        "folder_id": target_folder_id,
        "doc_type_folder_id": dt_folder_id,
        "slot_folder_id": slot_id,
        "source": source
    }
    if subfolder_id:
        result["subfolder_id"] = subfolder_id
        result["subfolder_name"] = str(subfolder_name).strip()
    return result
