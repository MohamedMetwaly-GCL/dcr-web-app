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
        chosen_comps = children

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
    f_tokens, f_norm = _tokens(folder_name), _norm(folder_name)
    c_norm, n_norm = _norm(code), _norm(name)
    c_tokens = _tokens(code)
    if not f_norm:
        return 0
    if f_norm in (c_norm, n_norm):
        score = SCORE_EXACT
    elif c_tokens and f_tokens[:len(c_tokens)] == c_tokens:
        score = SCORE_PREFIX
    elif c_tokens and any(f_tokens[i:i + len(c_tokens)] == c_tokens
                        for i in range(len(f_tokens) - len(c_tokens) + 1)):
        score = SCORE_TOKEN
    elif n_norm and (" " + n_norm + " ") in (" " + f_norm + " "):
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

    return score


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
import time
_FOLDER_USABLE_CACHE = {}   # folder_id -> (timestamp, bool)
_CHILD_FOLDERS_CACHE = {}   # parent_id -> (timestamp, list)


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
    _CHILD_FOLDERS_CACHE.pop(parent_id, None)
    body = {"name": name, "mimeType": FOLDER_MIME, "parents": [parent_id]}
    res = service.files().create(body=body, fields="id, name", supportsAllDrives=True).execute()
    if res and res.get("id"):
        _FOLDER_USABLE_CACHE[res["id"]] = (time.time(), True)
    return res


# ---------------- Orchestration ----------------
def resolve_upload_destination(pid, dt_id, slot, service, store, create_if_missing=False,
                               subfolder_name=None, flat=False, direction="", destination_folder_id="",
                               party="", from_party="", to_party="", doc_no=""):
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
    if root_id:
        m = re.search(r"/folders/([a-zA-Z0-9_-]+)", root_id)
        if m:
            root_id = m.group(1)
    if not root_id:
        raise RoutingError("project_has_no_drive_folder")
    dt = store.get_doc_type(pid, dt_id)
    if not dt:
        raise RoutingError("doc_type_not_found")

    source = "cache"
    patch = {}
    slot_cache = dict(dt.get("drive_slot_folders") or {})
    dir_cache = dict(dt.get("drive_direction_folders") or {})
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
            now_iso = datetime.datetime.utcnow().replace(microsecond=0).isoformat() + "Z"
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
        if cand_id and not (opposite_inc or opposite_out) and folder_is_usable(service, cand_id):
            dt_folder_id = cand_id
            source = "cache"

    if not dt_folder_id:
        source = "discovered"
        folder = match_doc_type_folder(list_child_folders(service, root_id), dt.get("code"), dt.get("name"), direction=direction)
        if folder is None:
            if not create_if_missing:
                raise RoutingError("doc_type_folder_not_found")
            folder = create_folder(service, root_id, "%s - %s" % (dt.get("code"), dt.get("name")))
            source = "created"
        dt_folder_id = folder["id"]
        slot_cache = {}            # slot folders belong to the old doc type folder
        now_iso = datetime.datetime.utcnow().replace(microsecond=0).isoformat() + "Z"
        if dir_key:
            dir_cache[dir_key] = dt_folder_id
            patch.update({"drive_direction_folders": dir_cache,
                          "folder_synced_at": now_iso})
        else:
            patch.update({"drive_folder_id": dt_folder_id, "drive_folder_name": folder.get("name", ""),
                          "drive_slot_folders": {},
                          "folder_synced_at": now_iso})

    # 2) Submit / Replied sub-folder (skipped when flat=True: single-party documents such as PR / Letters)
    slot_id = str(slot_cache.get(slot) or "").strip()
    if destination_folder_id and folder_is_usable(service, destination_folder_id):
        slot_id = destination_folder_id
    elif flat:
        slot_id = resolve_hierarchical_subfolder(
            service, dt_folder_id, project,
            direction=direction, party=party,
            from_party=from_party, to_party=to_party, doc_no=doc_no
        )
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
