import os
import re
from attachments import classify_path_kind

# ==========================================
# 1. Configuration & Setup
# ==========================================
SCOPES = ['https://www.googleapis.com/auth/drive.readonly']
SERVICE_ACCOUNT_FILE = 'service_account.json'  # Will be placed in D:\DCR\

def get_drive_service():
    """Authenticate and return the Google Drive API service.
    First checks the GOOGLE_CREDENTIALS_JSON environment variable (for Production).
    Falls back to local service_account.json file (for Local Development).
    """
    import json
    from google.oauth2 import service_account
    from googleapiclient.discovery import build
    env_creds = os.environ.get('GOOGLE_CREDENTIALS_JSON')
    
    if env_creds:
        try:
            creds_dict = json.loads(env_creds)
            creds = service_account.Credentials.from_service_account_info(
                creds_dict, scopes=SCOPES)
            return build('drive', 'v3', credentials=creds)
        except Exception as e:
            print(f"Error parsing GOOGLE_CREDENTIALS_JSON: {e}")
            return None
            
    # Fallback to local JSON file
    if not os.path.exists(SERVICE_ACCOUNT_FILE):
        print(f"Warning: {SERVICE_ACCOUNT_FILE} not found and no Env Var set. Drive Integration disabled.")
        return None
        
    creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    return build('drive', 'v3', credentials=creds)

# ==========================================
# 2. Sync Logic & Priority Algorithm
# ==========================================
def extract_doc_no(name):
    """
    Extracts the document number from a file/folder name.
    Supports robust dynamic segments (e.g., MAT-EASTMAIN-B08-FF-001 REV00).
    Requires at least one dash and ending with digits (optionally a letter).
    """
    # 1. Match WITH Revision (e.g., ...-001 REV00)
    m = re.search(r'((?:[A-Z0-9]+-)+\d+[A-Z]?\s+REV\d+)', name, re.IGNORECASE)
    if m: 
        return m.group(1).upper()
    
    # 2. Match WITHOUT Revision (e.g., ...-001)
    m = re.search(r'((?:[A-Z0-9]+-)+\d+[A-Z]?)', name, re.IGNORECASE)
    return m.group(1).upper() if m else None

def get_item_priority(full_path, mime_type):
    """
    Priority Logic (lower = Higher priority), inspecting directory segments ONLY (excluding basename):
    For NOC 4-Part structure: Part D > Part C > Part B > Part A
    For Dual-slot structure: Reply Folder (1) > Reply File (2) > Submit Folder (3) > Submit File (4)
    """
    is_folder = (mime_type == 'application/vnd.google-apps.folder')
    kind = classify_path_kind(full_path)

    if kind in ("part_a", "part_b", "part_c", "part_d"):
        part_letter = kind.split("_")[-1]
        part_rank = {'d': 1, 'c': 3, 'b': 5, 'a': 7}.get(part_letter, 7)
        return part_rank if is_folder else (part_rank + 1)

    if kind == 'approval':
        return 1 if is_folder else 2
    if kind == 'submission':
        return 3 if is_folder else 4

    # If neither, treat as lowest priority
    return 9

def is_drive_item_in_replied_folder(service, item_id, folder_cache=None, max_depth=6):
    """Inspect a Drive file/folder's ancestor chain to verify whether it resides inside a `Replied` folder.

    Returns `(in_replied_bool, item_meta_or_None)`. Never raises on Drive API errors.
    """
    clean_id = str(item_id or "").strip()
    if not service or not clean_id:
        return False, None
    cache = folder_cache if isinstance(folder_cache, dict) else {}

    def _get_meta(fid):
        if fid in cache:
            return cache[fid]
        try:
            meta = service.files().get(
                fileId=fid,
                fields="id, name, mimeType, parents, webViewLink, trashed",
                supportsAllDrives=True,
            ).execute()
        except Exception:
            meta = None
        cache[fid] = meta
        return meta

    item_meta = _get_meta(clean_id)
    if not isinstance(item_meta, dict) or item_meta.get("trashed"):
        return False, item_meta

    parents = [str(p).strip() for p in (item_meta.get("parents") or []) if str(p).strip()]
    if not parents:
        return False, item_meta

    for start_parent in parents:
        chain_names = []
        visited = {clean_id}
        cur_id = start_parent
        depth = 0
        while cur_id and cur_id not in visited and depth < max_depth:
            visited.add(cur_id)
            p_meta = _get_meta(cur_id)
            if not isinstance(p_meta, dict) or p_meta.get("trashed"):
                break
            pname = str(p_meta.get("name") or "").strip()
            if pname:
                chain_names.append(pname)
            next_parents = [str(p).strip() for p in (p_meta.get("parents") or []) if str(p).strip()]
            cur_id = next_parents[0] if next_parents else ""
            depth += 1

        if chain_names:
            # chain_names is leaf-to-root; reverse to root-to-leaf with trailing '/' so all segments are dirs
            dir_path = " / ".join(reversed(chain_names)) + " /"
            kind = classify_path_kind(dir_path)
            if kind == "approval":
                return True, item_meta
            if kind == "submission":
                return False, item_meta

    return False, item_meta


def reclassify_project_legacy_reply_links(
    project_id,
    service=None,
    replied_items_by_id=None,
    submit_item_ids=None,
    inspect_parents=True,
):
    """Promote legacy/un-typed `fileLocation`/`drive_link`/`submission` links in `project_id`
    to `attachments['approval']` (`role='reply'`) when verified to reside inside a `Replied` folder.

    Returns summary dict: `{'scanned', 'reclassified', 'skipped', 'errors', 'record_ids'}`.
    """
    import attachments as _att
    import db

    pid = str(project_id or "").strip()
    summary = {"scanned": 0, "reclassified": 0, "skipped": 0, "errors": 0, "record_ids": []}
    if not pid:
        return summary

    if service is None and inspect_parents:
        service = get_drive_service()

    replied_map = replied_items_by_id if isinstance(replied_items_by_id, dict) else {}
    submit_ids = set(submit_item_ids or ())
    folder_cache = {}

    rows = db.get_project_records_for_drive_reclassify(pid) or []
    for row in rows:
        rec_id = row.get("id")
        rec_data = row.get("data")
        if not rec_id or not isinstance(rec_data, dict):
            continue
        summary["scanned"] += 1

        att = rec_data.get("attachments") if isinstance(rec_data.get("attachments"), dict) else {}
        if any(isinstance(att.get(ns), dict) for ns in _att.NOC_SLOTS):
            summary["skipped"] += 1
            continue

        # Collect candidate un-typed/legacy/submission link targets on this record
        candidates = []
        sub_entry = att.get(_att.SLOT_SUBMISSION) if isinstance(att.get(_att.SLOT_SUBMISSION), dict) else None
        if sub_entry:
            sub_url = _att._clean_link(sub_entry.get("url"))
            sub_id = str(
                sub_entry.get("file_id")
                or sub_entry.get("folder_id")
                or _att.extract_drive_item_id(sub_url)
                or ""
            ).strip()
            if sub_id:
                candidates.append((sub_id, sub_url))

        for key in _att.LEGACY_LINK_KEYS:
            leg_url = _att._clean_link(rec_data.get(key))
            if not leg_url:
                continue
            leg_id = _att.extract_drive_item_id(leg_url)
            if leg_id and not any(cid == leg_id for cid, _ in candidates):
                candidates.append((leg_id, leg_url))

        if not candidates:
            summary["skipped"] += 1
            continue

        promoted_any = False
        cur_data = rec_data
        for item_id, item_url in candidates:
            # Safety: if this item also exists in a Submit folder, never reclassify it to reply
            if item_id in submit_ids:
                continue

            matched_item = None
            if item_id in replied_map:
                matched_item = replied_map[item_id]
            elif inspect_parents and service is not None:
                try:
                    in_replied, item_meta = is_drive_item_in_replied_folder(
                        service, item_id, folder_cache=folder_cache
                    )
                    if in_replied:
                        matched_item = item_meta or {"id": item_id}
                except Exception:
                    summary["errors"] += 1
                    continue

            if matched_item is None:
                continue

            new_data, changed = _att.promote_legacy_link_to_reply(
                cur_data,
                matched_item=matched_item,
                reply_url=item_url,
            )
            if changed:
                cur_data = new_data
                promoted_any = True

        if promoted_any:
            db.update_record_data_scoped(pid, rec_id, cur_data)
            summary["reclassified"] += 1
            summary["record_ids"].append(rec_id)
        else:
            summary["skipped"] += 1

    return summary


def process_drive_folder(folder_id, root_folder_name="Root", project_id=None, reclassify_legacy_links=True):
    """
    Scans a Google Drive folder recursively using Batched BFS to avoid rate limits,
    builds the full path for each item, applies Priority Logic, and updates the DB.
    """
    service = get_drive_service()
    if not service: return
    
    # BFS Queue to hold folders we need to scan: (folder_id, folder_full_path)
    queue = [(folder_id, root_folder_name)]
    doc_map = {}
    slot_map = {}      # doc_no -> {'submission': {...}, 'approval': {...}}
    part_rev_map = {}  # doc_no -> {slot: {rev_int: {'url': ..., 'priority': ..., 'rev': rev_int, 'rev_label': ...}}}
    replied_items_by_id = {}  # item_id -> item metadata for all files/folders inside Replied/
    submit_item_ids = set()   # item_ids for all files/folders inside Submit/
    
    while queue:
        # Batch up to 15 folders per API query to dramatically reduce API calls
        batch = queue[:15]
        queue = queue[15:]
        
        # Build compound query: ('id1' in parents or 'id2' in parents)
        parents_query = " or ".join([f"'{f_id}' in parents" for f_id, _ in batch])
        query = f"({parents_query}) and trashed = false"
        
        # Dictionary to quickly look up the parent's path
        path_map = {f_id: path for f_id, path in batch}
        
        page_token = None
        while True:
            results = service.files().list(
                q=query,
                fields="nextPageToken, files(id, name, mimeType, parents, webViewLink)",
                pageSize=1000,
                pageToken=page_token,
                supportsAllDrives=True,
                includeItemsFromAllDrives=True,
            ).execute()
            
            items = results.get('files', [])
            
            for item in items:
                # Identify which folder in our batch is the parent of this item
                parent_id = next((p for p in item.get('parents', []) if p in path_map), None)
                if not parent_id: continue

                # Skip _DCR_Deleted archive folders/items before traversal or doc_no extraction
                if str(item.get('name') or '').strip().lower() == '_dcr_deleted':
                    continue
                
                # Build the Full Path
                item_full_path = f"{path_map[parent_id]} / {item['name']}"
                
                is_folder = (item['mimeType'] == 'application/vnd.google-apps.folder')
                if is_folder:
                    queue.append((item['id'], item_full_path))

                # Track every item discovered inside Replied vs Submit folders (even if filename has no docNo)
                kind = classify_path_kind(item_full_path)
                if kind == 'approval' and item.get('id'):
                    replied_items_by_id[item['id']] = {
                        'id': item['id'],
                        'name': item.get('name') or '',
                        'mimeType': item.get('mimeType') or '',
                        'webViewLink': item.get('webViewLink') or '',
                    }
                elif kind == 'submission' and item.get('id'):
                    submit_item_ids.add(item['id'])
                
                # Extract Document Number
                doc_no = extract_doc_no(item['name'])
                if doc_no:
                    # Apply Priority Logic using the FULL PATH
                    priority = get_item_priority(item_full_path, item['mimeType'])
                    
                    if doc_no not in doc_map or priority < doc_map[doc_no]['priority']:
                        doc_map[doc_no] = {
                            'link': item['webViewLink'],
                            'priority': priority
                        }

                    # Additive: also track the best Submit/Replied and per-part revisions separately
                    if kind:
                        if kind in ("part_a", "part_b", "part_c", "part_d"):
                            from attachments import format_rev_label
                            rev_matches = re.findall(r"\b(?:rev|r)\s*[-_.]?\s*(\d+)\b", item_full_path, re.IGNORECASE)
                            rev_num = max((int(m) for m in rev_matches), default=0)
                            by_slot = part_rev_map.setdefault(doc_no, {}).setdefault(kind, {})
                            prev = by_slot.get(rev_num)
                            if prev is None or priority < prev['priority']:
                                by_slot[rev_num] = {
                                    'url': item['webViewLink'],
                                    'priority': priority,
                                    'rev': rev_num,
                                    'rev_label': format_rev_label(rev_num),
                                }
                        else:
                            slot = slot_map.setdefault(doc_no, {})
                            if kind not in slot or priority < slot[kind]['priority']:
                                slot[kind] = {'link': item['webViewLink'], 'priority': priority}
            
            page_token = results.get('nextPageToken')
            if not page_token:
                break
                
    # Update Database with the strongest links found, strictly scoped to the folder's project(s)
    import db
    if project_id and str(project_id).strip():
        target_pids = [str(project_id).strip()]
    else:
        target_pids = db.get_project_ids_for_drive_folder(folder_id)

    if doc_map:
        for pid in target_pids:
            for doc_no, data in doc_map.items():
                sync_links = {k: v['link'] for k, v in slot_map.get(doc_no, {}).items()}
                for part_slot, rev_dict in part_rev_map.get(doc_no, {}).items():
                    ordered_revs = sorted(rev_dict.values(), key=lambda x: (-x['rev'], x['priority']))
                    if ordered_revs:
                        top_entry = dict(ordered_revs[0])
                        top_entry.pop('priority', None)
                        if len(ordered_revs) > 1:
                            top_entry['history'] = [
                                {k: v for k, v in older.items() if k != 'priority'}
                                for older in ordered_revs[1:]
                            ]
                        sync_links[part_slot] = top_entry
                db.update_record_link(doc_no, data['link'], sync_links, project_id=pid)

    if reclassify_legacy_links and target_pids:
        for pid in target_pids:
            try:
                reclassify_project_legacy_reply_links(
                    pid,
                    service=service,
                    replied_items_by_id=replied_items_by_id,
                    submit_item_ids=submit_item_ids,
                    inspect_parents=True,
                )
            except Exception as exc:
                print(f"[Drive Legacy Reclassify Warning] project={pid}: {exc}")

# ==========================================
# 3. Webhook Registration (Run once per folder)
# ==========================================
def register_webhook(folder_id, webhook_url, channel_id):
    """
    Registers a Push Notification Webhook for a specific Drive Folder.
    """
    service = get_drive_service()
    if not service: return
    
    body = {
        "id": channel_id, # A unique UUID for this channel
        "type": "web_hook",
        "address": webhook_url
    }
    return service.files().watch(fileId=folder_id, body=body).execute()


# ==========================================
# 4. Write-capable service (uploads / folder routing) - separate from the read-only sync
# ==========================================
WRITE_SCOPES = ['https://www.googleapis.com/auth/drive']

def get_drive_write_service():
    """Return a Drive API service with WRITE scope, or None if credentials are unavailable.

    Uses the same credential sources as get_drive_service() but a separate scope, so the
    existing read-only sync is unaffected. Callers must pass supportsAllDrives=True on
    Drive calls (done in drive_routing.py).
    """
    import json
    from google.oauth2 import service_account
    from googleapiclient.discovery import build
    env_creds = os.environ.get('GOOGLE_CREDENTIALS_JSON')
    try:
        if env_creds:
            creds = service_account.Credentials.from_service_account_info(
                json.loads(env_creds), scopes=WRITE_SCOPES)
        elif os.path.exists(SERVICE_ACCOUNT_FILE):
            creds = service_account.Credentials.from_service_account_file(
                SERVICE_ACCOUNT_FILE, scopes=WRITE_SCOPES)
        else:
            print("Warning: no Drive credentials found. Drive write service disabled.")
            return None
        return build('drive', 'v3', credentials=creds, cache_discovery=False)
    except Exception as e:
        print(f"Error creating Drive write service: {e}")
        return None
