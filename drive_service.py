import os
import re
from google.oauth2 import service_account
from googleapiclient.discovery import build
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
    Priority Logic (lower = Higher priority):
    For NOC 4-Part structure: Part D > Part C > Part B > Part A
    For Dual-slot structure: Reply Folder (1) > Reply File (2) > Submit Folder (3) > Submit File (4)
    """
    path_lower = full_path.lower()
    is_folder = (mime_type == 'application/vnd.google-apps.folder')

    m = re.search(r'\bpart[\s_-]*([abcd])\b', path_lower)
    if m:
        part_letter = m.group(1)
        part_rank = {'d': 1, 'c': 3, 'b': 5, 'a': 7}.get(part_letter, 7)
        return part_rank if is_folder else (part_rank + 1)

    if 'reply' in path_lower:
        return 1 if is_folder else 2
    elif 'submit' in path_lower:
        return 3 if is_folder else 4

    # If neither, treat as lowest priority
    return 9

def process_drive_folder(folder_id, root_folder_name="Root"):
    """
    Scans a Google Drive folder recursively using Batched BFS to avoid rate limits,
    builds the full path for each item, applies Priority Logic, and updates the DB.
    """
    service = get_drive_service()
    if not service: return
    
    # BFS Queue to hold folders we need to scan: (folder_id, folder_full_path)
    queue = [(folder_id, root_folder_name)]
    doc_map = {}
    slot_map = {}  # doc_no -> {'submission': {...}, 'approval': {...}}
    
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
                pageToken=page_token
            ).execute()
            
            items = results.get('files', [])
            
            for item in items:
                # Identify which folder in our batch is the parent of this item
                parent_id = next((p for p in item.get('parents', []) if p in path_map), None)
                if not parent_id: continue
                
                # Build the Full Path
                item_full_path = f"{path_map[parent_id]} / {item['name']}"
                
                is_folder = (item['mimeType'] == 'application/vnd.google-apps.folder')
                if is_folder:
                    queue.append((item['id'], item_full_path))
                
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

                    # Additive: also track the best Submit and best Replied link separately
                    kind = classify_path_kind(item_full_path)
                    if kind:
                        slot = slot_map.setdefault(doc_no, {})
                        if kind not in slot or priority < slot[kind]['priority']:
                            slot[kind] = {'link': item['webViewLink'], 'priority': priority}
            
            page_token = results.get('nextPageToken')
            if not page_token:
                break
                
    # Update Database with the strongest links found
    if doc_map:
        import db
        for doc_no, data in doc_map.items():
            sync_links = {k: v['link'] for k, v in slot_map.get(doc_no, {}).items()}
            db.update_record_link(doc_no, data['link'], sync_links)

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
