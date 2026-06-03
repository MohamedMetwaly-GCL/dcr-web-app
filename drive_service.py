import os
import re
from google.oauth2 import service_account
from googleapiclient.discovery import build

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
    Priority Logic (1 = Highest, 4 = Lowest):
    1: Reply Folder (path contains 'reply', is folder)
    2: Reply File   (path contains 'reply', is file)
    3: Submit Folder (path contains 'submit', is folder)
    4: Submit File   (path contains 'submit', is file)
    """
    path_lower = full_path.lower()
    is_folder = (mime_type == 'application/vnd.google-apps.folder')
    
    if 'reply' in path_lower:
        return 1 if is_folder else 2
    elif 'submit' in path_lower:
        return 3 if is_folder else 4
    
    # If neither, treat as lowest priority
    return 5

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
            
            page_token = results.get('nextPageToken')
            if not page_token:
                break
                
    # Update Database with the strongest links found
    if doc_map:
        import db
        for doc_no, data in doc_map.items():
            db.update_record_link(doc_no, data['link'])

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
