"""get_oauth_refresh_token.py - ONE-TIME helper (run by you, locally) to obtain a Google OAuth refresh token.

Prerequisites (Google Cloud Console, same project as the Service Account is fine):
  1. APIs & Services > Library > enable "Google Drive API"
  2. OAuth consent screen: User type External, add your Gmail as a test user,
     then click "PUBLISH APP" (status "In production"). IMPORTANT: apps left in "Testing"
     get refresh tokens that EXPIRE AFTER 7 DAYS. An unverified app in production is fine for personal use.
  3. Credentials > Create credentials > OAuth client ID > "Desktop app" > download JSON.

Run:
    pip install google-auth-oauthlib
    python get_oauth_refresh_token.py path\\to\\client_secret.json

A browser opens; sign in with the Google account that OWNS the Drive folders and approve.
The script prints three values. Put them in environment variables (locally, and later in Railway
variables yourself) - never commit them:
    GOOGLE_OAUTH_CLIENT_ID, GOOGLE_OAUTH_CLIENT_SECRET, GOOGLE_OAUTH_REFRESH_TOKEN
    DRIVE_UPLOAD_AUTH=oauth
"""
import json
import sys

SCOPES = ["https://www.googleapis.com/auth/drive"]


def main():
    if len(sys.argv) != 2:
        print(__doc__)
        sys.exit(2)
    try:
        from google_auth_oauthlib.flow import InstalledAppFlow
    except ImportError:
        print("Missing dependency. Run:  pip install google-auth-oauthlib")
        sys.exit(2)
    with open(sys.argv[1], "r", encoding="utf-8") as f:
        cfg = json.load(f)
    block = cfg.get("installed") or cfg.get("web") or {}
    flow = InstalledAppFlow.from_client_secrets_file(sys.argv[1], SCOPES)
    # access_type=offline + prompt=consent guarantees a refresh token is returned
    creds = flow.run_local_server(port=0, access_type="offline", prompt="consent")
    if not creds.refresh_token:
        print("No refresh token returned. Remove the app from https://myaccount.google.com/permissions and retry.")
        sys.exit(1)
    print("\nSet these environment variables (keep them secret):\n")
    print("GOOGLE_OAUTH_CLIENT_ID=%s" % block.get("client_id", ""))
    print("GOOGLE_OAUTH_CLIENT_SECRET=%s" % block.get("client_secret", ""))
    print("GOOGLE_OAUTH_REFRESH_TOKEN=%s" % creds.refresh_token)
    print("DRIVE_UPLOAD_AUTH=oauth")


if __name__ == "__main__":
    main()
