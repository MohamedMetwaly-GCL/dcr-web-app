"""get_oauth_refresh_token.py - Helper to obtain a Google OAuth refresh token.

Prerequisites in Google Cloud Console:
  1. APIs & Services > OAuth consent screen:
     - Set User Type to 'External'
     - CRITICAL: Click "PUBLISH APP" to switch status from 'Testing' to 'In production'.
       * If left in 'Testing', Google automatically expires refresh tokens after exactly 7 days!
       * Setting it to 'In production' ensures the refresh token NEVER expires.
  2. APIs & Services > Credentials:
     - Ensure you have an OAuth 2.0 Client ID (type: Desktop application).

How to run:
  Option A (with JSON file):
     python get_oauth_refresh_token.py path\\to\\client_secret.json

  Option B (interactive - just enter Client ID and Client Secret from Railway / Google Cloud):
     python get_oauth_refresh_token.py
"""
import json
import os
import sys

SCOPES = ["https://www.googleapis.com/auth/drive"]


def main():
    try:
        from google_auth_oauthlib.flow import InstalledAppFlow
    except ImportError:
        print("Missing dependency. Run:  pip install google-auth-oauthlib")
        sys.exit(2)

    client_id = ""
    client_secret = ""

    if len(sys.argv) == 2:
        json_path = sys.argv[1]
        if not os.path.exists(json_path):
            print(f"Error: File '{json_path}' does not exist.")
            sys.exit(1)
        with open(json_path, "r", encoding="utf-8") as f:
            cfg = json.load(f)
        block = cfg.get("installed") or cfg.get("web") or {}
        client_id = block.get("client_id", "").strip()
        client_secret = block.get("client_secret", "").strip()
    else:
        print("=" * 60)
        print(" Google OAuth Refresh Token Generator")
        print("=" * 60)
        print("You can find Client ID and Secret in Google Cloud Console or Railway Variables.\n")
        client_id = input("Enter GOOGLE_OAUTH_CLIENT_ID: ").strip()
        client_secret = input("Enter GOOGLE_OAUTH_CLIENT_SECRET: ").strip()

    if not client_id or not client_secret:
        print("\nError: Both Client ID and Client Secret are required.")
        sys.exit(1)

    client_config = {
        "installed": {
            "client_id": client_id,
            "client_secret": client_secret,
            "auth_uri": "https://accounts.google.com/o/oauth2/auth",
            "token_uri": "https://oauth2.googleapis.com/token",
            "redirect_uris": ["http://localhost"]
        }
    }

    print("\nOpening your browser for authorization...")
    print("Please sign in with the Google Account that owns or manages the Drive folders.")
    
    flow = InstalledAppFlow.from_client_config(client_config, SCOPES)
    # access_type=offline + prompt=consent guarantees a fresh refresh token is returned
    creds = flow.run_local_server(port=0, access_type="offline", prompt="consent")

    if not creds.refresh_token:
        print("\n[!] No refresh token returned by Google.")
        print("Go to https://myaccount.google.com/permissions, remove access for this app, and run again.")
        sys.exit(1)

    print("\n" + "=" * 60)
    print(" SUCCESS! Copy this new token to Railway Variables:")
    print("=" * 60)
    print(f"GOOGLE_OAUTH_REFRESH_TOKEN={creds.refresh_token}")
    print("=" * 60)
    print("\nReminder: Make sure OAuth Consent Screen in Google Cloud Console is published to 'In production' so this token never expires!")


if __name__ == "__main__":
    main()
