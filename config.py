"""config.py - DCR application configuration.

Centralises all environment-driven settings so app.py stays clean.
Imported once at startup; all values are plain Python — no Flask context needed.

Step 1 of the incremental refactor. Only configuration lives here;
no routes, no business logic, no database calls.
"""
import os
import datetime

# ── Deployment environment ────────────────────────────────────
# True when running on Railway (or any host that sets RENDER=true).
# Controls secure cookies and debug mode.
IS_RENDER = bool(os.environ.get("RENDER"))

# ── Secret key ───────────────────────────────────────────────
# Must be set as a Railway environment variable before deploying.
# Generate a value with:
#   python -c "import secrets; print(secrets.token_hex(32))"
# Then add to Railway → service → Variables → SECRET_KEY
_secret_key = os.environ.get("SECRET_KEY")
if not _secret_key:
    raise RuntimeError(
        "SECRET_KEY environment variable is not set. "
        "Generate a value with:  python -c \"import secrets; print(secrets.token_hex(32))\" "
        "and add it to your Railway service Variables before deploying."
    )
SECRET_KEY = _secret_key

# ── Session cookie configuration ─────────────────────────────
# Applied to the Flask app via app.config.update(**SESSION_CONFIG).
SESSION_CONFIG = dict(
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE="Lax",
    SESSION_COOKIE_SECURE=IS_RENDER,
    PERMANENT_SESSION_LIFETIME=datetime.timedelta(hours=8),
)
