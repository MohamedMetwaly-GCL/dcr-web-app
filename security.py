"""security.py - Security hardening module for DCR.

Handles:
1. Username sanitization and SQL/Command injection prevention.
2. Brute-force protection and Rate Limiting (6 failed attempts threshold, 15-minute lockout).
3. Direct file upload isolation, magic-byte verification, and antivirus/malware heuristic scanner.
"""
import base64
import io
import os
import re
import time
import zipfile
import threading
from typing import Tuple, List, Optional

# =====================================================================
# 1. USERNAME VALIDATION & INJECTION PREVENTION
# =====================================================================

USERNAME_REGEX = re.compile(r"^[a-zA-Z0-9_.-]{3,40}$")

# Suspicious patterns or characters that strongly suggest injection or malicious input
SUSPICIOUS_USERNAME_PATTERNS = [
    re.compile(r"['\"`]"),                   # Quotes
    re.compile(r"[;<>()/\\|=&$]"),           # Shell / SQL separators and redirections
    re.compile(r"--"),                       # SQL line comment
    re.compile(r"/\*|\*/"),                  # SQL block comment
    re.compile(r"\b(union|select|insert|update|delete|drop|alter|exec|system|script)\b", re.IGNORECASE),
    re.compile(r"[\x00-\x1f\x7f-\x9f]"),    # Non-printable control characters / null bytes
]

def validate_username(username: str) -> Tuple[bool, str, bool]:
    """
    Validates username format and blocks injection patterns.
    Returns: (is_valid, error_message, is_suspicious_injection)
    """
    if not username:
        return False, "Username is required.", False

    uname = str(username).strip()

    # Check for suspicious injection patterns
    for pat in SUSPICIOUS_USERNAME_PATTERNS:
        if pat.search(uname):
            return False, "Invalid username: prohibited characters or injection patterns detected.", True

    if len(uname) < 3 or len(uname) > 40:
        return False, "Username must be between 3 and 40 characters.", False

    if not USERNAME_REGEX.match(uname):
        return False, "Username can only contain letters, numbers, dots, hyphens, and underscores.", False

    return True, "", False


# =====================================================================
# 2. BRUTE FORCE PROTECTION & RATE LIMITING
# =====================================================================

def get_client_ip(req=None) -> str:
    """Extracts client IP address safely from Flask request headers."""
    try:
        if req is None:
            from flask import request as flask_req, has_request_context
            if has_request_context():
                req = flask_req
        if req:
            cf_ip = req.headers.get("CF-Connecting-IP")
            if cf_ip:
                return cf_ip.strip()
            x_fwd = req.headers.get("X-Forwarded-For")
            if x_fwd:
                return x_fwd.split(",")[0].strip()
            return req.remote_addr or ""
    except Exception:
        pass
    return ""


class LoginRateLimiter:
    """
    In-memory, thread-safe rate limiter for login attempts.
    Threshold: 6 failed attempts within window_seconds (default 10 mins).
    Lockout duration: lockout_seconds (default 15 mins).
    """
    def __init__(self, max_attempts: int = 6, window_seconds: int = 600, lockout_seconds: int = 900):
        self.max_attempts = max_attempts
        self.window_seconds = window_seconds
        self.lockout_seconds = lockout_seconds
        self._lock = threading.Lock()
        self._records = {}  # key -> {"attempts": [ts, ...], "locked_until": ts}

    def is_locked(self, ip: str, username: str = "") -> Tuple[bool, int]:
        """
        Checks if the IP or username is currently locked out.
        Returns: (is_locked, remaining_seconds)
        """
        now = time.time()
        with self._lock:
            keys = [f"ip:{ip}"]
            if username:
                keys.append(f"user:{username.strip().lower()}")
            for key in keys:
                rec = self._records.get(key)
                if rec and rec.get("locked_until", 0) > now:
                    remaining = int(rec["locked_until"] - now)
                    return True, max(1, remaining)
        return False, 0

    def record_failure(self, ip: str, username: str = "") -> Tuple[bool, int, int]:
        """
        Records a failed login attempt for the IP and username.
        Returns: (is_locked_now, remaining_attempts, lockout_seconds)
        """
        now = time.time()
        with self._lock:
            self._cleanup(now)
            keys = [f"ip:{ip}"]
            if username:
                keys.append(f"user:{username.strip().lower()}")

            is_locked_now = False
            min_remaining = self.max_attempts
            lockout_sec = 0

            for key in keys:
                rec = self._records.setdefault(key, {"attempts": [], "locked_until": 0})
                # Keep only attempts within sliding window
                rec["attempts"] = [t for t in rec["attempts"] if now - t < self.window_seconds]
                rec["attempts"].append(now)

                if len(rec["attempts"]) >= self.max_attempts:
                    rec["locked_until"] = now + self.lockout_seconds
                    rec["attempts"] = []
                    is_locked_now = True
                    lockout_sec = self.lockout_seconds
                else:
                    rem = self.max_attempts - len(rec["attempts"])
                    if rem < min_remaining:
                        min_remaining = rem

            return is_locked_now, (0 if is_locked_now else min_remaining), lockout_sec

    def record_success(self, ip: str, username: str = ""):
        """Clears failed attempt records upon successful login."""
        with self._lock:
            self._records.pop(f"ip:{ip}", None)
            if username:
                self._records.pop(f"user:{username.strip().lower()}", None)

    def reset(self, ip: Optional[str] = None, username: Optional[str] = None):
        """Resets rate limiting records (useful for testing or admin unblock)."""
        with self._lock:
            if ip:
                self._records.pop(f"ip:{ip}", None)
            if username:
                self._records.pop(f"user:{username.strip().lower()}", None)
            if not ip and not username:
                self._records.clear()

    def _cleanup(self, now: float):
        """Purges old expired entries to prevent memory growth."""
        if len(self._records) > 2000:
            keys_to_del = [
                k for k, v in self._records.items()
                if v.get("locked_until", 0) <= now
                and not any(now - t < self.window_seconds for t in v.get("attempts", []))
            ]
            for k in keys_to_del:
                self._records.pop(k, None)


# Global rate limiter instance (6 attempts limit, 10 min window, 15 min lockout)
login_limiter = LoginRateLimiter(max_attempts=6, window_seconds=600, lockout_seconds=900)


# =====================================================================
# 3. DIRECT FILE UPLOAD ISOLATION & ANTIVIRUS SCANNING
# =====================================================================

# Base64-encoded to avoid triggering local antivirus scanners on source code
EICAR_SIGNATURE = base64.b64decode(b"WDVPIVAlQEFQWzRcUFpYNTQoUF4pN0NDKTd9JEVJQ0FSLVNUQU5EQVJELUFOVElWSVJVUy1URVNULUZJTEUhJEgrSCo=")

DANGEROUS_EXTS_INSIDE_ARCHIVE = {
    ".exe", ".dll", ".bat", ".cmd", ".sh", ".vbs", ".ps1", ".scr", ".jar",
    ".elf", ".com", ".pif", ".hta", ".cpl", ".msi", ".vbe", ".wsf", ".wsh"
}

DANGEROUS_MALWARE_PATTERNS = [
    re.compile(b"<\\?php", re.IGNORECASE),
    re.compile(b"<%[\\s=]", re.IGNORECASE),
    re.compile(b"<script[^>]*>", re.IGNORECASE),
    re.compile(b"\\bpowershell(\\s|\\.exe)", re.IGNORECASE),
    re.compile(b"/bin/(ba)?sh", re.IGNORECASE),
]

# Magic bytes
MAGIC_ZIP = (b"PK\x03\x04", b"PK\x05\x06")
MAGIC_OLE2 = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
MAGIC_PNG = b"\x89PNG\r\n\x1a\n"
MAGIC_JPEG = b"\xff\xd8\xff"
MAGIC_GIF = (b"GIF87a", b"GIF89a")


def scan_file_for_malware(raw_bytes: bytes) -> Tuple[bool, str]:
    """
    Heuristic and signature scan for malware or dangerous payloads.
    Returns: (is_clean, reason)
    """
    if not raw_bytes:
        return False, "File is empty."

    # 1. EICAR Standard Antivirus Test String check
    if EICAR_SIGNATURE in raw_bytes:
        return False, "Threat detected: EICAR test signature found."

    # 2. Check for Windows PE Executable Header (MZ header)
    if raw_bytes.startswith(b"MZ"):
        return False, "Threat detected: Executable binary (PE/MZ) header detected."

    # 3. Check for Linux ELF Header
    if raw_bytes.startswith(b"\x7fELF"):
        return False, "Threat detected: Linux ELF executable binary detected."

    # 4. Check for Mach-O Executable Headers
    if len(raw_bytes) >= 4 and raw_bytes[:4] in (
        b"\xfe\xed\xfa\xce", b"\xfe\xed\xfa\xcf", b"\xce\xfa\xed\xfe", b"\xcf\xfa\xed\xfe"
    ):
        return False, "Threat detected: Mach-O binary detected."

    # 5. Check for embedded webshell / script patterns in raw bytes
    sample = raw_bytes[:65536]  # Scan first 64KB for embedded webshell signatures
    for pat in DANGEROUS_MALWARE_PATTERNS:
        if pat.search(sample):
            return False, "Threat detected: Malicious embedded script/webshell signature detected."

    return True, ""


def scan_and_validate_file(
    raw_bytes: bytes,
    declared_ext: str,
    allowed_exts: List[str],
    max_size_mb: int = 15
) -> Tuple[bool, str]:
    """
    Validates file size, scans for malware signatures, and verifies magic bytes/container structure.
    Returns: (is_safe, error_message)
    """
    ext = (declared_ext or "").strip().lower().lstrip(".")
    allowed = [e.lower().lstrip(".") for e in allowed_exts]

    if ext not in allowed:
        return False, f"File extension '.{ext}' is not permitted. Allowed: {', '.join(allowed)}."

    # Size check
    size_bytes = len(raw_bytes)
    max_bytes = max_size_mb * 1024 * 1024
    if size_bytes > max_bytes:
        return False, f"File size ({size_bytes / (1024*1024):.1f} MB) exceeds maximum limit of {max_size_mb} MB."

    if size_bytes == 0:
        return False, "File is empty."

    # Antivirus & malware signature scan
    is_clean, av_reason = scan_file_for_malware(raw_bytes)
    if not is_clean:
        return False, f"Antivirus scan failed: {av_reason}"

    # Container and Magic Byte verification
    if ext == "xlsx":
        # XLSX must be a ZIP container
        if not (raw_bytes.startswith(MAGIC_ZIP[0]) or raw_bytes.startswith(MAGIC_ZIP[1])):
            return False, "Invalid XLSX format: file is not a valid OpenXML spreadsheet."

        try:
            with zipfile.ZipFile(io.BytesIO(raw_bytes), "r") as zf:
                # 1. Zip Bomb check: verify uncompressed size and compression ratio
                total_uncompressed = 0
                for info in zf.infolist():
                    total_uncompressed += info.file_size
                    # Check individual file dangerous extensions
                    _, item_ext = os.path.splitext(info.filename.lower())
                    if item_ext in DANGEROUS_EXTS_INSIDE_ARCHIVE:
                        return False, f"Security threat: executable payload '{info.filename}' found inside archive."

                    # Check for VBA Macros (.xlsm disguised as .xlsx)
                    if "vbaproject.bin" in info.filename.lower() or "vba" in info.filename.lower():
                        return False, "Security violation: VBA macros / scripts are not allowed in document imports."

                if total_uncompressed > 100 * 1024 * 1024:  # 100MB uncompressed limit
                    return False, "Security violation: Decompression bomb protection triggered (uncompressed data exceeds 100MB)."

                # Must contain [Content_Types].xml
                namelist = zf.namelist()
                if not any(name == "[Content_Types].xml" for name in namelist):
                    return False, "Invalid XLSX structure: missing [Content_Types].xml."

        except zipfile.BadZipFile:
            return False, "Corrupted or invalid XLSX archive."
        except Exception as e:
            return False, f"Error validating XLSX structure: {str(e)}"

    elif ext == "xls":
        # Legacy XLS must start with OLE2 header
        if not raw_bytes.startswith(MAGIC_OLE2):
            return False, "Invalid XLS format: missing OLE2 compound document header."

    elif ext == "csv":
        # CSV must be valid text, not binary
        # Check first 4096 bytes for excessive null bytes
        prefix = raw_bytes[:4096]
        if b"\x00" in prefix:
            return False, "Invalid CSV format: binary content detected."

        # Verify text decodability
        decoded = False
        text = ""
        for encoding in ("utf-8", "utf-8-sig", "latin-1", "cp1256", "utf-16"):
            try:
                text = raw_bytes[:16384].decode(encoding)
                decoded = True
                break
            except Exception:
                continue

        if not decoded:
            return False, "Invalid CSV format: unable to decode text."

        # Check for Formula Injection / DDE command execution patterns
        # e.g., =cmd|, =SYSTEM|, =DDE|, +cmd, @SUM(cmd|...) anywhere in a cell (start of line or after comma/delimiter)
        if re.search(r"(?:^|[,;\"\t])\s*[=@+\-]\s*(cmd|powershell|mshta|cscript|wscript|system|dde)", text, re.IGNORECASE | re.MULTILINE):
            return False, "Security violation: Dangerous spreadsheet formula / command execution pattern detected in CSV."

    return True, ""


def scan_and_validate_image(raw_bytes: bytes, max_size_mb: int = 3) -> Tuple[bool, str, str]:
    """
    Validates uploaded images (logos, avatars).
    Returns: (is_safe, error_message, detected_mime)
    """
    size_bytes = len(raw_bytes)
    max_bytes = max_size_mb * 1024 * 1024
    if size_bytes > max_bytes:
        return False, f"Image size exceeds {max_size_mb} MB limit.", ""

    if size_bytes == 0:
        return False, "Image file is empty.", ""

    # Malware scan
    is_clean, av_reason = scan_file_for_malware(raw_bytes)
    if not is_clean:
        return False, f"Security threat in image: {av_reason}", ""

    # Check magic bytes
    if raw_bytes.startswith(MAGIC_PNG):
        return True, "", "image/png"
    elif raw_bytes.startswith(MAGIC_JPEG):
        return True, "", "image/jpeg"
    elif raw_bytes.startswith(MAGIC_GIF[0]) or raw_bytes.startswith(MAGIC_GIF[1]):
        return True, "", "image/gif"
    elif raw_bytes.startswith(b"RIFF") and len(raw_bytes) >= 12 and raw_bytes[8:12] == b"WEBP":
        return True, "", "image/webp"

    # Reject SVGs explicitly to avoid Stored XSS
    if b"<svg" in raw_bytes[:2048].lower() or b"<?xml" in raw_bytes[:2048].lower():
        return False, "SVG images are not allowed for security reasons (use PNG, JPG, or WEBP).", ""

    return False, "Unsupported image format. Allowed formats: PNG, JPEG, WEBP, GIF.", ""
