"""Run reviewed attachment and fake Drive checks in an offline Windows session."""

from contextlib import redirect_stderr, redirect_stdout
from hashlib import sha256
from importlib import metadata
import json
import os
from pathlib import Path
import runpy
import sys
import tempfile
import traceback
from unittest.mock import patch


ROOT = Path(__file__).resolve().parent
WORKTREE = Path("~/.codex/worktrees/dcr-security-integration/DCR").expanduser().resolve()
DB_SHA256 = "d141d7067b459bf9e68fcfcbca7ccd1e501a526f7a82524d51a2fb292e77a507"
REVIEWED = {
    "security_test_support.py": "69d0fac797580c5603a84e849f0e9c59f108a159ad5a67b240738955b4dcdc80",
    "attachments.py": "4ba52e3a3259c05f45e25110b59e5eb5d0279ba9d6f080b2b0c157efe4744839",
    "drive_routing.py": "aa080cebfd3dc7e34cb0951afaf96130cfda7acdeb071217ca08f6a3a35f2a6e",
    "drive_upload.py": "286f33a1182a91798e40c89fb5d4e2204e097d1648b07dbec02c64ec9ef4fd77",
    "test_attachments.py": "d749b7a232364f49766e2b041f1d72deff0e40624f704d79b961b5b987ce26be",
    "test_drive_routing.py": "4848b0a3d7cefb7ca9aed576312873066f7fe9f334a28ff20f5591ec049af222",
    "test_drive_upload.py": "06754d4e3218d6f6e26c9ead57739b38c79225ad61992ffd0c189e8874918d02",
}


def verify_hash(path, expected):
    if sha256(path.read_bytes()).hexdigest() != expected:
        raise RuntimeError(f"SHA-256 mismatch: {path.name}")


def verify_candidate(isolated):
    manifest_path = isolated / "release_candidate_sec5_3fbdc24" / "SHA256_MANIFEST.json"
    verify_hash(manifest_path, "e0896460e9aca30a2ad01894c206a624cfda3981aa20bd70d8e07a80e0122aad")
    manifest = json.loads(manifest_path.read_text(encoding="utf-8-sig"))
    for entry in manifest["modified_source_files"]:
        expected = DB_SHA256 if entry["path"] == "db.py" else entry["rc_sha256"]
        verify_hash(ROOT / entry["path"], expected)
    for name, expected in REVIEWED.items():
        verify_hash(ROOT / name, expected)
    print(f"Verified {len(manifest['modified_source_files'])} candidate files and the auxiliary test inputs")


def check_runtime(venv):
    if sys.version_info[:3] != (3, 11, 9):
        raise RuntimeError(f"Python 3.11.9 is required. Found {sys.version.split()[0]}")
    config = (venv / "pyvenv.cfg").read_text(encoding="utf-8").lower()
    if "include-system-site-packages = false" not in config.splitlines():
        raise RuntimeError("Use a private venv with system packages disabled")
    packages = venv / "Lib/site-packages"
    versions = {d.metadata["Name"].lower(): d.version for d in metadata.distributions(path=[str(packages)])}
    for name in ("flask", "requests", "google-auth"):
        if name not in versions:
            raise RuntimeError(f"Missing dependency in the private venv: {name}")
        print(f"{name} {versions[name]}")
    major, minor, *_ = versions["flask"].split(".")
    if int(major) != 2 or int(minor) < 3:
        raise RuntimeError("Flask >=2.3,<3.0 is required")
    return packages


def run_checks(scratch, local):
    isolated = local / "DCR_DEV_ISOLATED"
    verify_candidate(isolated)
    venv = Path(sys.executable).resolve().parent.parent
    if any(venv.is_relative_to(p.resolve()) for p in (isolated, Path("D:/DCR"))):
        raise RuntimeError("Use a private runtime outside the protected checkouts")
    if any(ROOT.glob(".env*")) or (ROOT / "service_account.json").exists():
        raise RuntimeError("Credential-like filename present in worktree")
    packages = check_runtime(venv)

    environment = {name: os.environ[name] for name in ("SystemRoot", "WINDIR") if name in os.environ}
    for name in ("TEMP", "TMP", "USERPROFILE", "LOCALAPPDATA", "APPDATA", "HOME"):
        environment[name] = str(scratch)
    os.environ.clear()
    os.environ.update(environment)
    os.chdir(scratch)

    runner = runpy.run_path(str(ROOT / "security_test_support.py"))
    offline = runner["_require_offline_windows"]
    offline()
    sys.addaudithook(runner["_guard_sockets"])
    sys.path[:0] = [str(ROOT), str(packages)]

    with patch("drive_upload.get_shared_http", side_effect=runner["_blocked"]), patch(
        "drive_upload.access_token", side_effect=runner["_blocked"]
    ):
        for name in ("test_attachments.py", "test_drive_routing.py", "test_drive_upload.py"):
            offline()
            print(f"Running {name}", flush=True)
            runpy.run_path(str(ROOT / name), run_name="__main__")
            print(f"PASS {name}", flush=True)
    offline()


def main():
    if os.name != "nt" or ROOT != WORKTREE:
        raise SystemExit("Run from the DCR integration worktree on Windows")
    if not (sys.flags.isolated and sys.flags.no_site and sys.dont_write_bytecode):
        raise SystemExit("Start the private Python with -I -S -B")
    local = Path(os.environ["LOCALAPPDATA"]).resolve()
    temp_root = (local / "Temp").resolve(strict=True)
    if any(temp_root.is_relative_to(p.resolve()) for p in (local / "DCR_DEV_ISOLATED", Path("D:/DCR"))):
        raise SystemExit("Temporary directory overlaps a protected checkout")
    scratch = Path(tempfile.mkdtemp(prefix="dcr-offline-validation-", dir=temp_root))
    log_path = scratch / "validation.log"
    print(f"Validation log: {log_path}", flush=True)
    with log_path.open("x", encoding="utf-8") as log, redirect_stdout(log), redirect_stderr(log):
        try:
            print(f"Python {sys.version.split()[0]} at {sys.executable}")
            print(f"Worktree: {ROOT}")
            run_checks(scratch, local)
        except (Exception, SystemExit):
            traceback.print_exc()
            return 1
        print("Attachment and fake Drive checks passed")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
