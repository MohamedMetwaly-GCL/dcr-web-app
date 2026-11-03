"""test_drive_upload_smoke.py - MANUAL smoke test for Drive uploads. NOT part of the automatic tests.

Uploads a small, clearly-named test file into ONE test folder you specify, verifies it, then
deletes ONLY that file. It touches nothing else and never reads DATABASE_URL.

Usage (PowerShell, from d:\\DCR):
    python test_drive_upload_smoke.py --folder-id <TEST_FOLDER_ID> --auth service_account
    python test_drive_upload_smoke.py --folder-id <TEST_FOLDER_ID> --auth oauth
Options:
    --size-kb N    payload size (default 1). Use 1024 with --chunk-kb 256 to exercise multi-chunk upload.
    --chunk-kb N   chunk size in KB, multiple of 256 (default 256)
    --keep         do not delete the test file afterwards
Credentials come from the same env vars / service_account.json the app uses.
"""
import argparse
import io
import os
import sys
import time

import drive_upload as U


def step(msg):
    print("\n[%s] %s" % (time.strftime("%H:%M:%S"), msg))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--folder-id", required=True)
    ap.add_argument("--auth", choices=["service_account", "oauth"], default=None)
    ap.add_argument("--size-kb", type=int, default=1)
    ap.add_argument("--chunk-kb", type=int, default=256)
    ap.add_argument("--keep", action="store_true")
    a = ap.parse_args()

    name = "DCR_SMOKE_TEST_%d.txt" % int(time.time())
    payload = (("DCR smoke test %s\n" % name) * 4096).encode()[: a.size_kb * 1024].ljust(a.size_kb * 1024, b".")
    file_id, token, ok = None, None, False
    try:
        step("1/6 Authenticating (mode: %s)" % (a.auth or os.environ.get("DRIVE_UPLOAD_AUTH") or "service_account"))
        creds = U.get_upload_credentials(a.auth)
        token = U.access_token(creds)
        print("    OK - access token obtained")

        step("2/6 Checking test folder")
        f = U.get_file(token, a.folder_id)
        print("    folder: %r  type-ok: %s  can-add-children: %s" % (
            f.get("name"), f.get("mimeType") == "application/vnd.google-apps.folder",
            (f.get("capabilities") or {}).get("canAddChildren")))
        if f.get("mimeType") != "application/vnd.google-apps.folder":
            raise U.DriveUploadError("The given ID is not a folder", reason="not_a_folder")

        step("3/6 Creating resumable upload session")
        url = U.create_resumable_session(token, a.folder_id, name, "text/plain", len(payload))
        print("    OK - session created")

        step("4/6 Uploading %d KB in %d KB chunks" % (a.size_kb, a.chunk_kb))
        res = U.upload_stream_resumable(url, io.BytesIO(payload), len(payload), chunk_size=a.chunk_kb * 1024,
                                        progress=lambda d, t: print("    %d/%d bytes" % (d, t)))
        file_id = res["id"]
        print("    OK - file id: %s" % file_id)

        step("5/6 Verifying uploaded file")
        meta = U.get_file(token, file_id, fields=U.FILE_FIELDS)
        owners = [o.get("emailAddress") for o in meta.get("owners", [])]
        print("    name: %s | size: %s | owner: %s" % (meta.get("name"), meta.get("size"), owners))
        assert int(meta.get("size", -1)) == len(payload), "size mismatch"
        print("    link: %s" % meta.get("webViewLink"))
        ok = True
    except U.DriveUploadError as e:
        print("\n    FAILED: %s" % e)
        print("    reason: %s | http: %s" % (e.reason, e.status))
        if e.hint:
            print("    HINT: %s" % e.hint)
    except Exception as e:
        print("\n    FAILED (%s): %s" % (type(e).__name__, str(e)[:300]))
    finally:
        if file_id and token:
            if a.keep:
                step("6/6 --keep given: leaving test file %s in place" % file_id)
            else:
                step("6/6 Deleting the test file")
                try:
                    U.delete_file(token, file_id)
                    print("    OK - test file deleted")
                except U.DriveUploadError as e:
                    print("    could not delete (%s) - remove %s manually" % (e, name))
    print("\n=== RESULT: %s ===" % ("PASS" if ok else "FAIL"))
    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()
