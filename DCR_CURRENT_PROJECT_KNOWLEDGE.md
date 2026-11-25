# DCR Current Project Knowledge

**Repository:** `D:\DCR-dev` (synced with `D:\DCR` and `origin/main`)  
**Active Local Branch:** `main`  
**Latest Merged Milestone Commit:** `6e03a6a` (merged via `fix/restore-gcl-swirl-logo` into `main`)  
**Release / Deployment Status:** `main` updated with GCL swirl logo restoration (`fix/restore-gcl-swirl-logo`, `6e03a6a`) and pushed to `origin/main` via `D:\DCR` for Railway Auto-Deploy (prior release `9069abe` verified live on `2026-11-24`).

---

## Completed Branding Fix: Restore Original GCL Swirl Logo (`fix/restore-gcl-swirl-logo`)
- Restored `static/favicon.ico` (`316,265` bytes, multi-size transparent RGBA `GCL_Logo_2019_Transparent_Swirl_icoconvert.ico` from `4ff6449`).
- Regenerated transparent RGBA PNG icons (`static/apple-touch-icon.png` `180x180`, `static/icon-192.png` `192x192`, `static/icon-512.png` `512x512`) from the `256x256` ICO frame with ~6% transparent padding.
- Added `?v=gcl2019` cache-busting query parameters across `html_render.py` (`_head_icon_tags()`, `printHseQrLabel()`), `templates/executive_summary.html`, and `static/manifest.json`.

---

## Completed Milestone: HSE Sub-Tabs & PTW Modularization (`feat/ptw-subtypes-structure`)

### 1. PTW Permit Subtypes (`HWP`–`RGP`) & Legacy Preservation
- Standardized the 8 active `PTW` (`hse_ptw_subtype`) permit types and default `Doc No.` codes in `hse_defaults.py` / `db.py`:
  - `Hot Work Permit` (`HWP`)
  - `Cold Work Permit` (`CWP`)
  - `Work at Height Permit` (`WAH`)
  - `Confined Space Permit` (`CSP`)
  - `Excavation Permit` (`EXP`)
  - `Electrical / Isolation Permit` (`EIP`)
  - `Lifting Operations Permit` (`LOP`)
  - `Radiography Permit` (`RGP`)
- Single-select `<select id="f-legacySubtype">` enforcement in the record modal (replacing legacy multi-select checkboxes), plus fix for the `remarks` (`ملاحظات`) `undefined` placeholder.
- Existing `Doc No.` values and Google Drive file links on legacy `PTW` records are preserved untouched.

### 2. Independent Per-Subtype `Doc No.` Auto-Numbering
- Added per-subtype `Doc No.` sequence generation in `hse_defaults.py` and `GET /api/next_doc_no/<pid>/<dt_id>?subtype=...` in `app.py` (e.g., `PEM058-HWP-001`, `PEM058-CWP-001`), using configurable per-subtype codes stored in `dropdown_lists.meta`.

### 3. Google Drive Subfolder Routing (`PTW/<Subtype>/`)
- Updated `drive_routing.py` (`resolve_upload_destination`) and `blueprints/drive_uploads.py` (`api_create_upload_session`) so `PTW` uploads route into `<Project HSE Root>/PTW/<Full Subtype Name>/` (or `<Project HSE Root>/PTW/` when no subtype is selected), auto-creating/reusing subtype subfolders without `Submit`/`Replied` wrappers.

### 4. Generalized Sub-Tab Strip, Inline Creation & Auto-Wiring
- Generalized `#ptw-subtabs` in `html_render.py` across all document types that have subtype lists (`PTW`, `OBS`, `INC`, `SVP`, `FEI`, and custom `*_subtype` lists).
- Added inline circular `+` sub-tab creation button (`#ptw-subtab-add-btn`) and localized `#add-subtype-modal` (with `Doc No.` code input for `PTW`).
- Added toolbar `+ Sub-Tab` / `+ تابة فرعية` button (`#btn-add-subtab`) for document types with zero subtypes configured; `POST /api/lists/<pid>` auto-wires the `legacySubtype` column in `columns_config` via `db.ensure_doc_type_subtype_column`.

### 5. Sub-Tab Right-Click Context Menu
- Added `#subtab-ctx-menu` in `html_render.py` on subtype pills (excluding `All` and `+`) for editors (`CAN_EDIT`), matching main document-type tabs (`#tabctx`) with EN/AR localization in `i18n.py`:
  - **Edit Sub-Tab Settings** (`PATCH /api/lists/<pid>`, preserving active filter on rename)
  - **Move Left / Move Right** (RTL/LTR-aware visual ordering, persisted via `POST /api/lists/<pid>/reorder`)
  - **Delete Sub-Tab** (`DELETE /api/lists/<pid>`, removing only from `dropdown_lists` without deleting or altering `records`)

### 6. Security & Validation Hardening (`M1–M3`, `L1`)
- **M1 (`db.py`):** Guarded legacy `"Maintenance Work"` removal in `_sync_all_hse_ptw_subtypes()` with persistent `app_settings` marker `ptw_maintenance_work_cleanup_v1` so it runs once and never deletes admin-readded items on subsequent restarts.
- **M2 (`blueprints/lists.py`):** Before auto-wiring `legacySubtype` on `POST /api/lists/<pid>`, verified `dt_id` exists (`404`), `can_access_doc_type` is authorized (`403`), and `list_name` matches the document type's expected subtype list pattern (`400`).
- **M3 (`db.py`, `blueprints/lists.py`):** Enforced server-side subtype name validation (non-empty, `<= 80` chars, case-insensitive uniqueness on add/rename) and `PTW` code validation (`[A-Z0-9]{2,8}`, `400` on explicit duplicate code, auto-disambiguation for colliding fallback codes).
- **L1 (`drive_routing.py`, `blueprints/drive_uploads.py`):** Added `sanitize_drive_folder_segment()` to strip NUL/control characters, `..` traversal sequences, and slashes (`/`, `\`) and cap Drive subfolder names at `80` characters.
