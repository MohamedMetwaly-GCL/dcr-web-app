"""blueprints/hse.py - Scoped HSE Command Dashboard, QR Verification, and Batch Inspection routes."""

from datetime import date
from html import escape
import json
import re
import secrets

from flask import Blueprint, Response, jsonify, make_response, request

import db
from auth import can_access_doc_type, can_edit, can_view_project, current_user
from hse_analytics import DEFAULT_WARN_DAYS, get_scoped_hse_analytics, normalize_hse_status, parse_hse_date
from hse_qr import generate_qr_svg
from i18n import get_lang, html_dir, t, t_dt, t_status

hse_bp = Blueprint("hse", __name__)

_QR_TOKEN_RE = re.compile(r"^[A-Za-z0-9_-]{8,64}$")


def _parse_int_arg(name, default, minimum, maximum):
    raw = request.args.get(name)
    if raw in (None, ""):
        return default
    try:
        val = int(raw)
    except (TypeError, ValueError):
        return default
    return max(minimum, min(maximum, val))


def _load_scoped_payload():
    u = current_user()
    if not u:
        return None, (jsonify(error="LOGIN_REQUIRED"), 403)
    pid = request.args.get("pid") or request.args.get("project_id")
    year = request.args.get("year")
    warn_days = _parse_int_arg("warn_days", DEFAULT_WARN_DAYS, 0, 365)
    months = _parse_int_arg("months", 12, 1, 24)
    alert_filter = request.args.get("type") or request.args.get("alert_type")
    payload, err = get_scoped_hse_analytics(
        u,
        pid=pid,
        year=year,
        warn_days=warn_days,
        months=months,
        alert_filter=alert_filter,
    )
    if err:
        msg, code = err
        return None, (jsonify(error=msg), code)
    return payload, None


@hse_bp.route("/api/hse/kpis")
def api_hse_kpis():
    payload, err_resp = _load_scoped_payload()
    if err_resp:
        return err_resp
    kpis = payload["kpis"]
    return jsonify(
        ok=True,
        **kpis,
        projects=payload["projects"],
        alerts_counts=payload["alerts_counts"],
    )


@hse_bp.route("/api/hse/trends")
def api_hse_trends():
    payload, err_resp = _load_scoped_payload()
    if err_resp:
        return err_resp
    return jsonify(
        ok=True,
        monthly_trends=payload["monthly_trends"],
        ptw_status=payload["ptw_status"],
        register_breakdown=payload["register_breakdown"],
        available_years=payload["available_years"],
    )


@hse_bp.route("/api/hse/alerts")
def api_hse_alerts():
    payload, err_resp = _load_scoped_payload()
    if err_resp:
        return err_resp
    return jsonify(
        ok=True,
        alerts=payload["alerts"],
        counts=payload["alerts_counts"],
        warn_days=payload["warn_days"],
    )


@hse_bp.route("/api/hse/overview")
def api_hse_overview():
    payload, err_resp = _load_scoped_payload()
    if err_resp:
        return err_resp
    return jsonify(ok=True, **payload)


def _resolve_hse_record_context(record_id):
    """Fetch (rec, proj, dt_code) for an existing record ID, else None tuple."""
    rid = str(record_id or "").strip()
    if not rid:
        return None, None, ""
    rec = db.get_record_by_id(rid)
    if not rec:
        return None, None, ""
    pid = rec.get("_project_id")
    dt_id = rec.get("_dt_id")
    proj = db.get_project(pid) if pid else None
    dt = db.get_doc_type(pid, dt_id) if (pid and dt_id) else None
    dt_code = str((dt or {}).get("code") or dt_id or "").strip().upper()
    return rec, proj, dt_code


def _ensure_qr_token(record_id, project_id, existing_token=None):
    """Return an immutable qr_token stored inside records.data without overwriting an existing one."""
    cur = str(existing_token or "").strip()
    if cur and _QR_TOKEN_RE.fullmatch(cur):
        return cur
    candidate = secrets.token_urlsafe(16)
    db.exe(
        """UPDATE records
           SET data = COALESCE(data, '{}'::jsonb) || %s::jsonb,
               updated_at = NOW()
           WHERE id = %s AND project_id = %s AND COALESCE(data->>'qr_token', '') = ''""",
        (json.dumps({"qr_token": candidate}), record_id, project_id),
    )
    row = db.q("SELECT data->>'qr_token' AS qr_token FROM records WHERE id=%s AND project_id=%s", (record_id, project_id), one=True)
    stored = str((row or {}).get("qr_token") or "").strip()
    return stored or candidate


@hse_bp.route("/api/hse/qr/<record_id>")
def api_hse_qr(record_id):
    """Generate or fetch an immutable QR verification badge for an HSE FEI/PTW record."""
    u = current_user()
    if not u:
        return jsonify(error="LOGIN_REQUIRED"), 403

    rec, proj, dt_code = _resolve_hse_record_context(record_id)
    if not rec or not proj:
        return jsonify(error="Record not found"), 404

    pid = rec.get("_project_id")
    dt_id = rec.get("_dt_id")
    if not can_view_project(pid, u):
        return jsonify(error="Forbidden"), 403
    if not can_access_doc_type(pid, dt_id, u):
        return jsonify(error="Document type access denied"), 403
    if str(proj.get("workspace_id") or "").strip().upper() != "HSE" or dt_code not in ("FEI", "PTW"):
        return jsonify(error="QR badges are only supported for HSE FEI and PTW records"), 400

    qr_token = _ensure_qr_token(rec.get("_id"), pid, rec.get("qr_token"))
    verify_path = f"/qr/{qr_token}"
    verify_url = f"{request.url_root.rstrip('/')}{verify_path}"
    svg = generate_qr_svg(verify_url)

    if str(request.args.get("format") or "").strip().lower() == "svg":
        resp = make_response(svg, 200)
        resp.headers["Content-Type"] = "image/svg+xml; charset=utf-8"
        resp.headers["Cache-Control"] = "no-store"
        return resp

    return jsonify(
        ok=True,
        record_id=rec.get("_id"),
        project_id=pid,
        project_code=proj.get("code") or pid,
        project_name=proj.get("name") or pid,
        dt_code=dt_code,
        doc_no=rec.get("docNo") or rec.get("_id"),
        title=rec.get("title") or rec.get("equipmentRef") or "",
        subtype=rec.get("legacySubtype") or "",
        location=rec.get("location") or "",
        qr_token=qr_token,
        verify_path=verify_path,
        verify_url=verify_url,
        svg=svg,
    )


_ALLOWED_FEI_BATCH_STATUSES = (
    "Valid",
    "Needs Maintenance",
    "Expired",
    "Under Refill",
    "Out of Service",
)

_FEI_STATUS_CANON = {
    "valid": "Valid",
    "ok": "Valid",
    "صالحة": "Valid",
    "needs maintenance": "Needs Maintenance",
    "maintenance": "Needs Maintenance",
    "صيانة": "Needs Maintenance",
    "under refill": "Under Refill",
    "refill": "Under Refill",
    "expired": "Expired",
    "empty": "Expired",
    "فارغة": "Expired",
    "out of service": "Out of Service",
}

_LEGACY_TEXT_EN = {
    "بودرة جافة": "Dry Powder",
    "ثاني أكسيد الكربون (co2)": "CO2",
    "ثاني أكسيد الكربون": "CO2",
    "رغوة": "Foam",
    "مياه": "Water",
    "فحص لوحة الكهرباء.docx": "Electrical Panel Inspection",
    "تقرير فحص حزام الامان.docx": "Safety Harness Inspection",
    "تقرير فحص ماكينات اللحام.docx": "Welding Machine Inspection",
}


def _normalize_fei_status(raw_status, default=""):
    raw = str(raw_status or "").strip()
    if not raw:
        return default
    return _FEI_STATUS_CANON.get(raw.lower())


def _normalize_qr_text(raw_val, fallback="—"):
    txt = str(raw_val or "").strip()
    if not txt:
        return fallback
    return _LEGACY_TEXT_EN.get(txt.lower(), txt)


def _compute_public_qr_status(dt_code, data, today=None):
    ref_day = today or date.today()
    if dt_code == "PTW":
        st = normalize_hse_status(data.get("status"))
        valid_until = parse_hse_date(data.get("validUntil") or data.get("valid_until"))
        if st == "Closed":
            return "Closed", "closed"
        if st == "Open":
            if valid_until and valid_until < ref_day:
                return "Overdue", "bad"
            return "Open", "open"
        return (st if st in ("Open", "Closed") else "Open"), "open"

    # FEI verification status normalized strictly to English
    expiry = parse_hse_date(data.get("expiryDate") or data.get("expiry_date") or data.get("validUntil") or data.get("valid_until"))
    raw_cond = str(data.get("extinguisherStatus") or data.get("status") or "").strip()
    canon_cond = _normalize_fei_status(raw_cond)
    if canon_cond in ("Expired", "Out of Service"):
        return canon_cond, "bad"
    if expiry is not None and expiry < ref_day:
        return "Expired", "bad"
    if canon_cond in ("Needs Maintenance", "Under Refill"):
        return canon_cond, "open"
    if expiry is not None or canon_cond == "Valid":
        return "Valid", "ok"
    norm_cond = normalize_hse_status(raw_cond)
    if norm_cond in ("Open", "Closed"):
        return norm_cond, ("ok" if norm_cond == "Closed" else "open")
    return "Valid", "ok"


@hse_bp.route("/qr/<qr_token>")
def public_qr_verify(qr_token):
    """Strictly read-only public QR verification card for FEI and PTW records (no PII / no Drive URLs)."""
    lang = get_lang(request)
    dir_attr = html_dir(lang)
    switch_label = escape(t("lang.switch_label", lang))
    switch_target = escape(t("lang.switch_target", lang))
    switch_title = escape(t("lang.switch_title", lang))
    lang_toggle_js = (
        "function switchDcrLang(nextLang){var n=(String(nextLang||'').trim().toLowerCase()==='ar')?'ar':'en';"
        "try{document.cookie='dcr_lang='+n+'; Path=/; Max-Age=31536000; SameSite=Lax';}catch(e){}"
        "try{var u=new URL(window.location.href);u.searchParams.set('lang',n);window.location.href=u.toString();}catch(e){window.location.reload();}}"
    )
    cairo_font_css = (
        "@font-face{font-family:'Cairo';src:url('/static/fonts/Cairo-Regular.ttf') format('truetype');font-weight:400 800;font-display:swap}"
        ".bidi-ltr{direction:ltr!important;unicode-bidi:isolate!important;display:inline-block}"
        ".lang-btn{padding:4px 10px;border-radius:8px;border:1px solid #cbd5e1;background:#f8fafc;color:#0f172a;font-size:12px;font-weight:700;cursor:pointer}"
        ".lang-btn:hover{background:#e2e8f0}"
    )

    token = str(qr_token or "").strip()
    row = None
    if _QR_TOKEN_RE.fullmatch(token):
        row = db.q(
            """SELECT r.id, r.project_id, r.dt_id, r.data,
                      p.name AS proj_name, p.code AS proj_code, COALESCE(p.workspace_id, 'ENG') AS workspace_id,
                      COALESCE(dt.code, r.dt_id) AS dt_code
               FROM records r
               LEFT JOIN projects p ON p.id = r.project_id
               LEFT JOIN doc_types dt ON dt.id = r.dt_id AND dt.project_id = r.project_id
               WHERE r.data->>'qr_token' = %s
               LIMIT 1""",
            (token,),
            one=True,
        )

    dt_code = str((row or {}).get("dt_code") or "").strip().upper()
    ws_id = str((row or {}).get("workspace_id") or "").strip().upper()
    if not row or ws_id != "HSE" or dt_code not in ("FEI", "PTW"):
        from html_render import _head_icon_tags
        icon_tags = _head_icon_tags()
        not_found_html = f"""<!DOCTYPE html><html lang="{lang}" dir="{dir_attr}"><head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>{escape(t("qr.not_found_title", lang))}</title>
{icon_tags}
<style>
  {cairo_font_css}
  body{{font-family:system-ui,-apple-system,sans-serif;background:#f8fafc;color:#0f172a;margin:0;padding:24px;display:flex;justify-content:center;align-items:center;min-height:90vh}}
  html[dir="rtl"] body{{font-family:'Cairo',system-ui,-apple-system,sans-serif;direction:rtl}}
  .card{{background:#fff;border:1px solid #e2e8f0;border-radius:14px;padding:28px;max-width:420px;width:100%;text-align:center;box-shadow:0 10px 25px rgba(15,23,42,.06)}}
  .top-toggle{{display:flex;justify-content:flex-end;margin-bottom:8px}}
  .foot{{margin-top:16px;padding-top:10px;border-top:1px dashed #cbd5e1;font-size:11px;color:#64748b;text-align:center}}
</style></head><body><div class="card">
  <div class="top-toggle"><button type="button" id="langSwitcherBtn" class="lang-btn" onclick="switchDcrLang('{switch_target}')" title="{switch_title}" aria-label="{switch_title}">{switch_label}</button></div>
  <div style="font-size:36px;margin-bottom:8px;">⚠️</div>
  <h2 style="margin:0 0 8px;font-size:18px;">{escape(t("qr.not_found_heading", lang))}</h2>
  <p style="margin:0;color:#64748b;font-size:13px;">{escape(t("qr.not_found_desc", lang))}</p>
  <div class="foot">{escape(t("qr.footer", lang))}</div>
</div><script>{lang_toggle_js}</script></body></html>"""
        resp = make_response(not_found_html, 404)
        resp.headers["Cache-Control"] = "no-store"
        resp.headers["X-Robots-Tag"] = "noindex, nofollow"
        return resp

    data = row.get("data") or {}
    if isinstance(data, str):
        try:
            data = json.loads(data)
        except Exception:
            data = {}

    status_label, status_tone = _compute_public_qr_status(dt_code, data)
    badge_colors = {
        "ok": ("#dcfce7", "#166534"),
        "open": ("#fef3c7", "#92400e"),
        "closed": ("#dcfce7", "#166534"),
        "bad": ("#fee2e2", "#991b1b"),
    }
    bg_hex, fg_hex = badge_colors.get(status_tone, ("#f1f5f9", "#334155"))

    default_reg_type = t("qr.reg_type_fei", "en") if dt_code == "FEI" else t("qr.reg_type_ptw", "en")
    proj_code_raw = str(row.get("proj_code") or row.get("project_id") or "")
    proj_name_raw = str(row.get("proj_name") or row.get("project_id") or "")

    # Strict whitelist: ONLY non-PII verification fields (never employee PII or raw Drive URLs)
    safe_card = {
        "project_code": proj_code_raw,
        "project_name": proj_name_raw,
        "register_type": t_dt(dt_code, default_reg_type, lang),
        "doc_no": str(data.get("docNo") or row.get("id") or "—"),
        "subtype": _normalize_qr_text(data.get("legacySubtype") or data.get("extinguisherType")),
        "location": str(data.get("location") or "—"),
        "status": t_status(status_label, lang),
        "valid_expiry": str(
            (data.get("validUntil") or data.get("valid_until") or "—")
            if dt_code == "PTW"
            else (data.get("expiryDate") or data.get("expiry_date") or data.get("validUntil") or "—")
        ),
        "last_inspection": str(data.get("lastInspection") or data.get("last_inspection") or data.get("issuedDate") or "—"),
    }

    from html_render import _head_icon_tags
    icon_tags = _head_icon_tags()
    html = f"""<!DOCTYPE html><html lang="{lang}" dir="{dir_attr}"><head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>{escape(t("qr.page_title_prefix", lang))} — {escape(safe_card['doc_no'])}</title>
{icon_tags}
<style>
  {cairo_font_css}
  body{{font-family:system-ui,-apple-system,sans-serif;background:#f1f5f9;color:#0f172a;margin:0;padding:18px;display:flex;justify-content:center;align-items:center;min-height:92vh}}
  html[dir="rtl"] body{{font-family:'Cairo',system-ui,-apple-system,sans-serif;direction:rtl;text-align:right}}
  .card{{background:#fff;border:1px solid #cbd5e1;border-top:5px solid #0d9488;border-radius:14px;padding:22px;max-width:440px;width:100%;box-shadow:0 12px 30px rgba(15,23,42,.08)}}
  .hdr{{display:flex;justify-content:space-between;align-items:center;gap:8px;margin-bottom:14px;flex-wrap:wrap}}
  .hdr-actions{{display:flex;align-items:center;gap:8px}}
  .code{{font-family:monospace;font-weight:800;font-size:15px;color:#0f766e}}
  .badge{{padding:4px 12px;border-radius:999px;font-size:12px;font-weight:800;background:{bg_hex};color:{fg_hex}}}
  .row{{display:flex;justify-content:space-between;gap:12px;padding:9px 0;border-bottom:1px solid #f1f5f9;font-size:13px}}
  .row:last-child{{border-bottom:none}}
  .lbl{{color:#64748b;font-weight:600;flex-shrink:0}}
  .val{{color:#0f172a;font-weight:700;text-align:right;word-break:break-word}}
  html[dir="rtl"] .val{{text-align:left}}
  .foot{{margin-top:14px;padding-top:10px;border-top:1px dashed #cbd5e1;font-size:11px;color:#64748b;text-align:center}}
</style></head><body>
<div class="card">
  <div class="hdr">
    <div>
      <div style="font-size:11px;font-weight:700;color:#64748b;text-transform:uppercase;">🛡️ {escape(t("qr.badge_header", lang))}</div>
      <div class="code bidi-ltr">{escape(safe_card['doc_no'])}</div>
    </div>
    <div class="hdr-actions">
      <span class="badge">{escape(safe_card['status'])}</span>
      <button type="button" id="langSwitcherBtn" class="lang-btn" onclick="switchDcrLang('{switch_target}')" title="{switch_title}" aria-label="{switch_title}">{switch_label}</button>
    </div>
  </div>
  <div>
    <div class="row"><span class="lbl">{escape(t("qr.lbl_project", lang))}</span><span class="val"><span class="bidi-ltr">{escape(safe_card['project_code'])}</span> — {escape(safe_card['project_name'])}</span></div>
    <div class="row"><span class="lbl">{escape(t("qr.lbl_register_type", lang))}</span><span class="val">{escape(safe_card['register_type'])}</span></div>
    <div class="row"><span class="lbl">{escape(t("qr.lbl_register_no", lang))}</span><span class="val"><span class="bidi-ltr">{escape(safe_card['doc_no'])}</span></span></div>
    <div class="row"><span class="lbl">{escape(t("qr.lbl_subtype", lang))}</span><span class="val">{escape(safe_card['subtype'])}</span></div>
    <div class="row"><span class="lbl">{escape(t("qr.lbl_location", lang))}</span><span class="val">{escape(safe_card['location'])}</span></div>
    <div class="row"><span class="lbl">{escape(t("qr.lbl_valid_expiry", lang))}</span><span class="val"><span class="bidi-ltr">{escape(safe_card['valid_expiry'])}</span></span></div>
    <div class="row"><span class="lbl">{escape(t("qr.lbl_last_inspection", lang))}</span><span class="val"><span class="bidi-ltr">{escape(safe_card['last_inspection'])}</span></span></div>
  </div>
  <div class="foot">{escape(t("qr.footer", lang))}</div>
</div>
<script>{lang_toggle_js}</script>
</body></html>"""
    resp = make_response(html, 200)
    resp.headers["Cache-Control"] = "no-store"
    resp.headers["X-Robots-Tag"] = "noindex, nofollow"
    return resp


@hse_bp.route("/api/hse/fei/batch-inspect", methods=["POST"])
def api_hse_fei_batch_inspect():
    """Atomically batch-update inspection fields across selected HSE FEI records."""
    u = current_user()
    if not u:
        return jsonify(error="LOGIN_REQUIRED"), 403

    data = request.get_json(silent=True) or {}
    raw_ids = data.get("record_ids")
    if not isinstance(raw_ids, list) or not raw_ids:
        return jsonify(ok=False, error="record_ids must be a non-empty list"), 400

    record_ids = []
    seen = set()
    for item in raw_ids:
        rid = str(item or "").strip()
        if rid and rid not in seen:
            seen.add(rid)
            record_ids.append(rid)
    if not record_ids or len(record_ids) > 200:
        return jsonify(ok=False, error="Provide between 1 and 200 valid record_ids"), 400

    raw_last = data.get("last_inspection") if data.get("last_inspection") not in (None, "") else data.get("lastInspection")
    if raw_last in (None, ""):
        last_insp_date = date.today()
    else:
        last_insp_date = parse_hse_date(raw_last)
        if last_insp_date is None:
            return jsonify(ok=False, error="Invalid last_inspection date format (expected YYYY-MM-DD)"), 400
    last_insp_iso = last_insp_date.isoformat()

    raw_exp = data.get("expiry_date") if data.get("expiry_date") not in (None, "") else data.get("expiryDate")
    expiry_iso = None
    if raw_exp not in (None, ""):
        exp_date = parse_hse_date(raw_exp)
        if exp_date is None:
            return jsonify(ok=False, error="Invalid expiry_date format (expected YYYY-MM-DD)"), 400
        expiry_iso = exp_date.isoformat()

    raw_status = str(data.get("status") if data.get("status") not in (None, "") else (data.get("extinguisherStatus") or "")).strip()
    if not raw_status:
        status_val = "Valid"
    else:
        status_val = _normalize_fei_status(raw_status)
        if not status_val or status_val not in _ALLOWED_FEI_BATCH_STATUSES:
            return jsonify(
                ok=False,
                error=f"Invalid FEI status '{raw_status}'. Allowed: {', '.join(_ALLOWED_FEI_BATCH_STATUSES)}",
            ), 400
    notes_val = str(data.get("notes") or "").strip()

    # Validate existence, FEI doc type, and per-record RBAC BEFORE performing any write
    targets = []
    for rid in record_ids:
        rec, proj, dt_code = _resolve_hse_record_context(rid)
        if not rec or not proj:
            return jsonify(ok=False, error=f"Record '{rid}' not found"), 404
        pid = rec.get("_project_id")
        dt_id = rec.get("_dt_id")
        if not can_edit(pid, u):
            return jsonify(ok=False, error="Forbidden"), 403
        if not can_access_doc_type(pid, dt_id, u):
            return jsonify(ok=False, error="Document type access denied"), 403
        if str(proj.get("workspace_id") or "").strip().upper() != "HSE" or dt_code != "FEI":
            return jsonify(ok=False, error=f"Record '{rid}' is not an HSE FEI record"), 400
        targets.append((rid, pid, dt_id, rec))

    patch_data = {
        "lastInspection": last_insp_iso,
        "last_inspection": last_insp_iso,
    }
    if expiry_iso is not None:
        patch_data["expiryDate"] = expiry_iso
        patch_data["expiry_date"] = expiry_iso
    if status_val:
        patch_data["status"] = status_val
        patch_data["extinguisherStatus"] = status_val
    if notes_val:
        patch_data["notes"] = notes_val
        patch_data["conditionNotes"] = notes_val

    patch_json = json.dumps(patch_data, ensure_ascii=False)
    affected_pids = set()
    with db.transaction():
        for rid, pid, dt_id, rec in targets:
            db.exe(
                """UPDATE records
                   SET data = COALESCE(data, '{}'::jsonb) || %s::jsonb,
                       updated_at = NOW()
                   WHERE id = %s AND project_id = %s""",
                (patch_json, rid, pid),
            )
            doc_no = str(rec.get("docNo") or rid)
            db.log_action(
                u.get("username", ""),
                "BATCH_INSPECT",
                project_id=pid,
                dt_id=dt_id,
                record_id=rid,
                doc_no=doc_no,
                field_name="FEI",
                new_value=status_val or last_insp_iso,
                detail=f"Batch inspected FEI '{doc_no}' (Last Inspection: {last_insp_iso}"
                + (f", Next Expiry: {expiry_iso}" if expiry_iso else "")
                + (f", Status: {status_val}" if status_val else "")
                + ")",
            )
            affected_pids.add(pid)

    for pid in affected_pids:
        try:
            db.invalidate_dashboard_cache(pid)
        except Exception:
            pass

    return jsonify(ok=True, updated=len(targets), record_ids=record_ids, patch=patch_data)
