"""html_render.py - DCR HTML page rendering.

Contains all server-side HTML generation for the three application pages
and their shared building blocks:

  _user_info_html(u)      -- topbar buttons/labels for the current user
  BASE_CSS                -- shared <style> block injected into every page
  SHARED_JS               -- shared toast/modal/apiFetch JS injected into every page
  render_login()          -- the login page (no auth required)
  render_dashboard(u)     -- the main dashboard / analytics page
  render_register(u,proj) -- the per-project document register page

Dependency map (no circular imports):
  html_render -> auth  (for can_edit)
  html_render -> db    (for get_doc_types, get_logo)
  html_render -> utils (for STATUS_COLORS)
  auth        -> db
  db          -> (psycopg2 only)

Step 3 of the incremental refactor. Only HTML rendering lives here.
No routes, no business logic, no database writes.
"""
import json
import os

import db
from flask import url_for
from auth import can_edit
from utils import STATUS_COLORS

def _user_info_html(u):
    if not u:
        return ('<a href="/login"><button class="tb-btn glow">🔐 Login</button></a>', "guest", "GUEST", "#fff3")
    role = u["role"]
    colors = {"superadmin":"rgba(240,165,0,.35)","admin":"rgba(255,255,255,.2)",
              "editor":"rgba(99,102,241,.3)","viewer":"rgba(255,255,255,.15)","guest":"rgba(255,255,255,.1)"}
    rbg = colors.get(role, "#fff3")
    labels = {"superadmin":"SUPER ADMIN","admin":"ADMIN","editor":"EDITOR","viewer":"VIEWER"}
    rlbl = labels.get(role, role.upper())
    btns = ""
    if role == "superadmin":
        btns += '<button class="tb-btn" onclick="openAdmin()">⚙ Admin</button>'
    btns += '<button class="tb-btn" onclick="changePw()">🔑</button>'
    btns += '<form action="/logout" method="post" style="display:inline"><button type="submit" class="tb-btn">⏻</button></form>'
    name = u["username"]
    return btns, name, rlbl, rbg

BASE_CSS = """
<style>
:root{--pr:#1a3a5c;--pl:#2563a8;--ac:#f0a500;--bg:#f0f4f8;--wh:#fff;--bd:#dde3ed;
  --tx:#1e2a3a;--mu:#6b7a94;--ok:#16a34a;--er:#ef4444;--wa:#f59e0b;--rd:6px}
*{box-sizing:border-box;margin:0;padding:0}
html{scroll-behavior:smooth}
body{font-family:'Segoe UI',Arial,sans-serif;background:var(--bg);color:var(--tx);font-size:13px}
body.dark{--pr:#17314d;--pl:#60a5fa;--ac:#f6c453;--bg:#0f172a;--wh:#162132;--bd:#304257;
  --tx:#e2e8f0;--mu:#9fb0c6}
#topbar{background:var(--pr);color:#fff;height:46px;display:flex;align-items:center;
  padding:0 14px;gap:8px;box-shadow:0 2px 8px rgba(0,0,0,.25);flex-shrink:0;position:relative;z-index:100}
body.dark #topbar{background:#0d1f33}
#topbar .sp{flex:1}
.tb-btn{background:rgba(255,255,255,.15);border:1px solid rgba(255,255,255,.25);color:#fff;
  padding:5px 11px;border-radius:var(--rd);cursor:pointer;font-size:12px;font-family:inherit;
  text-decoration:none;display:inline-block;transition:background .15s}
.tb-btn:hover{background:rgba(255,255,255,.28)}
.tb-btn.glow{background:rgba(240,165,0,.3);border-color:rgba(240,165,0,.7);font-weight:700}
.topbar-title-short{display:none}
.topbar-user{display:inline-flex;align-items:center;gap:6px;color:rgba(255,255,255,.85);font-size:11px}
.overlay{position:fixed;inset:0;background:rgba(0,0,0,.55);z-index:1000;
  display:flex;align-items:center;justify-content:center;backdrop-filter:blur(3px)}
.overlay.hidden{display:none!important}
.modal{background:#fff;border-radius:10px;box-shadow:0 24px 64px rgba(0,0,0,.3);
  width:92%;max-width:600px;max-height:90vh;display:flex;flex-direction:column;
  animation:mIn .18s ease}
body.dark .modal{background:var(--wh);color:var(--tx);box-shadow:0 24px 64px rgba(0,0,0,.48)}
@keyframes mIn{from{transform:translateY(-14px);opacity:0}}
.mhdr{background:var(--pr);color:#fff;padding:12px 18px;font-weight:700;font-size:13px;
  display:flex;justify-content:space-between;align-items:center;flex-shrink:0;border-radius:10px 10px 0 0}
body.dark .mhdr{background:#17314d}
.mbody{padding:16px 18px;overflow-y:auto;flex:1}
.mfoot{padding:10px 18px;border-top:1px solid var(--bd);display:flex;justify-content:flex-end;
  gap:8px;background:var(--bg);flex-shrink:0;border-radius:0 0 10px 10px}
body.dark .mfoot{background:#101a29}
.xbtn{background:none;border:none;color:#fff;font-size:20px;cursor:pointer;opacity:.7;line-height:1}
.xbtn:hover{opacity:1}
.fgrid{display:grid;grid-template-columns:1fr 1fr;gap:12px}
.fg{display:flex;flex-direction:column;gap:4px}
.fg.full{grid-column:1/-1}
.fg label{font-size:10px;font-weight:700;color:var(--mu);text-transform:uppercase;letter-spacing:.4px}
.fg input,.fg select,.fg textarea{width:100%;padding:7px 10px;border:1.5px solid var(--bd);
  border-radius:var(--rd);font-family:inherit;font-size:12px;outline:none;transition:border-color .2s;
  background:var(--wh);color:var(--tx)}
.fg input:focus,.fg select:focus,.fg textarea:focus{border-color:var(--pl);box-shadow:0 0 0 2px rgba(37,99,168,.1)}
.btn{padding:7px 16px;border-radius:var(--rd);cursor:pointer;font-family:inherit;
  font-size:12px;font-weight:600;border:1px solid transparent;transition:all .15s}
.btn-pr{background:var(--pr);color:#fff}.btn-pr:hover{background:var(--pl)}
.btn-sc{background:var(--bg);color:var(--tx);border-color:var(--bd)}.btn-sc:hover{background:var(--bd)}
.btn-ok{background:var(--ok);color:#fff}
.btn-er{background:var(--er);color:#fff}
.btn-sm{padding:4px 10px;font-size:11px}
.stitle{font-size:10.5px;font-weight:700;color:var(--pr);text-transform:uppercase;
  letter-spacing:.5px;margin:10px 0 6px;padding-bottom:4px;border-bottom:1.5px solid var(--pr)}
#tab-overview > .stitle:first-child{display:none}
.record-form-shell{display:flex;flex-direction:column;gap:12px}
.form-section{background:linear-gradient(180deg,#fcfdff,#f7fafe);border:1px solid var(--bd);border-radius:12px;
  padding:12px 12px 10px;box-shadow:0 2px 8px rgba(15,23,42,.04)}
.form-section-header{display:flex;justify-content:space-between;gap:10px;align-items:flex-start;margin-bottom:10px}
.form-section-title{font-size:10px;font-weight:800;color:var(--pr);text-transform:uppercase;letter-spacing:.45px}
.form-section-sub{font-size:10px;color:var(--mu);line-height:1.45;margin-top:3px}
.form-section-grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:10px 14px;align-items:start}
.form-section-grid .fg{min-width:0}
.form-section-grid .fg label{margin-bottom:4px}
.form-section-grid textarea{min-height:82px;resize:vertical}
.form-section-grid .fg.full textarea{min-height:98px}
body.dark .form-section{background:linear-gradient(180deg,#162132,#101a28);border-color:#2b3c4f;box-shadow:none}
body.dark .stitle,body.dark .form-section-title{color:#93c5fd}
body.dark .form-section-sub,body.dark .fg label{color:#cbd5e1}
body.dark .modal .mbody,body.dark .modal .fg label,body.dark .modal .sitem .nm{color:#e2e8f0}
body.dark .modal .stitle{border-bottom-color:#3b82f6}
body.dark .btn-sc{background:#101a29}
body.dark .btn-sc:hover{background:#223246}
body.dark .fg input,body.dark .fg select,body.dark .fg textarea{background:#0f1a29;border-color:#314558;color:#e2e8f0}
body.dark .fg input::placeholder,body.dark .fg textarea::placeholder{color:#7f93ac}
@media(max-width:768px){
  .modal{width:94vw;max-height:88vh;border-radius:10px}
  .mhdr{padding:10px 14px;font-size:12px}
  .mbody{padding:10px 12px}
  .mfoot{padding:8px 12px;gap:7px;position:sticky;bottom:0}
  .btn{padding:7px 12px;font-size:11px}
  .form-section{padding:9px 10px 8px;border-radius:10px}
  .record-form-shell{gap:8px}
  .form-section-header{margin-bottom:7px}
  .form-section-title{font-size:9px}
  .form-section-sub{font-size:9px;line-height:1.25;margin-top:2px}
  .form-section-grid{gap:8px 10px}
  .form-section-grid .fg label{font-size:9px;margin-bottom:2px}
  .fg input,.fg select,.fg textarea{padding:6px 8px;font-size:11px}
  .form-section-grid textarea{min-height:68px}
  .form-section-grid .fg.full textarea{min-height:78px}
  .stitle{font-size:9px;margin:7px 0 4px;padding-bottom:3px}
  .slist{max-height:150px;padding:3px;gap:2px}
  .sitem{gap:5px;padding:3px 5px;font-size:10px}
  .sitem button{padding:2px 6px;font-size:9px}
  .addrow{gap:5px;margin-top:5px}
  .addrow input{padding:5px 7px;font-size:10px}
}
.badge{display:inline-block;border-radius:10px;padding:2px 9px;font-size:10px;font-weight:700}
#toast{position:fixed;bottom:28px;right:18px;background:var(--pr);color:#fff;
  padding:10px 18px;border-radius:var(--rd);font-size:12px;z-index:9999;
  box-shadow:0 4px 16px rgba(0,0,0,.2);transform:translateY(80px);opacity:0;
  transition:all .3s;pointer-events:none;max-width:320px}
#toast.show{transform:none;opacity:1}
#toast.ok{background:#16a34a}#toast.er{background:#ef4444}#toast.wa{background:#f59e0b;color:#000}
</style>"""

SHARED_JS = """
<div id="toast"></div>
<script>
function applyDarkMode(on){
  document.body.classList.toggle('dark',!!on);
  const btn=document.getElementById('darkBtn');
  if(btn)btn.textContent=on?'☀️':'🌙';
}
function toggleDark(){
  const on=!document.body.classList.contains('dark');
  applyDarkMode(on);
  localStorage.setItem('dcr_dark',on?'1':'');
}
if(localStorage.getItem('dcr_dark'))applyDarkMode(true);
function toast(msg,type=''){
  const t=document.getElementById('toast');
  t.textContent=msg;t.className='show '+(type||'');
  clearTimeout(t._t);t._t=setTimeout(()=>t.className='',3200);
}
function openM(id){document.getElementById(id).classList.remove('hidden')}
function closeM(id){document.getElementById(id).classList.add('hidden')}
function toggleProjectInfo(){
  const bar=document.getElementById('projbar');
  const btn=document.getElementById('projbar-toggle');
  if(!bar||!btn)return;
  const open=bar.classList.toggle('show-details');
  btn.textContent=open?'Hide Info':'Project Info';
}
async function apiFetch(url,opts={}){
  const r=await fetch(url,{credentials:'include',headers:{'Content-Type':'application/json'},...opts});
  if(r.status===403){const d=await r.json().catch(()=>({}));
    if(d.error==='LOGIN_REQUIRED'){window.location='/login';return null;}
    throw new Error(d.error||'Forbidden');}
  if(!r.ok)throw new Error(await r.text());
  return r.json();
}
async function changePw(){
  const pw=prompt('New password (min 4 chars):');
  if(!pw||pw.length<4)return;
  const r=await apiFetch('/api/change_password',{method:'POST',body:JSON.stringify({password:pw})});
  if(r&&r.ok)toast('✔ Password changed','ok');else toast((r&&r.error)||'Error','er');
}
</script>"""


def render_login():
    static_dir = os.path.join(os.path.dirname(__file__), "static")
    logo_name = "logo-login.png"
    logo_file = os.path.join(static_dir, logo_name)
    logo_ver = str(int(os.path.getmtime(logo_file))) if os.path.exists(logo_file) else "1"
    logo_src = url_for("static", filename=logo_name) + f"?v={logo_ver}"
    return f"""<!DOCTYPE html><html lang="en"><head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>DCR — Login</title>
<style>
*{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:'Segoe UI',Arial,sans-serif;height:100vh;display:flex;align-items:center;
  justify-content:center;margin:0;overflow:hidden;background:linear-gradient(135deg,#8BC34A 0%,#5f9f56 24%,#2F4F64 100%)}}
.shell{{width:100%;height:100%;display:flex;align-items:center;justify-content:center;padding:50px 16px 48px}}
.card{{background:rgba(255,255,255,.98);border-radius:16px;box-shadow:0 26px 70px rgba(18,34,46,.28);
  width:100%;max-width:360px;overflow:hidden;backdrop-filter:blur(6px)}}
.brand-band{{height:10px;background:linear-gradient(90deg,#8BC34A,#2F4F64)}}
.cbody{{padding:17px 18px 13px}}
.logo-wrap{{display:flex;justify-content:center;align-items:center;margin-bottom:2px;padding:2px;background:transparent;border-radius:12px}}
.logo-wrap img{{width:min(100%,262px);height:auto;display:block;background:transparent;filter:drop-shadow(0 8px 16px rgba(47,79,100,.1))}}
.title{{font-size:27px;font-weight:800;letter-spacing:.35px;color:#2F4F64;text-align:center;margin-bottom:2px}}
.subtitle{{font-size:13px;color:#6d7b87;text-align:center;margin-bottom:9px}}
.fld{{margin-bottom:7px}}
.fld label{{display:block;font-size:11px;font-weight:700;color:#4f6370;text-transform:uppercase;letter-spacing:.55px;margin-bottom:7px}}
.fld input{{width:100%;padding:13px 15px;border:1.5px solid #d9e2e8;border-radius:11px;font-family:inherit;font-size:14px;
  outline:none;transition:border-color .2s, box-shadow .2s, background .2s;background:#fbfdff;color:#1f2f3b}}
.fld input:focus{{border-color:#8BC34A;box-shadow:0 0 0 4px rgba(139,195,74,.18);background:#fff}}
.err{{background:#fef2f2;border:1px solid #fecaca;color:#b91c1c;padding:10px 12px;border-radius:10px;font-size:12px;margin-bottom:16px;display:none}}
.btn-login{{width:100%;padding:14px 16px;background:#2F4F64;color:#fff;border:none;border-radius:11px;font-family:inherit;
  font-size:14px;font-weight:800;letter-spacing:.25px;cursor:pointer;transition:transform .18s, box-shadow .18s, background .18s}}
.btn-login:hover{{transform:translateY(-1px);box-shadow:0 12px 28px rgba(47,79,100,.26);background:#284355}}
.btn-login:active{{transform:translateY(0)}}
.hint{{text-align:center;color:#7d8a95;font-size:11px;margin-top:7px}}
@media(max-width:480px){{
  .shell{{padding:38px 12px 34px}}
  .cbody{{padding:15px 15px 13px}}
  .title{{font-size:24px}}
  .subtitle{{margin-bottom:9px}}
  .fld{{margin-bottom:7px}}
  .logo-wrap img{{width:min(100%,214px)}}
}}
</style></head><body>
<div class="shell">
  <div class="card">
    <div class="brand-band"></div>
    <div class="cbody">
      <div class="logo-wrap"><img src="{logo_src}" alt="Gas Chill"></div>
      <div class="title">Document Control System</div>
      <div class="subtitle">Secure Project Access</div>
      <div class="err" id="err"></div>
      <div class="fld"><label>Username</label><input id="un" type="text" autofocus autocomplete="username" placeholder="Enter your username"></div>
      <div class="fld"><label>Password</label><input id="pw" type="password" autocomplete="current-password" placeholder="Enter your password"></div>
      <button class="btn-login" onclick="login()">Sign In</button>
      <p class="hint">Gas Chill Document Control Register</p>
    </div>
  </div>
</div>
<script>
document.getElementById('pw').onkeydown=e=>{{if(e.key==='Enter')login()}};
document.getElementById('un').onkeydown=e=>{{if(e.key==='Enter')document.getElementById('pw').focus()}};
async function login(){{
  const un=document.getElementById('un').value.trim();
  const pw=document.getElementById('pw').value;
  const err=document.getElementById('err');
  err.style.display='none';
  if(!un||!pw){{err.textContent='Please enter username and password';err.style.display='block';return;}}
  const r=await fetch('/login',{{method:'POST',credentials:'include',
    headers:{{'Content-Type':'application/json'}},body:JSON.stringify({{username:un,password:pw}})}});
  const d=await r.json();
  if(d.ok) window.location='/';
  else{{err.textContent=d.error||'Invalid credentials';err.style.display='block';}}
}}
</script></body></html>"""




def render_dashboard(u):
    btns, uname, rlbl, rbg = _user_info_html(u)
    role = u["role"] if u else "guest"
    audit_btn = '<button class="tbtn" onclick="showTab(&#39;audit&#39;)">📜 Audit Log</button>' if role in ('superadmin','admin') else ''

    return f"""<!DOCTYPE html><html lang="en"><head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>DCR — Dashboard</title>
{BASE_CSS}
<style>
/* ── Layout ── */
body{{display:flex;flex-direction:column;min-height:100vh;overflow-x:hidden}}
.wrap{{max-width:1480px;margin:0 auto;padding:12px 12px;flex:1;width:100%}}

/* ── Dark Mode ── */
body.dark{{--bg:#0f172a;--wh:#162132;--tx:#e2e8f0;--mu:#9fb0c6;--bd:#304257;--pr:#60a5fa;--pl:#93c5fd}}
body.dark .kpi,body.dark .ccard,body.dark .pcard,body.dark .panel,
body.dark .psel-bar,body.dark .tbl-wrap{{background:#162132;color:#e2e8f0}}
body.dark .kval{{color:#93c5fd}}
body.dark .kpi.ok .kval{{color:#86efac}}
body.dark .kpi.wa .kval{{color:#fcd34d}}
body.dark .kpi.er .kval{{color:#fca5a5}}
body.dark .kpi.pu .kval{{color:#c4b5fd}}
body.dark .klbl,body.dark .ktrend,body.dark .psel-bar label{{color:#b8c8da}}
body.dark .clbl{{color:#93c5fd}}
body.dark .stitle{{color:#93c5fd;border-color:#93c5fd}}
body.dark .dt-tbl th{{background:#1e3a5f;color:#e2e8f0}}
body.dark .dt-tbl td{{border-color:#334155;color:#cbd5e1}}
body.dark .dt-tbl tr:hover td{{background:#142235}}
body.dark .dt-tbl .alt td{{background:#132031}}
body.dark #topbar{{background:#0f2640}}
body.dark .pchdr{{background:#1e3a5f}}
body.dark .pcbody{{color:#dbe7f3}}
body.dark select,body.dark input{{background:#1e293b;color:#e2e8f0;border-color:#334155}}
body.dark .tbtn{{background:#1e293b;color:#e2e8f0;border-color:#334155}}
body.dark .ov-row{{border-color:#334155;color:#cbd5e1}}
body.dark .prog{{background:#334155}}
body.dark .addcard{{background:linear-gradient(180deg,#132031,#0f1a29);border-color:#304257;color:#b8c8da}}
body.dark .addcard:hover{{background:linear-gradient(180deg,#17263a,#122033);border-color:#60a5fa;color:#dbe7f3}}
body.dark .pr-block,body.dark .ltr-summary-card,body.dark .ltr-parties-wrap,body.dark .ltr-party-card{{border-color:#304257!important}}
body.dark .pr-block,body.dark .ltr-summary-card,body.dark .ltr-parties-wrap{{background:#101a29!important}}
body.dark .ltr-party-card{{background:#162132!important}}

/* ── KPIs ── */
.kpi-grid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(120px,1fr));gap:7px;margin-bottom:10px}}
.kpi{{background:var(--wh);border-radius:10px;padding:9px 11px;min-height:82px;
  box-shadow:0 1px 4px rgba(0,0,0,.07);border-left:4px solid var(--pr);cursor:default;
  transition:transform .15s,box-shadow .15s;display:flex;flex-direction:column;justify-content:center}}
.kpi:hover{{transform:translateY(-2px);box-shadow:0 4px 14px rgba(0,0,0,.12)}}
.kpi.ok{{border-left-color:var(--ok)}}.kpi.wa{{border-left-color:var(--wa)}}
.kpi.er{{border-left-color:var(--er)}}.kpi.pu{{border-left-color:#7c3aed}}
.kval{{font-size:28px;font-weight:800;color:var(--pr);line-height:1;margin-bottom:4px}}
.kpi.ok .kval{{color:var(--ok)}}.kpi.wa .kval{{color:var(--wa)}}
.kpi.er .kval{{color:var(--er)}}.kpi.pu .kval{{color:#7c3aed}}
.klbl{{font-size:9px;color:#54657d;font-weight:700;text-transform:uppercase;letter-spacing:.5px;margin-top:0}}
.ktrend{{font-size:10px;color:var(--mu);margin-top:3px}}

/* ── Toolbar ── */
.psel-bar{{display:flex;align-items:center;gap:8px;background:var(--wh);padding:7px 10px;
  border-radius:8px;box-shadow:0 1px 4px rgba(0,0,0,.07);margin-bottom:10px;flex-wrap:wrap}}
.psel-bar label{{font-size:11px;font-weight:700;color:var(--mu);white-space:nowrap}}
.psel-bar select,.psel-bar input{{padding:5px 10px;border:1.5px solid var(--bd);
  border-radius:var(--rd);font-family:inherit;font-size:12px;outline:none}}
.tbtn{{padding:5px 11px;border:1.5px solid var(--bd);border-radius:var(--rd);
  background:var(--wh);cursor:pointer;font-size:11px;font-weight:600;font-family:inherit;
  transition:all .15s;color:var(--tx)}}
.tbtn:hover{{background:var(--pr);color:#fff;border-color:var(--pr)}}
.tbtn.active{{background:var(--pr);color:#fff;border-color:var(--pr)}}

/* ── Charts grid ── */
.charts-grid{{display:grid;grid-template-columns:1fr 1fr;gap:9px;margin-bottom:10px}}
.charts-grid-3{{display:grid;grid-template-columns:1fr 1fr 1fr;gap:9px;margin-bottom:10px}}
.ccard{{background:var(--wh);border-radius:10px;padding:10px 11px;
  box-shadow:0 1px 4px rgba(0,0,0,.07);transition:box-shadow .15s,transform .15s}}
.ccard:hover{{box-shadow:0 4px 14px rgba(0,0,0,.08);transform:translateY(-1px)}}
.clbl{{font-size:10px;font-weight:700;color:var(--pr);text-transform:uppercase;
  letter-spacing:.5px;margin-bottom:6px;display:flex;align-items:center;gap:6px}}
canvas{{max-height:174px}}

/* ── Project cards ── */
.overview-stack{{display:grid;gap:10px;margin-bottom:10px}}
.overview-projects-panel{{padding:11px 12px;margin-bottom:0;min-width:0}}
.overview-projects-panel .panel-title{{margin-bottom:8px}}
.overview-pr-panel{{padding:11px 12px;margin-bottom:0;min-width:0}}
.overview-pr-panel .panel-title{{margin-bottom:8px}}
.overview-wide-panel{{padding:11px 12px;margin-bottom:10px}}
.overview-wide-panel .panel-title{{margin-bottom:8px}}
.pr-analytics-grid{{display:grid;grid-template-columns:minmax(170px,.85fr) minmax(240px,1.18fr) minmax(290px,1.5fr);gap:10px;align-items:stretch}}
.pr-analytics-grid > div{{min-width:0}}
.pr-block,.ltr-summary-card,.ltr-parties-wrap,.ltr-party-card,.charts-grid > *,.charts-grid-3 > *,#pr-panel,#ltr-panel{{min-width:0}}
.ltr-summary-grid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));gap:9px}}
.ltr-parties-grid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:8px}}
.pgrid{{display:grid;grid-template-columns:repeat(auto-fill,minmax(200px,1fr));gap:8px;margin-bottom:0}}
.pcard{{background:var(--wh);border-radius:10px;box-shadow:0 2px 6px rgba(0,0,0,.08);
  overflow:hidden;text-decoration:none;color:inherit;display:block;
  transition:transform .15s,box-shadow .15s;position:relative;min-height:132px}}
.pcard:hover{{transform:translateY(-2px);box-shadow:0 6px 18px rgba(0,0,0,.12)}}
.pchdr{{background:var(--pr);padding:8px 10px;display:flex;align-items:center;justify-content:space-between}}
.pcbody{{padding:7px 10px 8px;display:flex;flex-direction:column;gap:2px}}
.prow{{display:flex;justify-content:space-between;align-items:center;margin-bottom:1px}}
.prog{{height:5px;background:#eef1f7;border-radius:99px;overflow:hidden;margin-top:4px}}
.progf{{height:100%;border-radius:99px;transition:width .6s ease}}
.addcard{{background:linear-gradient(180deg,#fbfcfe,#f5f8fc);border-radius:10px;border:1.5px dashed #c9d4e2;min-height:132px;
  display:flex;align-items:center;justify-content:center;flex-direction:column;gap:6px;
  cursor:pointer;transition:all .15s;color:#71829a;font-size:12px}}
.addcard:hover{{border-color:var(--pr);color:var(--pr);background:linear-gradient(180deg,#ffffff,#f3f8ff)}}

/* ── Tables ── */
.tbl-wrap{{background:var(--wh);border-radius:10px;box-shadow:0 1px 4px rgba(0,0,0,.07);
  overflow-x:auto;overflow-y:hidden;margin-bottom:12px;transition:box-shadow .15s;
  -webkit-overflow-scrolling:touch}}
.tbl-wrap:hover{{box-shadow:0 4px 14px rgba(0,0,0,.07)}}
.dt-tbl{{width:100%;min-width:760px;border-collapse:collapse;font-size:12px}}
.dt-tbl th{{background:#183754;color:#fff;padding:7px 10px;
  text-align:left;font-weight:600;white-space:nowrap;font-size:11px}}
.dt-tbl td{{padding:4px 10px;border-bottom:1px solid #dde5ef;font-variant-numeric:tabular-nums;transition:background .12s ease}}
.dt-tbl tr:hover td{{background:#f1f6fb}}
.dt-tbl .alt td{{background:#fbfcfe}}
.dt-summary-row.warn td{{background:#fff4e8}}
.dt-summary-row.warn:hover td{{background:#feeccc}}
body.dark .dt-summary-row.warn td{{background:#35271e}}
body.dark .dt-summary-row.warn:hover td{{background:#453224}}
#overview-pane-discipline .dt-tbl{{table-layout:fixed}}
#overview-pane-discipline .dt-tbl th:first-child,
#overview-pane-discipline .dt-tbl td:first-child{{white-space:nowrap}}
#overview-pane-discipline .dt-tbl th:nth-child(2),
#overview-pane-discipline .dt-tbl td:nth-child(2){{white-space:nowrap}}
.disc-group-row td{{background:#f9fbfd;font-weight:600}}
.disc-group-row.alt td{{background:#f6f9fc}}
.disc-group-row:hover td{{background:#edf3f8}}
.disc-group-row.warn td{{background:#fff4e8}}
.disc-group-row.warn:hover td{{background:#feeccc}}
.disc-child-row{{display:none}}
.disc-child-row.open{{display:table-row}}
.disc-child-row td{{background:#f7fafe}}
.disc-child-row:hover td{{background:#eef3f8}}
.disc-child-spacer{{width:14px;padding:0 0 0 10px!important}}
.disc-child-label{{padding-left:18px!important;position:relative;font-weight:600;color:#2f4f64}}
.disc-child-label::before{{content:'';position:absolute;left:8px;top:50%;width:6px;height:6px;border-radius:999px;background:#c7d2de;transform:translateY(-50%)}}
.disc-badge{{display:inline-flex;align-items:center;justify-content:center;min-width:26px;padding:2px 8px;border-radius:999px;background:#e8eef6;color:#2f4f64;font-size:10px;font-weight:700}}
.disc-expander{{display:inline-flex;align-items:center;justify-content:center;width:24px;height:24px;border:1px solid #d7e0ea;border-radius:999px;background:#fff;color:#2f4f64;cursor:pointer;font-size:11px;transition:all .15s}}
.disc-expander:hover{{border-color:#2f4f64;background:#f5f8fc}}
.disc-expander.open{{background:#2f4f64;color:#fff;border-color:#2f4f64}}
.disc-meta{{font-size:10px;color:var(--mu);font-weight:600}}
.disc-num-cell{{text-align:center!important;font-weight:700;font-variant-numeric:tabular-nums}}
.disc-mobile-list{{display:none}}
.disc-mobile-card{{border:1px solid var(--bd);border-radius:8px;background:var(--wh);overflow:hidden;margin-bottom:8px}}
.disc-mobile-card.warn{{border-color:#f2c48d;background:#fff8ef}}
.disc-mobile-head{{display:flex;justify-content:space-between;gap:10px;align-items:flex-start;padding:9px 10px;border-bottom:1px solid var(--bd)}}
.disc-mobile-title{{font-weight:800;color:var(--pr);font-size:12px}}
.disc-mobile-sub{{font-size:10px;color:var(--mu);margin-top:2px}}
.disc-mobile-grid{{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:6px 10px;padding:9px 10px}}
.disc-mobile-cell{{min-width:0}}
.disc-mobile-lbl{{font-size:9px;font-weight:700;color:var(--mu);text-transform:uppercase;letter-spacing:.24px}}
.disc-mobile-val{{font-size:13px;font-weight:800;margin-top:1px;font-variant-numeric:tabular-nums;color:var(--tx)}}
.disc-mobile-children{{display:none;border-top:1px solid var(--bd);background:rgba(255,255,255,.55)}}
.disc-mobile-card.open .disc-mobile-children{{display:block}}
.disc-mobile-child{{padding:8px 10px;border-bottom:1px solid var(--bd)}}
.disc-mobile-child:last-child{{border-bottom:none}}
#overview-pane-discipline .disc-group-row td,#overview-pane-discipline .disc-child-row td{{vertical-align:middle}}
#overview-pane-discipline .disc-group-row td:nth-child(3) > div{{display:flex;align-items:center;justify-content:flex-start;gap:8px;min-width:0}}
#overview-pane-discipline .disc-group-row td:nth-child(3),#overview-pane-discipline .disc-child-row td:nth-child(3){{min-width:170px}}
#overview-pane-discipline .disc-num-cell{{min-width:84px}}
body.dark .disc-group-row td{{background:#162132;color:#dbe7f3}}
body.dark .disc-group-row.alt td{{background:#132031}}
body.dark .disc-group-row:hover td{{background:#1a2a3d}}
body.dark .disc-group-row.warn td{{background:#3a2a1f}}
body.dark .disc-group-row.warn:hover td{{background:#4a3425}}
body.dark .disc-child-row td{{background:#101a29}}
body.dark .disc-child-row:hover td{{background:#162132}}
body.dark .disc-badge{{background:#223246;color:#dbe7f3}}
body.dark .disc-expander{{background:#101a29;border-color:#304257;color:#dbe7f3}}
body.dark .disc-expander.open{{background:#3b82f6;border-color:#3b82f6;color:#fff}}
body.dark .disc-meta,body.dark .disc-child-label{{color:#b8c8da}}
body.dark .disc-mobile-card{{background:#162132;border-color:#304257}}
body.dark .disc-mobile-card.warn{{background:#2d241a;border-color:#5b3b20}}
body.dark .disc-mobile-children{{background:#101a29;border-color:#304257}}
body.dark .disc-mobile-child{{border-color:#304257}}
body.dark .disc-mobile-sub,body.dark .disc-mobile-lbl{{color:#b8c8da}}
body.dark .disc-mobile-val{{color:#e2e8f0}}
@media(max-width:640px){{
  .disc-mobile-grid{{grid-template-columns:repeat(4,minmax(0,1fr));gap:5px;padding:8px}}
  .disc-mobile-cell{{display:flex;flex-direction:column;gap:1px;min-width:0;background:rgba(241,245,249,.62);border-radius:6px;padding:5px 6px}}
  .disc-mobile-lbl{{font-size:7.5px;line-height:1;white-space:nowrap}}
  .disc-mobile-val{{font-size:11px;line-height:1.05;margin-top:0;white-space:nowrap}}
  .disc-mobile-title{{font-size:11px}}
  .disc-mobile-sub{{font-size:9px}}
  body.dark .disc-mobile-cell{{background:#101a29}}
}}
.overview-table-switch{{display:flex;gap:6px;align-items:center;flex-wrap:wrap;margin:8px 0 8px}}
.overview-table-btn{{padding:6px 11px;border:1.5px solid var(--bd);border-radius:999px;background:var(--bg);
  color:var(--mu);cursor:pointer;font-size:10px;font-weight:700;letter-spacing:.35px;font-family:inherit;
  transition:all .15s}}
.overview-table-btn:hover{{border-color:#2F4F64;color:#2F4F64;background:#fff}}
.overview-table-btn.active{{background:#2F4F64;color:#fff;border-color:#2F4F64;box-shadow:0 2px 6px rgba(47,79,100,.16)}}
.overview-table-pane{{display:none}}
.overview-table-pane.active{{display:block}}
.overview-table-shell{{max-height:350px;overflow:auto;-webkit-overflow-scrolling:touch;padding-right:2px;scroll-behavior:smooth}}
.overview-table-shell .dt-tbl th{{position:sticky;top:0;z-index:2;box-shadow:0 1px 0 rgba(255,255,255,.08)}}
.overview-table-shell .tbl-wrap{{margin-bottom:0}}
.pr-toggle{{padding:2px 7px;border:1px solid var(--bd);background:#fff;border-radius:3px;cursor:pointer;font-size:11px}}
.pr-toggle:hover{{background:var(--pr);color:#fff;border-color:var(--pr)}}
.pr-items-row{{display:none}}
.pr-items-row.open{{display:table-row}}
.pr-items-row>td{{padding:0!important;border-bottom:1px solid #d9e2ec!important;border-right:none!important;max-width:none!important;background:transparent!important}}
.pr-items-panel{{display:block;width:100%;box-sizing:border-box;padding:16px;background:#f8fafc;border-top:1px solid #d9e2ec;border-bottom:1px solid #d9e2ec;position:relative;z-index:1}}
.pr-items-grid-wrap{{display:block;width:100%;max-width:100%;overflow-x:auto;padding-bottom:2px}}
.pr-items-title{{display:flex;align-items:center;gap:8px;margin-bottom:8px;font-size:11px;font-weight:800;color:var(--pr);text-transform:uppercase;letter-spacing:.45px}}
.pr-items-title:before{{content:"";width:4px;height:16px;border-radius:999px;background:var(--ac);display:inline-block}}
.pr-items-grid{{display:grid;grid-template-columns:minmax(520px,1fr) 120px 100px minmax(240px,.6fr);min-width:1100px;width:100%;border:1px solid #dbe3ed;border-radius:6px;overflow:hidden;background:#fff;font-size:11px}}
.pr-items-grid-head{{background:#173f66;color:#fff;font-weight:800;padding:10px 12px;border-right:1px solid rgba(255,255,255,.18)}}
.pr-items-grid-head:last-child{{border-right:none}}
.pr-items-cell{{padding:10px 12px;border-top:1px solid #edf2f7;border-right:1px solid #edf2f7;background:#fff;color:#10233a;white-space:normal;word-break:normal;line-height:1.35}}
.pr-items-cell:nth-child(4n){{border-right:none}}
.pr-items-cell.qty{{text-align:center;font-variant-numeric:tabular-nums}}
.pr-items-cell.unit{{white-space:nowrap}}
.pr-items-section{{grid-column:1/-1;background:#e8eef6;color:var(--pr);font-weight:800;text-transform:uppercase;letter-spacing:.35px;padding:9px 12px;border-top:1px solid #d7dee8}}
.pr-items-empty{{color:var(--mu);font-size:11px;padding:6px 0}}
body.dark .pr-items-row>td{{background:transparent!important;border-color:#304257!important}}
body.dark .pr-items-panel{{background:#101a29;border-color:#304257}}
body.dark .pr-items-grid{{background:#162132;border-color:#304257}}
body.dark .pr-items-grid-head{{background:#17314d;color:#dbeafe;border-color:#304257}}
body.dark .pr-items-cell{{background:#162132;color:#d7e1ec;border-color:#253648}}
body.dark .pr-items-section{{background:#1e3147;color:#dbeafe;border-color:#304257}}
.pr-items-editor table{{width:100%;border-collapse:collapse;font-size:11px}}
.pr-items-editor th{{background:#e2e8f0;color:#334155;padding:6px 8px;text-align:left;font-weight:700}}
.pr-items-editor td{{padding:6px 8px;border-bottom:1px solid #e5e7eb}}
.pr-items-editor input{{width:100%;padding:6px 8px;border:1.5px solid var(--bd);border-radius:var(--rd);font-family:inherit;font-size:12px}}
.pr-items-editor tr.pr-head-edit td{{background:#f8fafc}}
.pr-items-editor .pr-head-label{{font-size:10px;font-weight:700;color:var(--pr);text-transform:uppercase;letter-spacing:.35px;margin-bottom:4px}}

/* ── Overdue panel ── */
.ov-row{{display:flex;align-items:center;gap:8px;padding:7px 10px;
  border-bottom:1px solid var(--bd);font-size:11px}}
.ov-row:last-child{{border-bottom:none}}
.ov-badge{{background:#fef2f2;color:#ef4444;border-radius:6px;
  padding:2px 8px;font-weight:700;font-size:10px;white-space:nowrap}}
.ov-badge.warn{{background:#fffbeb;color:#f59e0b}}

/* ── Executive Summary panel ── */
.panel{{background:var(--wh);border-radius:10px;padding:12px 14px;
  box-shadow:0 1px 4px rgba(0,0,0,.07);margin-bottom:12px;transition:box-shadow .15s,transform .15s}}
.panel:hover{{box-shadow:0 4px 14px rgba(0,0,0,.08);transform:translateY(-1px)}}
.panel-title{{font-size:11px;font-weight:700;color:var(--pr);
  text-transform:uppercase;letter-spacing:.48px;margin-bottom:8px;
  border-bottom:1.5px solid var(--pr);padding-bottom:5px}}

/* ── Tabs ── */
.tab-bar{{display:flex;gap:4px;margin-bottom:12px;border-bottom:2px solid var(--bd);padding-bottom:0}}
.tab{{padding:8px 16px;cursor:pointer;font-size:12px;font-weight:600;color:var(--mu);
  border-bottom:2px solid transparent;margin-bottom:-2px;transition:all .15s}}
.tab.active{{color:var(--pr);border-bottom-color:var(--pr)}}
.tab:hover:not(.active){{color:var(--tx)}}
.tab-pane{{display:none}}.tab-pane.active{{display:block}}

/* ── Loading spinner ── */
#ld{{position:fixed;inset:0;background:rgba(15,38,64,.92);z-index:500;
  display:flex;align-items:center;justify-content:center;flex-direction:column;gap:14px}}
.spin{{width:40px;height:40px;border:4px solid rgba(255,255,255,.2);border-top-color:#f0a500;
  border-radius:50%;animation:spin .6s linear infinite}}
@keyframes spin{{to{{transform:rotate(360deg)}}}}

/* ── Mobile ── */
@media(max-width:1200px){{
  .wrap{{padding:12px 10px}}
  .pgrid{{grid-template-columns:repeat(auto-fill,minmax(190px,1fr))}}
  .pr-analytics-grid{{grid-template-columns:minmax(160px,.9fr) minmax(220px,1.08fr) minmax(250px,1.32fr)}}
}}
@media(min-width:901px) and (max-width:1080px){{
  .pr-analytics-grid{{grid-template-columns:minmax(160px,.95fr) minmax(210px,1.08fr) minmax(240px,1.28fr)}}
}}
@media(max-width:900px){{
  .wrap{{padding:10px}}
  .overview-stack{{gap:10px}}
  .pr-analytics-grid{{grid-template-columns:1fr}}
  .charts-grid,.charts-grid-3{{grid-template-columns:1fr}}
  .kpi-grid{{grid-template-columns:repeat(3,minmax(0,1fr))}}
  .tab{{padding:7px 10px;font-size:11px}}
  canvas{{max-height:210px}}
  .overview-table-shell{{max-height:320px}}
}}
@media(max-width:640px){{
  .wrap{{padding:8px}}
  .kpi-grid{{grid-template-columns:repeat(2,minmax(0,1fr));gap:6px}}
  .kpi{{min-height:74px}}
  .kval{{font-size:22px}}
  .pgrid{{grid-template-columns:1fr}}
  .overview-projects-panel,.overview-pr-panel,.overview-wide-panel{{padding:10px 10px}}
  .psel-bar{{align-items:stretch;gap:6px}}
  .psel-bar select{{min-width:0;flex:1 1 160px}}
  .overview-table-shell{{max-height:300px}}
  .dt-tbl{{min-width:760px}}
  #overview-pane-discipline .tbl-wrap{{display:none}}
  #overview-pane-discipline .disc-mobile-list{{display:block}}
  .ltr-summary-grid,.ltr-parties-grid{{grid-template-columns:1fr}}
  .pr-block{{padding:10px!important}}
}}
@media(max-width:768px){{
  .dt-tbl{{min-width:600px}}
  .psel-bar{{flex-wrap:wrap;gap:6px}}
  .psel-bar select{{min-width:140px;flex:1}}
}}
@media(max-width:480px){{
  .psel-bar{{flex-direction:column;align-items:stretch}}
  .psel-bar .tbtn{{width:100%;justify-content:center}}
  .kpi-grid{{grid-template-columns:repeat(2,1fr);gap:5px}}
  .charts-grid{{gap:8px}}
  canvas{{max-height:190px}}
}}

/* ── Del button ── */
.del-pbtn{{background:none;border:none;cursor:pointer;color:#ef4444;
  font-size:11px;padding:2px 5px;border-radius:3px}}
.del-pbtn:hover{{background:#fef2f2}}
</style>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
</head><body>

<div id="ld">
  <div class="spin"></div>
  <div style="color:rgba(255,255,255,.7);font-size:13px;font-weight:600">Loading dashboard...</div>
</div>

<div id="topbar">
  <span style="font-size:20px">📋</span>
  <span style="font-weight:700;font-size:14px">Document Control Register</span>
  <div class="sp"></div>
  {btns}
  <button class="tb-btn" onclick="toggleDark()" id="darkBtn" title="Toggle dark mode">🌙</button>
  <span style="color:rgba(255,255,255,.45);padding:0 4px">|</span>
  <span style="color:rgba(255,255,255,.8);font-size:11px">👤 {uname}
    <span style="background:{rbg};border-radius:3px;padding:1px 7px;font-size:9px;font-weight:700">{rlbl}</span>
  </span>
</div>

<div class="wrap">

  <!-- KPIs -->
  <div class="kpi-grid">
    <div class="kpi"><div class="kval" id="kpi-total">—</div><div class="klbl">Total Docs</div></div>
    <div class="kpi ok"><div class="kval" id="kpi-approved">—</div><div class="klbl">Approved</div></div>
    <div class="kpi wa"><div class="kval" id="kpi-pending">—</div><div class="klbl">Under Review</div></div>
    <div class="kpi er"><div class="kval" id="kpi-overdue">—</div><div class="klbl">Overdue</div></div>
    <div class="kpi pu"><div class="kval" id="kpi-rejected">—</div><div class="klbl">Rejected/Revise</div></div>
    <div class="kpi"><div class="kval" id="kpi-pct">—</div><div class="klbl">Completion %</div></div>
    <div class="kpi"><div class="kval" id="kpi-projs">—</div><div class="klbl">Projects</div></div>
  </div>

  <!-- Toolbar -->
  <div class="psel-bar">
    <label>🔍 Project:</label>
    <select id="proj-sel" onchange="filterProject(this.value)">
      <option value="">All Projects</option>
    </select>
    <label>🏗 Discipline:</label>
    <select id="disc-sel" onchange="filterDisc(this.value)">
      <option value="">All Disciplines</option>
    </select>
    <div class="sp" style="flex:1"></div>
    <button class="tbtn" onclick="showTab('overview')">📊 Overview</button>
    <button class="tbtn" onclick="showTab('analytics')">📈 Analytics</button>
    <button class="tbtn" onclick="showTab('overdue')">⚠️ Overdue</button>
    <button class="tbtn" onclick="showTab('executive')">📋 Executive</button>
    {audit_btn}
  </div>

    <!-- TAB: OVERVIEW -->
    <div id="tab-overview" class="tab-pane active">
      <div class="stitle">🗂 Projects</div>
      <div class="overview-stack">
        <div class="panel overview-projects-panel">
          <div class="panel-title">Projects</div>
          <div class="pgrid" id="pgrid"></div>
        </div>
        <div class="panel overview-pr-panel">
          <div class="panel-title">PR Analytics</div>
          <div id="pr-panel" style="display:grid;grid-template-columns:1fr;gap:10px">
            <div style="font-size:11px;color:var(--mu)">Loading...</div>
          </div>
        </div>
      </div>

      <div class="panel overview-wide-panel">
        <div class="panel-title">Letters Overview</div>
        <div id="ltr-panel" style="display:grid;grid-template-columns:1fr;gap:10px">
          <div style="font-size:11px;color:var(--mu)">Loading...</div>
        </div>
      </div>

    <div class="charts-grid">
      <div class="ccard"><div class="clbl">Documents by Project</div><canvas id="cProj"></canvas></div>
      <div class="ccard"><div class="clbl">Status Distribution</div><canvas id="cStatus"></canvas></div>
    </div>

    <div class="panel overview-wide-panel">
      <div class="panel-title">Operational Breakdown</div>
      <div class="overview-table-switch">
        <button type="button" id="ovtab-doc-types" class="overview-table-btn active" onclick="setOverviewTableTab('docTypes')">Document Types Summary</button>
        <button type="button" id="ovtab-discipline" class="overview-table-btn" onclick="setOverviewTableTab('discipline')">Discipline Breakdown</button>
      </div>
      <div class="overview-table-shell">
        <div id="overview-pane-docTypes" class="overview-table-pane active">
          <div class="tbl-wrap">
            <table class="dt-tbl">
              <thead><tr>
                <th>Project</th><th>Code</th><th>Type</th>
                <th style="text-align:center">Total</th>
                <th style="text-align:center">Approved</th>
                <th style="text-align:center">Pending</th>
                <th style="text-align:center">Rejected</th>
                <th style="text-align:center">Overdue</th>
              </tr></thead>
              <tbody id="dt-tbody"></tbody>
            </table>
            <div id="dt-empty" style="text-align:center;padding:24px;color:var(--mu);display:none">No data</div>
          </div>
        </div>
        <div id="overview-pane-discipline" class="overview-table-pane">
          <div class="tbl-wrap">
            <table class="dt-tbl">
              <colgroup>
                <col style="width:120px">
                <col style="width:90px">
                <col>
                <col style="width:88px">
                <col style="width:88px">
                <col style="width:88px">
                <col style="width:88px">
                <col style="width:88px">
                <col style="width:52px">
              </colgroup>
              <thead><tr>
                <th>Project</th><th>Doc Type</th><th>Disciplines</th>
                <th style="text-align:center">Total</th>
                <th style="text-align:center">Approved</th>
                <th style="text-align:center">Pending</th>
                <th style="text-align:center">Rejected</th>
                <th style="text-align:center">Overdue</th>
                <th style="text-align:center;width:42px">View</th>
              </tr></thead>
              <tbody id="disc-tbody"></tbody>
            </table>
            <div id="disc-empty" style="text-align:center;padding:24px;color:var(--mu);display:none">No data</div>
          </div>
          <div id="disc-mobile-list" class="disc-mobile-list"></div>
        </div>
      </div>
    </div>
  </div>

  <!-- TAB: ANALYTICS -->
  <div id="tab-analytics" class="tab-pane">
    <div class="charts-grid">
      <div class="ccard">
        <div class="clbl">📅 Monthly Trend (last 6 months)</div>
        <div style="position:relative;height:200px"><canvas id="cTrend"></canvas></div>
      </div>
      <div class="ccard">
        <div class="clbl">⏳ Aging — Pending Docs (Days Lapsed)</div>
        <div style="position:relative;height:200px"><canvas id="cAging"></canvas></div>
      </div>
    </div>
    <div class="charts-grid">
      <div class="ccard">
        <div class="clbl">🔄 Document Quality (Revisions Needed)</div>
        <div style="position:relative;height:200px"><canvas id="cQuality"></canvas></div>
      </div>
      <div class="ccard">
        <div class="clbl">✅ Approval Rate by Doc Type</div>
        <div style="position:relative;height:200px"><canvas id="cApprRate"></canvas></div>
      </div>
    </div>
  </div>

  <!-- TAB: OVERDUE -->
  <div id="tab-overdue" class="tab-pane">
    <div class="panel">
      <div class="panel-title">⚠️ Overdue Documents — Awaiting Reply</div>
      <div id="ov-filter" style="display:flex;gap:8px;margin-bottom:10px;flex-wrap:wrap">
        <input id="ov-search" placeholder="🔍 Search..." oninput="filterOverdue()"
          style="padding:6px 10px;border:1.5px solid var(--bd);border-radius:var(--rd);font-size:12px;outline:none;flex:1;min-width:160px">
        <select id="ov-sort" onchange="filterOverdue()"
          style="padding:6px 10px;border:1.5px solid var(--bd);border-radius:var(--rd);font-size:12px;outline:none">
          <option value="days">Sort: Most Overdue</option>
          <option value="doc">Sort: Doc No.</option>
          <option value="disc">Sort: Discipline</option>
        </select>
      </div>
      <div id="ov-list" style="max-height:450px;overflow-y:auto"></div>
      <div id="ov-count" style="font-size:11px;color:var(--mu);margin-top:8px;text-align:right"></div>
    </div>
  </div>

  <!-- TAB: EXECUTIVE SUMMARY -->
  <div id="tab-executive" class="tab-pane">
    <div id="exec-content">
      <div style="text-align:center;padding:40px;color:var(--mu)">Loading...</div>
    </div>
  </div>

  <!-- TAB: AUDIT LOG -->
  <div id="tab-audit" class="tab-pane">
    <div class="panel">
      <div class="panel-title">📜 Activity Log — All Changes</div>
      <!-- Filters -->
      <div style="display:flex;gap:8px;margin-bottom:12px;flex-wrap:wrap;align-items:center">
        <select id="aud-pid" onchange="loadAudit()"
          style="padding:6px 10px;border:1.5px solid var(--bd);border-radius:var(--rd);font-size:12px;outline:none;min-width:160px">
          <option value="">All Projects</option>
        </select>
        <select id="aud-user" onchange="loadAudit()"
          style="padding:6px 10px;border:1.5px solid var(--bd);border-radius:var(--rd);font-size:12px;outline:none">
          <option value="">All Users</option>
        </select>
        <select id="aud-action" onchange="loadAudit()"
          style="padding:6px 10px;border:1.5px solid var(--bd);border-radius:var(--rd);font-size:12px;outline:none">
          <option value="">All Actions</option>
        </select>
        <button class="tbtn" onclick="loadAudit(true)" style="margin-left:auto">🔄 Refresh</button>
      </div>
      <!-- Table -->
      <div style="overflow-x:auto">
        <table class="dt-tbl" id="aud-tbl" style="min-width:800px">
          <thead><tr>
            <th style="width:140px">Time</th>
            <th style="width:100px">User</th>
            <th style="width:80px">Action</th>
            <th style="width:100px">Project</th>
            <th style="width:120px">Document</th>
            <th style="width:100px">Field</th>
            <th>Old Value</th>
            <th>New Value</th>
            <th>Detail</th>
          </tr></thead>
          <tbody id="aud-tbody"></tbody>
        </table>
      </div>
      <div id="aud-footer" style="display:flex;justify-content:space-between;align-items:center;margin-top:10px;font-size:11px;color:var(--mu)">
        <span id="aud-count"></span>
        <div style="display:flex;gap:8px">
          <button id="aud-prev" class="tbtn" onclick="auditPage(-1)" disabled>← Prev</button>
          <span id="aud-page" style="padding:5px 10px;font-size:11px">Page 1</span>
          <button id="aud-next" class="tbtn" onclick="auditPage(1)">Next →</button>
        </div>
      </div>
    </div>
  </div>

</div>

<!-- Modals -->
<div class="overlay hidden" id="newproj-modal">
  <div class="modal" style="max-width:420px">
    <div class="mhdr"><span>➕ New Project</span><button class="xbtn" onclick="closeM('newproj-modal')">✕</button></div>
    <div class="mbody">
      <div class="fgrid">
        <div class="fg"><label>Project ID</label><input id="np-id" placeholder="e.g. 24CP01"></div>
        <div class="fg"><label>Short Code</label><input id="np-code" placeholder="e.g. CP01"></div>
        <div class="fg full"><label>Project Name</label><input id="np-name" placeholder="Full project name"></div>
      </div>
    </div>
    <div class="mfoot">
      <button class="btn btn-sc" onclick="closeM('newproj-modal')">Cancel</button>
      <button class="btn btn-pr" id="cpbtn" onclick="createProject()">Create</button>
    </div>
  </div>
</div>

<!-- ADMIN MODAL (dashboard) -->
<div class="overlay hidden" id="admin-modal">
  <div class="modal" style="max-width:780px">
    <div class="mhdr"><span>⚙ Admin Panel</span><button class="xbtn" onclick="closeM('admin-modal')">✕</button></div>
    <div class="mbody" id="admin-body"></div>
    <div class="mfoot"><button class="btn btn-sc" onclick="closeM('admin-modal')">Close</button></div>
  </div>
</div>

{SHARED_JS}
<script>
const ROLE='{role}';
let STATS=[],OVERDUE_DATA=[],EXEC_DATA=null;
let PR_ANALYTICS=null;
let pChart,sChart,tChart,aChart,qChart,arChart,prTradeChart;
let _currentDisc='',_currentPid='';

// ── Tab switching ──────────────────────────────────────────
function showTab(name){{
  document.querySelectorAll('.tab-pane').forEach(p=>p.classList.remove('active'));
  const pane=document.getElementById('tab-'+name);
  if(pane)pane.classList.add('active');
  if(name==='analytics'){{analyticsLoaded=false;setTimeout(()=>loadAnalytics(),50);}}
  if(name==='overdue'){{overdueLoaded=false;loadOverdue();}}
  if(name==='executive'){{EXEC_DATA=null;loadExecutive();}}
  if(name==='audit'){{loadAudit(true);}}
}}

// ── Init ──────────────────────────────────────────────────
  async function init(){{
    try{{
    const [stats]=await Promise.all([
        apiFetch('/api/dashboard_stats'),
      ]);
      STATS=stats;
      if(!STATS)return;
      // Populate project filter
      document.getElementById('proj-sel').innerHTML=
        '<option value="">All Projects</option>'+
      STATS.map(p=>`<option value="${{p.id}}">${{p.name}} (${{p.code}})</option>`).join('');
    // Populate discipline filter
    const discs=new Set();
    STATS.forEach(p=>(p.dt_stats||[]).forEach(dt=>
      (dt.disc_breakdown||[]).forEach(d=>discs.add(d.disc))));
        document.getElementById('disc-sel').innerHTML=
          '<option value="">All Disciplines</option>'+
          [...discs].sort().map(d=>`<option value="${{d}}">${{d}}</option>`).join('');
        await refreshOverviewPanels('');
        renderAll('','');
    }}finally{{document.getElementById('ld').style.display='none';}}
  }}

function filterProject(pid){{
  _currentPid=pid;refreshOverviewPanels(pid);renderAll(pid,_currentDisc);
  // Reload analytics/overdue if their tab is active
  const activeTab=document.querySelector('.tab-pane.active');
  if(activeTab?.id==='tab-analytics'){{analyticsLoaded=false;loadAnalytics();}}
  if(activeTab?.id==='tab-overdue'){{overdueLoaded=false;loadOverdue();}}
  if(activeTab?.id==='tab-executive'){{EXEC_DATA=null;loadExecutive();}}
}}
function filterDisc(disc){{_currentDisc=disc;renderAll(_currentPid,disc);}}

async function refreshOverviewPanels(pid){{
  const qs=pid?`?project_id=${{encodeURIComponent(pid)}}`:'';
  PR_ANALYTICS=await apiFetch('/api/pr_analytics_summary'+qs).catch(()=>null);
  renderPrAnalytics(PR_ANALYTICS);
}}

function getFiltered(pid,disc){{
  let d=pid?STATS.filter(s=>s.id===pid):STATS;
  if(disc){{
    d=d.map(p=>{{
      const dtFiltered=(p.dt_stats||[]).map(dt=>{{
        const discF=(dt.disc_breakdown||[]).filter(ds=>ds.disc===disc);
        const t=discF.reduce((s,r)=>s+r.total,0);
        const ap=discF.reduce((s,r)=>s+r.approved,0);
        const pe=discF.reduce((s,r)=>s+r.pending,0);
        const rj=discF.reduce((s,r)=>s+r.rejected,0);
        const ov=discF.reduce((s,r)=>s+r.overdue,0);
        return {{...dt,total:t,approved:ap,pending:pe,rejected:rj,overdue:ov,disc_breakdown:discF}};
      }}).filter(dt=>dt.total>0);
      const t=dtFiltered.reduce((s,dt)=>s+dt.total,0);
      const ap=dtFiltered.reduce((s,dt)=>s+dt.approved,0);
      return {{...p,dt_stats:dtFiltered,total:t,approved:ap,
        pct:t?Math.round(ap/t*100):0}};
    }}).filter(p=>p.total>0);
  }}
  return d;
}}

function renderAll(pid,disc){{
  const d=getFiltered(pid,disc);
  const lettersScope=pid?STATS.filter(s=>s.id===pid):STATS;
  updateKPIs(d);renderCards(d,pid);renderLettersOverview(lettersScope);renderCharts(d);renderDTTable(d);renderDiscTable(d);
  ensureOverviewTableControls();setOverviewTableTab(window._overviewTableTab||'docTypes');
}}

// ── KPIs ──────────────────────────────────────────────────

function renderPrAnalytics(data){{
  const el=document.getElementById('pr-panel');
  if(!el)return;
  if(!data){{
    el.innerHTML='<div style="font-size:11px;color:var(--mu)">Unavailable</div>';
    if(prTradeChart){{prTradeChart.destroy();prTradeChart=null;}}
    return;
  }}
  const topProjects=(data.top_projects||[]).length
    ? (data.top_projects||[]).map((p,idx)=>`<div style="display:flex;justify-content:space-between;gap:10px;padding:${{idx===0?'7px 0 8px':'6px 0'}};border-bottom:1px solid var(--bd);align-items:center">
        <div style="min-width:0">
          <div style="font-size:${{idx===0?'11.5px':'11px'}};font-weight:${{idx===0?'800':'600'}};color:${{idx===0?'var(--pr)':'var(--tx)'}};white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${{p.project_name}}</div>
          ${{idx===0?'<div style="font-size:9px;color:var(--mu);margin-top:2px;text-transform:uppercase;letter-spacing:.28px">Most active project</div>':''}}
        </div>
        <span style="font-size:${{idx===0?'12px':'11px'}};font-weight:800;color:var(--pr);white-space:nowrap">${{p.pr_count}}</span>
      </div>`).join('')
    : `<div style="font-size:11px;color:var(--mu)">No project data</div>`;
  el.innerHTML=`<div class="pr-analytics-grid">
    <div class="pr-block pr-total" style="background:var(--bg);border-radius:8px;padding:10px 12px;border:1px solid #e3eaf2;display:flex;flex-direction:column;justify-content:center">
      <div style="font-size:10px;font-weight:700;color:var(--tx);text-transform:uppercase;letter-spacing:.4px">Total PRs</div>
      <div style="font-size:24px;font-weight:800;color:var(--pr);line-height:1.05;margin-top:4px">${{data.total_pr_records||0}}</div>
      <div style="font-size:9px;color:var(--mu);margin-top:6px;text-transform:uppercase;letter-spacing:.28px">Current dashboard scope</div>
      <div style="font-size:10px;color:var(--mu);margin-top:3px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">Top project: ${{data.top_project_name||'No project data'}}</div>
    </div>
    <div class="pr-block pr-projects" style="background:var(--bg);border-radius:8px;padding:10px 12px;border:1px solid #e3eaf2;display:flex;flex-direction:column;min-height:156px">
      <div style="display:flex;justify-content:space-between;gap:10px;align-items:flex-start;margin-bottom:8px">
        <div style="font-size:10px;font-weight:700;color:var(--tx);text-transform:uppercase;letter-spacing:.4px">Most Active Projects</div>
        <div style="font-size:10px;color:var(--mu);white-space:nowrap">${{(data.top_projects||[]).length}} shown</div>
      </div>
      <div style="flex:1;overflow:auto;padding-right:2px">
        ${{topProjects}}
      </div>
    </div>
    <div class="pr-block pr-trades" style="background:var(--bg);border-radius:8px;padding:10px 12px;border:1px solid #e3eaf2;display:flex;flex-direction:column;min-height:156px">
      <div style="display:flex;justify-content:space-between;gap:10px;align-items:flex-start;margin-bottom:8px">
        <div>
          <div style="font-size:10px;font-weight:700;color:var(--tx);text-transform:uppercase;letter-spacing:.4px">Most Active Trades</div>
          <div style="font-size:15px;font-weight:800;color:var(--pr);line-height:1.1;margin-top:5px">${{data.top_trade_name||'No trade data'}}</div>
          <div style="font-size:9px;color:var(--mu);margin-top:2px;text-transform:uppercase;letter-spacing:.28px">Most active trade</div>
        </div>
        <div style="text-align:right">
          <div style="font-size:15px;font-weight:800;color:var(--pr);line-height:1">${{data.top_trade_count||0}}</div>
          <div style="font-size:10px;color:var(--mu);margin-top:3px">${{data.trade_count_total||0}} trades</div>
        </div>
      </div>
      <div style="flex:1;min-height:150px;max-height:160px"><canvas id="pr-trades-chart"></canvas></div>
    </div>
  </div>`;
  const tradeCanvas=document.getElementById('pr-trades-chart');
  if(prTradeChart){{prTradeChart.destroy();prTradeChart=null;}}
  if(!tradeCanvas)return;
  const trades=data.top_trades||[];
  const dark=document.body.classList.contains('dark');
  prTradeChart=new Chart(tradeCanvas,{{
    type:'bar',
    data:{{
      labels:trades.map(t=>t.trade),
      datasets:[{{
        label:'PR Count',
        data:trades.map(t=>t.pr_count),
        backgroundColor:'#2563a8',
        borderRadius:4
      }}]
    }},
    options:{{
      responsive:true,
      maintainAspectRatio:false,
      indexAxis:'y',
      plugins:{{legend:{{display:false}}}},
      scales:{{
        x:{{beginAtZero:true,grid:{{color:dark?'rgba(159,176,198,.16)':'rgba(0,0,0,.06)'}},ticks:{{precision:0,font:{{size:10}},color:dark?'#cbd5e1':'#54657d'}}}},
        y:{{grid:{{display:false}},ticks:{{font:{{size:10}},color:dark?'#dbe7f3':'#334155'}}}}
      }}
    }}
  }});
}}

function updateKPIs(d){{
  const t=d.reduce((s,p)=>s+p.total,0),
    ap=d.reduce((s,p)=>s+p.approved,0),
    pe=d.reduce((s,p)=>s+p.pending,0),
    ov=d.reduce((s,p)=>s+p.overdue,0),
    rj=d.reduce((s,p)=>s+(p.rejected||0),0);
  document.getElementById('kpi-total').textContent=t;
  document.getElementById('kpi-approved').textContent=ap;
  document.getElementById('kpi-pending').textContent=pe;
  document.getElementById('kpi-overdue').textContent=ov;
  const rjEl=document.getElementById('kpi-rejected');
  if(rjEl)rjEl.textContent=rj;
  document.getElementById('kpi-pct').textContent=(t?Math.round(ap/t*100):0)+'%';
  document.getElementById('kpi-projs').textContent=d.length;
}}

// ── Project Cards ─────────────────────────────────────────
function renderCards(d,pid){{
  const g=document.getElementById('pgrid');g.innerHTML='';
  d.forEach(p=>{{
    const col=p.pct>=80?'#16a34a':p.pct>=50?'#f59e0b':'#ef4444';
    const a=document.createElement('a');a.href='/app?p='+p.id;a.className='pcard';
    a.innerHTML=`<div class="pchdr">
      <div><div style="color:rgba(255,255,255,.6);font-size:10px;font-weight:700">${{p.code}}</div>
        <div style="color:#fff;font-weight:700;font-size:13px">${{p.name}}</div></div>
      ${{p.overdue>0?`<span style="background:#ef4444;color:#fff;border-radius:10px;padding:1px 7px;font-size:10px;font-weight:700">⚠${{p.overdue}}</span>`:''}}</div>
    <div class="pcbody">
      ${{p.client?`<div style="font-size:10px;color:var(--mu);margin-bottom:5px">👤 ${{p.client}}</div>`:''}}<div class="prow">
        <span style="font-size:11px;color:var(--mu)">Total</span><b>${{p.total}}</b></div>
      <div class="prow"><span style="font-size:11px;color:var(--mu)">Approved</span>
        <b style="color:#16a34a">${{p.approved}}</b></div>
      <div class="prow"><span style="font-size:11px;color:var(--mu)">Pending</span>
        <b style="color:#f59e0b">${{p.pending}}</b></div>
      <div class="prog"><div class="progf" style="width:${{p.pct}}%;background:${{col}}"></div></div>
      <div style="font-size:10px;color:${{col}};font-weight:700;margin-top:3px">${{p.pct}}%</div>
      ${{ROLE==='superadmin'?`<button class="del-pbtn" onclick="event.preventDefault();delProject('${{p.id}}','${{p.name}}')">🗑</button>`:''}}
    </div>`;
    g.appendChild(a);
  }});
  if(ROLE==='superadmin'&&!pid){{
    const add=document.createElement('div');add.className='addcard';
    add.innerHTML='<span style="font-size:28px">➕</span><span>New Project</span>';
    add.onclick=()=>openM('newproj-modal');g.appendChild(add);
  }}
}}

// ── Overview Charts ───────────────────────────────────────
function renderCharts(d){{
  const dark=document.body.classList.contains('dark');
  const gridColor=dark?'rgba(159,176,198,.16)':'rgba(0,0,0,.05)';
  const tickColor=dark?'#cbd5e1':'#54657d';
  const legendColor=dark?'#dbe7f3':'#334155';
  if(pChart)pChart.destroy();
  pChart=new Chart(document.getElementById('cProj'),{{type:'bar',
    data:{{labels:d.map(s=>s.code),datasets:[
      {{label:'Approved',data:d.map(s=>s.approved),backgroundColor:'#16a34a',borderRadius:4}},
      {{label:'Pending', data:d.map(s=>s.pending), backgroundColor:'#f59e0b',borderRadius:4}},
      {{label:'Overdue', data:d.map(s=>s.overdue), backgroundColor:'#ef4444',borderRadius:4}}]}},
    options:{{responsive:true,plugins:{{legend:{{position:'bottom',labels:{{boxWidth:10,font:{{size:10}},color:legendColor}}}}}},
      scales:{{y:{{beginAtZero:true,grid:{{color:gridColor}},ticks:{{color:tickColor}}}},
        x:{{grid:{{display:false}},ticks:{{color:tickColor}}}}}}}}}});
  const t=d.reduce((s,p)=>s+p.total,0),ap=d.reduce((s,p)=>s+p.approved,0),
    pe=d.reduce((s,p)=>s+p.pending,0),ov=d.reduce((s,p)=>s+p.overdue,0),
    rj=d.reduce((s,p)=>s+(p.rejected||0),0);
  if(sChart)sChart.destroy();
  sChart=new Chart(document.getElementById('cStatus'),{{type:'doughnut',
    data:{{labels:['Approved','Pending','Rejected','Overdue'],
      datasets:[{{data:[ap,pe,rj,ov],
        backgroundColor:['#16a34a','#f59e0b','#7c3aed','#ef4444'],
        borderWidth:3,borderColor:dark?'#162132':'#fff',hoverOffset:6}}]}},
    options:{{responsive:true,cutout:'65%',
      plugins:{{legend:{{position:'bottom',labels:{{boxWidth:10,font:{{size:10}},color:legendColor}}}},
        tooltip:{{callbacks:{{label:ctx=>` ${{ctx.label}}: ${{ctx.raw}} (${{t?Math.round(ctx.raw/t*100):0}}%)`}}}}}}}}}});
}}

// ── DT Table ──────────────────────────────────────────────
function renderLettersOverview(d){{
  const el=document.getElementById('ltr-panel');
  if(!el)return;
  try{{
    const safeHtml=v=>String(v??'').replace(/[&<>"']/g,m=>({{'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}}[m]));
    const partyMap={{}};
    const stats=(Array.isArray(d)?d:[]).reduce((acc,p)=>{{
      const l=(p&&typeof p==='object'&&p.ltr&&typeof p.ltr==='object')?p.ltr:{{}};
      acc.total+=Number(l.total||0);
      acc.sent+=Number(l.sent||0);
      acc.received+=Number(l.received||0);
      if(Number(l.total||0)>0)acc.project_count+=1;
      (Array.isArray(l.party_stats)?l.party_stats:[]).forEach(row=>{{
        const party=String((row&&row.party)||'').trim();
        if(!party)return;
        if(!partyMap[party])partyMap[party]={{party,total:0,sent:0,received:0}};
        partyMap[party].total+=Number((row&&row.total)||0);
        partyMap[party].sent+=Number((row&&row.sent)||0);
        partyMap[party].received+=Number((row&&row.received)||0);
      }});
      return acc;
    }},{{total:0,sent:0,received:0,project_count:0}});
    if(!stats.total){{
      el.innerHTML='<div style="font-size:11px;color:var(--mu)">No letter records in the current project scope.</div>';
      return;
    }}
    const topParties=Object.values(partyMap)
      .sort((a,b)=>(b.total-a.total)||(b.sent-a.sent)||(b.received-a.received)||a.party.localeCompare(b.party))
      .slice(0,6);
    const cards=[
      ['Total Letters',stats.total,'#2F4F64','Across '+stats.project_count+' project'+(stats.project_count===1?'':'s')],
      ['Sent',stats.sent,'#2563a8','Outgoing correspondence'],
      ['Received',stats.received,'#16a34a','Incoming correspondence'],
    ];
    const summaryHtml=`<div class="ltr-summary-grid">
      ${{cards.map(([label,value,color,hint])=>`<div class="ltr-summary-card" style="background:var(--bg);border-radius:8px;padding:11px 13px;border:1px solid #e3eaf2;border-left:4px solid ${{color}};box-shadow:0 1px 2px rgba(0,0,0,.03)">
        <div style="font-size:10px;font-weight:700;color:var(--tx);text-transform:uppercase;letter-spacing:.45px">${{label}}</div>
        <div style="font-size:26px;font-weight:800;color:${{color}};line-height:1.08;margin-top:5px">${{value}}</div>
        <div style="font-size:10px;color:var(--mu);margin-top:3px">${{hint}}</div>
      </div>`).join('')}}
    </div>`;
    const partiesHtml=topParties.length
      ? `<div class="ltr-parties-wrap" style="background:var(--bg);border-radius:8px;padding:9px 11px;border:1px solid #e3eaf2">
          <div style="display:flex;justify-content:space-between;align-items:center;gap:10px;margin-bottom:8px">
            <div style="font-size:10px;font-weight:700;color:var(--tx);text-transform:uppercase;letter-spacing:.45px">Most Active Parties</div>
            <div style="font-size:10px;color:var(--mu);font-weight:700">Top ${{topParties.length}} Parties</div>
          </div>
          <div class="ltr-parties-grid">
            ${{topParties.map((row, idx)=>{{
              const pct=stats.total?Math.round((Number(row.total||0)/stats.total)*100):0;
              const accent=idx<3?'rgba(47,79,100,.12)':'rgba(221,227,237,.95)';
              const bg=idx<3?'linear-gradient(180deg,rgba(255,255,255,.98),rgba(245,249,255,.94))':'rgba(255,255,255,.82)';
              const shadow=idx<3?'0 3px 10px rgba(47,79,100,.06)':'0 1px 3px rgba(15,23,42,.04)';
              return `<div class="ltr-party-card" style="border:1px solid ${{accent}};border-radius:9px;background:${{bg}};padding:9px 10px;box-shadow:${{shadow}};transition:background .15s,border-color .15s,transform .15s,box-shadow .15s"
                onmouseenter="this.style.background='linear-gradient(180deg,rgba(255,255,255,1),rgba(244,248,253,.98))';this.style.borderColor='#c7d2de';this.style.transform='translateY(-1px)';this.style.boxShadow='0 4px 12px rgba(47,79,100,.08)'"
                onmouseleave="this.style.background='${{bg}}';this.style.borderColor='${{accent}}';this.style.transform='translateY(0)';this.style.boxShadow='${{shadow}}'">
                <div style="display:flex;justify-content:space-between;gap:10px;align-items:flex-start">
                  <div style="min-width:0">
                    <div style="font-size:12px;font-weight:800;color:var(--tx);white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${{safeHtml(row.party)}}</div>
                    <div style="font-size:9px;color:var(--mu);margin-top:2px">Share of letters: ${{pct}}%</div>
                  </div>
                  <div style="text-align:right;flex-shrink:0">
                    <div style="font-size:14px;font-weight:800;color:#2F4F64;line-height:1">${{row.total}}</div>
                    <div style="font-size:9px;color:var(--mu);margin-top:2px">Total</div>
                  </div>
                </div>
                <div style="display:grid;grid-template-columns:1fr 1fr;gap:7px;margin-top:8px">
                  <div style="background:#eef4fb;border-radius:7px;padding:6px 8px;border:1px solid #dbe8f7">
                    <div style="font-size:10px;font-weight:800;color:#2563a8">Sent: ${{row.sent}}</div>
                    <div style="font-size:9px;color:#2563a8;letter-spacing:.2px;margin-top:2px">Outgoing</div>
                  </div>
                  <div style="background:#edf9f0;border-radius:7px;padding:6px 8px;border:1px solid #d7efdc">
                    <div style="font-size:10px;font-weight:800;color:#16a34a">Received: ${{row.received}}</div>
                    <div style="font-size:9px;color:#16a34a;letter-spacing:.2px;margin-top:2px">Incoming</div>
                  </div>
                </div>
              </div>`;
            }}).join('')}}
          </div>
        </div>`
      : '<div style="font-size:11px;color:var(--mu)">No party activity available in the current project scope.</div>';
    el.innerHTML=`<div style="display:grid;gap:10px">
      ${{summaryHtml}}
      ${{partiesHtml}}
    </div>`;
  }}catch(err){{
    console.error('renderLettersOverview failed', err);
    el.innerHTML='<div style="font-size:11px;color:var(--mu)">Letters overview is temporarily unavailable.</div>';
  }}
}}

function ensureOverviewTableControls(){{
  const overview=document.getElementById('tab-overview');
  if(!overview)return;
  const shell=overview.querySelector('.overview-table-shell');
  const docPane=document.getElementById('overview-pane-docTypes');
  const discPane=document.getElementById('overview-pane-discipline');
  if(!shell||!docPane||!discPane)return;
}}

function setOverviewTableTab(tab){{
  window._overviewTableTab=(tab==='discipline')?'discipline':'docTypes';
  const docPane=document.getElementById('overview-pane-docTypes');
  const discPane=document.getElementById('overview-pane-discipline');
  const docBtn=document.getElementById('ovtab-doc-types');
  const discBtn=document.getElementById('ovtab-discipline');
  if(docPane)docPane.classList.toggle('active',window._overviewTableTab==='docTypes');
  if(discPane)discPane.classList.toggle('active',window._overviewTableTab==='discipline');
  if(docBtn)docBtn.classList.toggle('active',window._overviewTableTab==='docTypes');
  if(discBtn)discBtn.classList.toggle('active',window._overviewTableTab==='discipline');
}}

function renderDTTable(d){{
  const tbody=document.getElementById('dt-tbody'),empty=document.getElementById('dt-empty');
  tbody.innerHTML='';let rows=[],i=0;
  d.forEach(p=>{{(p.dt_stats||[]).forEach(dt=>{{rows.push({{...dt,pcode:p.code,pid:p.id}});}});}});
  if(!rows.length){{empty.style.display='block';return;}}
  empty.style.display='none';
  rows.forEach(r=>{{
    const pct=r.total?Math.round(r.approved/r.total*100):0;
    const col=pct>=80?'#16a34a':pct>=50?'#f59e0b':'#ef4444';
    const tr=document.createElement('tr');
    tr.className=`dt-summary-row${{r.overdue>0?' warn':''}}${{i%2?' alt':''}}`;
    i++;
    tr.innerHTML=`<td style="font-size:10px;color:var(--mu)">${{r.pcode}}</td>
      <td style="font-weight:700;color:var(--pr)">${{r.code}}</td>
      <td>${{r.name}}</td>
      <td style="text-align:center;font-weight:700">${{r.total}}</td>
      <td style="text-align:center;color:#16a34a;font-weight:700">${{r.approved}}</td>
      <td style="text-align:center;color:#f59e0b;font-weight:700">${{r.pending}}</td>
      <td style="text-align:center;color:#7c3aed;font-weight:700">${{r.rejected||0}}</td>
      <td style="text-align:center;color:#ef4444;font-weight:700">${{r.overdue}}</td>`;
    tbody.appendChild(tr);
  }});
}}

// ── Disc Table ────────────────────────────────────────────
function renderDiscTable(data){{
  const tbody=document.getElementById('disc-tbody'),empty=document.getElementById('disc-empty');
  const mobile=document.getElementById('disc-mobile-list');
  tbody.innerHTML='';let rows=0,groupIdx=0;
  if(mobile)mobile.innerHTML='';
  const mk=(label,v,c)=>`<td class="disc-num-cell" data-label="${{label}}" style="color:${{c}}">${{v||0}}</td>`;
  const safe=v=>String(v??'').replace(/[&<>"']/g,m=>({{'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}}[m]));
  const metric=(label,value,color)=>`<div class="disc-mobile-cell"><div class="disc-mobile-lbl">${{label}}</div><div class="disc-mobile-val" style="color:${{color||'var(--tx)'}}">${{value||0}}</div></div>`;
  data.forEach(p=>{{(p.dt_stats||[]).forEach(dt=>{{
    const disc=dt.disc_breakdown||[];if(!disc.length)return;
    rows++;groupIdx++;
    const groupId=`discgrp-${{groupIdx}}`;
    const warn=(Number(dt.overdue||0)>0)||disc.some(ds=>Number(ds.overdue||0)>0);
    if(mobile){{
      const card=document.createElement('div');
      card.className=`disc-mobile-card${{warn?' warn':''}}`;
      card.dataset.group=groupId;
      card.innerHTML=`<div class="disc-mobile-head">
        <div>
          <div class="disc-mobile-title">${{safe(p.code)}} | ${{safe(dt.code)}}</div>
          <div class="disc-mobile-sub">${{disc.length}} ${{disc.length===1?'discipline':'disciplines'}}</div>
        </div>
        <button type="button" class="disc-expander" aria-expanded="false" onclick="toggleDiscMobileGroup('${{groupId}}', this)">â–¼</button>
      </div>
      <div class="disc-mobile-grid">
        ${{metric('Project',safe(p.code),'var(--pr)')}}
        ${{metric('Doc Type',safe(dt.code),'var(--pr)')}}
        ${{metric('Disciplines',disc.length,'#2f4f64')}}
        ${{metric('Total',dt.total,'var(--pr)')}}
        ${{metric('Approved',dt.approved,'#16a34a')}}
        ${{metric('Pending',dt.pending,'#f59e0b')}}
        ${{metric('Rejected',dt.rejected||0,'#7c3aed')}}
        ${{metric('Overdue',dt.overdue,'#ef4444')}}
      </div>
      <div class="disc-mobile-children">
        ${{disc.map(ds=>`<div class="disc-mobile-child">
          <div class="disc-mobile-title">${{safe(ds.disc)}}</div>
          <div class="disc-mobile-grid">
            ${{metric('Total',ds.total,'var(--pr)')}}
            ${{metric('Approved',ds.approved,'#16a34a')}}
            ${{metric('Pending',ds.pending,'#f59e0b')}}
            ${{metric('Rejected',ds.rejected||0,'#7c3aed')}}
            ${{metric('Overdue',ds.overdue,'#ef4444')}}
          </div>
        </div>`).join('')}}
      </div>`;
      mobile.appendChild(card);
    }}
    const tr=document.createElement('tr');
    tr.className=`disc-group-row${{warn?' warn':''}}${{rows%2===0?' alt':''}}`;
    tr.innerHTML=`<td data-label="Project" style="font-size:10px;color:var(--mu)">${{p.code}}</td>
      <td data-label="Doc Type" style="font-weight:700;color:var(--pr)">${{dt.code}}</td>
      <td data-label="Disciplines"><div style="display:flex;align-items:center;gap:8px"><span class="disc-badge">${{disc.length}}</span><span class="disc-meta">${{disc.length===1?'1 discipline':`${{disc.length}} disciplines`}}</span></div></td>
      ${{mk('Total',dt.total,'var(--pr)')}}
      ${{mk('Approved',dt.approved,'#16a34a')}}
      ${{mk('Pending',dt.pending,'#f59e0b')}}
      ${{mk('Rejected',dt.rejected||0,'#7c3aed')}}
      ${{mk('Overdue',dt.overdue,'#ef4444')}}
      <td style="text-align:center"><button type="button" class="disc-expander" data-group="${{groupId}}" aria-expanded="false" onclick="toggleDiscGroup('${{groupId}}', this)">▼</button></td>`;
    tbody.appendChild(tr);
    disc.forEach(ds=>{{
      const child=document.createElement('tr');
      child.className='disc-child-row';
      child.dataset.group=groupId;
      child.innerHTML=`<td class="disc-child-spacer"></td>
        <td style="font-size:10px;color:var(--mu)">${{dt.code}}</td>
        <td class="disc-child-label">${{ds.disc}}</td>
        ${{mk('Total',ds.total,'var(--pr)')}}
        ${{mk('Approved',ds.approved,'#16a34a')}}
        ${{mk('Pending',ds.pending,'#f59e0b')}}
        ${{mk('Rejected',ds.rejected||0,'#7c3aed')}}
        ${{mk('Overdue',ds.overdue,'#ef4444')}}
        <td></td>`;
      tbody.appendChild(child);
    }});
  }});}});
  [...tbody.querySelectorAll('.disc-group-row')].forEach(tr=>{{
    const cells=tr.children;
    if(cells[0])cells[0].dataset.label='Project';
    if(cells[1])cells[1].dataset.label='Doc Type';
    if(cells[2])cells[2].dataset.label='Disciplines';
    if(cells[8])cells[8].dataset.label='View';
  }});
  [...tbody.querySelectorAll('.disc-child-row')].forEach(tr=>{{
    const cells=tr.children;
    if(cells[0])cells[0].dataset.label='';
    if(cells[1])cells[1].dataset.label='Doc Type';
    if(cells[2])cells[2].dataset.label='Discipline';
    if(cells[8])cells[8].dataset.label='';
  }});
  if(empty)empty.style.display=rows?'none':'block';
}}

function toggleDiscGroup(groupId, btn){{
  const rows=document.querySelectorAll(`#disc-tbody tr[data-group="${{groupId}}"]`);
  const shouldOpen=!btn.classList.contains('open');
  rows.forEach(row=>row.classList.toggle('open', shouldOpen));
  btn.classList.toggle('open', shouldOpen);
  btn.setAttribute('aria-expanded', shouldOpen ? 'true' : 'false');
  btn.textContent=shouldOpen ? '▲' : '▼';
}}

// ── Analytics Tab ─────────────────────────────────────────
function toggleDiscMobileGroup(groupId, btn){{
  const card=document.querySelector(`#disc-mobile-list .disc-mobile-card[data-group="${{groupId}}"]`);
  if(!card)return;
  const shouldOpen=!card.classList.contains('open');
  card.classList.toggle('open', shouldOpen);
  btn.classList.toggle('open', shouldOpen);
  btn.setAttribute('aria-expanded', shouldOpen ? 'true' : 'false');
  btn.textContent=shouldOpen ? 'â–²' : 'â–¼';
}}

let analyticsLoaded=false;
async function loadAnalytics(forceReload=false){{
  if(analyticsLoaded&&!forceReload)return;
  analyticsLoaded=true;
  const pid=_currentPid;
  const [trend,aging,quality]=await Promise.all([
    apiFetch('/api/analytics/trend'+(pid?`?pid=${{pid}}`:'') ),
    apiFetch('/api/analytics/aging'+(pid?`?pid=${{pid}}`:'') ),
    apiFetch('/api/analytics/quality'+(pid?`?pid=${{pid}}`:'') ),
  ]);

  // Monthly Trend
  if(tChart)tChart.destroy();
  tChart=new Chart(document.getElementById('cTrend'),{{type:'bar',
    data:{{labels:trend.map(t=>t.month),datasets:[
      {{label:'Submitted',data:trend.map(t=>t.submitted),backgroundColor:'rgba(37,99,168,.25)',
        borderColor:'#2563a8',borderWidth:2,borderRadius:4,type:'bar'}},
      {{label:'Approved',data:trend.map(t=>t.approved),backgroundColor:'#16a34a',
        borderColor:'#16a34a',borderWidth:2,borderRadius:4,type:'bar'}}]}},
    options:{{responsive:true,plugins:{{legend:{{position:'bottom',labels:{{boxWidth:10,font:{{size:10}}}}}}}},
      scales:{{y:{{beginAtZero:true}},x:{{grid:{{display:false}}}}}}}}}});

  // Aging
  if(aChart)aChart.destroy();
  aChart=new Chart(document.getElementById('cAging'),{{type:'bar',
    data:{{labels:aging.map(a=>a.range+' Days'),
      datasets:[{{label:'Pending Docs',data:aging.map(a=>a.count),
        backgroundColor:['#16a34a','#f59e0b','#f97316','#ef4444'],borderRadius:6}}]}},
    options:{{responsive:true,plugins:{{legend:{{display:false}},
      tooltip:{{callbacks:{{label:ctx=>`${{ctx.raw}} documents`}}}}}},
      scales:{{y:{{beginAtZero:true}},x:{{grid:{{display:false}}}}}}}}}});

  // Quality
  if(qChart)qChart.destroy();
  qChart=new Chart(document.getElementById('cQuality'),{{type:'doughnut',
    data:{{labels:quality.map(q=>'Rev '+q.revisions),
      datasets:[{{data:quality.map(q=>q.count),
        backgroundColor:['#16a34a','#f59e0b','#f97316','#ef4444'],
        borderWidth:3,borderColor:'#fff',hoverOffset:6}}]}},
    options:{{responsive:true,cutout:'60%',
      plugins:{{legend:{{position:'bottom',labels:{{boxWidth:10,font:{{size:10}}}}}}}}}}}});

  // Approval Rate per Doc Type
  const d=getFiltered(_currentPid,_currentDisc);
  const dtLabels=[],dtRates=[];
  d.forEach(p=>(p.dt_stats||[]).forEach(dt=>{{
    if(dt.total>0){{dtLabels.push(dt.code);dtRates.push(dt.total?Math.round(dt.approved/dt.total*100):0);}}
  }}));
  if(arChart)arChart.destroy();
  arChart=new Chart(document.getElementById('cApprRate'),{{type:'bar',
    data:{{labels:dtLabels,datasets:[{{label:'Approval %',data:dtRates,
      backgroundColor:dtRates.map(r=>r>=80?'#16a34a':r>=50?'#f59e0b':'#ef4444'),
      borderRadius:4}}]}},
    options:{{responsive:true,indexAxis:'y',
      plugins:{{legend:{{display:false}},tooltip:{{callbacks:{{label:ctx=>`${{ctx.raw}}%`}}}}}},
      scales:{{x:{{beginAtZero:true,max:100,ticks:{{callback:v=>v+'%'}}}},y:{{grid:{{display:false}}}}}}}}}});
}}

// ── Overdue Tab ───────────────────────────────────────────
let overdueLoaded=false;
async function loadOverdue(){{
  if(overdueLoaded)return;
  overdueLoaded=true;
  const pid=_currentPid;
  OVERDUE_DATA=await apiFetch('/api/analytics/overdue'+(pid?`?pid=${{pid}}`:'') )||[];
  filterOverdue();
}}

function filterOverdue(){{
  const q=(document.getElementById('ov-search')?.value||'').toLowerCase();
  const sort=document.getElementById('ov-sort')?.value||'days';
  let data=[...OVERDUE_DATA];
  if(q)data=data.filter(r=>
    (r.docNo||'').toLowerCase().includes(q)||
    (r.title||'').toLowerCase().includes(q)||
    (r.discipline||'').toLowerCase().includes(q)||
    (r.dt_code||'').toLowerCase().includes(q));
  data.sort((a,b)=>sort==='days'?b.days_overdue-a.days_overdue:
    sort==='doc'?(a.docNo||'').localeCompare(b.docNo||''):
    (a.discipline||'').localeCompare(b.discipline||''));
  const list=document.getElementById('ov-list');
  if(!data.length){{list.innerHTML='<div style="text-align:center;padding:32px;color:var(--mu)">✅ No overdue documents</div>';return;}}
  list.innerHTML=data.map(r=>{{
    const bg=r.days_overdue>21?'er':r.days_overdue>14?'warn':'warn';
    return `<div class="ov-row">
      <span class="ov-badge ${{bg}}">${{r.days_overdue}}d</span>
      <div style="flex:1;min-width:0">
        <div style="font-weight:700;font-size:12px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${{r.docNo}}</div>
        <div style="color:var(--mu);font-size:10px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${{r.title||'—'}}</div>
      </div>
      <span style="font-size:10px;color:var(--mu);white-space:nowrap">${{r.discipline||'—'}}</span>
      <span style="font-size:10px;background:#f0f4f8;padding:2px 7px;border-radius:4px;white-space:nowrap">${{r.dt_code}}</span>
      <span style="font-size:10px;color:var(--mu);white-space:nowrap">${{(r.issuedDate||'').slice(0,10)}}</span>
    </div>`;
  }}).join('');
  document.getElementById('ov-count').textContent=`Showing ${{data.length}} overdue document${{data.length!==1?'s':''}}`;
}}

// ── Executive Summary ─────────────────────────────────────
async function loadExecutive(){{
  if(EXEC_DATA)return;
  EXEC_DATA=await apiFetch('/api/executive_summary');
  if(!EXEC_DATA)return;
  const s=EXEC_DATA.summary;
  const col=s.completion_pct>=80?'#16a34a':s.completion_pct>=50?'#f59e0b':'#ef4444';
  document.getElementById('exec-content').innerHTML=`
    <div class="panel">
      <div class="panel-title">📋 Executive Summary — Generated ${{EXEC_DATA.generated_at}}</div>
      <div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(120px,1fr));gap:8px;margin-bottom:16px">
        ${{[['Total Docs',s.total,'var(--pr)'],['Approved',s.approved,'#16a34a'],
           ['Pending',s.pending,'#f59e0b'],['Rejected',s.rejected,'#7c3aed'],
           ['Overdue',s.overdue,'#ef4444'],['Completion',s.completion_pct+'%',col]]
          .map(([l,v,c])=>`<div style="background:var(--bg);border-radius:8px;padding:10px;text-align:center">
            <div style="font-size:22px;font-weight:800;color:${{c}}">${{v}}</div>
            <div style="font-size:9px;color:var(--mu);font-weight:700;text-transform:uppercase">${{l}}</div>
          </div>`).join('')}}
      </div>
      <table class="dt-tbl" style="margin-bottom:16px">
        <thead><tr><th>Project</th><th>Code</th>
          <th style="text-align:center">Total</th>
          <th style="text-align:center">Approved</th>
          <th style="text-align:center">%</th>
          <th style="text-align:center">Overdue</th></tr></thead>
        <tbody>
          ${{EXEC_DATA.projects.map((p,i)=>{{
            const col2=p.pct>=80?'#16a34a':p.pct>=50?'#f59e0b':'#ef4444';
            return `<tr class="${{i%2?'alt':''}}">
              <td style="font-weight:600">${{p.name}}</td>
              <td style="color:var(--pr);font-weight:700">${{p.code}}</td>
              <td style="text-align:center;font-weight:700">${{p.total}}</td>
              <td style="text-align:center;color:#16a34a;font-weight:700">${{p.approved}}</td>
              <td style="text-align:center;font-weight:700;color:${{col2}}">${{p.pct}}%</td>
              <td style="text-align:center;color:#ef4444;font-weight:700">${{p.overdue||0}}</td>
            </tr>`;
          }}).join('')}}
        </tbody>
      </table>
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px">
        <div>
          <div class="panel-title" style="margin-top:0">⏳ Aging Summary</div>
          ${{EXEC_DATA.aging.map(a=>`
            <div style="display:flex;justify-content:space-between;align-items:center;
              padding:6px 0;border-bottom:1px solid var(--bd);font-size:12px">
              <span>${{a.range}} Days</span>
              <span style="font-weight:700;color:${{a.range==='>21'?'#ef4444':'#f59e0b'}}">${{a.count}} docs</span>
            </div>`).join('')}}
        </div>
        <div>
          <div class="panel-title" style="margin-top:0">⚠️ Top Overdue</div>
          ${{EXEC_DATA.top_overdue.slice(0,5).map(r=>`
            <div style="padding:5px 0;border-bottom:1px solid var(--bd);font-size:11px">
              <div style="font-weight:700">${{r.docNo}}</div>
              <div style="color:var(--mu)">${{r.discipline||'—'}} · <span style="color:#ef4444;font-weight:700">${{r.days_overdue}}d overdue</span></div>
            </div>`).join('')||'<div style="color:var(--mu);padding:10px 0">No overdue documents ✅</div>'}}
        </div>
      </div>
    </div>`;
}}

// ── Project management ────────────────────────────────────
async function createProject(){{
  const id=document.getElementById('np-id')?.value.trim();
  const code=document.getElementById('np-code')?.value.trim();
  const name=document.getElementById('np-name')?.value.trim();
  if(!id||!name){{toast('Fill in Project ID and Name','er');return;}}
  const btn=document.getElementById('cpbtn');
  btn.disabled=true;btn.textContent='Creating...';
  try{{
    const r=await fetch('/api/projects/create',{{method:'POST',credentials:'include',
      headers:{{'Content-Type':'application/json','X-Requested-With':'XMLHttpRequest'}},
      body:JSON.stringify({{id,code:code||id,name}})}});
    const d=await r.json();
    if(d.ok){{toast('✔ Project created','ok');closeM('newproj-modal');
      setTimeout(()=>location.reload(),800);}}
    else toast(d.error||'Error','er');
  }}catch(e){{toast('Network error','er');}}
  finally{{btn.disabled=false;btn.textContent='Create';}}
}}

async function delProject(pid,name){{
  if(!confirm(`Delete project "${{name}}"?\nThis will delete ALL records, documents and settings. This cannot be undone.`))return;
  const r=await apiFetch('/api/projects/delete/'+pid,{{method:'POST'}});
  if(r&&r.ok){{toast('Project deleted','ok');setTimeout(()=>location.reload(),800);}}
  else toast((r&&r.error)||'Error','er');
}}

// ── Admin Panel ─────────────────────────────────────────
async function openAdmin(){{
  const [users,projects]=await Promise.all([apiFetch('/api/users'),apiFetch('/api/projects')]);
  if(!users||!projects)return;
  const body=document.getElementById('admin-body');body.innerHTML='';
  const utitle=document.createElement('div');utitle.className='stitle';utitle.textContent='👥 Users';body.appendChild(utitle);
  for(const u of users){{
    const row=document.createElement('div');row.className='urow';
    row.innerHTML=`<span style="flex:1;font-weight:600">👤 ${{u.username}}</span>
      <span class="badge" style="background:#fef3c7;color:#92400e">${{u.role.toUpperCase()}}</span>
      ${{u.username!=='admin'?`<button class="btn btn-sc btn-sm" onclick="chgPw('${{u.username}}')">🔑 PW</button>
        <button class="btn btn-er btn-sm" onclick="delUsr('${{u.username}}')">✕</button>`:
        '<span style="font-size:10px;color:var(--mu)">(protected)</span>'}}`;
    body.appendChild(row);
    if(u.role!=='superadmin'){{
      const ad=document.createElement('div');
      ad.style.cssText='padding:4px 10px 10px 32px;border-bottom:1px solid var(--bd);margin-bottom:4px';
      ad.innerHTML='<div style="font-size:10px;color:var(--mu);margin-bottom:6px">Project access:</div>';
      const assigned=await apiFetch('/api/users/'+u.username+'/projects').catch(()=>[]);
      const pl=document.createElement('div');pl.style.cssText='display:flex;flex-wrap:wrap;gap:5px';
      projects.forEach(p=>{{
        const isOn=assigned.includes(p.id);
        const btn=document.createElement('button');
        btn.style.cssText='padding:3px 10px;border-radius:4px;cursor:pointer;font-size:11px;font-weight:700;font-family:inherit;transition:all .15s;border:2px solid '+(isOn?'#f0a500':'#e2e8f0')+';background:'+(isOn?'#1a3a5c':'#f8fafc')+';color:'+(isOn?'#fff':'#94a3b8');
        btn.textContent=p.code;btn.title=p.name;
        if(isOn)btn.dataset.on='1';
        btn.onclick=async()=>{{
          const on=!!btn.dataset.on;
          await apiFetch('/api/users',{{method:'POST',body:JSON.stringify({{action:on?'unassign':'assign',username:u.username,project_id:p.id}})}});
          if(on){{delete btn.dataset.on;btn.style.borderColor='#e2e8f0';btn.style.background='#f8fafc';btn.style.color='#94a3b8';}}
          else{{btn.dataset.on='1';btn.style.borderColor='#f0a500';btn.style.background='#1a3a5c';btn.style.color='#fff';}}
          toast((btn.dataset.on?'✔ Assigned: ':'Removed: ')+p.name,'ok');
        }};
        pl.appendChild(btn);
      }});
      ad.appendChild(pl);body.appendChild(ad);
    }}
  }}
  const at=document.createElement('div');at.className='stitle';at.textContent='➕ Add User';body.appendChild(at);
  const ar=document.createElement('div');
  ar.innerHTML=`<div style="display:grid;grid-template-columns:1fr 1fr 1fr auto;gap:8px;align-items:end">
    <div class="fg"><label>Username</label><input id="nu-name" placeholder="username"></div>
    <div class="fg"><label>Role</label><select id="nu-role">
      <option value="editor">Editor</option><option value="viewer">Viewer</option>
      <option value="admin">Admin</option></select></div>
    <div class="fg"><label>Password</label><input id="nu-pw" type="password"></div>
    <button class="btn btn-pr btn-sm" onclick="addUsr()">Add</button></div>`;
  body.appendChild(ar);openM('admin-modal');
}}
async function addUsr(){{
  const name=document.getElementById('nu-name')?.value.trim().toLowerCase();
  const role=document.getElementById('nu-role')?.value;
  const pw=document.getElementById('nu-pw')?.value;
  if(!name||!pw){{toast('Username and password required','er');return;}}
  const r=await apiFetch('/api/users',{{method:'POST',body:JSON.stringify({{action:'add',username:name,role,password:pw}})}});
  if(r&&r.ok){{toast('✔ User added','ok');closeM('admin-modal');openAdmin();}}
  else toast((r&&r.error)||'Error','er');
}}
async function delUsr(u){{
  if(!confirm('Delete user: '+u+'?'))return;
  const r=await apiFetch('/api/users',{{method:'POST',body:JSON.stringify({{action:'delete',username:u}})}});
  if(r&&r.ok){{toast('Deleted','wa');closeM('admin-modal');openAdmin();}}
}}
async function chgPw(u){{
  const pw=prompt('New password for '+u+':');if(!pw)return;
  await apiFetch('/api/users',{{method:'POST',body:JSON.stringify({{action:'change_password',username:u,password:pw}})}});
  toast('✔ Password changed','ok');
}}

// ── Audit Log ────────────────────────────────────────────────
let _auditOffset=0,_auditHasMore=true;
const ACTION_COLORS={{
  'ADD':'#166534','EDIT':'#1d4ed8','DELETE':'#991b1b',
  'LOGIN':'#6b7280','LOGIN_FAIL':'#dc2626','EXPORT_EXCEL':'#7c3aed'
}};

async function loadAudit(reset=false){{
  if(reset)_auditOffset=0;
  const pid=document.getElementById('aud-pid')?.value||'';
  const user=document.getElementById('aud-user')?.value||'';
  const action=document.getElementById('aud-action')?.value||'';
  const params=new URLSearchParams();
  if(pid)params.set('pid',pid);
  if(user)params.set('username',user);
  if(action)params.set('action',action);
  params.set('offset',_auditOffset);
  const data=await apiFetch('/api/audit?'+params);
  if(!data)return;

  // Populate filters on first load
  if(reset||_auditOffset===0){{
    const pSel=document.getElementById('aud-pid');
    const curPid=pSel.value;
    pSel.innerHTML='<option value="">All Projects</option>'+
      (data.projects||[]).map(p=>`<option value="${{p.id}}"${{p.id===curPid?' selected':''}}>${{p.code}} — ${{p.name}}</option>`).join('');

    const uSel=document.getElementById('aud-user');
    const curU=uSel.value;
    uSel.innerHTML='<option value="">All Users</option>'+
      (data.users||[]).map(u=>`<option value="${{u}}"${{u===curU?' selected':''}}>${{u}}</option>`).join('');

    const aSel=document.getElementById('aud-action');
    const curA=aSel.value;
    aSel.innerHTML='<option value="">All Actions</option>'+
      (data.actions||[]).map(a=>`<option value="${{a}}"${{a===curA?' selected':''}}>${{a}}</option>`).join('');
  }}

  const tbody=document.getElementById('aud-tbody');
  tbody.innerHTML='';
  const rows=data.rows||[];
  _auditHasMore=rows.length===100;

  if(!rows.length){{
    tbody.innerHTML='<tr><td colspan="9" style="text-align:center;padding:24px;color:var(--mu)">No activity found</td></tr>';
  }}else{{
    rows.forEach((r,i)=>{{
      const tr=document.createElement('tr');
      tr.className=i%2===0?'':'alt';
      const ts=new Date(r.ts);
      const tsStr=ts.toLocaleDateString('en-GB')+' '+ts.toLocaleTimeString('en-GB',{{hour:'2-digit',minute:'2-digit'}});
      const actionColor=ACTION_COLORS[r.action]||'#374151';
      const actionBg=r.action==='ADD'?'#bbf7d0':r.action==='EDIT'?'#dbeafe':r.action==='DELETE'?'#fee2e2':r.action==='LOGIN'?'#f3f4f6':'#fef3c7';
      tr.innerHTML=`
        <td style="font-size:10px;color:var(--mu);white-space:nowrap">${{tsStr}}</td>
        <td style="font-weight:600;font-size:11px">${{r.username||''}}</td>
        <td><span style="background:${{actionBg}};color:${{actionColor}};font-size:9px;font-weight:700;padding:2px 7px;border-radius:10px;white-space:nowrap">${{r.action||''}}</span></td>
        <td style="font-size:10px;color:var(--mu)">${{r.project_id||''}}</td>
        <td style="font-size:11px;font-weight:600;color:var(--pr)">${{r.doc_no||''}}</td>
        <td style="font-size:10px;color:var(--mu)">${{r.field_name||''}}</td>
        <td style="font-size:10px;color:#dc2626;max-width:150px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${{r.old_value||''}}">${{r.old_value||''}}</td>
        <td style="font-size:10px;color:#166534;max-width:150px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${{r.new_value||''}}">${{r.new_value||''}}</td>
        <td style="font-size:10px;color:var(--mu)">${{r.detail||''}}</td>`;
      tbody.appendChild(tr);
    }});
  }}

  const page=Math.floor(_auditOffset/100)+1;
  document.getElementById('aud-count').textContent=`Showing ${{rows.length}} entries`;
  document.getElementById('aud-page').textContent=`Page ${{page}}`;
  document.getElementById('aud-prev').disabled=_auditOffset===0;
  document.getElementById('aud-next').disabled=!_auditHasMore;
}}

function auditPage(dir){{
  _auditOffset=Math.max(0,_auditOffset+dir*100);
  loadAudit();
}}


init();
</script>
</body></html>"""



def render_register(u, proj):
    pid      = proj["id"]
    uname    = u["username"] if u else "guest"
    role     = u["role"]     if u else "guest"
    editable = can_edit(pid)
    btns, _, rlbl, rbg = _user_info_html(u)
    sc_json  = json.dumps(STATUS_COLORS)

    dts       = db.get_doc_types(pid)
    logo_l    = db.get_logo(pid,"logo_left")
    logo_r    = db.get_logo(pid,"logo_right")
    logo_html = ""
    if logo_l: logo_html += f'<img src="/api/logo/{pid}/logo_left" style="height:48px;object-fit:contain;flex-shrink:0">'
    if logo_r: logo_html += f'<img src="/api/logo/{pid}/logo_right" style="height:48px;object-fit:contain;flex-shrink:0;margin-left:8px">'

    DEFAULT_PROJ_FIELDS = [("code","Code"),("name","Project Name"),("startDate","Start"),("endDate","End"),
                   ("client","Client"),("landlord","Landlord"),("pmo","PMO"),
                   ("mainConsultant","Consultant"),("mepConsultant","MEP"),("contractor","Contractor")]
    custom_labels = proj.get("_labels") or {}
    PROJ_FIELDS = [(k, custom_labels.get(k, lbl)) for k, lbl in DEFAULT_PROJ_FIELDS]

    primary_keys = {"code", "name"}
    projbar_primary = "".join(
        f'<div class="pf primary"><span class="pf-lbl">{lbl}</span>'
        f'<span class="pf-val" data-key="{key}">{proj.get(key,"") or "—"}</span></div>'
        for key,lbl in PROJ_FIELDS
        if key in primary_keys and proj.get(key,"").strip())
    projbar_secondary = "".join(
        f'<div class="pf secondary"><span class="pf-lbl">{lbl}</span>'
        f'<span class="pf-val" data-key="{key}">{proj.get(key,"") or "—"}</span></div>'
        for key,lbl in PROJ_FIELDS
        if key not in primary_keys and proj.get(key,"").strip())
    projbar_toggle = ('<button id="projbar-toggle" class="tool-btn" type="button" '
                      'onclick="toggleProjectInfo()">Project Info</button>') if projbar_secondary else ''

    tabs_html = "".join(
        f'<button class="tab-btn" data-id="{dt["id"]}" onclick="switchTab(\'{dt["id"]}\')">'
        f'<span>{dt["code"]}</span><span class="tcnt" id="cnt-{dt["id"]}">0</span></button>'
        for dt in dts)

    _sa_only = role == 'superadmin'
    _hol_btn = " <button class='tool-btn purple' onclick='openSettings()'>🗓 Holidays</button>" if _sa_only else ''
    _col_btn = "<button class='tool-btn purple' onclick='manageColumns()'>⚙ Columns</button>" if _sa_only else ''
    _lst_btn = "<button class='tool-btn' onclick='openLists()'>📋 Lists</button>" if _sa_only else ''
    edit_btns = (f'<button class="tool-btn" onclick="addRecord()">➕ Add</button>'
                 f'{_col_btn}'
                 f'{_hol_btn}'
                 f'{_lst_btn}'
                 f'<button class="tool-btn" onclick="editProject()">🏗 Project</button>'
                 if editable else
                 '<span style="font-size:11px;color:rgba(255,255,255,.5);padding:4px 8px">'
                 '👁 Read-only — <a href="/login" style="color:#f0a500">login to edit</a></span>')

    return f"""<!DOCTYPE html><html lang="en"><head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>DCR — {proj.get("name","Register")}</title>
{BASE_CSS}
<style>
.hidden{{display:none!important}}
body{{height:100vh;display:flex;flex-direction:column;overflow:hidden}}
body.dark #projbar,body.dark #toolbar,body.dark #ltr-quickbar,body.dark #noc-summary{{background:#162132;color:#e2e8f0;border-color:#304257}}
body.dark #tabsbar{{background:#0d1f33}}
body.dark #tblwrap{{background:#111b2a}}
body.dark .pf-lbl{{color:#93c5fd}}
body.dark .pf-val{{color:#dbe7f3}}
body.dark #regtbl th{{background:#17314d}}
body.dark .frow th{{background:#101a29}}
body.dark .frow input,body.dark .frow select{{background:#0f1a29;border-color:#314558;color:#e2e8f0}}
body.dark #regtbl td{{border-color:#253648;color:#d7e1ec}}
body.dark #regtbl tr.alt td{{background:#132031}}
body.dark #regtbl tr.ov td{{background:#3a231f;color:#f8d8c7}}
body.dark #regtbl tr.rv td{{background:#101a29;color:#b8c8da}}
body.dark #regtbl tbody tr:hover td{{background:#1e3147;color:#eef5ff}}
body.dark #regtbl tbody tr.ov:hover td{{background:#4a2a24;color:#ffe1d2}}
body.dark #regtbl tbody tr.rv:hover td{{background:#1e3147;color:#eef5ff}}
body.dark #regtbl tbody tr.row-selected td{{background:#1d3a5a!important;color:#f3f8ff!important;box-shadow:inset 0 1px 0 rgba(147,197,253,.2),inset 0 -1px 0 rgba(147,197,253,.2)}}
body.dark #regtbl tbody tr.row-selected td:first-child{{box-shadow:inset 4px 0 0 #93c5fd,inset 0 1px 0 rgba(147,197,253,.2),inset 0 -1px 0 rgba(147,197,253,.2)}}
body.dark #regtbl tbody tr.row-selected:hover td{{background:#25496d!important;color:#fff!important}}
body.dark .flink{{color:#93c5fd}}
body.dark .tool-dd-menu{{background:#162132;border-color:#304257}}
body.dark .tool-dd-menu button{{color:#e2e8f0}}
body.dark .tool-dd-menu button:hover{{background:#101a29;color:#93c5fd}}
body.dark .tool-btn{{color:#e2e8f0;border-color:#304257;background:#101a29}}
body.dark .tab-btn{{color:rgba(255,255,255,.72)}}
body.dark .tab-btn:hover{{background:rgba(255,255,255,.08)}}
body.dark .tcnt{{background:rgba(255,255,255,.14)}}
body.dark #sbar{{background:#0d1f33;color:rgba(255,255,255,.72)}}
body.dark .sitem{{background:#101a29;color:#e2e8f0;border-color:#253648}}
body.dark .sitem button{{background:#f8fafc;color:#0f172a;border-color:#94a3b8}}
body.dark .addrow input{{background:#f8fafc;color:#0f172a;border-color:#cbd5e1}}
@media print{{
  #topbar,#tabbar,.toolrow,.bulkbar,#statusbar,.acts{{display:none!important}}
  body{{height:auto;overflow:visible}}
  #main{{overflow:visible;height:auto}}
  #tblwrap{{overflow:visible;height:auto}}
  #regtbl{{font-size:10px}}
  #regtbl th,#regtbl td{{padding:4px 6px!important;white-space:normal!important}}
  @page{{size:A4 landscape;margin:10mm}}
}}
#projbar{{background:#fff;border-bottom:2px solid var(--pr);padding:5px 10px;
  display:flex;align-items:center;flex-wrap:wrap;flex-shrink:0;gap:6px}}
#projbar-main{{display:flex;align-items:center;flex:1 1 auto;gap:4px;min-width:0;flex-wrap:wrap}}
#projbar-primary,#projbar-extra{{display:flex;align-items:center;gap:0;flex-wrap:wrap;min-width:0}}
#projbar-toggle{{display:none}}
#projbar img{{max-height:30px;max-width:108px;object-fit:contain;flex-shrink:0}}
.pf{{display:flex;flex-direction:column;padding:0 8px;border-right:1px solid var(--bd);min-width:0}}
.pf:last-of-type{{border-right:none}}
.pf-lbl{{font-size:9px;font-weight:700;color:var(--pr);text-transform:uppercase;letter-spacing:.4px}}
.pf-val{{font-size:11px;line-height:1.2;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:220px}}
#tabsbar{{background:#0f2640;display:flex;align-items:center;overflow-x:auto;flex-shrink:0;
  padding:4px 8px;scrollbar-width:thin;gap:6px}}
#tabsbar::-webkit-scrollbar{{height:3px}}
#tabsbar::-webkit-scrollbar-thumb{{background:rgba(255,255,255,.3)}}
.tab-btn{{display:flex;align-items:center;gap:5px;padding:9px 12px;background:transparent;
  border:none;border-bottom:3px solid transparent;color:rgba(255,255,255,.55);cursor:pointer;
  font-family:inherit;font-size:11px;font-weight:600;white-space:nowrap;transition:all .15s}}
.tab-btn:hover{{color:#fff;background:rgba(255,255,255,.07)}}
.tab-btn.active{{color:#fff;border-bottom-color:var(--ac)}}
.tcnt{{background:rgba(255,255,255,.2);border-radius:10px;padding:1px 7px;font-size:10px;font-weight:700}}
.tab-btn.active .tcnt{{background:var(--ac);color:#000}}
.tab-add{{padding:5px 10px;background:rgba(255,255,255,.1);border:1px dashed rgba(255,255,255,.35);
  color:rgba(255,255,255,.7);border-radius:4px;cursor:pointer;font-size:16px;margin-left:6px}}
.tab-add:hover{{background:rgba(255,255,255,.2)}}
#toolbar{{background:#fff;border-bottom:1px solid var(--bd);padding:4px 8px;
  display:flex;align-items:center;gap:5px;flex-shrink:0;flex-wrap:wrap}}
#toolbar-actions{{display:flex;align-items:center;gap:5px;flex-wrap:wrap;min-width:0}}
.tool-btn{{display:flex;align-items:center;gap:4px;padding:4px 9px;background:var(--bg);
  border:1px solid var(--bd);border-radius:var(--rd);cursor:pointer;font-size:11px;
  font-family:inherit;color:var(--tx);transition:all .15s;white-space:nowrap}}
.tool-btn:hover{{background:var(--pr);color:#fff;border-color:var(--pr)}}
.tool-btn.purple:hover{{background:#7c3aed;border-color:#7c3aed}}
.tool-btn.teal:hover{{background:#0891b2;border-color:#0891b2}}
.tool-dd{{position:relative}}
.tool-dd-menu{{position:absolute;top:calc(100% + 4px);left:0;background:#fff;border:1.5px solid var(--bd);border-radius:6px;box-shadow:0 8px 24px rgba(0,0,0,.15);z-index:300;min-width:210px;overflow:hidden}}
.tool-dd-menu button{{display:block;width:100%;text-align:left;padding:9px 14px;border:none;background:none;cursor:pointer;font-size:12px;font-family:inherit;color:#1e2a3a;white-space:nowrap}}
.tool-dd-menu button:hover{{background:#f0f4f8;color:var(--pr)}}
#srchbox{{flex:1;min-width:150px;max-width:260px;padding:4px 10px 4px 28px;border:1px solid var(--bd);
  border-radius:var(--rd);font-family:inherit;font-size:12px;outline:none;
  background:#fff url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='13' height='13' viewBox='0 0 24 24' fill='none' stroke='%236b7a94' stroke-width='2'%3E%3Ccircle cx='11' cy='11' r='8'/%3E%3Cpath d='m21 21-4.35-4.35'/%3E%3C/svg%3E") no-repeat 7px center}}
#srchbox:focus{{border-color:var(--pl);box-shadow:0 0 0 2px rgba(37,99,168,.1)}}
#main{{flex:1;overflow:hidden;display:flex;flex-direction:column;min-width:0}}
#tblwrap{{flex:1;overflow:auto;min-width:0;-webkit-overflow-scrolling:touch}}
#regtbl{{width:100%;border-collapse:collapse;min-width:980px;font-size:12px}}
#regtbl thead{{position:sticky;top:0;z-index:20}}
#regtbl th{{background:var(--pr);color:#fff;padding:8px;text-align:left;font-weight:600;
  white-space:nowrap;border-right:1px solid rgba(255,255,255,.1);cursor:pointer;user-select:none;position:relative}}
#thead tr:first-child th{{position:relative;top:auto;z-index:auto;border-bottom:2px solid rgba(255,255,255,.24);box-shadow:none}}
#regtbl th:hover{{background:var(--pl)}}
.frow{{position:static}}
.frow th{{background:#eef1f7;padding:3px 4px;cursor:default;position:static;top:auto;z-index:auto;box-shadow:0 1px 0 rgba(221,227,237,.95)}}
.frow th:hover{{background:#eef1f7}}
.frow input,.frow select{{width:100%;padding:3px 6px;border:1px solid var(--bd);
  border-radius:3px;font-size:10px;font-family:inherit;background:#fff;outline:none}}
#regtbl td{{padding:5px 8px;border-bottom:1px solid #edf0f5;border-right:1px solid #f3f4f6;
  vertical-align:middle;max-width:220px;word-break:break-word}}
#regtbl tr:hover td{{background:rgba(37,99,168,.10);transition:background .1s}}
#regtbl tr.ov td{{background:#fff5f5}}
#regtbl tr.rv td{{color:var(--mu)}}
#regtbl tr.alt td{{background:#fafbfd}}
#regtbl tbody tr:hover td{{background:#dce8f5;color:var(--tx)}}
#regtbl tbody tr.ov:hover td{{background:#ffe7e7;color:#3b1f1f}}
#regtbl tbody tr.rv:hover td{{background:#fff4d8;color:#3b2a12}}
#regtbl tbody tr.row-selected td{{background:#d7e8fb!important;color:#10233a!important;box-shadow:inset 0 1px 0 rgba(37,99,168,.18),inset 0 -1px 0 rgba(37,99,168,.18)}}
#regtbl tbody tr.row-selected td:first-child{{box-shadow:inset 4px 0 0 #2563a8,inset 0 1px 0 rgba(37,99,168,.18),inset 0 -1px 0 rgba(37,99,168,.18)}}
#regtbl tbody tr.row-selected:hover td{{background:#c8def6!important;color:#10233a!important}}
.sr{{text-align:center;color:var(--mu);font-size:10px;min-width:28px}}
.chkcell{{text-align:center;width:28px;padding:4px!important}}
.chkcell input{{width:14px;height:14px;cursor:pointer;accent-color:var(--pr)}}
.acts{{white-space:nowrap;width:64px}}
.act{{padding:2px 7px;border:1px solid var(--bd);background:#fff;border-radius:3px;cursor:pointer;font-size:11px}}
.act:hover{{background:var(--pr);color:#fff;border-color:var(--pr)}}
.act.del:hover{{background:var(--er);border-color:var(--er)}}
.pr-toggle{{padding:2px 7px;border:1px solid var(--bd);background:#fff;border-radius:3px;cursor:pointer;font-size:11px}}
.pr-toggle:hover{{background:var(--pr);color:#fff;border-color:var(--pr)}}
.pr-items-row{{display:none}}
.pr-items-row.open{{display:table-row}}
.pr-items-row>td{{padding:0!important;border-bottom:1px solid #d9e2ec!important;border-right:none!important;max-width:none!important;background:transparent!important}}
.pr-items-panel{{display:block;width:100%;box-sizing:border-box;padding:16px;background:#f8fafc;border-top:1px solid #d9e2ec;border-bottom:1px solid #d9e2ec;position:relative;z-index:1}}
.pr-items-grid-wrap{{display:block;width:100%;max-width:100%;overflow-x:auto;padding-bottom:2px}}
.pr-items-title{{display:flex;align-items:center;gap:8px;margin-bottom:8px;font-size:11px;font-weight:800;color:var(--pr);text-transform:uppercase;letter-spacing:.45px}}
.pr-items-title:before{{content:"";width:4px;height:16px;border-radius:999px;background:var(--ac);display:inline-block}}
.pr-items-grid{{display:grid;grid-template-columns:minmax(520px,1fr) 120px 100px minmax(240px,.6fr);min-width:1100px;width:100%;border:1px solid #dbe3ed;border-radius:6px;overflow:hidden;background:#fff;font-size:11px}}
.pr-items-grid-head{{background:#173f66;color:#fff;font-weight:800;padding:10px 12px;border-right:1px solid rgba(255,255,255,.18)}}
.pr-items-grid-head:last-child{{border-right:none}}
.pr-items-cell{{padding:10px 12px;border-top:1px solid #edf2f7;border-right:1px solid #edf2f7;background:#fff;color:#10233a;white-space:normal;word-break:normal;line-height:1.35}}
.pr-items-cell:nth-child(4n){{border-right:none}}
.pr-items-cell.qty{{text-align:center;font-variant-numeric:tabular-nums}}
.pr-items-cell.unit{{white-space:nowrap}}
.pr-items-section{{grid-column:1/-1;background:#e8eef6;color:var(--pr);font-weight:800;text-transform:uppercase;letter-spacing:.35px;padding:9px 12px;border-top:1px solid #d7dee8}}
.pr-items-empty{{color:var(--mu);font-size:11px;padding:6px 0}}
body.dark .pr-items-row>td{{background:transparent!important;border-color:#304257!important}}
body.dark .pr-items-panel{{background:#101a29;border-color:#304257}}
body.dark .pr-items-grid{{background:#162132;border-color:#304257}}
body.dark .pr-items-grid-head{{background:#17314d;color:#dbeafe;border-color:#304257}}
body.dark .pr-items-cell{{background:#162132;color:#d7e1ec;border-color:#253648}}
body.dark .pr-items-section{{background:#1e3147;color:#dbeafe;border-color:#304257}}
.sbadge{{display:inline-block;border-radius:10px;padding:2px 9px;font-size:10px;font-weight:700}}
.flink{{color:var(--pl);text-decoration:underline;cursor:pointer;font-size:11px}}
.ovdate{{color:#dc2626;font-weight:700}}
.mlcell{{white-space:pre-line!important;word-break:break-word}}
#bulkbar{{display:none;background:#1a3a5c;color:#fff;padding:5px 14px;align-items:center;gap:10px;font-size:12px;flex-shrink:0}}
#bulkbar.show{{display:flex}}
#sbar{{background:var(--pr);color:rgba(255,255,255,.75);padding:2px 12px;font-size:10px;display:flex;gap:12px;flex-wrap:wrap;flex-shrink:0}}
.rz{{position:absolute;right:0;top:0;bottom:0;width:6px;cursor:col-resize;z-index:1}}
.rz:hover,.rz.rzg{{background:var(--ac)}}
.ms-con{{border:1px solid var(--bd);border-radius:var(--rd);min-height:34px;padding:3px;
  background:#fff;cursor:pointer;position:relative}}
.ms-tag{{display:inline-flex;align-items:center;gap:3px;background:var(--pr);color:#fff;
  border-radius:3px;padding:2px 7px;font-size:10px;margin:2px}}
.ms-rm{{cursor:pointer;opacity:.7}}
.ms-ph{{color:var(--mu);font-size:11px;padding:3px 5px}}
.ms-dd{{position:absolute;left:0;right:0;top:100%;background:#fff;border:1.5px solid var(--bd);
  border-radius:var(--rd);z-index:500;max-height:260px;overflow-y:auto;
  box-shadow:0 8px 24px rgba(0,0,0,.12);margin-top:2px}}
.ms-opt{{padding:9px 12px;cursor:pointer;font-size:12px;display:flex;align-items:center;gap:8px;border-bottom:1px solid #f0f4f8}}
.ms-opt:last-child{{border-bottom:none}}
.ms-opt:hover{{background:#f0f4f8;color:var(--pr)}}
.ms-opt.sel{{background:#eff6ff}}
body.dark .ms-con{{background:#101a29;border-color:#304257;color:#e2e8f0}}
body.dark .ms-ph{{color:#9fb0c6}}
body.dark .ms-dd{{background:#101a29;border-color:#304257;box-shadow:0 14px 36px rgba(0,0,0,.42)}}
body.dark .ms-opt{{background:#101a29;color:#dbe7f3;border-bottom-color:#253648}}
body.dark .ms-opt:hover{{background:#1e3147;color:#eef5ff}}
body.dark .ms-opt.sel{{background:#17314d;color:#eef5ff}}
body.dark .ms-opt input{{accent-color:#60a5fa}}
.letter-flow{{position:relative;padding-left:10px}}
.letter-flow-line{{position:absolute;left:13px;top:0;bottom:0;width:2px;background:linear-gradient(180deg,#d7e0e6,#c7d2da)}}
.letter-node{{position:relative;padding-left:30px;margin-bottom:14px}}
.letter-dot{{position:absolute;left:8px;top:14px;width:12px;height:12px;border-radius:999px;background:#2F4F64;border:3px solid #fff;box-shadow:0 0 0 2px rgba(47,79,100,.18)}}
.letter-dot.current{{background:#8BC34A;box-shadow:0 0 0 2px rgba(139,195,74,.35)}}
.letter-card{{background:#fff;border:1px solid #d7e0e6;border-left:4px solid #d7e0e6;border-radius:12px;padding:12px 14px;box-shadow:0 6px 18px rgba(15,23,42,.05)}}
.letter-card.current{{background:#f6fbef;border-color:#8BC34A;border-left-color:#8BC34A;box-shadow:0 10px 26px rgba(139,195,74,.14)}}
.letter-card-kicker{{font-size:9px;color:#7d8a95;text-transform:uppercase;letter-spacing:.35px;margin-bottom:6px}}
.letter-card-ref{{font-size:12px;font-weight:800;color:var(--pr)}}
.letter-card-subject{{font-size:12px;color:var(--tx);margin-top:6px;line-height:1.45}}
.letter-card-meta{{font-size:10px;color:var(--mu);margin-top:8px;line-height:1.5}}
.letter-date-badge{{display:inline-flex;align-items:center;background:#eef4f8;color:#2F4F64;border-radius:999px;padding:4px 10px;font-size:10px;font-weight:800;letter-spacing:.2px}}
.letter-current-badge{{flex-shrink:0;background:#8BC34A;color:#17351b;border-radius:999px;padding:3px 8px;font-size:9px;font-weight:800;letter-spacing:.35px;text-transform:uppercase}}
.letter-branch-cue{{color:#7d8a95;font-weight:800;margin-right:8px}}
.letter-connector{{position:absolute;left:-24px;top:0;bottom:22px;width:2px;background:#cbd5df;border-radius:2px}}
.letter-connector-arm{{position:absolute;left:-24px;top:24px;width:18px;height:2px;background:#cbd5df;border-radius:2px}}
body.dark .letter-flow-line{{background:linear-gradient(180deg,#3a526d,#253648)}}
body.dark .letter-dot{{background:#60a5fa;border-color:#101a29;box-shadow:0 0 0 2px rgba(96,165,250,.28)}}
body.dark .letter-dot.current{{background:#8BC34A;box-shadow:0 0 0 2px rgba(139,195,74,.42)}}
body.dark .letter-card{{background:#162132;border-color:#304257;border-left-color:#496780;box-shadow:0 10px 28px rgba(0,0,0,.28)}}
body.dark .letter-card.current{{background:#1d2b1d;border-color:#8BC34A;border-left-color:#8BC34A;box-shadow:0 10px 28px rgba(0,0,0,.32)}}
body.dark .letter-card-kicker,body.dark .letter-card-meta{{color:#b8c8da}}
body.dark .letter-card-ref{{color:#dbeafe}}
body.dark .letter-card-subject{{color:#eef5ff}}
body.dark .letter-date-badge{{background:#233850;color:#bfdbfe}}
body.dark .letter-branch-cue{{color:#93c5fd}}
body.dark .letter-connector,body.dark .letter-connector-arm{{background:#496780}}
body.dark #thread-body [style*="background:#fff"],
body.dark #timeline-body [style*="background:#fff"]{{background:#162132!important;border-color:#304257!important;box-shadow:0 10px 28px rgba(0,0,0,.28)!important}}
body.dark #thread-body [style*="background:#f6fbef"],
body.dark #timeline-body [style*="background:#f6fbef"]{{background:#1d2b1d!important;border-color:#8BC34A!important;box-shadow:0 10px 28px rgba(0,0,0,.32)!important}}
body.dark #thread-body [style*="color:#7d8a95"],
body.dark #timeline-body [style*="color:#7d8a95"]{{color:#b8c8da!important}}
body.dark #thread-body [style*="background:#eef4f8"],
body.dark #timeline-body [style*="background:#eef4f8"]{{background:#233850!important;color:#bfdbfe!important}}
body.dark #thread-body [style*="border:3px solid #fff"],
body.dark #timeline-body [style*="border:3px solid #fff"]{{border-color:#101a29!important}}
body.dark #thread-body [style*="linear-gradient(180deg,#d7e0e6"],
body.dark #timeline-body [style*="linear-gradient(180deg,#d7e0e6"]{{background:linear-gradient(180deg,#3a526d,#253648)!important}}
.empty{{text-align:center;padding:60px 20px;color:var(--mu)}}
.slist{{list-style:none;display:flex;flex-direction:column;gap:3px;
  max-height:190px;overflow-y:auto;border:1px solid var(--bd);border-radius:var(--rd);padding:4px}}
.sitem{{display:flex;align-items:center;gap:8px;padding:4px 8px;background:var(--bg);
  border-radius:3px;font-size:11px}}
.sitem .nm{{flex:1}}
.sitem button{{padding:2px 8px;font-size:10px;border:1px solid var(--bd);background:#fff;
  border-radius:3px;cursor:pointer}}
.sitem button:hover{{background:var(--er);color:#fff;border-color:var(--er)}}
.rtl-txt{{direction:rtl;text-align:right}}
.addrow{{display:flex;gap:6px;margin-top:6px}}
.addrow input{{flex:1;padding:5px 8px;border:1px solid var(--bd);border-radius:3px;
  font-size:11px;font-family:inherit;outline:none}}
#rec-modal .record-modal{{width:min(96vw,980px)!important;max-width:980px!important}}
#rec-modal .mbody{{padding:14px 16px 12px}}
.record-modal-actions{{gap:10px;flex-wrap:wrap}}
.record-modal-actions .btn{{min-height:38px}}
#pr-items-editor{{overflow-x:auto;padding-bottom:2px}}
.pr-items-editor table{{min-width:560px}}
@media(max-width:1200px){{
  #projbar{{flex-wrap:wrap;gap:6px;padding:5px 10px}}
  .pf{{min-width:110px;flex:1 1 120px}}
  #toolbar{{gap:6px}}
  #toolbar-actions{{flex:1 1 auto}}
  #srchbox{{max-width:none;flex:1 1 220px}}
}}
@media(max-width:900px){{
  .tool-btn{{min-height:34px}}
  #tabsbar{{padding:6px 8px}}
  .tab-btn{{border:1px solid rgba(255,255,255,.16);border-bottom:none;border-radius:999px;padding:8px 11px;color:rgba(255,255,255,.82)}}
  .tab-btn.active{{background:#fff;color:#0f2640}}
  .tab-add{{margin-left:0;border-radius:999px;padding:6px 11px;font-size:13px}}
  #srchbox{{flex:1 1 100%;min-width:0}}
  #regtbl{{font-size:11px}}
  #regtbl td,#regtbl th{{padding:4px 5px}}
  #tblwrap{{padding-bottom:4px}}
}}
@media(max-width:768px){{
  #topbar{{height:34px;padding:0 8px;gap:5px;overflow:hidden}}
  #topbar .topbar-title-full{{display:none}}
  #topbar .topbar-title-short{{display:inline!important;font-size:13px!important}}
  #topbar .topbar-title-short::after{{content:" Register"}}
  #topbar .topbar-mark{{font-size:16px!important}}
  #topbar .sp{{flex:0 1 6px;min-width:2px}}
  #topbar .tb-btn{{padding:3px 6px;font-size:9.5px;line-height:1.1;border-radius:7px;min-height:25px;display:inline-flex;align-items:center;justify-content:center;white-space:nowrap}}
  #topbar form .tb-btn{{width:auto;padding:3px 6px}}
  #topbar > span[style*="padding:0 4px"]{{display:none}}
  #topbar .topbar-user{{font-size:8px;gap:0;flex-shrink:0}}
  #topbar .topbar-user-name{{display:none}}
  #topbar .topbar-user span:last-child{{padding:2px 5px!important;font-size:8px!important;line-height:1.05;max-width:62px;text-align:center}}
  #projbar{{padding:2px 8px;gap:5px;align-items:flex-start;border-bottom-width:1px}}
  #projbar img{{max-height:34px;max-width:82px;margin-top:0}}
  #projbar-main{{display:grid;grid-template-columns:minmax(130px,.9fr) minmax(0,1.1fr);gap:2px 10px;flex:1 1 auto;width:auto;min-width:0}}
  #projbar-primary,#projbar-extra{{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:1px 7px;width:100%;min-width:0}}
  #projbar-primary{{grid-template-columns:minmax(0,1fr)}}
  #projbar-primary .pf,#projbar-extra .pf{{padding:0;border-right:none;min-width:0;display:block}}
  #projbar-primary .pf-lbl,#projbar-extra .pf-lbl{{font-size:6px;line-height:.9;color:var(--mu);letter-spacing:.16px;margin-bottom:0}}
  #projbar-primary .pf-val,#projbar-extra .pf-val{{font-size:9px;line-height:1;max-width:none;white-space:normal;display:-webkit-box;-webkit-line-clamp:1;-webkit-box-orient:vertical;overflow:hidden}}
  #projbar-primary .pf.primary:first-child .pf-val{{font-size:10px;font-weight:800;color:var(--pr);text-transform:uppercase;-webkit-line-clamp:1}}
  #projbar-primary .pf.primary:last-child{{grid-column:1/-1;border-bottom:none;padding-bottom:0;margin-bottom:0}}
  #projbar-primary .pf.primary:last-child .pf-val{{font-size:10.5px;font-weight:800;line-height:1;-webkit-line-clamp:1}}
  #projbar-extra{{padding-top:0;border-top:none}}
  #projbar-toggle,.proj-edit-btn{{display:none!important}}
  #tabsbar{{padding:2px 7px;gap:3px;scroll-snap-type:x proximity}}
  .tab-btn{{padding:3px 7px;font-size:8.5px;min-height:22px;border-radius:999px;scroll-snap-align:start}}
  .tcnt{{padding:0 4px;font-size:7.5px;line-height:1.25}}
  .tab-add{{padding:2px 7px;font-size:10px;min-height:22px}}
  #toolbar{{padding:2px 7px;gap:3px;align-items:stretch}}
  #toolbar-actions{{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));width:100%;gap:3px}}
  .tool-btn{{justify-content:center;min-width:0;padding:2px 3px;font-size:8.3px;min-height:22px;border-radius:6px;gap:2px}}
  .tool-dd{{min-width:0}}
  .tool-dd .tool-btn{{width:100%}}
  #srchbox{{width:100%;flex:1 1 100%;max-width:none;min-height:25px;padding:3px 8px 3px 27px;font-size:10px;border-radius:7px}}
  #tblwrap{{min-height:56vh;padding-bottom:2px}}
  #regtbl{{min-width:1040px;font-size:10.5px}}
  #regtbl th{{padding:5px 6px!important;font-size:10px;line-height:1.15}}
  .frow th{{padding:2px 4px!important}}
  .frow input,.frow select{{height:27px;padding:2px 5px;font-size:10px;border-radius:4px}}
  #regtbl td{{padding:4px 6px!important;line-height:1.18;max-width:190px}}
  .chkcell{{width:26px;padding:3px!important}}
  .chkcell input{{width:13px;height:13px}}
  .sr{{font-size:9px;min-width:24px}}
  .sbadge{{padding:2px 7px;font-size:9px;line-height:1.15}}
  #sbar{{padding:2px 8px;font-size:9px;gap:9px;line-height:1.2;justify-content:space-between}}
  #ltr-quickbar{{padding:4px 8px!important;gap:4px!important;display:flex!important;flex-direction:row!important}}
  #ltr-quickbar .tool-btn{{min-height:22px!important;font-size:8.5px!important;padding:2px 6px!important;flex:1 1 0!important}}
}}
@media(max-width:640px){{
  #topbar{{height:32px;padding:0 5px;gap:3px}}
  #topbar .tb-btn{{padding:3px 5px;font-size:9px;min-height:24px}}
  #topbar .topbar-title-full{{display:none}}
  #topbar .topbar-title-short{{display:inline}}
  #topbar .topbar-mark{{font-size:15px}}
  #topbar .topbar-user{{font-size:9px;gap:4px}}
  #topbar .topbar-user-name{{display:none}}
  #tabsbar{{padding:2px 6px;gap:3px}}
  .tab-btn{{padding:3px 7px;font-size:8.5px;min-height:22px}}
  .tab-add{{padding:2px 7px;font-size:10px;min-height:22px}}
  #projbar{{padding:2px 7px;align-items:flex-start;gap:4px}}
  #projbar img{{max-height:32px;max-width:78px}}
  #projbar-main{{display:grid;grid-template-columns:minmax(120px,.9fr) minmax(0,1.1fr);align-items:start;gap:2px 8px;width:auto;flex:1 1 auto}}
  #projbar-primary,#projbar-extra{{display:grid;grid-template-columns:1fr 1fr;gap:1px 6px;width:100%;min-width:0}}
  #projbar-primary .pf,#projbar-extra .pf{{padding:0;border-right:none;border-bottom:none;min-width:0;display:block;flex:none}}
  #projbar-primary .pf-lbl,#projbar-extra .pf-lbl{{display:block;font-size:6px;line-height:.9;color:var(--mu);letter-spacing:.18px}}
  #projbar-primary .pf-val,#projbar-extra .pf-val{{font-size:9px;line-height:1;max-width:none;white-space:normal;display:-webkit-box;-webkit-line-clamp:1;-webkit-box-orient:vertical;overflow:hidden}}
  #projbar-primary .pf.primary:first-child .pf-val{{font-size:10px;font-weight:800;color:var(--pr);text-transform:uppercase;letter-spacing:.3px;-webkit-line-clamp:1}}
  #projbar-primary .pf.primary:last-child{{grid-column:1/-1}}
  #projbar-primary .pf.primary:last-child .pf-val{{font-size:10px;font-weight:800}}
  #projbar-extra{{padding-top:0;border-top:none}}
  #projbar-toggle{{display:none}}
  .proj-edit-btn{{display:none}}
  #toolbar{{padding:2px 6px;align-items:stretch;gap:3px}}
  #toolbar-actions{{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));width:100%;gap:3px}}
  .tool-btn{{justify-content:center;min-width:0;padding:2px 3px;font-size:8px;min-height:21px}}
  #rec-modal .record-modal{{width:96vw!important;max-width:96vw!important}}
  #rec-modal .mbody{{padding:12px}}
  .record-modal-actions{{padding:10px 12px;justify-content:stretch}}
  .record-modal-actions .btn{{flex:1 1 100%;width:100%;min-height:40px}}
  .modal{{width:96vw;max-height:96vh}}
  #srchbox{{width:100%;flex:1 1 100%;max-width:none;padding-top:2px;padding-bottom:2px;font-size:10px;min-height:24px}}
  #regtbl{{min-width:1040px;font-size:10.5px}}
  #regtbl th{{white-space:nowrap!important;word-break:normal!important}}
  #tblwrap{{min-height:56vh}}
  #sbar{{padding:2px 8px;font-size:9px;gap:8px}}
  #ltr-quickbar{{display:flex!important;flex-direction:row!important;padding:3px 6px!important;gap:3px!important}}
  #ltr-quickbar .tool-btn{{flex:1 1 0!important;min-height:21px!important;font-size:8px!important;padding:2px 5px!important}}
  .overview-pr-panel{{overflow:hidden}}
  .pr-analytics-grid{{grid-template-columns:minmax(0,1fr)!important;width:100%}}
  .pr-block{{width:100%;max-width:100%;min-width:0;overflow:hidden}}
  .pr-projects,.pr-trades{{min-height:0!important}}
  .pr-trades > div:last-child{{min-width:0;height:180px;max-height:none!important}}
  #overview-pane-discipline .tbl-wrap{{display:none}}
  #overview-pane-discipline .disc-mobile-list{{display:block}}
}}
@media(max-width:520px){{
  #topbar{{height:31px;padding:0 5px;gap:3px}}
  #topbar .tb-btn{{font-size:8.5px;width:auto;padding:2px 4px;overflow:visible;min-height:23px}}
  #topbar a.tb-btn{{width:auto;font-size:8.5px;padding:2px 5px}}
  #topbar .topbar-user span:last-child{{max-width:50px;font-size:7px!important;padding:2px 4px!important}}
  #projbar{{padding:2px 6px}}
  #projbar img{{max-height:32px;max-width:76px}}
  #projbar-primary,#projbar-extra{{gap:1px 6px}}
  #projbar-primary .pf-lbl,#projbar-extra .pf-lbl{{font-size:6px}}
  #projbar-primary .pf-val,#projbar-extra .pf-val{{font-size:8.8px;line-height:1;-webkit-line-clamp:1}}
  #projbar-primary .pf.primary:last-child .pf-val{{font-size:10px}}
  #tabsbar{{padding:2px 6px}}
  .tab-btn{{padding:3px 7px;font-size:8.5px;gap:3px}}
  #toolbar{{padding:2px 5px}}
  #toolbar-actions{{grid-template-columns:repeat(4,minmax(0,1fr));gap:3px}}
  .tool-btn{{font-size:7.8px;min-height:20px;padding:2px 3px}}
  #srchbox{{min-height:23px;font-size:9.8px}}
  #regtbl{{min-width:980px;font-size:10px}}
  #regtbl th{{padding:4px 5px!important;font-size:9.5px}}
  #regtbl td{{padding:3px 5px!important;line-height:1.15}}
  .frow input,.frow select{{height:25px;font-size:9.5px}}
  #sbar{{font-size:8.5px;gap:6px}}
}}
@media(max-width:480px){{
  #topbar{{height:31px}}
  .tab-btn{{padding:5px 8px;font-size:9px}}
  .tool-btn{{padding:3px 3px;font-size:8.5px}}
  #toolbar-actions{{grid-template-columns:repeat(4,minmax(0,1fr))}}
  #ltr-quickbar .tool-btn{{flex:1 1 0!important}}
  #projbar-extra{{grid-template-columns:1fr 1fr}}
}}
@media(max-width:900px) and (orientation:landscape){{
  #topbar{{height:28px;padding:0 6px;gap:4px}}
  #topbar .topbar-title-full{{display:none}}
  #topbar .topbar-title-short{{display:inline!important;font-size:12px!important}}
  #topbar .topbar-title-short::after{{content:" Register"}}
  #topbar .tb-btn{{min-height:21px;padding:2px 5px;font-size:8px}}
  #topbar .topbar-user span:last-child{{font-size:7px!important;padding:1px 5px!important;max-width:58px}}
  #projbar{{padding:2px 8px;gap:5px;min-height:0;align-items:center}}
  #projbar img{{max-height:30px;max-width:72px}}
  #projbar-main{{gap:1px;display:grid;grid-template-columns:minmax(0,1fr);flex:1 1 auto}}
  #projbar-primary,#projbar-extra{{grid-template-columns:repeat(3,minmax(0,1fr));gap:1px 8px}}
  #projbar-primary .pf-lbl,#projbar-extra .pf-lbl{{font-size:5.8px;line-height:.9;margin:0}}
  #projbar-primary .pf-val,#projbar-extra .pf-val{{font-size:8px;line-height:.95;-webkit-line-clamp:1}}
  #projbar-primary .pf.primary:last-child{{grid-column:auto;border-bottom:none;padding-bottom:0;margin-bottom:0}}
  #projbar-primary .pf.primary:last-child .pf-val{{font-size:8.5px}}
  #tabsbar{{padding:1px 7px;gap:3px}}
  .tab-btn{{min-height:20px;padding:2px 7px;font-size:8px}}
  .tcnt{{font-size:7px;padding:0 4px}}
  .tab-add{{min-height:20px;padding:1px 7px;font-size:9px}}
  #toolbar{{padding:2px 6px;gap:3px;display:grid;grid-template-columns:minmax(0,1fr) minmax(150px,220px);align-items:center}}
  #toolbar-actions{{grid-template-columns:repeat(8,minmax(0,1fr));gap:3px;width:auto}}
  .tool-btn{{min-height:20px;padding:1px 3px;font-size:7.5px;gap:1px}}
  #srchbox{{min-height:22px;padding-top:1px;padding-bottom:1px;font-size:9px;flex:0 0 auto;width:100%;max-width:none}}
  #ltr-quickbar{{display:flex!important;flex-direction:row!important;padding:2px 6px!important;gap:3px!important}}
  #ltr-quickbar .tool-btn{{min-height:19px!important;font-size:7.5px!important;padding:1px 4px!important;flex:1 1 0!important}}
  #tblwrap{{min-height:50vh;height:calc(100vh - 188px)}}
  #sbar{{padding:1px 8px;font-size:8px;line-height:1.1}}
}}
</style></head><body>

<div id="topbar">
  <span class="topbar-mark" style="font-size:20px">📋</span>
  <span class="topbar-title-full" style="font-weight:700;font-size:14px;letter-spacing:.3px">Document Control Register</span>
  <span class="topbar-title-short" style="font-weight:800;font-size:13px;letter-spacing:.35px">DCR</span>
  <div class="sp"></div>
  <a href="/" class="tb-btn">📊 Dashboard</a>
  <button class="tb-btn" onclick="toggleDark()" id="darkBtn" title="Toggle dark mode">🌙</button>
  {btns}
  <span style="color:rgba(255,255,255,.45);padding:0 4px">|</span>
  <span class="topbar-user"><span class="topbar-user-name">👤 {uname}</span>
    <span style="background:{rbg};border-radius:3px;padding:1px 7px;font-size:9px;font-weight:700">{rlbl}</span>
  </span>
</div>

<div id="projbar">
  {logo_html}
  <div id="projbar-main">
    <div id="projbar-primary">{projbar_primary}</div>
    {projbar_toggle}
    <div id="projbar-extra">{projbar_secondary}</div>
  </div>
  {'<button class="proj-edit-btn" onclick="editProject()" style="margin-left:auto;background:var(--pr);color:#fff;border:none;padding:5px 12px;border-radius:var(--rd);cursor:pointer;font-size:11px;font-family:inherit;flex-shrink:0">✏ Edit</button>' if editable else ''}
</div>

<div id="tabsbar">
  {tabs_html}
  {'<button class="tab-add" onclick="addDocType()" title="Add Type">＋</button>' if editable else ''}
</div>

  <div id="toolbar">
    <div id="toolbar-actions">
      {edit_btns}
      <div class="tool-dd" id="exp-dd">
      <button class="tool-btn teal" onclick="toggleExpDD(event)">📥 Export ▾</button>
      <div class="tool-dd-menu hidden" id="exp-menu">
        <button onclick="doExport();closeExpDD()">📊 Excel — This Tab</button>
        <button onclick="doExportAll();closeExpDD()">📊 Excel — All Tabs</button>
        <button onclick="doExportPDF();closeExpDD()">📄 PDF — This Tab</button>
        <button onclick="doExportAllPDF();closeExpDD()">📄 PDF — All Tabs</button>
      </div>
    </div>
    <button class="tool-btn teal" onclick="doPrint()">🖨 Print</button>
      {'<button class="tool-btn teal" onclick="openImport()">📤 Import</button>' if editable else ''}
    </div>
    <input type="text" id="srchbox" placeholder="Search..." oninput="doSearch()">
  </div>
  <div id="ltr-quickbar" style="display:none;padding:8px 10px;border-bottom:1px solid var(--bd);background:#f8fafc;gap:6px;align-items:center;flex-wrap:wrap">
    <span style="font-size:10px;font-weight:700;color:var(--mu);text-transform:uppercase;letter-spacing:.4px">Letter Views</span>
    <button class="tool-btn" type="button" data-ltr-view="" onclick="setLTRQuickView('')">All</button>
    <button class="tool-btn" type="button" data-ltr-view="sent" onclick="setLTRQuickView('sent')">Sent Letters</button>
    <button class="tool-btn" type="button" data-ltr-view="received" onclick="setLTRQuickView('received')">Received Letters</button>
  </div>
  
  <div id="bulkbar">
  <span id="bulkcnt">0 selected</span>
  <button class="btn btn-er btn-sm" onclick="bulkDel()">🗑 Delete</button>
  <button class="btn btn-sm" style="background:rgba(255,255,255,.15);color:#fff;border-color:rgba(255,255,255,.3)" onclick="clearSel()">✕ Cancel</button>
</div>

<div id="main">
  <div id="tblwrap">
    <table id="regtbl"><thead id="thead"></thead><tbody id="tbody"></tbody></table>
    <div class="empty hidden" id="empty" style="display:none"><div style="font-size:48px;margin-bottom:10px">📁</div><p style="color:var(--mu)">No records yet — click ➕ Add to create one</p></div>
  </div>
  <div id="noc-summary" style="display:none;padding:10px 14px;border-top:1px solid var(--bd);background:#f8fafc">
    <div id="noc-summary-inner" style="display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:10px"></div>
  </div>
</div>

<div id="sbar">
  <span id="s-total">Total: 0</span>
  <span id="s-show">Showing: 0</span>
  <span id="s-ov">Overdue: 0</span>
  <span style="margin-left:auto" id="s-clock"></span>
</div>

<!-- ADD/EDIT RECORD -->
<div class="overlay hidden" id="rec-modal">
  <div class="modal record-modal">
    <div class="mhdr"><span id="rec-title">Add Document</span>
      <button class="xbtn" onclick="closeM('rec-modal')">✕</button></div>
    <div class="mbody"><div class="record-form-shell" id="rec-form"></div></div>
    <div class="mfoot record-modal-actions">
      <button class="btn btn-sc" onclick="closeM('rec-modal')">Cancel</button>
      <button class="btn btn-pr" onclick="saveRecord()">Save</button>
    </div>
  </div>
</div>

<!-- LETTER THREAD -->
<div class="overlay hidden" id="thread-modal">
  <div class="modal" style="max-width:760px">
    <div class="mhdr"><span id="thread-title">Letter Thread</span>
      <button class="xbtn" onclick="closeM('thread-modal')">✕</button></div>
    <div class="mbody" id="thread-body"></div>
    <div class="mfoot">
      <button class="btn btn-sc" onclick="closeM('thread-modal')">Close</button>
    </div>
  </div>
</div>

<!-- LETTER TIMELINE -->
<div class="overlay hidden" id="timeline-modal">
  <div class="modal" style="max-width:760px">
    <div class="mhdr"><span id="timeline-title">Letter Timeline</span>
      <button class="xbtn" onclick="closeM('timeline-modal')">✕</button></div>
    <div class="mbody" id="timeline-body"></div>
    <div class="mfoot">
      <button class="btn btn-sc" onclick="closeM('timeline-modal')">Close</button>
    </div>
  </div>
</div>

<!-- PROJECT MODAL -->
<div class="overlay hidden" id="proj-modal">
  <div class="modal" style="max-width:820px">
    <div class="mhdr"><span>🏗 Project Information</span>
      <button class="xbtn" onclick="closeM('proj-modal')">✕</button></div>
    <div class="mbody" id="proj-body"></div>
    <div class="mfoot">
      <button class="btn btn-sc" onclick="closeM('proj-modal')">Cancel</button>
      <button class="btn btn-pr" onclick="saveProject()">💾 Save Project</button>
    </div>
  </div>
</div>

<!-- LISTS -->
<div class="overlay hidden" id="lists-modal">
  <div class="modal">
    <div class="mhdr"><span>📋 Dropdown Lists</span>
      <button class="xbtn" onclick="closeM('lists-modal')">✕</button></div>
    <div class="mbody" id="lists-body"></div>
    <div class="mfoot"><button class="btn btn-sc" onclick="closeM('lists-modal')">Close</button></div>
  </div>
</div>

<!-- ADD DOC TYPE -->
<div class="overlay hidden" id="dt-modal">
  <div class="modal" style="max-width:420px">
    <div class="mhdr"><span>Add Document Type</span>
      <button class="xbtn" onclick="closeM('dt-modal')">✕</button></div>
    <div class="mbody">
      <div class="fgrid">
        <div class="fg"><label>Code</label><input id="dt-code" placeholder="e.g. MS"></div>
        <div class="fg"><label>Name</label><input id="dt-name" placeholder="e.g. Method Statement"></div>
      </div>
    </div>
    <div class="mfoot">
      <button class="btn btn-sc" onclick="closeM('dt-modal')">Cancel</button>
      <button class="btn btn-pr" onclick="saveDocType()">Add</button>
    </div>
  </div>
</div>

<!-- COLUMNS -->
<div class="overlay hidden" id="col-modal">
  <div class="modal">
    <div class="mhdr"><span>⚙ Manage Columns</span>
      <button class="xbtn" onclick="closeM('col-modal')">✕</button></div>
    <div class="mbody" id="col-body"></div>
    <div class="mfoot">
      <button class="btn btn-sc" onclick="openAddCol()">+ Add Column</button>
      <button class="btn btn-pr" onclick="closeM('col-modal');loadRecords()">Done</button>
    </div>
  </div>
</div>

<!-- ADD COLUMN -->
<div class="overlay hidden" id="addcol-modal">
  <div class="modal" style="max-width:460px">
    <div class="mhdr"><span>Add Column</span>
      <button class="xbtn" onclick="closeM('addcol-modal')">✕</button></div>
    <div class="mbody">
      <div class="fgrid">
        <div class="fg full"><label>Column Name</label><input id="col-name" placeholder="e.g. Review Duration"></div>
        <div class="fg full"><label>Type</label>
          <select id="col-type" onchange="onColType(this.value)">
            <option value="text">📝 Text</option>
            <option value="number">🔢 Number</option>
            <option value="date">📅 Date</option>
            <option value="dropdown">📋 Dropdown</option>
            <option value="link">🔗 Hyperlink</option>
            <option value="duration_calc">⏱ Duration (working days)</option>
          </select>
        </div>
        <div class="fg full" id="cg-list" style="display:none">
          <label>Dropdown Source</label><select id="col-list"></select></div>
        <div class="fg full" id="cg-ds" style="display:none">
          <label>Start Date Column</label><select id="col-ds"></select></div>
        <div class="fg full" id="cg-de" style="display:none">
          <label>End Date Column</label><select id="col-de"></select></div>
        <div class="fg full" id="cg-info" style="display:none">
          <div style="background:#f0f4fa;border-radius:6px;padding:9px 12px;font-size:11px;color:#1a3a5c">
            ⏱ Working days between two date columns — excludes Fridays + Egyptian holidays
          </div>
        </div>
      </div>
    </div>
    <div class="mfoot">
      <button class="btn btn-sc" onclick="closeM('addcol-modal')">Cancel</button>
      <button class="btn btn-pr" onclick="saveAddCol()">Add</button>
    </div>
  </div>
</div>

<!-- IMPORT -->
<div class="overlay hidden" id="import-modal">
  <div class="modal" style="max-width:480px">
    <div class="mhdr"><span>📤 Import Excel / CSV</span>
      <button class="xbtn" onclick="closeM('import-modal')">✕</button></div>
    <div class="mbody">
      <p style="font-size:12px;color:var(--mu);margin-bottom:12px">Select .xlsx or .csv file. Headers must match register columns. For full-workbook import, Excel sheets will be matched to document types automatically.</p>
      <input type="file" id="imp-file" accept=".csv,.xlsx,.xls" style="font-size:12px">
    </div>
    <div class="mfoot">
      <button class="btn btn-sc" onclick="closeM('import-modal')">Cancel</button>
      <button class="btn btn-sc" id="imp-project-btn" onclick="doImportProject()">Import Full Workbook</button>
      <button class="btn btn-pr" onclick="doImport()">Import</button>
    </div>
  </div>
</div>

<!-- ADMIN MODAL (register page) -->
<div class="overlay hidden" id="admin-modal">
  <div class="modal" style="max-width:780px">
    <div class="mhdr"><span>⚙ Admin Panel</span><button class="xbtn" onclick="closeM('admin-modal')">✕</button></div>
    <div class="mbody" id="admin-body"></div>
    <div class="mfoot"><button class="btn btn-sc" onclick="closeM('admin-modal')">Close</button></div>
  </div>
</div>

{SHARED_JS}
<script>
const PID='{pid}', ROLE='{role}', CAN_EDIT={'true' if editable else 'false'};
const SC={{...{sc_json},
  'Approved':['166534','ffffff'],
  'Accepted & to proceed with Part C':['bbf7d0','166534'],
  'Rejected':['fecaca','7f1d1d'],
  'Cancelled':['e5e7eb','475569'],
  'Under Review':['fef9c3','854d0e'],
  'Information Required':['e0e7ff','312e81'],
  'Pending':['fde68a','92400e']
}};
const PROJ_FIELDS=[['code','Code'],['name','Project Name'],['startDate','Start Date'],['endDate','End Date'],
  ['client','Client'],['landlord','Landlord'],['pmo','PMO'],['mainConsultant','Consultant'],
  ['mepConsultant','MEP'],['contractor','Contractor']];
const state={{tab:null,cols:[],recs:null,visibleRows:[],selectedRowId:null,sortCol:null,sortDir:'asc',filters:{{}},editId:null,lists:{{}},prItemsCache:{{}}}};

function isPRTab(){{
  const dt=state.dtList?.find(d=>d.id===state.tab)||{{}};
  const code=(dt.code||dt.id||'').toString().toUpperCase();
  const name=(dt.name||'').toString().toLowerCase();
  return code==='PR' || name.includes('requisition') || name.includes('purchase request');
}}

function isNOCTab(){{
  const dt=state.dtList?.find(d=>d.id===state.tab)||{{}};
  const code=(dt.code||dt.id||'').toString().toUpperCase();
  const name=(dt.name||'').toString().toLowerCase();
  return code==='NOC' || name.includes('notice of change');
}}

function isLTRTab(){{
  const dt=state.dtList?.find(d=>d.id===state.tab)||{{}};
  const code=(dt.code||dt.id||'').toString().toUpperCase();
  const name=(dt.name||'').toString().toLowerCase();
  return code==='LTR' || name.includes('letter');
}}

function isHiddenSystemListName(name){{
  const ln=String(name||'').trim().toLowerCase();
  return ['letter_direction','letter_status'].includes(ln);
}}

function projectHasLTR(){{
  return (state.dtList||[]).some(dt=>{{
    const code=String(dt?.code||dt?.id||'').trim().toUpperCase();
    const name=String(dt?.name||'').trim().toLowerCase();
    return code==='LTR' || name.includes('letter') || name.includes('correspondence');
  }});
}}

function ensureLTRProjectLists(){{
  if(!projectHasLTR())return;
  if(!state.lists||typeof state.lists!=='object')state.lists={{}};
  if(!Object.prototype.hasOwnProperty.call(state.lists,'correspondence_parties'))state.lists.correspondence_parties=[];
}}

function isLTRSingleSelectField(col){{
  if(!isLTRTab())return false;
  const key=String(col?.col_key||'').trim().toLowerCase();
  return ['direction','fromparty','toparty'].includes(key);
}}

function isLTRExcludedField(col){{
  if(!isLTRTab())return false;
  const role=getLTRFieldRole(col);
  const key=String(col?.col_key||'').trim().toLowerCase();
  return ['description','fileLocation','status'].includes(role)||['description','filelocation','status'].includes(key);
}}

function isLTRInternalField(col){{
  if(!isLTRTab())return false;
  const key=String(col?.col_key||col||'').trim().toLowerCase();
  return ['parentletterid'].includes(key);
}}

function normLTRText(v){{
  return String(v||'').trim().toLowerCase().replace(/[^a-z0-9]+/g,'');
}}

function getLTRFieldRole(col){{
  const key=normLTRText(col?.col_key||'');
  const label=normLTRText(col?.label||'');
  if(key==='docno'||label==='letterref')return 'docNo';
  if(key==='title'||key==='subject'||label==='subject')return 'title';
  if(key==='description'||label==='description')return 'description';
  if(key==='filelocation'||label.includes('attachment')||label.includes('filelocation'))return 'fileLocation';
  if(key==='remarks'||label==='remarks')return 'remarks';
  if(key==='direction'||label==='direction')return 'direction';
  if(key==='fromparty'||key==='from'||label==='fromparty'||label==='from')return 'fromParty';
  if(key==='toparty'||key==='to'||label==='toparty'||label==='to')return 'toParty';
  if(key==='issueddate'||label==='issuedate')return 'issuedDate';
  if(key==='receiveddate'||label==='receiveddate')return 'receivedDate';
  if(key==='parentletterid'||label==='parentletterid')return 'parentLetterId';
  if(key==='parentletterref'||key==='responseref'||label==='responseref'||label==='parentletter')return 'parentLetterRef';
  if(key==='status'||label==='status')return 'status';
  return '';
}}

function getLTRRegisterKeys(){{
  return ['docNo','direction','fromParty','toParty','issuedDate','receivedDate','title','parentLetterRef'];
}}

function getLTRFormKeys(){{
  return ['docNo','title','direction','fromParty','toParty','issuedDate','receivedDate','remarks','parentLetterId','parentLetterRef'];
}}

function getLTRColsByRole(cols, roles, visibleOnly=false){{
  const out=[];
  const used=new Set();
  for(const role of roles){{
    const col=(cols||[]).find(c=>{{
      if(visibleOnly&&!c.visible)return false;
      const r=getLTRFieldRole(c);
      if(r!==role||used.has(c.col_key))return false;
      return role!=='parentLetterId';
    }});
    if(col){{
      out.push(col);
      used.add(col.col_key);
    }}
  }}
  return out;
}}

function getLTRColForRole(cols, role){{
  return (cols||[]).find(c=>getLTRFieldRole(c)===role)||null;
}}

function getLTRValue(row, cols, role){{
  const matches=(cols||[]).filter(c=>getLTRFieldRole(c)===role);
  for(const col of matches){{
    const val=row?.[col.col_key];
    if(String(val??'').trim())return val;
  }}
  const first=matches[0];
  return first?row?.[first.col_key]:'';
}}

function formatLTRParentLabel(opt){{
  const doc=String(opt?.doc_no||'').trim();
  const subject=String(opt?.subject||'').trim();
  if(doc&&subject)return `${{doc}} | ${{subject}}`;
  return doc||subject||'Untitled letter';
}}

function formatLTRParentMeta(opt){{
  if(!opt)return 'No parent letter selected';
  const parts=[];
  if(opt.direction)parts.push(opt.direction);
  if(opt.from_party)parts.push('From: '+opt.from_party);
  if(opt.to_party)parts.push('To: '+opt.to_party);
  if(opt.date)parts.push(opt.date);
  return parts.join(' • ')||'No additional details';
}}

function getLTRVisibleColsFallback(cols, roles){{
  const picked=getLTRColsByRole(cols, roles, true);
  return picked.length?picked:(cols||[]).filter(c=>c.visible&&!isLTRInternalField(c)&&!isLTRExcludedField(c));
}}

function getLTRFormColsFallback(cols, roles){{
  const picked=getLTRColsByRole(cols, roles, false);
  return picked.length?picked:(cols||[]).filter(c=>!isLTRInternalField(c)&&!isLTRExcludedField(c));
}}

function updateLTRQuickBar(){{
  const bar=document.getElementById('ltr-quickbar');
  if(!bar)return;
  const on=isLTRTab();
  bar.style.display=on?'flex':'none';
  if(!on)delete state.filters._ltrQuick;
  bar.querySelectorAll('[data-ltr-view]').forEach(btn=>{{
    const active=(state.filters._ltrQuick||'')===btn.dataset.ltrView;
    btn.style.background=active?'var(--pr)':'';
    btn.style.color=active?'#fff':'';
    btn.style.borderColor=active?'var(--pr)':'';
  }});
}}

function setLTRQuickView(view){{
  state.filters._ltrQuick=view||'';
  updateLTRQuickBar();
  renderRows();
}}

function getPrDetailsColKey(){{
  const cols=state.cols||[];
  const norm=v=>String(v||'').toLowerCase().replace(/[^a-z0-9]+/g,'');
  for(const c of cols){{
    const key=String(c.col_key||'');
    const label=String(c.label||'').trim();
    if(label==='PR Details'||key==='prDetails'||key==='pr_details')return c.col_key;
  }}
  for(const c of cols){{
    const nk=norm(c.col_key),nl=norm(c.label);
    if(nk==='prdetails'||nl==='prdetails')return c.col_key;
  }}
  for(const c of cols){{
    const nk=norm(c.col_key),nl=norm(c.label);
    if(nk.includes('pr')&&nk.includes('detail'))return c.col_key;
    if(nl.includes('pr')&&nl.includes('detail'))return c.col_key;
  }}
  return null;
}}

function getPrManualDetails(row){{
  const detailsKey=getPrDetailsColKey();
  return detailsKey?String(row?.[detailsKey]||'').replace(/\s+/g,' ').trim():'';
}}

function getPrAutoSummary(row){{
  const items=state.prItemsCache[row?._id]||[];
  if(items.length){{
    const names=items
      .filter(it=>String(it?.row_type||'item').toLowerCase()!=='header')
      .map(it=>String(it?.item_name||'').trim())
      .filter(Boolean);
    if(names.length){{
      let summary=names.slice(0,2).join(', ');
      if(names.length>2)summary+=' ...';
      return summary;
    }}
  }}
  return '';
}}

function getPrSummary(row){{
  const manual=getPrManualDetails(row);
  if(manual)return manual;
  const auto=getPrAutoSummary(row);
  if(!auto)return '';
  return auto.length>80?auto.slice(0,77)+'...':auto;
}}

function escHtml(v){{
  return String(v==null?'':v)
    .replaceAll('&','&amp;')
    .replaceAll('<','&lt;')
    .replaceAll('>','&gt;')
    .replaceAll('"','&quot;')
    .replaceAll("'",'&#39;');
}}

// Init
(async()=>{{
  await Promise.all([loadDTs(), loadLists()]);
  updateClock(); setInterval(updateClock,60000);
}})();

function updateClock(){{document.getElementById('s-clock').textContent=new Date().toLocaleString('en-GB');}}

async function loadDTs(keepTab=false){{
  const dts=await apiFetch('/api/doc_types/'+PID); if(!dts)return;
  renderTabs(dts);
  ensureLTRProjectLists();
  await refreshCounts();
  if(!keepTab){{
    const tab=new URLSearchParams(location.search).get('tab');
    if(dts.length) switchTab(tab&&dts.find(d=>d.id===tab)?tab:dts[0].id);
  }}
}}

async function refreshCounts(){{
  const cnts=await apiFetch('/api/counts/'+PID); if(!cnts)return;
  Object.entries(cnts).forEach(([id,n])=>{{const el=document.getElementById('cnt-'+id);if(el)el.textContent=n;}});
}}

async function loadLists(force=false){{
  if(!force&&Object.keys(state.lists).length) return;
  const d=await apiFetch('/api/lists/'+PID); if(d) state.lists=d;
  ensureLTRProjectLists();
}}

function renderTabs(dts){{
  state.dtList=dts;
  const bar=document.getElementById('tabsbar');
  bar.querySelectorAll('.tab-btn').forEach(b=>b.remove());
  const addBtn=bar.querySelector('.tab-add');
  dts.forEach(dt=>{{
    const btn=document.createElement('button');
    btn.className='tab-btn'+(dt.id===state.tab?' active':'');
    btn.dataset.id=dt.id; btn.title=dt.name;
    btn.innerHTML=`<span>${{dt.code}}</span><span class="tcnt" id="cnt-${{dt.id}}">0</span>`;
    btn.onclick=()=>switchTab(dt.id);
    if(CAN_EDIT) btn.oncontextmenu=e=>{{e.preventDefault();tabMenu(dt.id,e);}};
    bar.insertBefore(btn,addBtn||null);
  }});
}}

function switchTab(id){{
  state.tab=id;state.recs=null;state.visibleRows=[];state.selectedRowId=null;state.filters={{}};state.sortCol=null;
  document.getElementById('srchbox').value='';
  document.querySelectorAll('.tab-btn').forEach(b=>b.classList.toggle('active',b.dataset.id===id));
  updateLTRQuickBar();
  loadRecords();
}}

function isMobileRegister(){{
  return window.matchMedia&&window.matchMedia('(max-width: 768px)').matches;
}}

function mobileColumnWidth(col){{
  if(!isMobileRegister()||!col)return 0;
  const key=String(col.col_key||'').toLowerCase();
  const label=String(col.label||'').toLowerCase();
  const text=(key+' '+label).replace(/[_./-]+/g,' ');
  const role=isLTRTab()?getLTRFieldRole(col):'';
  if(role==='title'||role==='subject'||role==='description'||text.includes('subject'))return 230;
  if(role==='fromParty'||role==='toParty')return 112;
  if(role==='parentLetterRef'||role==='parentLetterId')return 138;
  if(role==='direction')return 82;
  if(role==='issuedDate'||role==='receivedDate')return 106;
  if(text.includes('document no')||key==='docno'||key==='doc_no')return 152;
  if(text.includes('title')||text.includes('description')||text.includes('scope'))return 220;
  if(text.includes('remarks')||text.includes('comment')||text.includes('note'))return 180;
  if(text.includes('prepared')||text.includes('company')||text.includes('client')||text.includes('consultant')||text.includes('contractor')||text.includes('person'))return 132;
  if(text.includes('from')||text.includes(' to ')||text==='to'||key==='to'||key==='from')return 112;
  if(text.includes('ms ref')||text.includes('dwg')||text.includes('item ref')||text.includes('reference'))return 150;
  if(text.includes('discipline'))return 112;
  if(text.includes('sub trade')||text.includes('sub-trade'))return 122;
  if(text.includes('status'))return 118;
  if(text.includes('date'))return 104;
  if(text.includes('duration'))return 72;
  if(text.includes('file')||text.includes('link'))return 82;
  if(text.includes('floor')||text.includes('location'))return 116;
  if(text.includes('revision')||text.includes('rev'))return 82;
  if(text.includes('qty')||text.includes('quantity')||text.includes('amount'))return 88;
  return 108;
}}

function mobileColumnStyle(col){{
  const w=mobileColumnWidth(col);
  return w?`width:${{w}}px;min-width:${{w}}px;`: '';
}}

function tabMenu(id,e){{
  const old=document.getElementById('tabctx');if(old)old.remove();
  const m=document.createElement('div');m.id='tabctx';
  m.style.cssText=`position:fixed;left:${{e.clientX}}px;top:${{e.clientY}}px;background:#fff;
    border:1px solid var(--bd);border-radius:6px;box-shadow:0 4px 12px rgba(0,0,0,.15);z-index:9999;min-width:140px;overflow:hidden`;
  const dt=state.dtList?.find(d=>d.id===id)||{{}};
  const idx=(state.dtList||[]).findIndex(d=>d.id===id);
  m.innerHTML=`
    <div onclick="renameDT('${{id}}','${{dt.code||id}}','${{(dt.name||'').replace(/'/g,\"\\\\'\")}}')"
      style="padding:8px 14px;cursor:pointer;font-size:12px"
      onmouseover="this.style.background='#f0f4f8'" onmouseout="this.style.background=''">✏ Rename</div>
    <div onclick="moveDT('${{id}}',-1)" style="padding:8px 14px;cursor:${{idx>0?'pointer':'not-allowed'}};font-size:12px;opacity:${{idx>0?'1':'.45'}}"
      onmouseover="if(${{idx>0}})this.style.background='#f0f4f8'" onmouseout="this.style.background=''">⬅ Move Left</div>
    <div onclick="moveDT('${{id}}',1)" style="padding:8px 14px;cursor:${{idx>=0&&idx<(state.dtList||[]).length-1?'pointer':'not-allowed'}};font-size:12px;opacity:${{idx>=0&&idx<(state.dtList||[]).length-1?'1':'.45'}}"
      onmouseover="if(${{idx>=0&&idx<(state.dtList||[]).length-1}})this.style.background='#f0f4f8'" onmouseout="this.style.background=''">➡ Move Right</div>
    <div onclick="delDT('${{id}}')" style="padding:8px 14px;cursor:pointer;font-size:12px;color:#ef4444"
      onmouseover="this.style.background='#fef2f2'" onmouseout="this.style.background=''">🗑 Delete Type</div>`;
  document.body.appendChild(m);
  setTimeout(()=>document.addEventListener('click',()=>m.remove(),{{once:true}}),10);
}}

async function renameDT(id,oldCode,oldName){{
  const newCode=prompt('Tab code (short):',oldCode);
  if(newCode===null)return;
  const newName=prompt('Full name:',oldName);
  if(newName===null)return;
  await apiFetch('/api/doc_types/'+PID+'/'+id,{{method:'PATCH',
    body:JSON.stringify({{code:newCode.trim().toUpperCase(),name:newName.trim()}})}});
  await loadDTs(true);toast('✔ Renamed','ok');
}}

async function delDT(id){{
  if(!confirm('Delete type and ALL its records?'))return;
  await apiFetch('/api/doc_types/'+PID+'/'+id,{{method:'DELETE'}});
  if(state.tab===id)state.tab=null;
  await loadDTs(); toast('Deleted','wa');
}}

async function moveDT(id,delta){{
  const order=(state.dtList||[]).map(d=>d.id);
  const idx=order.indexOf(id);
  const next=idx+delta;
  if(idx<0||next<0||next>=order.length)return;
  [order[idx],order[next]]=[order[next],order[idx]];
  await apiFetch('/api/doc_types/'+PID+'/reorder',{{method:'POST',body:JSON.stringify({{order}})}});
  await loadDTs(true);
  toast('✔ Tab order updated','ok');
}}

async function loadRecords(){{
  if(!state.tab)return;
  const search=document.getElementById('srchbox').value.trim();
  const [data, widths]=await Promise.all([
    apiFetch('/api/records/'+PID+'/'+state.tab+(search?'?search='+encodeURIComponent(search):'')),
    apiFetch('/api/col_width/'+PID+'/'+state.tab)
  ]);
  if(!data)return;
  state.recs=data.records; state.allTabCols=data.columns||[]; state.cols=data.columns.filter(c=>c.visible);
  if(isLTRTab())state.cols=state.cols.filter(c=>!isLTRInternalField(c)&&!isLTRExcludedField(c));
  if(!state.recs.some(r=>String(r._id)===String(state.selectedRowId)))state.selectedRowId=null;
  state.prItemsCache=data.pr_items_map||{{}};
  state.colWidths=widths||{{}};
  const cnt=document.getElementById('cnt-'+state.tab); if(cnt)cnt.textContent=data.count;
  updateLTRQuickBar();
  buildHead(); requestAnimationFrame(()=>initRz()); renderRows();
}}

function buildHead(){{
  const head=document.getElementById('thead'); head.innerHTML='';
  const hr=document.createElement('tr');
  const chk=document.createElement('th'); chk.className='chkcell';
  if(CAN_EDIT)chk.innerHTML='<input type="checkbox" id="chkall" onchange="selAll(this.checked)">';
  hr.appendChild(chk);
  const sr=document.createElement('th');sr.textContent='Sr.';
  const _srW=state.colWidths&&state.colWidths['_sr'];
  sr.dataset.key='_sr';
  sr.style.cssText=(isMobileRegister()?'width:26px;min-width:26px;max-width:26px;':(_srW?'width:'+_srW+'px;min-width:'+_srW+'px;max-width:'+_srW+'px;':'width:34px;'))
    +'white-space:normal;word-break:normal;cursor:default';
  hr.appendChild(sr);
  state.cols.forEach(col=>{{
    const th=document.createElement('th');th.dataset.key=col.col_key;
    const w=state.colWidths&&state.colWidths[col.col_key];
    const labelTxt=String(col.label||'').toLowerCase();
    const minW=labelTxt.includes('sub-trade')?150:
      labelTxt.includes('discipline')?130:
      labelTxt.includes('title')?220:
      labelTxt.includes('subject')?220:0;
    const mw=mobileColumnWidth(col);
    if(mw)th.style.cssText='width:'+mw+'px;min-width:'+mw+'px;white-space:normal;word-break:normal';
    else if(w)th.style.cssText='width:'+w+'px;min-width:'+w+'px;max-width:'+w+'px;white-space:normal;word-break:normal';
    else th.style.cssText=(minW?'min-width:'+minW+'px;':'')+'white-space:normal;word-break:normal';
    const sortInd=state.sortCol===col.col_key?(state.sortDir==='asc'?' ↑':' ↓'):'';
    th.innerHTML='<span class="th-lbl">'+col.label+sortInd+'</span>';
    if(!['auto_date','auto_num'].includes(col.col_type))th.onclick=()=>sortBy(col.col_key);
    else th.style.cursor='default';
    hr.appendChild(th);
  }});
  const at=document.createElement('th');at.textContent='Actions';at.style.cssText='width:64px;cursor:default';hr.appendChild(at);
  head.appendChild(hr);
  const fr=document.createElement('tr');fr.className='frow';
  fr.appendChild(document.createElement('th'));fr.appendChild(document.createElement('th'));
  state.cols.forEach(col=>{{
    const th=document.createElement('th');
    const mw=mobileColumnWidth(col);
    if(mw)th.style.cssText='width:'+mw+'px;min-width:'+mw+'px';
    if(['auto_date','auto_num'].includes(col.col_type)){{fr.appendChild(th);return;}}
    if(col.col_type==='dropdown'&&col.list_name){{
      const sel=document.createElement('select');
      const ltrRole=isLTRTab()?getLTRFieldRole(col):'';
      const listName=ltrRole==='fromParty'||ltrRole==='toParty'
        ? 'correspondence_parties'
        : (ltrRole==='direction' ? 'letter_direction' : (ltrRole==='status' ? 'letter_status' : col.list_name));
      sel.innerHTML='<option value="">All</option>'+(state.lists[listName]||[]).map(o=>`<option ${{state.filters[col.col_key]===o?'selected':''}}>${{o}}</option>`).join('');
      sel.onchange=()=>{{state.filters[col.col_key]=sel.value;renderRows();}};
      th.appendChild(sel);
    }}else{{
      const inp=document.createElement('input');inp.value=state.filters[col.col_key]||'';
      inp.oninput=()=>{{state.filters[col.col_key]=inp.value;renderRows();}};
      th.appendChild(inp);
    }}
    fr.appendChild(th);
  }});
  fr.appendChild(document.createElement('th'));
  head.appendChild(fr);
}}

function parseDocNo(docNo){{
  const raw=String(docNo||'').trim();
  const revMatch=raw.match(/\\bREV\\s*([0-9]+)\\b/i);
  const rev=revMatch?parseInt(revMatch[1],10):0;
  const withoutRev=raw.replace(/\\s*\\bREV\\s*[0-9]+\\b\\s*$/i,'').trim();
  const m=withoutRev.match(/^(.*?)-(\\d+)$/);
  if(m)return{{prefix:m[1],num:parseInt(m[2],10),width:m[2].length,rev,base:withoutRev,raw}};
  return{{prefix:raw,num:0,width:3,rev,base:withoutRev||raw,raw}};
}}

function buildDocNoFromParts(p,num,rev){{
  return `${{p.prefix}}-${{String(num).padStart(p.width||3,'0')}} REV${{String(rev).padStart(2,'0')}}`;
}}

function incrementRevisionNumber(docNo){{
  const raw=String(docNo||'').trim();
  const m=raw.match(/\\bREV\\s*([0-9]+)\\b/i);
  if(!m)return '';
  const next=String(parseInt(m[1],10)+1).padStart(m[1].length,'0');
  return raw.replace(/\\bREV\\s*[0-9]+\\b/i,'REV'+next);
}}

function suggestNextDocNoFromRows(rows){{
  const parsed=(rows||[]).map(r=>parseDocNo(r.docNo||'')).filter(p=>p&&p.num>0);
  if(!parsed.length)return '';
  const families=new Map();
  parsed.forEach(p=>{{
    const key=String(p.prefix||'').toLowerCase();
    if(!families.has(key))families.set(key,[]);
    families.get(key).push(p);
  }});
  if(families.size!==1)return '';
  const family=[...families.values()][0];
  if(!family.length)return '';
  family.sort((a,b)=>a.num-b.num);
  const base=family[family.length-1];
  return buildDocNoFromParts(base,base.num+1,0);
}}

function suggestNextDocNoFromVisibleRows(){{
  return suggestNextDocNoFromRows(state.visibleRows||[]);
}}

function sortByDocNo(rows){{
  // Group by base (prefix+num), then sort revs within group
  const groups={{}};
  rows.forEach(r=>{{
    const p=parseDocNo(r.docNo||'');
    const key=p.prefix+'-'+String(p.num).padStart(6,'0');
    if(!groups[key])groups[key]={{baseNum:p.num,prefix:p.prefix,rows:[]}};
    groups[key].rows.push({{...r,_parsedNum:p.num,_parsedRev:p.rev}});
  }});
  // Sort groups by base number, then rows within group by rev
  const sorted=[];
  Object.values(groups)
    .sort((a,b)=>a.prefix.localeCompare(b.prefix)||a.baseNum-b.baseNum)
    .forEach(g=>{{
      g.rows.sort((a,b)=>a._parsedRev-b._parsedRev);
      g.rows.forEach(r=>sorted.push(r));
    }});
  return sorted;
}}

function validateDocNo(docNo,existingRecs,editId){{
  if(!docNo)return'Document No. is required';
  const p=parseDocNo(docNo);
  if(!p.num&&!docNo.includes('-'))return'Invalid format. Use: CODE-001 REV00';
  // Duplicate check
  const dup=existingRecs.find(r=>r._id!==editId&&(r.docNo||'').toLowerCase()===docNo.toLowerCase());
  if(dup)return'Document No. already exists: '+docNo;
  // REV check: base must exist
  if(p.rev>0){{
    const hasBase=existingRecs.some(r=>{{
      const bp=parseDocNo(r.docNo||'');
      return bp.prefix===p.prefix&&bp.num===p.num&&bp.rev===0&&r._id!==editId;
    }});
    if(!hasBase)return'Cannot add REV'+String(p.rev).padStart(2,'0')+' — REV00 not found for this document';
  }}
  // Sequence gap check (warning only)
  if(p.rev===0&&p.num>1){{
    const nums=existingRecs
      .filter(r=>{{const bp=parseDocNo(r.docNo||'');return bp.prefix===p.prefix&&bp.rev===0&&r._id!==editId;}})
      .map(r=>parseDocNo(r.docNo||'').num);
    const maxExist=nums.length?Math.max(...nums):0;
    if(p.num>maxExist+1)return'GAP: '+p.prefix+'-'+String(maxExist+1).padStart(3,'0')+' to '+p.prefix+'-'+String(p.num-1).padStart(3,'0')+' are missing';
  }}
  return null;
}}

function checkGap(docNo,recs,editId){{
  const p=parseDocNo(docNo);if(!p||p.rev>0)return null;
  // Find max existing base num for same prefix
  let maxN=0;
  recs.forEach(r=>{{
    if(r._id===editId)return;
    const rp=parseDocNo(r.docNo||'');
    if(rp&&rp.prefix===p.prefix&&rp.rev===0)maxN=Math.max(maxN,rp.num);
  }});
  if(p.num>maxN+1)return`Sequence gap detected: Next expected is ${{p.prefix}}-${{String(maxN+1).padStart(3,'0')}} REV00, but you entered ${{p.prefix}}-${{String(p.num).padStart(3,'0')}} REV00`;
  return null;
}}

function getSelectedRegisterRow(){{
  if(!state.selectedRowId)return null;
  return (state.recs||[]).find(r=>String(r._id)===String(state.selectedRowId))||null;
}}

function setSelectedRegisterRow(id){{
  state.selectedRowId=id||null;
  document.querySelectorAll('#regtbl tbody tr.row-selected').forEach(tr=>tr.classList.remove('row-selected'));
  if(state.selectedRowId){{
    const tr=document.querySelector(`#regtbl tbody tr[data-rec-id="${{CSS.escape(String(state.selectedRowId))}}"]`);
    if(tr)tr.classList.add('row-selected');
  }}
}}

function isInteractiveRowTarget(target){{
  return !!target.closest('button,a,input,select,textarea,label,.pr-items-row');
}}

function renderRows(){{
  const body=document.getElementById('tbody');body.innerHTML='';
  const isPrTab=isPRTab();
  const isLtrTab=isLTRTab();
  const prDetailsKey=isPrTab?getPrDetailsColKey():null;
  let rows=state.recs.filter(r=>{{
    for(const[k,v]of Object.entries(state.filters)){{
      if(k.startsWith('_'))continue;
      if(v&&!String(r[k]||'').toLowerCase().includes(v.toLowerCase()))return false;
    }}
    if(isLtrTab){{
      const quick=String(state.filters._ltrQuick||'').toLowerCase();
      const direction=String(getLTRValue(r,state.allTabCols,'direction')||'').toLowerCase();
      if(quick==='sent'&&direction!=='sent')return false;
      if(quick==='received'&&direction!=='received')return false;
    }}
    return true;
  }});
  if(state.sortCol){{
    rows.sort((a,b)=>{{
      const va=String(a[state.sortCol]||'').toLowerCase(),vb=String(b[state.sortCol]||'').toLowerCase();
      return state.sortDir==='asc'?(va>vb?1:-1):(va<vb?1:-1);
    }});
  }}else{{
    rows=sortByDocNo(rows);
  }}
  state.visibleRows=rows;
  const emptyEl=document.getElementById('empty');
  const tblEl=document.getElementById('regtbl');
  if(rows.length===0){{
    emptyEl.style.display='none';
    tblEl.style.display='';
    const tr=document.createElement('tr');
    const td=document.createElement('td');
    td.colSpan=state.cols.length+3;
    td.style.cssText='padding:28px 12px;text-align:center;color:var(--mu);font-size:12px';
    td.textContent=(state.recs&&state.recs.length)?'No records found':'No records found';
    tr.appendChild(td);
    body.appendChild(tr);
  }}else{{
    emptyEl.style.display='none';
    tblEl.style.display='';
  }}
  if(rows.length===0){{
    const ov=state.recs.filter(r=>r._overdue).length;
    document.getElementById('s-total').textContent='Total: '+state.recs.length;
    document.getElementById('s-show').textContent='Showing: 0';
    document.getElementById('s-ov').textContent='Overdue: '+ov;
    renderNocSummary(rows);
    return;
  }}
  let sr=1;
  rows.forEach((row,idx)=>{{
    const tr=document.createElement('tr');
    tr.dataset.recId=row._id;
    if(String(state.selectedRowId||'')===String(row._id))tr.classList.add('row-selected');
    tr.onclick=e=>{{
      if(isInteractiveRowTarget(e.target))return;
      setSelectedRegisterRow(row._id);
    }};
    if(row._overdue)tr.classList.add('ov');
    else if(row._isRev)tr.classList.add('rv');
    else if(idx%2===1)tr.classList.add('alt');
    const tc=document.createElement('td');tc.className='chkcell';
    if(CAN_EDIT){{const cb=document.createElement('input');cb.type='checkbox';cb.dataset.id=row._id;cb.onchange=updBulk;tc.appendChild(cb);}}
    tr.appendChild(tc);
    const tsr=document.createElement('td');tsr.className='sr';
    const _srTdW=state.colWidths&&state.colWidths['_sr'];
    if(isMobileRegister())tsr.style.cssText='width:26px;min-width:26px;max-width:26px';
    else if(_srTdW)tsr.style.cssText='width:'+_srTdW+'px;min-width:'+_srTdW+'px;max-width:'+_srTdW+'px';
    tsr.textContent=row._isRev?'':sr;tr.appendChild(tsr);
    state.cols.forEach(col=>{{
      const td=document.createElement('td');const key=col.col_key;let val='';
      const mcs=mobileColumnStyle(col);
      if(mcs)td.style.cssText=mcs;
      const ltrRole=isLtrTab?getLTRFieldRole(col):'';
      const k=key.toLowerCase();
      const longTextMeta=getLongTextMeta(col);
      if(longTextMeta||isFloorField(col))td.classList.add('mlcell');
      if(key==='expectedReplyDate'){{val=row._expectedReplyDate||'';if(row._overdue&&val)td.classList.add('ovdate');}}
      else if(key==='duration')val=row._duration||'';
      else if(col.col_type==='duration_calc'){{
        const[ds,de]=(col.list_name||'issuedDate,actualReplyDate').split(',');
        val=calcWD(row[ds.trim()]||'',row[de.trim()]||'');
      }}
      else if(key==='issuedDate')val=row._issuedFmt||'';
      else if(key==='actualReplyDate')val=row._replyFmt||'';
      else if(col.col_type==='date'||col.col_type==='auto_date'){{
        val=row['_fmt_'+key]||row[key]||'';  // use pre-formatted version
      }}
      else if(key==='status'||isStatusLikeField(col)){{
        val=row[key]||((isLtrTab&&ltrRole)?getLTRValue(row,state.allTabCols,ltrRole):'');
        if(val){{
          td.innerHTML=val.split(',').map(s=>{{s=s.trim();const[bg,fg]=SC[s]||['e5e7eb','374151'];
            return `<span class="sbadge" style="background:#${{bg}};color:#${{fg}}">${{s}}</span>`;}}).join('');
          tr.appendChild(td);return;
        }}
      }}
      else if(key==='fileLocation'){{
        const url=row[key]||'';
        if(url){{td.innerHTML=`<a class="flink" href="${{url}}" target="_blank">View</a>`;tr.appendChild(td);return;}}
      }}
      else if(isPrTab&&prDetailsKey&&key===prDetailsKey)val=getPrSummary(row);
      else val=String(row[key]||'');
      if(isLtrTab&&ltrRole&&!String(val||'').trim())val=String(getLTRValue(row,state.allTabCols,ltrRole)||'');
      let displayVal=formatDisplayValue(col,val);
      if(typeof displayVal==='string'&&longTextMeta){{
        const NL=String.fromCharCode(10);
        displayVal=displayVal
          .replaceAll(' / ',NL)
          .replaceAll(' /',NL)
          .replaceAll('/ ',NL);
      }}
      td.textContent=displayVal;
      applyDisplayDirection(td,displayVal);
      tr.appendChild(td);
    }});
    const ta=document.createElement('td');ta.className='acts';
    let acts='';
    if(isPrTab)acts+=`<button class="pr-toggle" onclick="togglePrItems('${{row._id}}', this)">Items</button> `;
    if(isLtrTab)acts+=`<button class="pr-toggle" onclick="openLetterThread('${{row._id}}')">Thread</button> <button class="pr-toggle" onclick="openLetterTimeline('${{row._id}}')">Timeline</button> `;
    if(CAN_EDIT)acts+=`<button class="act" onclick="editRec('${{row._id}}')">✏</button> <button class="act del" onclick="delRec('${{row._id}}')">🗑</button>`;
    else if(!acts)acts='<span style="color:var(--mu);font-size:10px">—</span>';
    ta.innerHTML=acts;
    tr.appendChild(ta);body.appendChild(tr);
    if(!row._isRev)sr++;
  }});
  const ov=state.recs.filter(r=>r._overdue).length;
  document.getElementById('s-total').textContent='Total: '+state.recs.length;
  document.getElementById('s-show').textContent='Showing: '+rows.length;
  document.getElementById('s-ov').textContent='Overdue: '+ov;
  renderNocSummary(rows);
}}

function calcWD(s,e){{
  if(!s||!e)return'';
  try{{let a=new Date(s),b=new Date(e);if(isNaN(a)||isNaN(b)||b<=a)return'0';
    let n=0,c=new Date(a);c.setDate(c.getDate()+1);
    while(c<=b){{if(c.getDay()!==5)n++;c.setDate(c.getDate()+1);}}return String(n);}}
  catch{{return'';}}
}}

function isFloorField(col){{
  const key=String(col?.col_key||'').toLowerCase();
  const label=String(col?.label||'').toLowerCase();
  return key==='floor'||label==='floor'||label==='floors';
}}

function isItemRefField(col){{
  const key=String(col?.col_key||'').toLowerCase();
  const label=String(col?.label||'').toLowerCase().replace(/[_./-]+/g,' ');
  return key==='itemref'||(label.includes('item ref')&&label.includes('dwg'));
}}

function isStatusLikeField(col){{
  const key=String(col?.col_key||'').toLowerCase();
  const label=String(col?.label||'').toLowerCase();
  return key==='status'||key.includes('status')||label.includes('status');
}}

async function openLetterThread(id){{
  const data=await apiFetch('/api/letters/thread/'+PID+'/'+id);if(!data)return;
  renderLetterThread(data);
  openM('thread-modal');
}}

async function openLetterTimeline(id){{
  const data=await apiFetch('/api/letters/timeline/'+PID+'/'+id);if(!data)return;
  renderLetterTimeline(data);
  openM('timeline-modal');
}}

function renderLetterThread(data){{
  const body=document.getElementById('thread-body');
  const title=document.getElementById('thread-title');
  const items=Array.isArray(data?.items)?data.items:[];
  const currentId=String(data?.current_id||'').trim();
  const current=items.find(it=>it.id===currentId)||items[0]||null;
  title.textContent=current?.doc_no?('Letter Thread — '+current.doc_no):'Letter Thread';
  if(!items.length){{
    body.innerHTML='<div style="padding:12px;color:var(--mu);font-size:12px">This letter has no linked thread yet.</div>';
    return;
  }}
  body.innerHTML=items.map(it=>{{
    const isCurrent=it.id===currentId;
    const margin=it.level*42;
    const border=isCurrent?'#8BC34A':'#d7e0e6';
    const bg=isCurrent?'#f6fbef':'#fff';
    const shadow=isCurrent?'0 10px 26px rgba(139,195,74,.14)':'0 6px 18px rgba(15,23,42,.05)';
    const titleText=escHtml(it.subject||'');
    const refText=escHtml(it.doc_no||'Untitled');
    const metaParts=[
      it.direction?escHtml(it.direction):'',
      it.from_party?('From: '+escHtml(it.from_party)):'',
      it.to_party?('To: '+escHtml(it.to_party)):'',
      it.date?escHtml(it.date):''
    ].filter(Boolean);
    const currentBadge=isCurrent?'<span class="letter-current-badge">Current</span>':'';
    const subjectHtml=titleText||'<span style="color:var(--mu)">No subject</span>';
    const branchCue=it.level>0?'<span style="color:#7d8a95;font-weight:800;margin-right:8px">↳</span>':'';
    const levelLabel=it.level>0?('Reply level '+it.level):'Root letter';
    const metaHtml=metaParts.length?'<div style="font-size:10px;color:var(--mu);margin-top:8px;line-height:1.5">'+metaParts.join(' • ')+'</div>':'';
    return `<div style="margin-left:${{margin}}px;margin-bottom:10px;position:relative">
      ${{it.level>0?`<div style="position:absolute;left:-24px;top:0;bottom:22px;width:2px;background:${{connectorColor}};border-radius:2px"></div>
      <div style="position:absolute;left:-24px;top:24px;width:18px;height:2px;background:${{connectorColor}};border-radius:2px"></div>`:''}}
      <div style="background:${{bg}};border:1px solid ${{border}};border-left:6px solid ${{border}};border-radius:12px;padding:12px 14px;box-shadow:${{shadow}}">
        <div style="font-size:9px;color:#7d8a95;text-transform:uppercase;letter-spacing:.35px;margin-bottom:6px">${{levelLabel}}</div>
        <div style="display:flex;justify-content:space-between;gap:12px;align-items:flex-start">
          <div style="min-width:0">
            <div style="font-size:12px;font-weight:800;color:var(--pr);display:flex;align-items:center" dir="auto">${{branchCue}}<span>${{refText}}</span></div>
            <div style="font-size:12px;color:var(--tx);margin-top:4px;line-height:1.45" dir="auto">${{subjectHtml}}</div>
          </div>
          ${{currentBadge}}
        </div>
        ${{metaHtml}}
      </div>
    </div>`;
  }}).join('');
}}

function renderLetterTimeline(data){{
  const body=document.getElementById('timeline-body');
  const title=document.getElementById('timeline-title');
  const items=Array.isArray(data?.items)?data.items:[];
  const currentId=String(data?.current_id||'').trim();
  const current=items.find(it=>it.id===currentId)||items[0]||null;
  title.textContent=current?.doc_no?('Letter Timeline — '+current.doc_no):'Letter Timeline';
  if(!items.length){{
    body.innerHTML='<div style="padding:12px;color:var(--mu);font-size:12px">This letter has no linked timeline yet.</div>';
    return;
  }}
  const cards=items.map(it=>{{
    const isCurrent=String(it.id||'')===currentId;
    const border=isCurrent?'#8BC34A':'#d7e0e6';
    const bg=isCurrent?'#f6fbef':'#fff';
    const shadow=isCurrent?'0 10px 26px rgba(139,195,74,.14)':'0 6px 18px rgba(15,23,42,.05)';
    const refText=escHtml(it.doc_no||'Untitled');
    const subjectHtml=escHtml(it.subject||'')||'<span style="color:var(--mu)">No subject</span>';
    const dateText=escHtml(it.date||it.created_at||'');
    const metaParts=[
      it.direction?escHtml(it.direction):'',
      it.from_party?('From: '+escHtml(it.from_party)):'',
      it.to_party?('To: '+escHtml(it.to_party)):''
    ].filter(Boolean);
    const metaHtml=metaParts.length?'<div style="font-size:10px;color:var(--mu);margin-top:8px;line-height:1.5">'+metaParts.join(' • ')+'</div>':'';
    const badge=isCurrent?'<span style="flex-shrink:0;background:#8BC34A;color:#17351b;border-radius:999px;padding:3px 8px;font-size:9px;font-weight:800;letter-spacing:.35px;text-transform:uppercase">Current</span>':'';
    const dateBadge=dateText?'<div style="display:inline-flex;align-items:center;background:#eef4f8;color:#2F4F64;border-radius:999px;padding:4px 10px;font-size:10px;font-weight:800;letter-spacing:.2px">'+dateText+'</div>':'';
    return `<div style="position:relative;padding-left:30px;margin-bottom:14px">
      <div style="position:absolute;left:8px;top:14px;width:12px;height:12px;border-radius:999px;background:${{isCurrent?'#8BC34A':'#2F4F64'}};border:3px solid #fff;box-shadow:0 0 0 2px ${{isCurrent?'rgba(139,195,74,.35)':'rgba(47,79,100,.18)'}}"></div>
      <div style="background:${{bg}};border:1px solid ${{border}};border-left:4px solid ${{border}};border-radius:12px;padding:12px 14px;box-shadow:${{shadow}}">
        <div style="display:flex;justify-content:space-between;gap:12px;align-items:flex-start">
          <div style="min-width:0">
            <div style="font-size:12px;font-weight:800;color:var(--pr)" dir="auto">${{refText}}</div>
            <div style="margin-top:5px">${{dateBadge}}</div>
            <div style="font-size:12px;color:var(--tx);margin-top:6px;line-height:1.45" dir="auto">${{subjectHtml}}</div>
          </div>
          ${{badge}}
        </div>
        ${{metaHtml}}
      </div>
    </div>`;
  }}).join('');
  body.innerHTML=`<div style="position:relative;padding-left:10px">
    <div style="position:absolute;left:13px;top:0;bottom:0;width:2px;background:linear-gradient(180deg,#d7e0e6,#c7d2da)"></div>
    ${{cards}}
  </div>`;
}}

function isCurrencyField(col){{
  const key=String(col?.col_key||'').toLowerCase();
  const label=String(col?.label||'').toLowerCase();
  return key.includes('cost')||key.includes('value')||label.includes('(egp)');
}}

function formatCurrencyValue(val){{
  const raw=String(val??'').trim();
  if(!raw)return '';
  const num=Number(raw.toString().replace(/,/g,''));
  if(Number.isNaN(num))return raw;
  return new Intl.NumberFormat('en-US',{{minimumFractionDigits:2,maximumFractionDigits:2}}).format(num)+' EGP';
}}

function parseSafeNumber(val){{
  const raw=String(val??'').trim();
  if(!raw)return null;
  const num=Number(raw.replace(/,/g,''));
  return Number.isFinite(num)?num:null;
}}

function getNocStageProgress(data){{
  const d=data||{{}};
  if(d.partDReturnDate||d.partDStatus||d.finalApprovedCost||d.voNo||d.voIssueDate||d.voValueWithSIAndVAT)return 'Stage D';
  if(d.partCIssueDate||d.submittedCost)return 'Stage C';
  if(d.partBReturnDate||d.partBStatus)return 'Stage B';
  if(d.partAIssueDate||d.docNo||d.title||d.nocDescription||d.originatingDocument)return 'Stage A';
  return 'Not Started';
}}

function getNocAutoBaseValue(data){{
  const finalVal=String(data?.voValueWithSIAndVAT??'').trim();
  if(!finalVal)return '';
  const num=Number(finalVal.replace(/,/g,''));
  if(Number.isNaN(num))return '';
  return (num/1.05).toFixed(2);
}}

function formatDisplayValue(col,val){{
  let text=String(val??'');
  if(!text)return '';
  const NL=String.fromCharCode(10);
  if(isFloorField(col))return text.split(',').map(s=>s.trim()).filter(Boolean).join(NL);
  if(isItemRefField(col))return text.replace(/\\r\\n/g,NL).replace(/\\r/g,NL);
  if(isCurrencyField(col))return formatCurrencyValue(text);
  return text;
}}

function getTextDirectionMeta(val){{
  const text=String(val??'').trim();
  if(!text)return null;
  const arabic=(text.match(/[\u0600-\u06FF\u0750-\u077F\u08A0-\u08FF]/g)||[]).length;
  const latin=(text.match(/[A-Za-z]/g)||[]).length;
  if(arabic&&arabic>=latin)return {{dir:'rtl',align:'right'}};
  return null;
}}

function applyDisplayDirection(el,val){{
  const meta=getTextDirectionMeta(val);
  el.classList.remove('rtl-txt');
  el.style.direction='';
  el.style.textAlign='';
  if(meta){{
    el.classList.add('rtl-txt');
    el.style.direction=meta.dir;
    el.style.textAlign=meta.align;
  }}
}}

function bindDirectionalInput(el){{
  if(!el)return;
  const sync=()=>applyDisplayDirection(el,el.value);
  el.addEventListener('input',sync);
  el.addEventListener('change',sync);
  sync();
}}

function getLongTextMeta(col){{
  const key=String(col?.col_key||'');
  const label=String(col?.label||'');
  const k=key.toLowerCase();
  const l=label.toLowerCase();
  const text=(k+' '+l).replace(/[_-]+/g,' ').replace(/\\s+/g,' ').trim();
  const isContentLike=['content','description'].some(v=>text.includes(v));
  const isRemarksLike=['remarks','notes'].some(v=>text.includes(v));
  const isMsLike=text.includes('ms ref')||text.includes('msref')||text.includes('ms reference')||
    (/\\bms\\b/.test(text)&&(text.includes('ref')||text.includes('reference')||text.includes('no')));
  const isItemRefLike=isItemRefField(col);
  if(isContentLike)return {{
    rows:5,
    style:'resize:vertical; min-height:100px',
    placeholder:'Use Enter to put each item on a separate line'
  }};
  if(isRemarksLike)return {{
    rows:3,
    style:'resize:vertical; min-height:76px',
    placeholder:'Use Enter for multiline remarks'
  }};
  if(isMsLike)return {{
    rows:3,
    style:'resize:vertical; min-height:76px',
    placeholder:'Use Enter to put each MS on a separate line'
  }};
  if(isItemRefLike)return {{
    rows:3,
    style:'resize:vertical; min-height:76px',
    placeholder:'Use Enter to put each item reference / DWG No. on a separate line'
  }};
  return null;
}}

function sortBy(key){{
  state.sortDir=state.sortCol===key?(state.sortDir==='asc'?'desc':'asc'):'asc';
  state.sortCol=key;buildHead();renderRows();
}}

async function fetchPrItems(recordId){{
  if(Object.prototype.hasOwnProperty.call(state.prItemsCache,recordId))return state.prItemsCache[recordId];
  try{{
    const r=await apiFetch('/api/pr_items/'+recordId);
    state.prItemsCache[recordId]=(r&&r.ok&&Array.isArray(r.items))?r.items:[];
  }}catch(e){{
    console.warn('PR items load failed for',recordId,e);
    state.prItemsCache[recordId]=[];
  }}
  return state.prItemsCache[recordId];
}}

function renderPrItemsTable(items, legacyText){{
  if(!items||!items.length){{
    const legacy=legacyText?`<div style="margin-top:8px"><div style="font-size:10px;font-weight:700;color:var(--mu);text-transform:uppercase;letter-spacing:.4px;margin-bottom:4px">Legacy PR Details</div><div style="color:var(--mu);font-size:11px;white-space:pre-line">${{escHtml(legacyText)}}</div></div>`:'';
    return `<div class="pr-items-title">PR Items Breakdown</div><div class="pr-items-empty">No items added</div>${{legacy}}`;
  }}
  const rows=items.map(it=>String(it?.row_type||'item').toLowerCase()==='header'
    ? `<div class="pr-items-section">${{escHtml(it.item_name||'')}}</div>`
    : `<div class="pr-items-cell item">${{escHtml(it.item_name||'')}}</div>
       <div class="pr-items-cell unit">${{escHtml(it.unit||'')}}</div>
       <div class="pr-items-cell qty">${{escHtml(it.quantity??'')}}</div>
       <div class="pr-items-cell remarks">${{escHtml(it.remarks||'')}}</div>`).join('');
  return `<div class="pr-items-title">PR Items Breakdown</div>
    <div class="pr-items-grid-wrap">
      <div class="pr-items-grid">
        <div class="pr-items-grid-head">Item</div>
        <div class="pr-items-grid-head">Unit</div>
        <div class="pr-items-grid-head">Qty</div>
        <div class="pr-items-grid-head">Remarks</div>
        ${{rows}}
      </div>
    </div>`;
}}

function getRenderedRegisterColumnCount(tr){{
  const headRow=document.querySelector('#regtbl thead tr:first-child');
  const headCount=headRow?headRow.children.length:0;
  const rowCount=tr?tr.children.length:0;
  return Math.max(headCount,rowCount,state.cols.length+3);
}}

function sizePrItemsPanel(panel){{
  const wrap=document.getElementById('tblwrap');
  if(!panel||!wrap)return;
  panel.style.width=Math.max(wrap.clientWidth,320)+'px';
}}

async function togglePrItems(recordId, btn){{
  const tr=btn.closest('tr');
  if(!tr)return;
  let nxt=tr.nextElementSibling;
  if(!nxt||!nxt.classList||!nxt.classList.contains('pr-items-row')||nxt.dataset.id!==recordId){{
    const colSpan=getRenderedRegisterColumnCount(tr);
    const panel=document.createElement('div');panel.className='pr-items-panel';
    panel.innerHTML='<div class="pr-items-empty">Loading...</div>';
    const td=document.createElement('td');td.colSpan=colSpan;td.appendChild(panel);
    const row=document.createElement('tr');row.className='pr-items-row';row.dataset.id=recordId;row.appendChild(td);
    row.style.display='none';
    tr.parentNode.insertBefore(row, tr.nextSibling);
    sizePrItemsPanel(panel);
    nxt=row;
  }}
  const open=!nxt.classList.contains('open');
  if(open){{
    const items=await fetchPrItems(recordId);
    const rowData=state.recs.find(r=>r._id===recordId)||{{}};
    const detailsKey=getPrDetailsColKey();
    const colIdx=state.cols.findIndex(c=>c.col_key===detailsKey);
    if(colIdx>=0){{
      const summaryCell=tr.children[2+colIdx];
      if(summaryCell)summaryCell.textContent=getPrSummary(rowData);
    }}
    const legacyKey=getPrDetailsColKey();
    const legacyText=legacyKey?String((state.recs.find(r=>r._id===recordId)||{{}})[legacyKey]||'').trim():'';
    const panel=nxt.querySelector('.pr-items-panel');
    if(panel){{
      panel.innerHTML=renderPrItemsTable(items, legacyText);
      sizePrItemsPanel(panel);
    }}
    nxt.style.display='table-row';
    nxt.classList.add('open');
    btn.textContent='Hide';
  }}else{{
    nxt.classList.remove('open');
    nxt.style.display='none';
    btn.textContent='Items';
  }}
}}

let _rzDrag=null;
document.addEventListener('mousemove',e=>{{
  if(!_rzDrag)return;
  const w=Math.max(40,_rzDrag.sw+e.clientX-_rzDrag.sx);
  _rzDrag.th.style.width=w+'px';_rzDrag.th.style.minWidth=w+'px';
}});
document.addEventListener('mouseup',()=>{{
  if(!_rzDrag)return;
  _rzDrag.rz.classList.remove('rzg');
  document.body.style.cursor='';document.body.style.userSelect='none';
  _rzDrag=null;
}});
let _rzActive=null;
function initRz(){{
  document.querySelectorAll('#regtbl thead tr:first-child th[data-key]').forEach(th=>{{
    if(th.querySelector('.rz'))return;
    const rz=document.createElement('div');rz.className='rz';th.appendChild(rz);
    rz.addEventListener('mousedown',e=>{{
      e.stopPropagation();e.preventDefault();
      _rzActive={{th,sx:e.clientX,sw:th.offsetWidth,key:th.dataset.key}};
      rz.classList.add('rzg');
      document.body.style.cursor='col-resize';
      document.body.style.userSelect='none';
    }});
  }});
}}
document.addEventListener('mousemove',e=>{{
  if(!_rzActive)return;
  const w=Math.max(40,_rzActive.sw+(e.clientX-_rzActive.sx));
  _rzActive.th.style.width=w+'px';
  _rzActive.th.style.minWidth=w+'px';
  _rzActive.th.style.maxWidth=w+'px';
}});
let _rzSaveTimer=null;
document.addEventListener('mouseup',()=>{{
  if(!_rzActive)return;
  const {{th,key}}=_rzActive;
  const w=th.offsetWidth;
  th.querySelector('.rz')?.classList.remove('rzg');
  _rzActive=null;
  document.body.style.cursor='';
  document.body.style.userSelect='';
  // Save width if superadmin
  if(ROLE==='superadmin'&&key){{
    clearTimeout(_rzSaveTimer);
    _rzSaveTimer=setTimeout(()=>{{
      apiFetch('/api/col_width/'+PID+'/'+state.tab,{{
        method:'POST',body:JSON.stringify({{col_key:key,width_px:w}})
      }}).then(r=>{{if(r&&r.ok){{if(!state.colWidths)state.colWidths={{}};state.colWidths[key]=w;}}}});
    }},400);
  }}
}});

let _st;
function doSearch(){{clearTimeout(_st);_st=setTimeout(()=>loadRecords(),250);}}

// Add/Edit Record
function buildRevisionDraftFromSelected(){{
  if(isLTRTab())return null;
  const selected=getSelectedRegisterRow();
  if(!selected)return null;
  const nextDocNo=incrementRevisionNumber(selected.docNo||'');
  if(!nextDocNo)return null;
  return {{...selected,docNo:nextDocNo,_cloneSourceId:selected._id}};
}}

function addRecord(){{
  state.editId=null;
  const draft=buildRevisionDraftFromSelected();
  if(draft){{
    document.getElementById('rec-title').textContent='Add Revision Draft';
    buildForm(draft,{{mode:'revisionDraft'}});
  }}else{{
    document.getElementById('rec-title').textContent='Add Document';
    buildForm(null,{{suggestedDocNo:suggestNextDocNoFromVisibleRows()}});
  }}
  openM('rec-modal');
}}
function editRec(id){{state.editId=id;const row=state.recs.find(r=>r._id===id);if(!row)return;document.getElementById('rec-title').textContent='Edit Document';buildForm(row);openM('rec-modal');}}

async function buildForm(row,opts={{}}){{
  const allCols=await apiFetch('/api/columns/'+PID+'/'+state.tab);if(!allCols)return;
  const AUTO=new Set(['expectedReplyDate','duration','_duration','_duration_today']);
  const formRoot=document.getElementById('rec-form');formRoot.innerHTML='';
  const sectionBodies={{}};
  let nextNo='';
  if(!row){{
    nextNo=String(opts.suggestedDocNo||'').trim();
    if(!nextNo){{const r=await apiFetch('/api/next_doc_no/'+PID+'/'+state.tab);nextNo=r?.next||'';}}
  }}
  const isPrTab=isPRTab();
  const isNocTab=isNOCTab();
  const isLtrTab=isLTRTab();
  const formCols=isLtrTab?allCols.filter(c=>c.visible&&!isLTRInternalField(c)&&!isLTRExcludedField(c)):allCols;
  const ltrParentIdCol=isLtrTab?getLTRColForRole(allCols,'parentLetterId'):null;
  let ltrParentOptions=[];
  if(isLtrTab){{
    const q=row?('?record_id='+encodeURIComponent(row._id)):'';
    const parentData=await apiFetch('/api/letters/parent-options/'+PID+q);
    ltrParentOptions=Array.isArray(parentData?.options)?parentData.options:[];
  }}
  const prevCols=state.cols;
  state.cols=formCols.filter(c=>c.visible);
  const prDetailsKey=isPrTab?getPrDetailsColKey():null;
  state.cols=prevCols;
  const nocSections={{
    'Basic Info':['docNo','title','nocDescription','originatingDocument','remarks'],
    'Part A':['partAIssueDate'],
    'Part B':['partBReturnDate','partBStatus'],
    'Part C':['partCIssueDate','submittedCost'],
    'Part D':['partDReturnDate','partDStatus','finalApprovedCost'],
    'Variation Order':['voNo','voIssueDate','voBaseValue','voValueWithSIAndVAT'],
  }};
  const nocSectionForKey=(key)=>Object.entries(nocSections).find(([,keys])=>keys.includes(key))?.[0]||null;
  let nocStageInp=null,nocBaseInp=null,nocTotalInp=null;
  if(isNocTab){{
    const sf=makeReadOnlyField('Stage Progress',getNocStageProgress(row||{{}}));
    nocStageInp=sf.inp;
    getOrCreateFormSection(formRoot,sectionBodies,'Basic Info').appendChild(sf.grp);
  }}
  if(isLtrTab){{
    const hid=document.createElement('input');
    hid.type='hidden';
    hid.id='f-'+(ltrParentIdCol?.col_key||'parentLetterId');
    hid.value=String(ltrParentIdCol?(row?.[ltrParentIdCol.col_key]||''):(getLTRValue(row,allCols,'parentLetterId')||''));
    formRoot.appendChild(hid);
  }}
  for(const col of formCols){{
    if(AUTO.has(col.col_key))continue;
    const key=col.col_key;
    const longTextMeta=getLongTextMeta(col);
    const full=['title','fileLocation','itemRef'].includes(key)||!!longTextMeta;
    const grp=document.createElement('div');grp.className='fg'+(full?' full':'');
    const targetSection=isNocTab?(nocSectionForKey(key)||'Additional Details'):getDynamicFormSection(col,{{isPrTab,isLtrTab,isNocTab}});
    const lbl=document.createElement('label');lbl.textContent=col.label;grp.appendChild(lbl);
    const ltrRole=isLtrTab?getLTRFieldRole(col):'';
    const val=(isLtrTab&&ltrRole)?(row?.[key]||getLTRValue(row,allCols,ltrRole)||''):(row?.[key]||'');
    if(isPrTab&&prDetailsKey&&key===prDetailsKey){{
      const ta=document.createElement('textarea');ta.id='f-'+key;ta.value=val;
      ta.rows=3;
      ta.style.cssText='resize:vertical;min-height:80px';
      ta.placeholder='Leave blank to auto-generate from items';
      bindDirectionalInput(ta);
      grp.appendChild(ta);
      const hint=document.createElement('div');
      hint.style.cssText='font-size:10px;color:var(--mu);margin-top:4px';
      hint.textContent='Leave blank to auto-generate from PR items.';
      grp.appendChild(hint);
      getOrCreateFormSection(formRoot,sectionBodies,targetSection).appendChild(grp);
      continue;
    }}
    if(isLtrTab&&getLTRFieldRole(col)==='parentLetterRef'){{
      lbl.textContent='Parent Letter';
      const currentParentId=String(ltrParentIdCol?(row?.[ltrParentIdCol.col_key]||''):(getLTRValue(row,allCols,'parentLetterId')||'')).trim();
      const currentParentRef=String(val||'').trim();
      const options=[...ltrParentOptions];
      if(currentParentId&&!options.some(o=>o.id===currentParentId)){{
        options.unshift({{
          id:currentParentId,
          doc_no:currentParentRef,
          subject:'',
          direction:'',
          from_party:'',
          to_party:'',
          date:''
        }});
      }}
      const sel=document.createElement('select');sel.id='f-'+key+'-selector';
      sel.innerHTML='<option value="">No parent letter</option>'+options.map(o=>`<option value="${{escHtml(o.id)}}">${{escHtml(formatLTRParentLabel(o))}}</option>`).join('');
      sel.value=currentParentId||'';
      grp.appendChild(sel);
      const meta=document.createElement('div');
      meta.style.cssText='font-size:10px;color:var(--mu);margin-top:5px;line-height:1.4';
      grp.appendChild(meta);
      const syncParent=()=>{{
        const picked=options.find(o=>o.id===sel.value)||null;
        const hid=document.getElementById('f-'+(ltrParentIdCol?.col_key||'parentLetterId'));
        if(hid)hid.value=picked?picked.id:'';
        meta.textContent=formatLTRParentMeta(picked);
      }};
      sel.onchange=syncParent;
      syncParent();
      getOrCreateFormSection(formRoot,sectionBodies,targetSection).appendChild(grp);
      continue;
    }}
    if(col.col_type==='date'){{const inp=document.createElement('input');inp.type='date';inp.id='f-'+key;inp.value=val;grp.appendChild(inp);}}
    else if(isLtrTab&&col.col_type==='dropdown'&&col.list_name){{
      const sel=document.createElement('select');sel.id='f-'+key;
      const ltrRole=getLTRFieldRole(col);
      const listName=ltrRole==='fromParty'||ltrRole==='toParty'
        ? 'correspondence_parties'
        : (ltrRole==='direction' ? 'letter_direction' : (ltrRole==='status' ? 'letter_status' : col.list_name));
      const opts=[...(state.lists[listName]||[])];
      if(val&&!opts.includes(val))opts.push(val);
      sel.innerHTML='<option value=""></option>'+opts.map(o=>`<option value="${{escHtml(o)}}">${{escHtml(o)}}</option>`).join('');
      sel.value=val||'';
      grp.appendChild(sel);
    }}
    else if(isLtrTab&&['title','description','remarks'].includes(getLTRFieldRole(col))){{
      const ta=document.createElement('textarea');ta.id='f-'+key;ta.value=val||'';
      const role=getLTRFieldRole(col);
      ta.rows=role==='description'?4:3;
      ta.style.cssText=role==='description'?'resize:vertical; min-height:96px':'resize:vertical; min-height:76px';
      ta.placeholder=role==='title'?'Use Enter for a multiline subject':'Use Enter for multiline text';
      bindDirectionalInput(ta);
      grp.appendChild(ta);
    }}
    else if(col.col_type==='dropdown'&&col.list_name){{
      // Add free-text input below multiselect
      const wrapper=document.createElement('div');
      wrapper.appendChild(buildMS(key,state.lists[col.list_name]||[],val));
      const freeInp=document.createElement('input');
      freeInp.id='f-free-'+key;freeInp.placeholder='Or type custom value...';
      freeInp.style.cssText='margin-top:4px;width:100%;padding:5px 8px;border:1px dashed var(--bd);border-radius:var(--rd);font-size:11px;color:var(--mu);outline:none';
      freeInp.title='Type a custom value not in the list';
      bindDirectionalInput(freeInp);
      freeInp.onchange=()=>{{
        const v=freeInp.value.trim();
        if(!v)return;
        const ms=document.getElementById('f-'+key);
        if(ms&&ms.classList.contains('ms-con')){{
          const cur=ms.dataset.value||'';
          const arr=cur?cur.split(', ').filter(Boolean):[];
          if(!arr.includes(v))arr.push(v);
          ms.dataset.value=arr.join(', ');
          // Trigger re-render
          ms.dispatchEvent(new Event('_update'));
          freeInp.value='';
        }}
      }};
      wrapper.appendChild(freeInp);
      grp.appendChild(wrapper);
    }}
    else if(col.col_type==='docno'){{
      const inp=document.createElement('input');inp.id='f-'+key;
      if(row){{inp.value=val;}}
      else{{inp.value=nextNo;inp.placeholder=nextNo?'':state.tab+'-001 REV00';}}
      inp.style.cssText='font-family:Consolas,monospace;font-weight:600';
      bindDirectionalInput(inp);
      grp.appendChild(inp);
    }}
    else if(col.col_type==='number'){{
      const inp=document.createElement('input');inp.type='number';inp.step='0.01';inp.id='f-'+key;inp.value=val;
      if(isCurrencyField(col))inp.placeholder='0.00';
      grp.appendChild(inp);
      if(key==='voBaseValue')nocBaseInp=inp;
      if(key==='voValueWithSIAndVAT')nocTotalInp=inp;
    }}
    else if(longTextMeta){{
      const ta=document.createElement('textarea');ta.id='f-'+key;ta.value=val;
      ta.rows=longTextMeta.rows;
      ta.style.cssText=longTextMeta.style;
      ta.placeholder=longTextMeta.placeholder;
      bindDirectionalInput(ta);
      grp.appendChild(ta);
    }}
    else{{const inp=document.createElement('input');inp.id='f-'+key;inp.value=val;if(col.col_type==='link')inp.placeholder='https://...';bindDirectionalInput(inp);grp.appendChild(inp);}}
    if(isNocTab&&isCurrencyField(col))grp.style.gridColumn='span 1';
    getOrCreateFormSection(formRoot,sectionBodies,targetSection).appendChild(grp);
  }}
  if(isNocTab){{
    const syncNoc=()=>{{
      const data={{}};
      allCols.forEach(c=>{{
        const el=document.getElementById('f-'+c.col_key);
        if(!el)return;
        data[c.col_key]=el.classList?.contains?.('ms-con')?(el.dataset.value||''):el.value;
      }});
      if(nocStageInp)nocStageInp.value=getNocStageProgress(data);
      if(nocBaseInp&&nocTotalInp&&!String(nocBaseInp.value||'').trim())nocBaseInp.value=getNocAutoBaseValue(data);
    }};
    formRoot.querySelectorAll('input, textarea, select').forEach(el=>el.addEventListener('input',syncNoc));
    formRoot.querySelectorAll('input, textarea, select').forEach(el=>el.addEventListener('change',syncNoc));
    syncNoc();
  }}
  if(isPrTab){{
    const prSection=getOrCreateFormSection(formRoot,sectionBodies,'PR Items');
    const grp=document.createElement('div');grp.className='fg full pr-items-editor';
    const lbl=document.createElement('label');lbl.textContent='PR Items';grp.appendChild(lbl);
    const wrap=document.createElement('div');wrap.id='pr-items-editor';
    wrap.innerHTML=`
      <table>
        <thead><tr><th>Item Name</th><th>Unit</th><th>Qty</th><th>Remarks</th><th></th></tr></thead>
        <tbody id="pr-items-body"></tbody>
      </table>
      <div style="margin-top:8px;display:flex;gap:8px;align-items:center">
        <button class="btn btn-sc btn-sm" onclick="addPrItemRow()">+ Add Item</button>
        <button class="btn btn-sc btn-sm" onclick="addPrHeaderRow()">+ Add Section Header</button>
        <div id="pr-legacy" style="font-size:11px;color:var(--mu);white-space:pre-line;display:none"></div>
      </div>`;
    grp.appendChild(wrap);prSection.appendChild(grp);
    const legacyText=prDetailsKey?String(row?.[prDetailsKey]||'').trim():'';
    initPrItemsEditor([], legacyText);
    if(row&&row._id)loadPrItemsForEdit(row._id, legacyText);
  }}
}}

function addPrItemRow(item={{}}){{
  const body=document.getElementById('pr-items-body');if(!body)return;
  const tr=document.createElement('tr');
  tr.dataset.rowType='item';
  tr.innerHTML=`
    <td><input class="pri-item"></td>
    <td><input class="pri-unit"></td>
    <td><input class="pri-qty"></td>
    <td><input class="pri-remarks"></td>
    <td><button type="button" class="btn btn-er btn-sm">✕</button></td>`;
  tr.querySelector('.pri-item').value=item.item_name||'';
  tr.querySelector('.pri-unit').value=item.unit||'';
  tr.querySelector('.pri-qty').value=item.quantity??'';
  tr.querySelector('.pri-remarks').value=item.remarks||'';
  bindDirectionalInput(tr.querySelector('.pri-item'));
  bindDirectionalInput(tr.querySelector('.pri-unit'));
  bindDirectionalInput(tr.querySelector('.pri-remarks'));
  tr.querySelector('button').onclick=()=>tr.remove();
  body.appendChild(tr);
}}

function addPrHeaderRow(item={{}}){{
  const body=document.getElementById('pr-items-body');if(!body)return;
  const tr=document.createElement('tr');
  tr.className='pr-head-edit';
  tr.dataset.rowType='header';
  tr.innerHTML=`
    <td colspan="4">
      <div class="pr-head-label">Section Header</div>
      <input class="pri-header" placeholder="Section title / description">
    </td>
    <td><button type="button" class="btn btn-er btn-sm">✕</button></td>`;
  tr.querySelector('.pri-header').value=item.item_name||'';
  bindDirectionalInput(tr.querySelector('.pri-header'));
  tr.querySelector('button').onclick=()=>tr.remove();
  body.appendChild(tr);
}}

function initPrItemsEditor(items, legacyText){{
  const body=document.getElementById('pr-items-body');if(!body)return;
  body.innerHTML='';
  if(items&&items.length)items.forEach(it=>String(it?.row_type||'item').toLowerCase()==='header'?addPrHeaderRow(it):addPrItemRow(it));
  else addPrItemRow();
  const legacy=document.getElementById('pr-legacy');
  if(legacy){{
    legacy.textContent=legacyText?('Legacy Details:\\n'+legacyText):'';
    legacy.style.display=(!items||!items.length)&&legacyText?'block':'none';
  }}
}}

async function loadPrItemsForEdit(recordId, legacyText){{
  const items=await fetchPrItems(recordId);
  initPrItemsEditor(items, legacyText);
}}

function getPrItemsFromEditor(){{
  const body=document.getElementById('pr-items-body');if(!body)return [];
  const rows=[...body.querySelectorAll('tr')];
  return rows.map(r=>{{
    const row_type=(r.dataset.rowType||'item').toLowerCase();
    if(row_type==='header'){{
      const item_name=r.querySelector('.pri-header')?.value.trim()||'';
      return {{row_type:'header', item_name}};
    }}
    const item_name=r.querySelector('.pri-item')?.value.trim()||'';
    const unit=r.querySelector('.pri-unit')?.value.trim()||'';
    const quantity=r.querySelector('.pri-qty')?.value.trim()||'';
    const remarks=r.querySelector('.pri-remarks')?.value.trim()||'';
    return {{row_type:'item', item_name, unit, quantity, remarks}};
  }}).filter(it=>it.row_type==='header'?it.item_name:(it.item_name||it.unit||it.quantity||it.remarks));
}}

function buildMS(key,options,init){{
  const sel=init?init.split(',').map(s=>s.trim()).filter(Boolean):[];
  const con=document.createElement('div');con.className='ms-con';con.id='f-'+key;con.dataset.value=init||'';
  function render(){{
    con.innerHTML='';
    sel.forEach(v=>{{
      const t=document.createElement('span');t.className='ms-tag';
      t.innerHTML=`${{v}} <span class="ms-rm" data-v="${{v}}">✕</span>`;
      t.querySelector('.ms-rm').onclick=e=>{{e.stopPropagation();sel.splice(sel.indexOf(v),1);con.dataset.value=sel.join(', ');render();}};
      con.appendChild(t);
    }});
    if(!sel.length)con.innerHTML='<span class="ms-ph">Select or type below...</span>';
    con.dataset.value=sel.join(', ');
  }}
  con.onclick=e=>{{
    if(e.target.classList.contains('ms-rm'))return;
    const ex=document.querySelector('.ms-dd');if(ex){{ex.remove();return;}}
    const dd=document.createElement('div');dd.className='ms-dd';
    dd.style.cssText='position:absolute;left:0;right:0;top:calc(100% + 2px);z-index:500;max-height:260px;overflow-y:auto';
    options.forEach(opt=>{{
      const it=document.createElement('div');
      it.className='ms-opt'+(sel.includes(opt)?' sel':'');
      const chk=document.createElement('input');chk.type='checkbox';chk.checked=sel.includes(opt);chk.style.cssText='flex-shrink:0;width:14px;height:14px;pointer-events:none';
      const lbl=document.createElement('span');lbl.textContent=opt;lbl.style.flex='1';
      it.appendChild(chk);it.appendChild(lbl);
      it.onclick=ev=>{{
        ev.stopPropagation();
        if(sel.includes(opt))sel.splice(sel.indexOf(opt),1);else sel.push(opt);
        con.dataset.value=sel.join(', ');render();
        chk.checked=sel.includes(opt);
        it.classList.toggle('sel',sel.includes(opt));
      }};
      dd.appendChild(it);
    }});
    con.style.position='relative';con.appendChild(dd);
  }};
  document.addEventListener('click',e=>{{if(!con.contains(e.target))con.querySelector('.ms-dd')?.remove();}},true);
  render();return con;
}}

function getFormSectionHint(title){{
  const hints={{
    'Document Details':'Core document identity and register metadata',
    'Dates & Status':'Timeline, approvals, and stage tracking',
    'Parties & Responsibility':'Origin, destination, and responsible stakeholders',
    'Narrative & References':'Subjects, remarks, descriptions, and supporting references',
    'Commercial & Quantities':'Values, quantities, and related numeric fields',
    'PR Items':'Line items and grouped procurement details',
    'Letter Information':'Reference and headline details for the letter',
    'Correspondence Routing':'Direction, parties, and linked parent correspondence',
    'Additional Details':'Remaining project-specific fields'
  }};
  return hints[title]||'Structured fields for this document section';
}}

function getDynamicFormSection(col, ctx){{
  const key=String(col?.col_key||'').toLowerCase();
  const label=String(col?.label||'').toLowerCase();
  const text=(key+' '+label).replace(/[_-]+/g,' ').trim();
  const longTextMeta=getLongTextMeta(col);
  if(ctx.isLtrTab){{
    if(['docno','title','reference'].some(v=>text.includes(v)))return 'Letter Information';
    if(['direction','fromparty','toparty','parentletter','issueddate','receiveddate'].some(v=>key.includes(v)||text.includes(v)))return 'Correspondence Routing';
    if(longTextMeta||['remarks','description','subject'].some(v=>text.includes(v)))return 'Narrative & References';
    return 'Additional Details';
  }}
  if(ctx.isPrTab){{
    if(['docno','title','reference','filelocation'].some(v=>text.includes(v)))return 'Document Details';
    if(['date','status','reply','review','issue'].some(v=>text.includes(v)))return 'Dates & Status';
    if(['qty','quantity','unit','amount','value','cost','price'].some(v=>text.includes(v)))return 'Commercial & Quantities';
    if(longTextMeta||['remarks','notes','description','itemref'].some(v=>text.includes(v)))return 'Narrative & References';
    return 'Document Details';
  }}
  if(['date','status','reply','review','issue','approval','return'].some(v=>text.includes(v)))return 'Dates & Status';
  if(['qty','quantity','unit','amount','value','cost','price'].some(v=>text.includes(v)))return 'Commercial & Quantities';
  if(['from','to','party','client','consultant','contractor','originator','recipient'].some(v=>text.includes(v)))return 'Parties & Responsibility';
  if(longTextMeta||['remarks','notes','description','subject','content','location','reference','originating'].some(v=>text.includes(v)))return 'Narrative & References';
  return 'Document Details';
}}

function getOrCreateFormSection(root, sectionMap, title){{
  if(sectionMap[title])return sectionMap[title];
  const sec=document.createElement('section');
  sec.className='form-section';
  const hdr=document.createElement('div');
  hdr.className='form-section-header';
  hdr.innerHTML=`<div><div class="form-section-title">${{escHtml(title)}}</div><div class="form-section-sub">${{escHtml(getFormSectionHint(title))}}</div></div>`;
  const body=document.createElement('div');
  body.className='form-section-grid';
  sec.appendChild(hdr);
  sec.appendChild(body);
  root.appendChild(sec);
  sectionMap[title]=body;
  return body;
}}

function appendSectionTitle(grid,title){{
  return getOrCreateFormSection(grid,{{}},title);
}}

function makeReadOnlyField(label,value){{
  const grp=document.createElement('div');grp.className='fg full';
  const lbl=document.createElement('label');lbl.textContent=label;grp.appendChild(lbl);
  const inp=document.createElement('input');inp.value=value||'';inp.readOnly=true;
  inp.style.cssText='background:var(--bg);font-weight:700;color:var(--pr)';
  grp.appendChild(inp);
  return {{grp,inp}};
}}

function renderNocSummary(rows){{
  const wrap=document.getElementById('noc-summary');
  const inner=document.getElementById('noc-summary-inner');
  if(!wrap||!inner)return;
  if(!isNOCTab()){{wrap.style.display='none';inner.innerHTML='';return;}}
  const totals={{
    submittedCost:0,
    finalApprovedCost:0,
    voBaseValue:0,
    voValueWithSIAndVAT:0,
  }};
  for(const row of (rows||[])){{
    for(const key of Object.keys(totals)){{
      const n=parseSafeNumber(row?.[key]);
      if(n!==null)totals[key]+=n;
    }}
  }}
  const items=[
    ['Visible NOCs', String((rows||[]).length), 'Filtered rows'],
    ['Submitted Cost', formatCurrencyValue(totals.submittedCost), 'Visible total'],
    ['Final Approved Cost', formatCurrencyValue(totals.finalApprovedCost), 'Visible total'],
    ['VO Base Value', formatCurrencyValue(totals.voBaseValue), 'Visible total'],
    ['VO Value Incl. SI & VAT', formatCurrencyValue(totals.voValueWithSIAndVAT), 'Visible total'],
  ];
  inner.innerHTML=items.map(([label,val,sub])=>`<div style="background:#fff;border:1px solid var(--bd);border-radius:8px;padding:10px 12px">
    <div style="font-size:10px;font-weight:700;color:var(--mu);text-transform:uppercase;letter-spacing:.4px">${{label}}</div>
    <div style="font-size:18px;font-weight:800;color:var(--pr);margin-top:4px;line-height:1.2">${{val}}</div>
    <div style="font-size:10px;color:var(--mu);margin-top:3px">${{sub}}</div>
  </div>`).join('');
  wrap.style.display='';
}}

function durChoice(docNo){{
  return new Promise(resolve=>{{
    const ov=document.createElement('div');ov.className='overlay';ov.style.cssText='position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:9999;display:flex;align-items:center;justify-content:center';
    ov.innerHTML=`<div style="background:#fff;border-radius:12px;padding:28px 24px;max-width:380px;width:90%;box-shadow:0 20px 60px rgba(0,0,0,.3)">
      <div style="font-size:16px;font-weight:700;margin-bottom:8px">⏱ Duration</div>
      <p style="font-size:13px;color:#64748b;margin-bottom:20px">Issue date is set but no reply date yet.<br>How should Duration be shown?</p>
      <div style="display:flex;flex-direction:column;gap:8px">
        <button id="dur-today" style="padding:10px 16px;background:#1a3a5c;color:#fff;border:none;border-radius:6px;cursor:pointer;font-size:13px">📅 Calculate from issue date → Today</button>
        <button id="dur-empty" style="padding:10px 16px;background:#f1f5f9;color:#374151;border:none;border-radius:6px;cursor:pointer;font-size:13px">— Leave Duration empty</button>
        <button id="dur-cancel" style="padding:8px;background:none;border:none;cursor:pointer;font-size:12px;color:#94a3b8">Cancel</button>
      </div></div>`;
    document.body.appendChild(ov);
    ov.querySelector('#dur-today').onclick=()=>{{document.body.removeChild(ov);resolve('today');}};
    ov.querySelector('#dur-empty').onclick=()=>{{document.body.removeChild(ov);resolve('empty');}};
    ov.querySelector('#dur-cancel').onclick=()=>{{document.body.removeChild(ov);resolve(null);}};
  }});
}}

async function saveRecord(){{
  const allCols=await apiFetch('/api/columns/'+PID+'/'+state.tab);if(!allCols)return;
  const AUTO=new Set(['expectedReplyDate','duration','_duration','_duration_today']);
  const data={{}};
  for(const col of allCols){{
    if(AUTO.has(col.col_key))continue;
    const el=document.getElementById('f-'+col.col_key);if(!el)continue;
      if(el.classList.contains('ms-con'))data[col.col_key]=el.dataset.value||'';
      else if(el.tagName==='TEXTAREA')data[col.col_key]=el.value.replace(/\\r\\n/g,String.fromCharCode(10)).replace(/\\r/g,String.fromCharCode(10));
      else data[col.col_key]=el.value.trim();
  }}
  if(isLTRTab()){{
    const direction=String(getLTRValue(data,allCols,'direction')||'').trim().toLowerCase();
    const issued=String(getLTRValue(data,allCols,'issuedDate')||'').trim();
    const received=String(getLTRValue(data,allCols,'receivedDate')||'').trim();
    const parentId=String(getLTRValue(data,allCols,'parentLetterId')||'').trim();
    if(direction==='sent'&&!issued){{toast('Issue Date is required for Sent letters','er');return;}}
    if(direction==='received'&&!received){{toast('Received Date is required for Received letters','er');return;}}
    if(state.editId&&parentId&&parentId===state.editId){{toast('A letter cannot reference itself as parent','er');return;}}
  }}
  if(isNOCTab()&&!String(data.voBaseValue||'').trim()&&String(data.voValueWithSIAndVAT||'').trim()){{
    const autoBase=getNocAutoBaseValue(data);
    if(autoBase)data.voBaseValue=autoBase;
  }}
  if(!data.docNo){{toast('Document No. required','er');return;}}
    // Duration is computed server-side automatically
  const valErr=validateDocNo(data.docNo,state.recs||[],state.editId);
  if(valErr&&valErr.startsWith('GAP')){{
    if(!confirm('⚠ Sequence gap detected: '+valErr.replace('GAP:','')+'. Continue anyway?'))return;
  }}else if(valErr){{toast('⚠ '+valErr,'er');return;}}
  if(state.editId)data._id=state.editId;
  const r=await apiFetch('/api/records/'+PID+'/'+state.tab,{{method:'POST',body:JSON.stringify(data)}});
  if(r&&r.ok){{
    if(isPRTab()){{
      const items=getPrItemsFromEditor();
      const recId=r.id||state.editId;
      try{{
        const prRes=await apiFetch('/api/pr_items/'+recId,{{method:'POST',body:JSON.stringify({{items}})}});
        if(prRes&&prRes.ok)state.prItemsCache[recId]=items;
        else {{toast('Items save failed','er');return;}}
      }}catch(e){{
        toast('Items save failed: '+e.message,'er');return;
      }}
    }}
    closeM('rec-modal');const savedTab=state.tab;await loadRecords();await refreshCounts();toast(state.editId?'Updated':'Added','ok');
  }}
  else toast('Error saving','er');
}}

async function delRec(id){{
  if(!confirm('Delete this record?'))return;
  await apiFetch('/api/records/'+id,{{method:'DELETE'}});
  await loadRecords();await refreshCounts();toast('Deleted','wa');
}}

// Bulk
function updBulk(){{
  const checked=document.querySelectorAll('.chkcell input[data-id]:checked');
  document.getElementById('bulkcnt').textContent=checked.length+' selected';
  document.getElementById('bulkbar').classList.toggle('show',checked.length>0);
  const all=document.querySelectorAll('.chkcell input[data-id]');
  const ca=document.getElementById('chkall');if(ca)ca.checked=all.length>0&&checked.length===all.length;
}}
function selAll(v){{document.querySelectorAll('.chkcell input[data-id]').forEach(cb=>cb.checked=v);updBulk();}}
function clearSel(){{document.querySelectorAll('.chkcell input').forEach(cb=>cb.checked=false);setSelectedRegisterRow(null);updBulk();}}
async function bulkDel(){{
  const ids=[...document.querySelectorAll('.chkcell input[data-id]:checked')].map(cb=>cb.dataset.id);
  if(!ids.length||!confirm('Delete '+ids.length+' records?'))return;
  const r=await apiFetch('/api/records/bulk_delete',{{method:'POST',body:JSON.stringify({{ids}})}});
  if(r&&r.ok){{
    clearSel();await loadRecords();await refreshCounts();toast('✔ Deleted '+(r.deleted||0),'ok');
  }}
}}

// Project Modal
async function editProject(){{
  const proj=await apiFetch('/api/project/'+PID);if(!proj)return;
  const body=document.getElementById('proj-body');body.innerHTML='';
  const customLabels=proj._labels||{{}};
  // Field visibility toggles
  const visNote=document.createElement('div');
  visNote.style.cssText='font-size:10px;color:var(--mu);margin-bottom:8px;padding:6px 8px;background:var(--bg);border-radius:4px';
  visNote.textContent='💡 Leave any field empty to hide it from the project bar';
  body.appendChild(visNote);
  const grid=document.createElement('div');grid.className='fgrid';
  PROJ_FIELDS.forEach(([key,defaultLbl])=>{{
    const curLbl=customLabels[key]||defaultLbl;
    const grp=document.createElement('div');grp.className='fg';
    // Label row: custom label input + value input
    const lblRow=document.createElement('div');lblRow.style.cssText='display:flex;gap:4px;margin-bottom:3px;align-items:center';
    const lblInp=document.createElement('input');
    lblInp.id='lbl-'+key;lblInp.value=curLbl;
    lblInp.placeholder='Field label';
    lblInp.title='Edit field label';
    lblInp.style.cssText='font-size:9px;font-weight:700;color:var(--pr);text-transform:uppercase;letter-spacing:.4px;border:1px dashed var(--bd);border-radius:3px;padding:2px 5px;width:100%;background:transparent;outline:none';
    lblRow.appendChild(lblInp);
    const valInp=document.createElement('input');valInp.id='pf-'+key;valInp.value=proj[key]||'';
    valInp.placeholder='(leave empty to hide)';
    grp.appendChild(lblRow);grp.appendChild(valInp);
    grid.appendChild(grp);
  }});
  body.appendChild(grid);
  // Logo section
  const lt=document.createElement('div');lt.className='stitle';lt.textContent='Company Logos';body.appendChild(lt);
  const lg=document.createElement('div');lg.style.cssText='display:grid;grid-template-columns:1fr 1fr;gap:12px';
  for(const[k,lbl]of[['logo_left','Left Logo'],['logo_right','Right Logo']]){{
    const d=document.createElement('div');d.style.cssText='border:1px solid var(--bd);border-radius:6px;padding:10px;text-align:center;background:var(--bg)';
    const lbel=document.createElement('div');lbel.style.cssText='font-size:10px;font-weight:700;color:var(--mu);text-transform:uppercase;margin-bottom:6px';lbel.textContent=lbl;d.appendChild(lbel);
    const img=document.createElement('img');img.id='lp-'+k;img.style.cssText='max-height:52px;max-width:100%;object-fit:contain;display:block;margin:0 auto 6px';d.appendChild(img);
    fetch('/api/logo/'+PID+'/'+k).then(r=>r.ok?r.blob():null).then(b=>{{if(b)img.src=URL.createObjectURL(b);}});
    const fi=document.createElement('input');fi.type='file';fi.accept='image/*';fi.style.cssText='width:100%;font-size:10px;margin-bottom:4px';
    fi.onchange=async e=>{{const f=e.target.files[0];if(!f)return;const b64=await new Promise(res=>{{const fr=new FileReader();fr.onload=e2=>res(e2.target.result);fr.readAsDataURL(f);}});img.src=b64;await apiFetch('/api/logo/'+PID+'/'+k,{{method:'POST',body:JSON.stringify({{data:b64.split(',')[1]}})}});toast('Logo saved','ok');}};
    d.appendChild(fi);
    const clr=document.createElement('button');clr.className='btn btn-sc btn-sm';clr.textContent='Remove';clr.style.fontSize='9px';
    clr.onclick=async()=>{{await apiFetch('/api/logo/'+PID+'/'+k,{{method:'POST',body:JSON.stringify({{data:''}})}});img.src='';toast('Removed','wa');}};
    d.appendChild(clr);lg.appendChild(d);
  }}
  body.appendChild(lg);
  openM('proj-modal');
}}

async function saveProject(){{
  const data={{}};
  const labels={{}};
  PROJ_FIELDS.forEach(([key,defaultLbl])=>{{
    const el=document.getElementById('pf-'+key);
    if(el)data[key]=el.value.trim();
    const lblEl=document.getElementById('lbl-'+key);
    if(lblEl&&lblEl.value.trim()&&lblEl.value.trim()!==defaultLbl)
      labels[key]=lblEl.value.trim();
  }});
  data._labels=labels;
  const r=await apiFetch('/api/project/'+PID,{{method:'POST',body:JSON.stringify(data)}});
  if(r===null)return;
  if(r.ok){{
    closeM('proj-modal');toast('✔ Project saved!','ok');
    setTimeout(()=>location.reload(),400);
  }}else toast('Save failed','er');
}}

// Doc Types
function addDocType(){{document.getElementById('dt-code').value='';document.getElementById('dt-name').value='';openM('dt-modal');}}
async function saveDocType(){{
  const code=document.getElementById('dt-code').value.trim().toUpperCase();
  const name=document.getElementById('dt-name').value.trim();
  if(!code||!name){{toast('Code and name required','er');return;}}
  await apiFetch('/api/doc_types/'+PID,{{method:'POST',body:JSON.stringify({{code,name}})}});
  closeM('dt-modal');await loadDTs();switchTab(code);toast('✔ Type added','ok');
}}

// Lists
async function openLists(){{
  await loadLists(true);
  ensureLTRProjectLists();
  const metaData=await apiFetch('/api/lists_meta/'+PID)||{{}};
  const META_LABELS={{approved:{{lbl:'Approved',bg:'#bbf7d0',fg:'#166534'}},rejected:{{lbl:'Rejected',bg:'#fce7f3',fg:'#831843'}},pending:{{lbl:'Pending',bg:'#fef9c3',fg:'#713f12'}},cancelled:{{lbl:'Cancelled',bg:'#f1f5f9',fg:'#94a3b8'}}}};
  const body=document.getElementById('lists-body');body.innerHTML='';
  for(const[ln,items]of Object.entries(state.lists).filter(([ln])=>!isHiddenSystemListName(ln))){{
    const isStatus=ln.toLowerCase().startsWith('status');
    const t=document.createElement('div');t.className='stitle';
    t.textContent=ln.charAt(0).toUpperCase()+ln.slice(1)+(isStatus?' 🏷':'');
    body.appendChild(t);
    const ul=document.createElement('ul');ul.className='slist';
    const metaItems=metaData[ln]||[];
    items.forEach(item=>{{
      const mi=metaItems.find(m=>m.value===item);
      const curMeta=mi?.meta||'pending';
      const ml=META_LABELS[curMeta]||META_LABELS.pending;
      const li=document.createElement('li');li.className='sitem';
      li.style.cssText='display:flex;align-items:center;gap:6px;padding:5px 8px';
      const nm=document.createElement('span');nm.style.flex='1';nm.textContent=item;
      li.appendChild(nm);
      if(isStatus){{
        const badge=document.createElement('span');
        badge.style.cssText=`background:${{ml.bg}};color:${{ml.fg}};font-size:9px;font-weight:700;padding:2px 8px;border-radius:10px;white-space:nowrap`;
        badge.textContent=ml.lbl;
        const sel=document.createElement('select');
        sel.style.cssText='font-size:10px;padding:2px 5px;border:1.5px solid #e2e8f0;border-radius:4px;background:#fff;cursor:pointer';
        ['approved','rejected','pending','cancelled'].forEach(m=>{{
          const o=document.createElement('option');o.value=m;
          o.textContent=META_LABELS[m].lbl;
          if(m===curMeta)o.selected=true;
          sel.appendChild(o);
        }});
        sel.onchange=async()=>{{
          const nm2=META_LABELS[sel.value];
          badge.style.background=nm2.bg;badge.style.color=nm2.fg;badge.textContent=nm2.lbl;
          await apiFetch('/api/lists_meta/'+PID,{{method:'POST',body:JSON.stringify({{list_name:ln,item_value:item,meta:sel.value}})}});
          toast('✔ Saved','ok');
        }};
        li.appendChild(badge);li.appendChild(sel);
      }}
      const idx=items.indexOf(item);
      const ren=document.createElement('button');ren.textContent='Rename';
      ren.onclick=()=>renameItem(ln,item);li.appendChild(ren);
      const up=document.createElement('button');up.textContent='↑';
      up.disabled=idx===0;
      up.style.opacity=idx===0?'.45':'1';
      up.onclick=()=>moveListItem(ln,item,-1);li.appendChild(up);
      const down=document.createElement('button');down.textContent='↓';
      down.disabled=idx===items.length-1;
      down.style.opacity=idx===items.length-1?'.45':'1';
      down.onclick=()=>moveListItem(ln,item,1);li.appendChild(down);
      const rb=document.createElement('button');rb.textContent='Remove';
      rb.onclick=()=>rmItem(ln,item,rb);li.appendChild(rb);
      ul.appendChild(li);
    }});
    body.appendChild(ul);
    const ar=document.createElement('div');ar.className='addrow';
    const ni=document.createElement('input');ni.id='new-'+ln;ni.placeholder='New item...';
    const ab=document.createElement('button');ab.className='btn btn-ok btn-sm';ab.textContent='Add';ab.onclick=()=>addItem(ln);
    ar.appendChild(ni);ar.appendChild(ab);
    if(isStatus){{
      const hint=document.createElement('div');
      hint.style.cssText='font-size:10px;color:#94a3b8;margin-top:4px';
      hint.textContent='🏷 Set category for each status — affects Dashboard counts';
      ar.appendChild(hint);
    }}
    body.appendChild(ar);
  }}
  const nl=document.createElement('div');nl.className='stitle';nl.textContent='New List';
  const nar=document.createElement('div');nar.className='addrow';
  nar.innerHTML=`<input id="new-list" placeholder="List name"><button class="btn btn-pr btn-sm" onclick="mkList()">Create</button>`;
  body.appendChild(nl);body.appendChild(nar);
  openM('lists-modal');
}}

async function addItem(ln){{
  const inp=document.getElementById('new-'+ln);const val=inp?.value.trim();if(!val)return;
  await apiFetch('/api/lists/'+PID,{{method:'POST',body:JSON.stringify({{list_name:ln,item:val}})}});
  // Auto-assign pending meta for new status items
  if(ln.toLowerCase().startsWith('status')){{
    await apiFetch('/api/lists_meta/'+PID,{{method:'POST',body:JSON.stringify({{list_name:ln,item_value:val,meta:'pending'}})}});
  }}
  await loadLists(true);openLists();
}}
async function renameItem(ln,item){{
  const next=prompt('Rename list item:',item);
  if(next===null)return;
  const new_item=next.trim();
  if(!new_item||new_item===item)return;
  await apiFetch('/api/lists/'+PID,{{method:'PATCH',body:JSON.stringify({{list_name:ln,old_item:item,new_item}})}});
  await loadLists(true);openLists();toast('✔ Item renamed','ok');
}}
async function moveListItem(ln,item,delta){{
  const items=[...(state.lists[ln]||[])];
  const idx=items.indexOf(item),next=idx+delta;
  if(idx<0||next<0||next>=items.length)return;
  [items[idx],items[next]]=[items[next],items[idx]];
  await apiFetch('/api/lists/'+PID+'/reorder',{{method:'POST',body:JSON.stringify({{list_name:ln,order:items}})}});
  await loadLists(true);openLists();toast('✔ List order updated','ok');
}}
async function rmItem(ln,item,btn){{
  await apiFetch('/api/lists/'+PID,{{method:'DELETE',body:JSON.stringify({{list_name:ln,item}})}});
  btn.closest('li').remove();await loadLists(true);
}}
async function mkList(){{
  const name=document.getElementById('new-list')?.value.trim().toLowerCase().replace(/[\\s]+/g,'_');if(!name)return;
  await apiFetch('/api/lists/'+PID,{{method:'POST',body:JSON.stringify({{list_name:name,item:'Item 1'}})}});
  await loadLists(true);openLists();
}}

// Column drag-to-reorder
function initColDrag(ul){{
  let drag=null,over=null;
  ul.querySelectorAll('li').forEach(li=>{{
    li.draggable=true;
    li.ondragstart=e=>{{drag=li;li.style.opacity='.4';e.dataTransfer.effectAllowed='move';}};
    li.ondragend=()=>{{drag.style.opacity='';drag=null;ul.querySelectorAll('li').forEach(l=>l.style.background='');}};
    li.ondragover=e=>{{e.preventDefault();if(li===drag)return;over=li;
      ul.querySelectorAll('li').forEach(l=>l.style.background='');
      li.style.background='#e0f2fe';}};
    li.ondrop=e=>{{e.preventDefault();if(!drag||drag===li)return;
      const items=[...ul.querySelectorAll('li')];
      const di=items.indexOf(drag),oi=items.indexOf(li);
      if(di<oi)ul.insertBefore(drag,li.nextSibling);else ul.insertBefore(drag,li);
      saveColOrder();}};
  }});
}}
async function saveColOrder(){{
  const ids=[...document.querySelectorAll('#col-sortable li')].map(li=>parseInt(li.dataset.id));
  await apiFetch('/api/columns/reorder/'+PID+'/'+state.tab,{{method:'POST',body:JSON.stringify({{order:ids}})}});
  toast('✔ Order saved','ok');
}}

// Columns
async function manageColumns(){{
  const cols=await apiFetch('/api/columns/'+PID+'/'+state.tab);if(!cols)return;
  const body=document.getElementById('col-body');body.innerHTML='';
  const ul=document.createElement('ul');ul.id='col-sortable';ul.className='slist';ul.style.maxHeight='400px';
  const visibleCols=isLTRTab()?cols.filter(c=>!isLTRInternalField(c)&&!isLTRExcludedField(c)):cols;
  visibleCols.forEach(col=>{{
    const li=document.createElement('li');li.className='sitem';li.dataset.id=col.id;li.draggable=true;
    li.style.cssText='cursor:default;align-items:center;gap:7px';
    const grip=document.createElement('span');grip.textContent='⠿';
    grip.style.cssText='cursor:grab;color:var(--mu);font-size:16px;flex-shrink:0';
    grip.title='Drag to reorder';
    const chk=document.createElement('input');chk.type='checkbox';chk.checked=col.visible;
    chk.onchange=()=>toggleCol(col.id,chk.checked);
    const nm=document.createElement('span');nm.textContent=col.label;nm.style.cssText='flex:1;font-size:12px';
    const tp=document.createElement('span');tp.textContent=col.col_type;
    tp.style.cssText='font-size:9px;background:#e0e7ff;color:#3730a3;padding:1px 6px;border-radius:3px;flex-shrink:0';
    const ren_btn=document.createElement('button');ren_btn.textContent='✏';
    ren_btn.title='Rename';ren_btn.className='btn btn-sc btn-sm';
    ren_btn.style.cssText='flex-shrink:0;padding:2px 6px;font-size:11px';
    ren_btn.onclick=async()=>{{
      const newLbl=prompt('New label for column:',col.label);
      if(!newLbl||!newLbl.trim())return;
      const r=await apiFetch('/api/columns/rename/'+col.id,{{method:'POST',body:JSON.stringify({{label:newLbl.trim()}})}});
      if(r&&r.ok){{nm.textContent=newLbl.trim();toast('✔ Renamed','ok');}}
      else toast('Error','er');
    }};
    const db_btn=document.createElement('button');db_btn.textContent='🗑';
    db_btn.title='Delete';db_btn.className='btn btn-er btn-sm';
    db_btn.style.cssText='flex-shrink:0;padding:2px 6px;font-size:11px';
    db_btn.onclick=()=>deleteCol(col.id,col.col_key,db_btn);
    li.appendChild(grip);li.appendChild(chk);li.appendChild(nm);li.appendChild(tp);li.appendChild(ren_btn);li.appendChild(db_btn);
    ul.appendChild(li);
  }});
  initColDrag(ul);
  body.appendChild(ul);openM('col-modal');
}}
async function toggleCol(id,v){{await apiFetch('/api/columns/visibility/'+id,{{method:'POST',body:JSON.stringify({{visible:v}})}});}}
async function deleteCol(id,key,btn){{const warn=key==='docNo'?'⚠ WARNING: Deleting Document No. column will break the register! Are you sure?':'Delete this column?';if(!confirm(warn))return;await apiFetch('/api/columns/'+id,{{method:'DELETE'}});btn.closest('li').remove();}}

function onColType(v){{
  document.getElementById('cg-list').style.display=v==='dropdown'?'flex':'none';
  document.getElementById('cg-ds').style.display=v==='duration_calc'?'flex':'none';
  document.getElementById('cg-de').style.display=v==='duration_calc'?'flex':'none';
  document.getElementById('cg-info').style.display=v==='duration_calc'?'block':'none';
}}
async function openAddCol(){{
  await loadLists();
  ensureLTRProjectLists();
  const ls=document.getElementById('col-list');
  ls.innerHTML=Object.keys(state.lists).filter(k=>!isHiddenSystemListName(k)).map(k=>`<option value="${{k}}">${{k}}</option>`).join('');
  const all=await apiFetch('/api/columns/'+PID+'/'+state.tab);
  const dates=(all||[]).filter(c=>['date','auto_date'].includes(c.col_type));
  const dopts=dates.map(c=>`<option value="${{c.col_key}}">${{c.label}}</option>`).join('')||'<option value="issuedDate">Issued Date</option>';
  document.getElementById('col-ds').innerHTML=dopts;
  document.getElementById('col-de').innerHTML=dopts;
  document.getElementById('col-name').value='';document.getElementById('col-type').value='text';onColType('text');
  openM('addcol-modal');
}}
async function saveAddCol(){{
  const name=document.getElementById('col-name').value.trim();
  const type=document.getElementById('col-type').value;
  const list=type==='dropdown'?document.getElementById('col-list').value:null;
  const ds=type==='duration_calc'?document.getElementById('col-ds').value:null;
  const de=type==='duration_calc'?document.getElementById('col-de').value:null;
  if(!name){{toast('Name required','er');return;}}
  if(type==='duration_calc'&&ds===de){{toast('Start and end must differ','er');return;}}
  await apiFetch('/api/columns/'+PID+'/'+state.tab,{{method:'POST',body:JSON.stringify({{label:name,col_type:type,list_name:type==='duration_calc'?(ds+','+de):list}})}});
  closeM('addcol-modal');closeM('col-modal');await loadRecords();toast('✔ Column added','ok');
}}

// Admin Panel (same as dashboard)
async function openAdmin(){{
  const [users,projects]=await Promise.all([apiFetch('/api/users'),apiFetch('/api/projects')]);
  if(!users||!projects) return;
  const body=document.getElementById('admin-body'); body.innerHTML='';
  const utitle=document.createElement('div');utitle.className='stitle';utitle.textContent='👥 Users';body.appendChild(utitle);
  for(const u of users){{
    const row=document.createElement('div');row.className='urow';
    row.innerHTML=`<span style="flex:1;font-weight:600">👤 ${{u.username}}</span>
      <span class="badge" style="background:#fef3c7;color:#92400e">${{u.role.toUpperCase()}}</span>
      ${{u.username!=='admin'?`<button class="btn btn-sc btn-sm" onclick="chgPw('${{u.username}}')">🔑 PW</button>
        <button class="btn btn-er btn-sm" onclick="delUsr('${{u.username}}')">✕</button>`:
        '<span style="font-size:10px;color:var(--mu)">(protected)</span>'}}`;
    body.appendChild(row);
    if(u.role!=='superadmin'){{
      const ad=document.createElement('div');ad.style.cssText='padding:4px 10px 10px 32px;border-bottom:1px solid var(--bd);margin-bottom:4px';
      ad.innerHTML='<div style="font-size:10px;color:var(--mu);margin-bottom:4px">Project access:</div>';
      const assigned=await apiFetch('/api/users/'+u.username+'/projects').catch(()=>[]);
      const pl=document.createElement('div');pl.style.cssText='display:flex;flex-wrap:wrap;gap:4px';
      projects.forEach(p=>{{
        const isOn=assigned.includes(p.id);
        const btn=document.createElement('button');
        btn.style.cssText=`padding:3px 10px;border-radius:4px;border:2px solid;cursor:pointer;font-size:11px;font-weight:700;font-family:inherit;transition:all .15s;background:${{isOn?'#1a3a5c':'#f1f5f9'}};color:${{isOn?'#fff':'#94a3b8'}};border-color:${{isOn?'#f0a500':'#e2e8f0'}}`;
        btn.textContent=p.code;btn.title=p.name;
        if(isOn) btn.dataset.on='1';
        btn.onclick=async()=>{{
          const on=!!btn.dataset.on;
          await apiFetch('/api/users',{{method:'POST',body:JSON.stringify({{action:on?'unassign':'assign',username:u.username,project_id:p.id}})}});
          if(on){{delete btn.dataset.on;btn.style.background='#f1f5f9';btn.style.color='#94a3b8';btn.style.borderColor='#e2e8f0';}}
          else{{btn.dataset.on='1';btn.style.background='#1a3a5c';btn.style.color='#fff';btn.style.borderColor='#f0a500';}}
          toast((btn.dataset.on?'✔ Assigned: ':'Removed: ')+p.name,'ok');
        }};
        pl.appendChild(btn);}});
      ad.appendChild(pl);body.appendChild(ad);
    }}
  }}
  const at=document.createElement('div');at.className='stitle';at.textContent='➕ Add User';body.appendChild(at);
  const ar=document.createElement('div');
  ar.innerHTML=`<div style="display:grid;grid-template-columns:1fr 1fr 1fr auto;gap:8px;align-items:end">
    <div class="fg"><label>Username</label><input id="nu-name" placeholder="username"></div>
    <div class="fg"><label>Role</label><select id="nu-role">
      <option value="editor">Editor</option><option value="viewer">Viewer</option>
      <option value="superadmin">Super Admin</option></select></div>
    <div class="fg"><label>Password</label><input id="nu-pw" type="password"></div>
    <button class="btn btn-pr btn-sm" style="margin-bottom:1px" onclick="addUsr()">Add</button></div>`;
  body.appendChild(ar);openM('admin-modal');
}}
async function addUsr(){{
  const name=document.getElementById('nu-name').value.trim().toLowerCase();
  const role=document.getElementById('nu-role').value;
  const pw=document.getElementById('nu-pw').value;
  if(!name||!pw){{toast('Username and password required','er');return;}}
  const r=await apiFetch('/api/users',{{method:'POST',body:JSON.stringify({{action:'add',username:name,role,password:pw}})}});
  if(r&&r.ok){{toast('✔ User added','ok');closeM('admin-modal');openAdmin();}}else toast((r&&r.error)||'Error','er');
}}
async function delUsr(u){{if(!confirm('Delete user: '+u+'?'))return;
  const r=await apiFetch('/api/users',{{method:'POST',body:JSON.stringify({{action:'delete',username:u}})}});
  if(r&&r.ok){{toast('Deleted','wa');closeM('admin-modal');openAdmin();}}
}}
async function chgPw(u){{const pw=prompt('New password for '+u+':');if(!pw)return;
  await apiFetch('/api/users',{{method:'POST',body:JSON.stringify({{action:'change_password',username:u,password:pw}})}});
  toast('✔ Password changed','ok');
}}

// Export/Import
function doExport(){{if(state.tab)window.location='/api/export/'+PID+'/'+state.tab;}}
function doExportPDF(){{if(state.tab)window.location='/api/export_pdf/'+PID+'/'+state.tab;}}
function doExportAllPDF(){{window.location='/api/export_pdf_all/'+PID;}}
function toggleExpDD(e){{e.stopPropagation();document.getElementById('exp-menu').classList.toggle('hidden');}}
function closeExpDD(){{document.getElementById('exp-menu').classList.add('hidden');}}
document.addEventListener('click',e=>{{if(!e.target.closest('#exp-dd'))closeExpDD();}});
function doPrint(){{
  const orig=document.title;
  document.title='DCR Print';
  window.print();
  document.title=orig;
}}
// Holidays Settings
async function openSettings(){{
  const r=await apiFetch('/api/settings/holidays');if(!r)return;
  let hols=[...(r.holidays||[])];
  const ov=document.createElement('div');ov.className='overlay';
  ov.style.cssText='position:fixed;inset:0;background:rgba(0,0,0,.55);z-index:9999;display:flex;align-items:center;justify-content:center';
  function buildUI(){{
    ov.innerHTML=`<div style="background:#fff;border-radius:12px;padding:24px;max-width:520px;width:95%;max-height:85vh;overflow:hidden;display:flex;flex-direction:column;box-shadow:0 20px 60px rgba(0,0,0,.3)">
      <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:14px">
        <div style="font-size:16px;font-weight:800;color:#1a3a5c">🗓 Public Holidays</div>
        <button id="hol-close" style="background:none;border:none;font-size:20px;cursor:pointer;color:#64748b">✕</button>
      </div>
      <p style="font-size:12px;color:#64748b;margin-bottom:12px">These dates are excluded from Duration and Expected Reply calculations (along with Fridays).</p>
      <div style="display:flex;gap:8px;margin-bottom:12px">
        <input type="date" id="hol-inp" style="flex:1;padding:6px 10px;border:1.5px solid #e2e8f0;border-radius:6px;font-family:inherit;font-size:12px">
        <button id="hol-add-btn" style="padding:6px 14px;background:#1a3a5c;color:#fff;border:none;border-radius:6px;cursor:pointer;font-size:12px;font-weight:600">+ Add</button>
      </div>
      <div id="hol-list" style="flex:1;overflow-y:auto;border:1px solid #e2e8f0;border-radius:6px;padding:8px;min-height:180px;max-height:340px"></div>
      <div style="margin-top:12px;font-size:11px;color:#94a3b8">${{hols.length}} holidays total</div>
      <div style="display:flex;gap:8px;margin-top:12px;justify-content:flex-end">
        <button id="hol-cancel" style="padding:7px 16px;border:1.5px solid #e2e8f0;background:#fff;border-radius:6px;cursor:pointer;font-size:12px">Cancel</button>
        <button id="hol-save" style="padding:7px 16px;background:#16a34a;color:#fff;border:none;border-radius:6px;cursor:pointer;font-size:12px;font-weight:600">💾 Save Holidays</button>
      </div></div>`;
    renderHL();
    ov.querySelector('#hol-close').onclick=()=>ov.remove();
    ov.querySelector('#hol-cancel').onclick=()=>ov.remove();
    ov.querySelector('#hol-add-btn').onclick=()=>{{
      const v=ov.querySelector('#hol-inp').value;
      if(!v)return;if(hols.includes(v)){{toast('Already added','wa');return;}}
      hols.push(v);renderHL();ov.querySelector('#hol-inp').value='';
    }};
    ov.querySelector('#hol-save').onclick=async()=>{{
      const btn=ov.querySelector('#hol-save');btn.disabled=true;btn.textContent='Saving...';
      const res=await apiFetch('/api/settings/holidays',{{method:'POST',body:JSON.stringify({{holidays:hols}})}});
      btn.disabled=false;btn.textContent='💾 Save Holidays';
      if(res&&res.ok){{toast('✔ '+res.count+' holidays saved','ok');ov.remove();}}
      else toast('Error','er');
    }};
  }}
  function renderHL(){{
    const ul=ov.querySelector('#hol-list');if(!ul)return;ul.innerHTML='';
    if(!hols.length){{ul.innerHTML='<div style="color:#94a3b8;text-align:center;padding:30px;font-size:12px">No holidays — click + Add to add dates</div>';return;}}
    [...hols].sort().forEach(d=>{{
      const row=document.createElement('div');
      row.style.cssText='display:flex;align-items:center;justify-content:space-between;padding:5px 8px;border-radius:4px;margin-bottom:3px;background:#f8fafc;font-size:12px';
      const dt=new Date(d+'T00:00:00');
      const fmt=dt.getDate().toString().padStart(2,'0')+'-'+(dt.getMonth()+1).toString().padStart(2,'0')+'-'+dt.getFullYear();
      const days=['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];
      const day=days[dt.getDay()];
      row.innerHTML=`<span style="color:#1e293b">${{fmt}} <span style="color:#94a3b8;font-size:10px">(${{day}})</span></span>
        <button data-d="${{d}}" style="background:none;border:none;color:#ef4444;cursor:pointer;font-size:15px;line-height:1;padding:2px 6px">✕</button>`;
      row.querySelector('button').onclick=e=>{{const dv=e.currentTarget.dataset.d;hols=hols.filter(h=>h!==dv);renderHL();}};
      ul.appendChild(row);
    }});
  }}
  buildUI();
  document.body.appendChild(ov);
}}

function doExportAll(){{window.location='/api/export_all/'+PID;}}
function openImport(){{openM('import-modal');}}
async function doImport(){{
  const file=document.getElementById('imp-file').files[0];if(!file)return;
  const ext=file.name.split('.').pop().toLowerCase();
  const btn=document.querySelector('#import-modal .btn-pr');
  btn.disabled=true;btn.textContent='⏳ Importing...';
  try{{
    const b64=await new Promise((res,rej)=>{{const fr=new FileReader();fr.onload=e=>res(e.target.result);fr.onerror=rej;fr.readAsDataURL(file);}});
    const r=await apiFetch('/api/import/'+PID+'/'+state.tab,{{method:'POST',body:JSON.stringify({{file_b64:b64,ext}})}});
    if(r===null)return;
    closeM('import-modal');switchTab(state.tab);await loadDTs();toast('✔ Imported '+r.imported+' records','ok');
  }}catch(e){{toast('Error: '+e.message,'er');}}
  finally{{btn.disabled=false;btn.textContent='Import';}}
}}
async function doImportProject(){{
  const file=document.getElementById('imp-file').files[0];if(!file)return;
  const ext=file.name.split('.').pop().toLowerCase();
  if(!['xlsx','xls'].includes(ext)){{toast('Full workbook import supports Excel only','er');return;}}
  const btn=document.getElementById('imp-project-btn');
  btn.disabled=true;btn.textContent='⏳ Importing...';
  try{{
    const b64=await new Promise((res,rej)=>{{const fr=new FileReader();fr.onload=e=>res(e.target.result);fr.onerror=rej;fr.readAsDataURL(file);}});
    const r=await apiFetch('/api/import_project/'+PID,{{method:'POST',body:JSON.stringify({{file_b64:b64,ext}})}});
    if(r===null)return;
    closeM('import-modal');switchTab(state.tab);await loadDTs();
    toast('✔ Imported '+r.imported_total+' records from '+(r.matched_sheets||[]).length+' sheet(s)','ok');
  }}catch(e){{toast('Error: '+e.message,'er');}}
  finally{{btn.disabled=false;btn.textContent='Import Full Workbook';}}
}}
</script></body></html>"""


