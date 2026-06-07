with open('D:\\DCR\\html_render.py', 'r', encoding='utf-8') as f:
    code = f.read()

css_old = '''/* 1. Remove Vertical Borders & Force Zebra Striping */
table, .dt-tbl, #regtbl { border-collapse: collapse !important; }
table td, table th, .dt-tbl td, .dt-tbl th, #regtbl td, #regtbl th { border-left: none !important; border-right: none !important; border-bottom: 1px solid #e2e8f0 !important; }
table tbody tr:nth-child(even) td, .dt-tbl tbody tr:nth-child(even) td, #regtbl tbody tr:nth-child(even) td { background-color: #f8fafc !important; }
table tbody tr:hover td, .dt-tbl tbody tr:hover td, #regtbl tbody tr:hover td { background-color: #eef5ff !important; }

/* 2. Frozen Columns (Checkbox, Sr, Document No) */
/* Force exact widths to prevent gaps between sticky columns */
#regtbl th:nth-child(1), #regtbl td:nth-child(1) { width: 32px !important; min-width: 32px !important; max-width: 32px !important; position: sticky !important; left: 0 !important; z-index: 5 !important; background-color: #ffffff !important; }
#regtbl th:nth-child(2), #regtbl td:nth-child(2) { width: 34px !important; min-width: 34px !important; max-width: 34px !important; position: sticky !important; left: 32px !important; z-index: 5 !important; background-color: #ffffff !important; }
#regtbl th.docno-cell, #regtbl td.docno-cell, #regtbl th:nth-child(3), #regtbl td:nth-child(3) { 
    position: sticky !important; left: 66px !important; z-index: 5 !important; background-color: #ffffff !important; 
    box-shadow: 2px 0 5px -2px rgba(0,0,0,0.1) !important;
    font-family: 'Consolas', 'Courier New', monospace !important; font-weight: 600 !important; color: var(--brand-teal) !important;
}

/* Zebra Striping & Hover for Frozen Columns */
#regtbl tbody tr:nth-child(even) td:nth-child(1), #regtbl tbody tr:nth-child(even) td:nth-child(2), #regtbl tbody tr:nth-child(even) td:nth-child(3), #regtbl tbody tr:nth-child(even) td.docno-cell { background-color: #f8fafc !important; }
#regtbl tbody tr:hover td:nth-child(1), #regtbl tbody tr:hover td:nth-child(2), #regtbl tbody tr:hover td:nth-child(3), #regtbl tbody tr:hover td.docno-cell { background-color: #eef5ff !important; }

/* Header z-index for Frozen Columns */
#regtbl th:nth-child(1), #regtbl th:nth-child(2), #regtbl th:nth-child(3), #regtbl th.docno-cell { z-index: 10 !important; background-color: var(--brand-navy) !important; }'''

css_new = '''/* 1. Remove Vertical Borders & Force Zebra Striping */
table, .dt-tbl, #regtbl { border-collapse: collapse !important; }
table td, table th, .dt-tbl td, .dt-tbl th, #regtbl td, #regtbl th { border-left: none !important; border-right: none !important; border-bottom: 1px solid #e2e8f0 !important; }

/* Proper background inherit for sticky columns */
#regtbl tbody tr td { background-color: inherit !important; }
#regtbl tbody tr { background-color: #ffffff; }
#regtbl tbody tr:nth-child(even) { background-color: #f8fafc; }
#regtbl tbody tr:hover { background-color: #f1f5f9; filter: brightness(0.97); }

body.dark #regtbl tbody tr { background-color: var(--bg); }
body.dark #regtbl tbody tr:nth-child(even) { background-color: #1e293b; }
body.dark #regtbl tbody tr:hover { background-color: #334155; }

/* 2. Frozen Columns (Checkbox, Sr, Document No) */
#regtbl th:nth-child(1), #regtbl td:nth-child(1) { width: 32px !important; min-width: 32px !important; max-width: 32px !important; position: sticky !important; left: 0 !important; z-index: 5 !important; background-color: inherit !important; }
#regtbl th:nth-child(2), #regtbl td:nth-child(2) { width: 34px !important; min-width: 34px !important; max-width: 34px !important; position: sticky !important; left: 32px !important; z-index: 5 !important; background-color: inherit !important; }
#regtbl th.docno-cell, #regtbl td.docno-cell, #regtbl th:nth-child(3), #regtbl td:nth-child(3) { 
    position: sticky !important; left: 66px !important; z-index: 5 !important; background-color: inherit !important; 
    box-shadow: 2px 0 5px -2px rgba(0,0,0,0.1) !important;
}

/* Specific styling for the data cells (not header) */
#regtbl td.docno-cell, #regtbl td:nth-child(3) {
    font-family: 'Consolas', 'Courier New', monospace !important; font-weight: 600 !important; color: var(--brand-teal) !important;
}

/* Header z-index for Frozen Columns */
#regtbl th:nth-child(1), #regtbl th:nth-child(2), #regtbl th:nth-child(3), #regtbl th.docno-cell { z-index: 10 !important; background-color: var(--brand-navy) !important; }
body.dark #regtbl th:nth-child(1), body.dark #regtbl th:nth-child(2), body.dark #regtbl th:nth-child(3), body.dark #regtbl th.docno-cell { background-color: #0b131e !important; }'''

polish_old = '''body.dark .mbody, body.dark .slist, body.dark .tool-dd-menu { scrollbar-color: rgba(255,255,255,0.2) transparent !important; }
</style>'''

polish_new = '''body.dark .mbody, body.dark .slist, body.dark .tool-dd-menu { scrollbar-color: rgba(255,255,255,0.2) transparent !important; }

/* === PHASE 5 UI POLISH === */
.sbadge::before { display: none !important; }
#projbar .pf { flex: 1 1 auto !important; min-width: 0 !important; }
#projbar .pf-val { white-space: nowrap !important; overflow: hidden !important; text-overflow: ellipsis !important; }

/* Topbar Project Info (Moved from projbar) */
#topbar-proj-info:hover { background: rgba(255,255,255,0.08); }
#topbar-proj-info .pf { display: flex; flex-direction: column; gap: 2px; }
#topbar-proj-info .pf-lbl { font-size: 9px; color: rgba(255,255,255,0.6); text-transform: uppercase; letter-spacing: 0.5px; font-weight: 700; }
#topbar-proj-info .pf-val { font-size: 15px; color: #ffffff; font-weight: 800; white-space: nowrap; }
#topbar-proj-info .pf.primary:first-child .pf-val { font-size: 14px; color: var(--brand-teal); }
</style>'''

html_old = '''<div id="topbar">
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
</div>'''

html_new = '''<div id="topbar">
  <span class="topbar-mark" style="font-size:20px">📋</span>
  <span class="topbar-title-full" style="font-weight:700;font-size:14px;letter-spacing:.3px">Document Control Register</span>
  <span class="topbar-title-short" style="font-weight:800;font-size:13px;letter-spacing:.35px">DCR</span>
  
  <div style="flex:1"></div>
  
  <div id="topbar-proj-info" onclick="if(typeof editProject === 'function') editProject()" style="display:flex;align-items:center;gap:24px;cursor:pointer;padding:4px 16px;border-radius:8px;transition:all 0.2s" title="Edit Project Information">
    {projbar_primary}
  </div>

  <div style="flex:1"></div>

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
    {projbar_toggle}
    <div id="projbar-extra" style="display:flex;gap:20px;flex:1 1 auto;overflow-x:auto">{projbar_secondary}</div>
  </div>
  {'<button class="proj-edit-btn" onclick="editProject()" style="margin-left:auto;background:var(--pr);color:#fff;border:none;padding:5px 12px;border-radius:var(--rd);cursor:pointer;font-size:11px;font-family:inherit;flex-shrink:0">✏ Edit</button>' if editable else ''}
</div>'''

c1, c2, c3 = code.count(css_old), code.count(polish_old), code.count(html_old)
if c1==1 and c2==1 and c3==1:
    code = code.replace(css_old, css_new).replace(polish_old, polish_new).replace(html_old, html_new)
    with open('D:\\DCR\\html_render.py', 'w', encoding='utf-8') as f:
        f.write(code)
    print('SUCCESS')
else:
    print(f'FAIL: counts {c1}, {c2}, {c3}')
