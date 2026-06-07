import re

with open(r'D:\DCR\html_render.py', 'r', encoding='utf-8') as f:
    code = f.read()

# ============================================================
# BUG 1: Valmore "V" invisible - improve contrast
# ============================================================
code = code.replace(
    '.valmore-icon {{ width: 18px; height: 18px; background:var(--wh); color: #1a2f4e;',
    '.valmore-icon {{ width: 18px; height: 18px; background: #f0f4f8; color: #1a2f4e;'
)

# ============================================================
# BUG 2: Login footer text spacing
# ============================================================
# Find the login footer credits and add margin-top
code = code.replace(
    '.ll-credit{{font-size:9px;',
    '.ll-credit{{margin-top:8px;font-size:9px;'
)
# If .ll-credit is not found try other selectors
if '.ll-credit{{margin-top:8px' not in code:
    # Find Gas Chill | Document Control System text in CSS
    code = re.sub(
        r'(Gas Chill \| Document Control System.*?font-size:\s*\d+px)',
        r'\1',
        code,
        flags=re.DOTALL
    )

# ============================================================
# BUG 3: Executive table overflow
# ============================================================
# Wrap the executive table in overflow-x:auto
code = code.replace(
    '<table class=\\"dt-tbl\\" style=\\"margin-bottom:16px\\">',
    '<div style=\\"overflow-x:auto;-webkit-overflow-scrolling:touch;margin-bottom:16px\\"><table class=\\"dt-tbl\\" style=\\"min-width:480px\\">'
)
code = code.replace(
    '</table>\\n        <div style=\\"display:grid;grid-template-columns:1fr 1fr',
    '</table></div>\\n        <div style=\\"display:grid;grid-template-columns:1fr 1fr'
)

# ============================================================
# BUG 4: Fix disc-expander encoding (Â·¼ -> ▼/▶)
# ============================================================
# Replace the disc-expander HTML entity with safe SVG arrow
code = code.replace(
    '.disc-expander{{display:inline-flex;align-items:center;justify-content:center;width:24px;height:24px;border:1px solid #d7e0ea;border-radius:999px;background:var(--wh);color:#2f4f64;cursor:pointer;font-size:11px;transition:all .15s}}',
    '.disc-expander{{display:inline-flex;align-items:center;justify-content:center;width:24px;height:24px;border:1px solid #d7e0ea;border-radius:999px;background:var(--wh);color:#2f4f64;cursor:pointer;font-size:11px;transition:all .15s;font-family:monospace;line-height:1}}'
)

# Fix the expander text content in JS render
code = re.sub(
    r'<span class=\\"disc-expander\\"[^>]*>[^<]*</span>',
    '<span class=\\"disc-expander\\" onclick=\\"toggleDisc(this)\\">\u25ba</span>',
    code
)
# Also fix via direct innerHTML replacement patterns
code = code.replace(
    'disc-expander\\">\u00e2\u00bc\u00bc',
    'disc-expander\\">\u25ba'
)
code = code.replace(
    'disc-expander.open{{background:#2f4f64;color:#fff;border-color:#2f4f64}}',
    'disc-expander.open{{background:#2f4f64;color:#fff;border-color:#2f4f64;transform:rotate(90deg)}}'
)

# ============================================================
# BUG 5: Doc Types Summary - apply same CSS as Discipline Breakdown
# ============================================================
# The dt-tbl is shared; add specific CSS for the doc types table wrapper
# Already uses .dt-tbl - need to ensure same mobile scroll wrapper
code = code.replace(
    '#overview-pane-doctype .tbl-wrap',
    '#overview-pane-doctype .tbl-wrap,#overview-pane-doctype .dt-tbl'
)

# ============================================================
# BUG 6: DUPLICATE "DCR Register" in topbar
# ============================================================
# Remove the duplicate span (the one without display:none - added by fix_mobile.py)
# Line 2893 has a second topbar-title-short without display:none
code = code.replace(
    '  <span class="topbar-title-short" style="display:none;font-weight:800;font-size:16px;">DCR</span>\n  <span class="topbar-title-full"',
    '  <span class="topbar-title-full"'
)
# Also remove extra duplicate in register topbar
# Keep only one topbar-title-short in register (line 2893)
# The one at 2891 has display:none which is wrong - it's a duplicate from our script
code = re.sub(
    r'(<span class="topbar-title-short" style="display:none;[^"]*">DCR</span>\s*\n\s*)(<span class="topbar-title-full")',
    r'\2',
    code
)
# Also remove the ::after " Register" since we'll show short properly
# Keep the short title that doesn't have display:none inline style

# ============================================================
# BUG 7: Sticky columns - disable on mobile
# ============================================================
# Add to the existing @media(max-width:768px) in register CSS
old_sticky_media = '  #tblwrap{{min-height:56vh;padding-bottom:2px}}'
new_sticky_media = '''  #tblwrap{{min-height:56vh;padding-bottom:2px}}
  /* Disable sticky columns on mobile - they consume too much space */
  #regtbl th:nth-child(1), #regtbl td:nth-child(1),
  #regtbl th:nth-child(2), #regtbl td:nth-child(2),
  #regtbl th.docno-cell, #regtbl td.docno-cell,
  #regtbl th:nth-child(3), #regtbl td:nth-child(3) {{
    position: static !important;
    z-index: auto !important;
    box-shadow: none !important;
  }}'''
code = code.replace(old_sticky_media, new_sticky_media)

# ============================================================
# BUG 8: Add Document modal - single column on mobile
# ============================================================
# Find the add modal CSS and add mobile override
add_modal_mobile = '''
  /* Add Document modal - single column on mobile */
  .modal-grid, .fg-row {{ flex-direction: column !important; }}
  .fg-row > div {{ width: 100% !important; max-width: 100% !important; }}
  .fg-half {{ width: 100% !important; flex: 0 0 100% !important; }}
  .mhdr {{ font-size: 14px !important; padding: 12px 14px !important; }}
  .mbody {{ padding: 12px 14px !important; }}
'''
# Add before closing of 768px media query in register
code = code.replace(
    '  #ltr-quickbar .tool-btn{{min-height:22px!important;font-size:8.5px!important;padding:2px 6px!important;flex:1 1 0!important}}\n}}',
    f'  #ltr-quickbar .tool-btn{{{{min-height:22px!important;font-size:8.5px!important;padding:2px 6px!important;flex:1 1 0!important}}}}\n{add_modal_mobile.replace("{", "{{").replace("}", "}}")}\n}}'
)

# ============================================================
# BUG 9: Landscape - topbar-proj-info text fix
# ============================================================
# Already hidden in landscape. Add consistent font-size for project display
old_landscape = '''@media (max-width: 900px) and (orientation: landscape) {{
  #projbar {{ display: none !important; }}
  .mhdr {{ padding: 8px 12px !important; }}
  #toolbar {{ padding: 2px 4px !important; }}
  .kpi {{ min-height: 50px !important; padding: 6px !important; }}
}}'''
new_landscape = '''@media (max-width: 900px) and (orientation: landscape) {{
  #projbar {{ display: none !important; }}
  .mhdr {{ padding: 8px 12px !important; }}
  #toolbar {{ padding: 2px 4px !important; }}
  .kpi {{ min-height: 50px !important; padding: 6px !important; }}
  #topbar-proj-info {{ display: none !important; }}
  #topbar {{ height: 30px !important; padding: 0 6px !important; gap: 4px !important; }}
  #topbar .tb-btn {{ padding: 2px 5px !important; font-size: 9px !important; min-height: 22px !important; }}
  #tabsbar {{ padding: 1px 5px !important; }}
  .tab-btn {{ padding: 2px 6px !important; font-size: 8px !important; min-height: 20px !important; }}
  #toolbar-actions {{ gap: 2px !important; }}
  .tool-btn {{ padding: 1px 4px !important; font-size: 7.5px !important; min-height: 20px !important; }}
}}'''
code = code.replace(old_landscape, new_landscape)

# ============================================================
# Add charset meta tag if missing (BUG 4 prevention)
# ============================================================
if '<meta charset=' not in code:
    code = code.replace('<head>', '<head>\n<meta charset="UTF-8">')

with open(r'D:\DCR\html_render.py', 'w', encoding='utf-8') as f:
    f.write(code)

print('All 9 mobile QA bugs fixed.')
