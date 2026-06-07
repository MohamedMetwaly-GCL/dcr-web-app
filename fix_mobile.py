import re

with open(r'D:\DCR\html_render.py', 'r', encoding='utf-8') as f:
    code = f.read()

# 1. Landscape Fix
landscape_css = '''
@media (max-width: 900px) and (orientation: landscape) {
  #projbar { display: none !important; }
  .mhdr { padding: 8px 12px !important; }
  #toolbar { padding: 2px 4px !important; }
  .kpi { min-height: 50px !important; padding: 6px !important; }
}
'''
code = code.replace('@media(max-width:768px){{\n', landscape_css.replace('{','{{').replace('}','}}') + '@media(max-width:768px){{\n')

# 2. Letter Views Bug
code = re.sub(r'display:flex!important;flex-direction:row!important;?', 'flex-direction:row!important;', code)
code = re.sub(r'#ltr-quickbar\{padding:4px 8px!important;gap:4px!important;display:flex!important;flex-direction:row!important\}', '#ltr-quickbar{padding:4px 8px!important;gap:4px!important;flex-direction:row!important}', code)

# 3. Header Fixes
# Short title
if '<span class="topbar-title-full"' in code:
    code = code.replace('<span class="topbar-title-full"', '<span class="topbar-title-short" style="display:none;font-weight:800;font-size:16px;">DCR</span>\n  <span class="topbar-title-full"')

mobile_fixes = '''
  .topbar-title-full { display:none!important; }
  .topbar-title-short { display:inline-block!important; }
  .valmore-logo img, .ll-brand img, .proj-logo { max-height:25px!important; }
  .pf-val { white-space:nowrap!important; overflow:hidden!important; text-overflow:ellipsis!important; max-width:140px!important; }
  .lists-modal-row { flex-wrap:wrap!important; }
  .slist { overflow-x: auto; }
  .admin-email-cell { word-break: break-all; }
  #toolbar { padding:2px 4px!important; }
  .tool-btn { padding: 4px 6px!important; font-size: 9px!important; min-height: 24px!important; }
'''
code = code.replace('@media(max-width:768px){{\n', f'@media(max-width:768px){{\n{mobile_fixes.replace("{","{{").replace("}","}}")}')

with open(r'D:\DCR\html_render.py', 'w', encoding='utf-8') as f:
    f.write(code)

print('Mobile fixes applied.')
