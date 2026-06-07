import re

with open("html_render.py", "r", encoding="utf-8") as f:
    c = f.read()

# Add button
btn_html = """    <button class="tbtn" onclick="showTab('executive')">📋 Executive</button>
    <button class="tbtn" onclick="showTab('daily-digest')">🌟 Daily Digest</button>"""
c = c.replace("""    <button class="tbtn" onclick="showTab('executive')">📋 Executive</button>""", btn_html)

# Add Container
container_html = """
    <!-- TAB: DAILY DIGEST -->
    <div id="tab-daily-digest" class="tab-pane">
      <div class="stitle">🌟 Daily Digest</div>
      <div id="daily-digest-content" style="display:grid;grid-template-columns:repeat(auto-fit,minmax(300px,1fr));gap:20px;">
        <div style="color:var(--mu);font-size:14px;">Loading Daily Digest...</div>
      </div>
    </div>
"""

# Insert container after tab-executive
exec_tab_end_idx = c.find("    <!-- TAB: ADMIN -->")
if exec_tab_end_idx != -1:
    c = c[:exec_tab_end_idx] + container_html + "\n" + c[exec_tab_end_idx:]
else:
    print("Could not find TAB: ADMIN")

# Add loadDailyDigest logic
script_js = """
async function loadDailyDigest() {
  const container = document.getElementById('daily-digest-content');
  container.innerHTML = '<div style="color:var(--mu);font-size:14px;">Loading...</div>';
  try {
    const targetPid = PID || 'all';
    const r = await apiFetch('/api/daily_digest/' + targetPid);
    if(!r) { container.innerHTML = '<div style="color:red">Failed to load digest.</div>'; return; }
    
    let html = '';
    
    // Received Card
    html += '<div class="panel" style="border-top:4px solid #3b82f6;"><div class="panel-title" style="color:#1e40af;font-size:16px;">📥 Received Today (' + r.received.length + ')</div><div style="margin-top:12px;">';
    if(r.received.length === 0) html += '<div style="color:var(--mu);font-size:13px;font-style:italic;">No documents received today.</div>';
    r.received.forEach(doc => {
      html += '<div style="padding:10px;border-bottom:1px solid var(--bd);"><div style="font-weight:700;color:var(--tx);font-size:13px;">'+doc.docNo+'</div><div style="color:var(--mu);font-size:12px;margin-top:4px;">'+doc.title+'</div></div>';
    });
    html += '</div></div>';

    // Issued Card
    html += '<div class="panel" style="border-top:4px solid #f97316;"><div class="panel-title" style="color:#c2410c;font-size:16px;">📤 Issued Today (' + r.issued.length + ')</div><div style="margin-top:12px;">';
    if(r.issued.length === 0) html += '<div style="color:var(--mu);font-size:13px;font-style:italic;">No documents issued today.</div>';
    r.issued.forEach(doc => {
      html += '<div style="padding:10px;border-bottom:1px solid var(--bd);"><div style="font-weight:700;color:var(--tx);font-size:13px;">'+doc.docNo+'</div><div style="color:var(--mu);font-size:12px;margin-top:4px;">'+doc.title+'</div></div>';
    });
    html += '</div></div>';

    // Replied Card
    html += '<div class="panel" style="border-top:4px solid #22c55e;"><div class="panel-title" style="color:#166534;font-size:16px;">✅ Replied Today (' + r.replied.length + ')</div><div style="margin-top:12px;">';
    if(r.replied.length === 0) html += '<div style="color:var(--mu);font-size:13px;font-style:italic;">No documents replied today.</div>';
    r.replied.forEach(doc => {
      html += '<div style="padding:10px;border-bottom:1px solid var(--bd);"><div style="font-weight:700;color:var(--tx);font-size:13px;">'+doc.docNo+'</div><div style="color:var(--mu);font-size:12px;margin-top:4px;">'+doc.title+'</div>';
      if(doc.status) html += '<div style="margin-top:4px;font-size:11px;font-weight:700;color:var(--tx2);">Status: ' + doc.status + '</div>';
      html += '</div>';
    });
    html += '</div></div>';

    container.innerHTML = html;
  } catch(e) {
    console.error(e);
    container.innerHTML = '<div style="color:red">Error loading daily digest.</div>';
  }
}
"""

c = c.replace("async function loadExecutiveTab()", script_js + "\nasync function loadExecutiveTab()")

# Call loadDailyDigest inside showTab
showtab_injection = """  if(name==='executive') loadExecutiveTab();
  if(name==='daily-digest') loadDailyDigest();"""
c = c.replace("  if(name==='executive') loadExecutiveTab();", showtab_injection)

# Also call loadDailyDigest when PID changes if tab is active
pid_change_injection = """  if(state.currentTab==='executive') loadExecutiveTab();
  if(state.currentTab==='daily-digest') loadDailyDigest();"""
c = c.replace("  if(state.currentTab==='executive') loadExecutiveTab();", pid_change_injection)


with open("html_render.py", "w", encoding="utf-8") as f:
    f.write(c)
