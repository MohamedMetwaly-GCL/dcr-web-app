import re

with open("html_render.py", "r", encoding="utf-8") as f:
    c = f.read()

container_html = """
    <!-- TAB: DAILY DIGEST -->
    <div id="tab-daily-digest" class="tab-pane">
      <div class="stitle">🌟 Daily Digest</div>
      <div id="daily-digest-content" style="display:grid;grid-template-columns:repeat(auto-fit,minmax(300px,1fr));gap:20px;">
        <div style="color:var(--mu);font-size:14px;">Loading Daily Digest...</div>
      </div>
    </div>
"""

# Find <!-- TAB: OVERVIEW --> and insert before it? 
# Or find <!-- TAB: SETTINGS -->
idx = c.find("    <!-- TAB: OVERVIEW -->")
c = c[:idx] + container_html + "\n" + c[idx:]

with open("html_render.py", "w", encoding="utf-8") as f:
    f.write(c)
