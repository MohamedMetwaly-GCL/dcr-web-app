import re
with open("html_render.py", "r", encoding="utf-8") as f:
    c = f.read()

# I need to completely replace openDistributionMatrix(pid)
# I'll find its start and end.
start_idx = c.find("async function openDistributionMatrix(pid) {")
if start_idx != -1:
    end_idx = c.find("async function syncDriveLinks(btn)", start_idx)
    
    new_func = """async function openDistributionMatrix(pid) {
  if(!pid){ toast('No project selected','er'); return; }

  try {
    const [docTypes, savedDist, projUsers] = await Promise.all([
      apiFetch('/api/doc_types/'+pid).catch(()=>[]),
      apiFetch('/api/distribution/'+pid).catch(()=>({})),
      apiFetch('/api/project_users/'+pid).catch(()=>[])
    ]);
    if(!docTypes || !docTypes.length){ toast('No doc types found for this project','wa'); return; }
    
    const body = document.getElementById('dist-body');
    body.innerHTML = '';

    function makeUserSelector(initialUsers, pid, dtId) {
      const wrap = document.createElement('div');
      wrap.style.cssText = 'display:flex;flex-wrap:wrap;gap:4px;padding:6px 8px;border:1px solid var(--bd);border-radius:6px;min-height:36px;background:var(--bg);cursor:text;';
      
      let users = [...(initialUsers||[])];
      
      function renderTags() {
        wrap.innerHTML='';
        users.forEach((em,i)=>{
          const chip=document.createElement('span');
          chip.style.cssText='display:inline-flex;align-items:center;gap:4px;padding:2px 8px;border-radius:99px;background:#dbeafe;color:#1e40af;font-size:11px;font-weight:600';
          chip.innerHTML=em+' <span style="cursor:pointer;font-size:14px;line-height:1" onclick="this.parentNode.remove();users.splice('+i+',1);saveDistRow()">✕</span>';
          wrap.appendChild(chip);
        });
        
        const sel=document.createElement('select');
        sel.style.cssText='border:none;outline:none;background:transparent;font-size:12px;min-width:180px;flex:1;color:var(--tx)';
        sel.innerHTML = '<option value="">+ Add User...</option>';
        projUsers.forEach(u => {
          if(!users.includes(u.username)) {
            sel.innerHTML += `<option value="${u.username}">${u.username} (${u.role})</option>`;
          }
        });
        
        sel.onchange=async(e)=>{
          const val = sel.value;
          if(val && !users.includes(val)) {
            users.push(val);
            await saveDistRow();
            renderTags();
          }
        };
        wrap.appendChild(sel);
      }
      
      async function saveDistRow() {
        try {
          const r = await apiFetch('/api/distribution/'+pid, {
            method:'POST',
            body: JSON.stringify({doc_type_id:dtId, event_type:'access', emails:users})
          });
          if(r&&r.ok) toast('✔ Saved','ok');
          else toast('Save failed','er');
        } catch(e){ toast('Save error','er'); console.error(e); }
      }
      
      renderTags();
      return wrap;
    }

    docTypes.forEach(dt=>{
      const dtSect=document.createElement('div');
      dtSect.style.cssText='margin-bottom:20px;border:1px solid var(--bd);border-radius:8px;overflow:hidden';

      const dtHdr=document.createElement('div');
      dtHdr.style.cssText='padding:10px 14px;background:var(--bg2,#f1f5f9);font-weight:700;font-size:13px;display:flex;align-items:center;justify-content:space-between;gap:8px';
      
      const titleSpan = document.createElement('div');
      titleSpan.innerHTML='<span style="font-size:16px">📄</span> '+dt.name+' <code style="font-size:10px;padding:1px 6px;border-radius:4px;background:#e2e8f0;color:#475569">'+dt.code+'</code>';
      
      const magicBtn = document.createElement('button');
      magicBtn.className = 'btn-ok';
      magicBtn.innerHTML = 'Generate Magic Link 🔗';
      magicBtn.style.padding = '4px 10px';
      magicBtn.style.fontSize = '12px';
      magicBtn.onclick = async () => {
        try {
          const r = await apiFetch(`/api/magic/generate/${pid}/${dt.id}`, {method:'POST'});
          if(r && r.ok) {
            navigator.clipboard.writeText(r.link);
            toast('Magic Link copied to clipboard!','ok');
          } else toast('Failed to generate link', 'er');
        } catch(e){ toast('Error','er'); }
      };
      
      dtHdr.appendChild(titleSpan);
      dtHdr.appendChild(magicBtn);
      dtSect.appendChild(dtHdr);

      const distRows=document.createElement('div');
      distRows.style.padding='12px 14px';
      
      const evtRow=document.createElement('div');
      evtRow.style.cssText='display:grid;grid-template-columns:200px 1fr;gap:10px;align-items:start;margin-bottom:12px;padding-bottom:12px;border-bottom:none';
      
      const evtLabel=document.createElement('div');
      evtLabel.innerHTML='<span style="font-size:12px;font-weight:700;color:#0f172a">Assigned Engineers</span><div style="font-size:10px;color:var(--mu);margin-top:2px">Users with access to daily digest</div>';
      
      const savedUsers = (savedDist[dt.id]||{})['access']||[];
      const tagInput = makeUserSelector(savedUsers, pid, dt.id);
      
      evtRow.appendChild(evtLabel);
      evtRow.appendChild(tagInput);
      distRows.appendChild(evtRow);
      
      dtSect.appendChild(distRows);
      body.appendChild(dtSect);
    });

    openM('dist-modal');
  } catch(e){
    console.error('[DistMatrix]', e);
    toast('Error loading distribution matrix','er');
  }
}

"""
    c = c[:start_idx] + new_func + c[end_idx:]

with open("html_render.py", "w", encoding="utf-8") as f:
    f.write(c)
