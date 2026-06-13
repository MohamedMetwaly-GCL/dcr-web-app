import re

with open("html_render.py", "r", encoding="utf-8") as f:
    c = f.read()

# 1. Update the UI section that renders the Distribution Matrix button
ui_start = c.find("// ── Distribution Matrix (visible to DC & Admins) ──────────────")
ui_end = c.find("const er={...DEFAULT_EXPECTED_REPLY_RULE", ui_start)

if ui_start != -1 and ui_end != -1:
    new_ui = """// ── Distribution Matrix & Master Magic Link ──────────────
  try {
    const isDCOrAdmin = _WHOAMI && (_WHOAMI.role==='superadmin'||_WHOAMI.role==='admin'||(_WHOAMI.dc_projects||[]).includes(PID));
    if(isDCOrAdmin) {
      const distTitle=document.createElement('div');distTitle.className='stitle';distTitle.style.marginTop='18px';
      distTitle.innerHTML='📇 Distribution Matrix & Master Link';body.appendChild(distTitle);
      
      const distNote=document.createElement('div');
      distNote.style.cssText='font-size:10px;color:var(--mu);margin:-4px 0 10px;padding:6px 8px;background:var(--bg);border-radius:4px';
      distNote.textContent='Map Disciplines to internal team members to drive the Smart Project Digest.';
      body.appendChild(distNote);
      
      const distBtn=document.createElement('button');
      distBtn.className='btn btn-sc';
      distBtn.style.cssText='width:100%;display:flex;align-items:center;justify-content:center;gap:8px;padding:10px;margin-bottom:8px';
      distBtn.innerHTML='📇 Open Distribution Matrix';
      distBtn.onclick=()=>openDistributionMatrix(PID);
      body.appendChild(distBtn);

      const masterLinkBtn=document.createElement('button');
      masterLinkBtn.className='btn btn-ok';
      masterLinkBtn.style.cssText='width:100%;display:flex;align-items:center;justify-content:center;gap:8px;padding:10px';
      masterLinkBtn.innerHTML='🔗 Generate Master Magic Link';
      masterLinkBtn.onclick=async ()=>{
        masterLinkBtn.disabled = true;
        masterLinkBtn.innerHTML = 'Generating...';
        try {
          const r = await apiFetch(`/api/magic/generate_master/${PID}`, {method:'POST'});
          if(r && r.ok) {
            navigator.clipboard.writeText(r.link);
            toast('Master Link copied to clipboard!','ok');
          } else toast('Failed to generate link', 'er');
        } catch(e){ toast('Error generating link','er'); }
        masterLinkBtn.disabled = false;
        masterLinkBtn.innerHTML = '🔗 Generate Master Magic Link';
      };
      body.appendChild(masterLinkBtn);
    }
  } catch(e){ console.error('[DistMatrix UI]',e); }

  """
    c = c[:ui_start] + new_ui + c[ui_end:]


# 2. Update the openDistributionMatrix function
func_start = c.find("async function openDistributionMatrix(pid) {")
if func_start != -1:
    func_end = c.find("async function syncDriveLinks(btn)", func_start)
    
    new_func = """async function openDistributionMatrix(pid) {
  if(!pid){ toast('No project selected','er'); return; }

  try {
    const [projects, projUsers] = await Promise.all([
      apiFetch('/api/projects').catch(()=>[]),
      apiFetch('/api/project_users/'+pid).catch(()=>[])
    ]);
    
    const proj = projects.find(p => p.id === pid || p.code === pid) || {};
    let matrix = (proj.data || {}).distribution_matrix || {"Project Management": []};
    
    if (!matrix["Project Management"]) {
      matrix["Project Management"] = [];
    }
    
    const body = document.getElementById('dist-body');
    body.innerHTML = '';
    
    const note = document.createElement('div');
    note.style.cssText = 'padding: 10px; background: #eff6ff; color: #1e3a8a; font-size: 12px; border-radius: 6px; margin-bottom: 15px; border: 1px solid #bfdbfe;';
    note.innerHTML = 'Define Role-Based distribution here. Add Disciplines and assign users. Users assigned to <b>Project Management</b> will bypass filters and see all documents.';
    body.appendChild(note);

    function makeUserSelector(discipline, initialUsers) {
      const wrap = document.createElement('div');
      wrap.style.cssText = 'display:flex;flex-wrap:wrap;gap:4px;padding:6px 8px;border:1px solid var(--bd);border-radius:6px;min-height:36px;background:var(--bg);cursor:text;';
      
      let users = [...(initialUsers||[])];
      
      function renderTags() {
        wrap.innerHTML='';
        users.forEach((em,i)=>{
          const chip=document.createElement('span');
          chip.style.cssText='display:inline-flex;align-items:center;gap:4px;padding:2px 8px;border-radius:99px;background:#dbeafe;color:#1e40af;font-size:11px;font-weight:600';
          chip.innerHTML=em+' <span style="cursor:pointer;font-size:14px;line-height:1" onclick="this.parentNode.remove();users.splice('+i+',1);updateMatrix()">✕</span>';
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
        
        sel.onchange=(e)=>{
          const val = sel.value;
          if(val && !users.includes(val)) {
            users.push(val);
            updateMatrix();
            renderTags();
          }
        };
        wrap.appendChild(sel);
      }
      
      function updateMatrix() {
        matrix[discipline] = users;
      }
      
      renderTags();
      return wrap;
    }

    const rowsContainer = document.createElement('div');
    
    function renderRows() {
      rowsContainer.innerHTML = '';
      for (const [discipline, users] of Object.entries(matrix)) {
        const row = document.createElement('div');
        row.style.cssText = 'margin-bottom: 12px; padding: 12px; border: 1px solid var(--bd); border-radius: 8px; background: var(--bg2, #f8fafc);';
        
        const header = document.createElement('div');
        header.style.cssText = 'display: flex; justify-content: space-between; align-items: center; margin-bottom: 8px;';
        
        const title = document.createElement('div');
        title.style.cssText = 'font-weight: bold; font-size: 14px; color: #0f172a;';
        title.textContent = discipline;
        if (discipline === "Project Management") title.innerHTML += ' <span style="font-size:10px; background:#fef08a; padding:2px 6px; border-radius:4px; margin-left:4px;">BYPASS FILTER</span>';
        
        const delBtn = document.createElement('button');
        delBtn.innerHTML = '🗑️';
        delBtn.style.cssText = 'background:none; border:none; cursor:pointer; font-size:12px;';
        if (discipline !== "Project Management") {
          delBtn.onclick = () => {
            delete matrix[discipline];
            renderRows();
          };
          header.appendChild(title);
          header.appendChild(delBtn);
        } else {
          header.appendChild(title);
        }
        
        row.appendChild(header);
        row.appendChild(makeUserSelector(discipline, users));
        rowsContainer.appendChild(row);
      }
    }
    renderRows();
    body.appendChild(rowsContainer);
    
    const addWrap = document.createElement('div');
    addWrap.style.cssText = 'display: flex; gap: 8px; margin-top: 15px;';
    const discInput = document.createElement('input');
    discInput.placeholder = 'New Discipline Name...';
    discInput.className = 'inp';
    discInput.style.flex = '1';
    
    const addBtn = document.createElement('button');
    addBtn.className = 'btn btn-ok';
    addBtn.textContent = 'Add Discipline';
    addBtn.onclick = () => {
      const v = discInput.value.trim();
      if(v && !matrix[v]) {
        matrix[v] = [];
        discInput.value = '';
        renderRows();
      }
    };
    
    addWrap.appendChild(discInput);
    addWrap.appendChild(addBtn);
    body.appendChild(addWrap);
    
    const saveWrap = document.createElement('div');
    saveWrap.style.cssText = 'margin-top: 20px; text-align: right; border-top: 1px solid var(--bd); padding-top: 15px;';
    const finalSaveBtn = document.createElement('button');
    finalSaveBtn.className = 'btn btn-pr';
    finalSaveBtn.textContent = '💾 Save Matrix';
    finalSaveBtn.onclick = async () => {
      finalSaveBtn.disabled = true;
      finalSaveBtn.textContent = 'Saving...';
      try {
        const r = await fetch(`/api/project/${pid}/settings`, {
          method: 'POST',
          headers: {'Content-Type': 'application/json'},
          body: JSON.stringify({ distribution_matrix: matrix })
        });
        const d = await r.json();
        if(d && d.ok) {
          toast('Matrix saved successfully', 'ok');
          closeM('dist-modal');
        } else {
          toast(d.error || 'Failed to save', 'er');
        }
      } catch(e) {
        toast('Network Error', 'er');
      }
      finalSaveBtn.disabled = false;
      finalSaveBtn.textContent = '💾 Save Matrix';
    };
    saveWrap.appendChild(finalSaveBtn);
    body.appendChild(saveWrap);

    openM('dist-modal');
  } catch(e){
    console.error('[DistMatrix]', e);
    toast('Error loading distribution matrix','er');
  }
}

"""
    c = c[:func_start] + new_func + c[func_end:]


with open("html_render.py", "w", encoding="utf-8") as f:
    f.write(c)

print("html_render.py patched successfully!")
