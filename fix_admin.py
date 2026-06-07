import sys

with open(r'D:\DCR\html_render.py', 'r', encoding='utf-8') as f:
    code = f.read()

# 1. Fix the dropdown height
code = code.replace("roleSel.style.cssText='height:28px;padding:2px 7px';", "roleSel.style.cssText='height:auto;padding:4px 8px;outline:none;';")

# 2. Fix openAdmin slowness
old_open1 = """  const [users,projects]=await Promise.all([apiFetch('/api/users'),apiFetch('/api/projects')]);
  body.innerHTML='<div class="stitle" style="margin-top:0">👥 USERS</div>';
  for (const u of users) {"""
new_open1 = """  const [users,projects]=await Promise.all([apiFetch('/api/users'),apiFetch('/api/projects')]);
  const assignments=await Promise.all(users.map(u=>u.role!=='superadmin'?apiFetch('/api/users/'+u.username+'/projects').catch(()=>[]):Promise.resolve([])));
  body.innerHTML='<div class="stitle" style="margin-top:0">👥 USERS</div>';
  let _i = 0;
  for (const u of users) {
    const assigned_cached = assignments[_i++];"""

code = code.replace(old_open1, new_open1)

old_fetch1 = """      const assigned=await apiFetch('/api/users/'+u.username+'/projects').catch(()=>[]);"""
new_fetch1 = """      const assigned=assigned_cached;"""

code = code.replace(old_fetch1, new_fetch1)

old_onclick = """        btn.onclick=async()=>{
          const on=!!btn.dataset.on;
          await apiFetch('/api/users',{method:'POST',body:JSON.stringify({action:on?'unassign':'assign',username:u.username,project_id:p.id,is_dc:false})});
          openAdmin();
        };"""

new_onclick = """        btn.onclick=async()=>{
          const on=!!btn.dataset.on;
          btn.dataset.on = on ? '' : '1';
          const isOn = !on;
          btn.style.borderColor = isDC?'#22c55e':(isOn?'#f0a500':'#e2e8f0');
          btn.style.background = isOn?'#1a3a5c':'#f8fafc';
          btn.style.color = isOn?'#fff':'#94a3b8';
          btn.textContent = p.code + (isDC ? ' (DC)' : '');
          try {
            await apiFetch('/api/users',{method:'POST',body:JSON.stringify({action:on?'unassign':'assign',username:u.username,project_id:p.id,is_dc:false})});
          } catch(e) { openAdmin(); }
        };"""

code = code.replace(old_onclick, new_onclick)

old_oncontext = """        btn.oncontextmenu = async (e) => {
          e.preventDefault();
          if(!isOn) return toast('User must be assigned to project first','wa');
          await apiFetch('/api/users', {method:'POST', body:JSON.stringify({action:'assign',username:u.username,project_id:p.id, is_dc: !isDC})});
          openAdmin();
        };"""

new_oncontext = """        btn.oncontextmenu = async (e) => {
          e.preventDefault();
          const currIsOn = !!btn.dataset.on;
          if(!currIsOn) return toast('User must be assigned to project first','wa');
          await apiFetch('/api/users', {method:'POST', body:JSON.stringify({action:'assign',username:u.username,project_id:p.id, is_dc: !isDC})});
          openAdmin(); // Reload to refresh isDC constant state accurately
        };"""
code = code.replace(old_oncontext, new_oncontext)

with open(r'D:\DCR\html_render.py', 'w', encoding='utf-8') as f:
    f.write(code)

print("SUCCESS")
