with open(r'D:\DCR\html_render.py', 'r', encoding='utf-8') as f:
    code = f.read()

old_str = """async function openAdmin(){{
  const [users,projects]=await Promise.all([apiFetch('/api/users'),apiFetch('/api/projects')]);
  if(!users||!projects)return;
  const body=document.getElementById('admin-body');body.innerHTML='';
  const utitle=document.createElement('div');utitle.className='stitle';utitle.textContent='👥 Users';body.appendChild(utitle);
  for(const u of users){{"""

new_str = """async function openAdmin(){{
  const [users,projects]=await Promise.all([apiFetch('/api/users'),apiFetch('/api/projects')]);
  if(!users||!projects)return;
  const assignments=await Promise.all(users.map(u=>u.role!=='superadmin'?apiFetch('/api/users/'+u.username+'/projects').catch(()=>[]):Promise.resolve([])));
  const body=document.getElementById('admin-body');body.innerHTML='';
  const utitle=document.createElement('div');utitle.className='stitle';utitle.textContent='👥 Users';body.appendChild(utitle);
  let _i = 0;
  for(const u of users){{
    const assigned_cached = assignments[_i++];"""

if old_str in code:
    code = code.replace(old_str, new_str)
    with open(r'D:\DCR\html_render.py', 'w', encoding='utf-8') as f:
        f.write(code)
    print("SUCCESS")
else:
    print("FAILED")
