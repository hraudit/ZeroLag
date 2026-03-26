// ============================================================
//  ZeroLag — Portal App Logic
//  Talks to Apps Script API via fetch (JSONP for cross-origin)
// ============================================================

const API_URL = 'https://script.google.com/a/macros/cars24.com/s/AKfycbwxP4rm4NEFpazXIrWyj8FQ6AtuQ-so7dzRvPEtZf5vjIW3Is6C5y3qtwRXtrKtVyz4/exec';

// ── API CALLER ───────────────────────────────────────────────
function api(action, params) {
  return new Promise((resolve, reject) => {
    const cb = '_zl_' + Date.now();
    const qs = new URLSearchParams({action, cb, ...(params||{})}).toString();
    const url = API_URL + '?' + qs;
    window[cb] = (data) => {
      delete window[cb];
      document.head.removeChild(script);
      resolve(data);
    };
    const script = document.createElement('script');
    script.src = url;
    script.onerror = () => { delete window[cb]; reject(new Error('Network error')); };
    document.head.appendChild(script);
  });
}

function apiPost(action, body) {
  return fetch(API_URL, {
    method: 'POST',
    headers: {'Content-Type':'application/json'},
    body: JSON.stringify({action, ...body}),
  }).then(r => r.json());
}

// ── STATE ────────────────────────────────────────────────────
let USER = null;
let FY = 'FY2025-26';
let CUR_PAGE = 'dashboard';
const TITLES = {
  dashboard:'Dashboard', reports:'Reports', approvals:'Approval Emails',
  upload:'Document Upload', rules:'Sync Rules', synclog:'Sync Log',
  observations:'Observations', auditlog:'Audit Log', users:'User Management',
  settings:'Settings'
};

// ── LOGIN ────────────────────────────────────────────────────
async function zlSignIn(e) {
  e.preventDefault();
  const email = document.getElementById('em').value.trim().toLowerCase();
  const btn = document.getElementById('btn');
  const bt = document.getElementById('bt');
  const sp = document.getElementById('sp');
  const ar = document.getElementById('ar');
  const err = document.getElementById('err');
  const errt = document.getElementById('errt');

  if (!/^[^\s@]+@cars24\.com$/i.test(email)) {
    err.classList.add('on');
    errt.textContent = 'Only @cars24.com addresses allowed';
    return;
  }

  btn.disabled = true;
  sp.style.display = 'block';
  bt.textContent = 'Checking…';
  ar.style.display = 'none';
  err.classList.remove('on');

  try {
    const res = await api('auth', {email});
    sp.style.display = 'none';
    ar.style.display = 'block';

    if (!res.ok) {
      btn.disabled = false;
      bt.textContent = 'Sign in';
      err.classList.add('on');
      errt.textContent = res.error === 'not_whitelisted'
        ? 'Access denied. Contact People Ops.'
        : res.error === 'deactivated'
        ? 'Account deactivated. Contact admin.'
        : 'Error: ' + res.error;
      return;
    }

    USER = res.user;
    bt.textContent = '✓ Access granted';
    btn.style.background = 'var(--green)';

    // Update sidebar
    const initials = USER.name.split(' ').map(w=>w[0]).join('').slice(0,2).toUpperCase();
    el('uav').textContent = initials;
    el('uname').textContent = USER.name;
    el('urole').textContent = USER.role + ' · P&C';

    // Apply permissions
    applyPerms();

    setTimeout(() => {
      const ls = el('login-screen');
      ls.style.transition = 'opacity 0.35s';
      ls.style.opacity = '0';
      setTimeout(() => { ls.style.display = 'none'; showFY(); }, 350);
    }, 500);

  } catch(ex) {
    sp.style.display = 'none';
    ar.style.display = 'block';
    btn.disabled = false;
    bt.textContent = 'Sign in';
    err.classList.add('on');
    errt.textContent = 'Connection error. Try again.';
  }
}

function zlCheck(input) {
  const v = input.value.trim();
  const ok = /^[^\s@]+@cars24\.com$/i.test(v);
  el('btn').disabled = !ok;
  const err = el('err');
  const hint = el('hint');
  if (v && /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(v) && !ok) {
    err.classList.add('on');
    el('errt').textContent = 'Only @cars24.com addresses allowed';
    if (hint) hint.style.visibility = 'hidden';
  } else {
    err.classList.remove('on');
    if (hint) hint.style.visibility = 'visible';
  }
}

function applyPerms() {
  if (!USER) return;
  const hide = (id) => { const e = el(id); if(e) e.style.display = 'none'; };
  if (!toBool(USER.can_config))        { hide('nb-rules'); hide('nb-synclog'); }
  if (!toBool(USER.can_manage_users))  hide('nb-users');
  if (!toBool(USER.can_config))        hide('nb-settings');
}

// ── FY SELECTOR ──────────────────────────────────────────────
let selFYVal = 'FY2025-26';
function selFY(el_,fy) {
  document.querySelectorAll('.fy-opt').forEach(o=>o.classList.remove('sel'));
  el_.classList.add('sel');
  selFYVal = fy;
}
function showFY() {
  const fyo = el('fy-overlay');
  if(fyo) fyo.style.display = 'flex';
}
function confirmFY() {
  FY = selFYVal;
  const d = FY.replace('FY','FY ');
  // Show portal shell FIRST before calling nav
  const shell = el('portal-shell');
  if(shell) shell.style.display = '';
  const fyo = el('fy-overlay');
  if(fyo) fyo.style.display = 'none';
  const fyd = el('fy-display');
  if(fyd) fyd.textContent = d;
  // Small delay to ensure DOM is visible before nav
  setTimeout(function() {
    nav('dashboard', el('nb-dashboard'));
  }, 50);
}

// ── NAVIGATION ───────────────────────────────────────────────
function nav(id, navEl) {
  document.querySelectorAll('.page').forEach(p=>p.classList.remove('on'));
  const pg = el('pg-'+id); if(pg) pg.classList.add('on');
  document.querySelectorAll('.nb').forEach(n=>n.classList.remove('on'));
  if (navEl) navEl.classList.add('on');
  CUR_PAGE = id;
  const tbt = el('tb-title'); if(tbt) tbt.innerHTML = (TITLES[id]||id) + ' <span id="tb-sub" style="font-size:12px;color:var(--t3);font-weight:400;margin-left:6px">' + FY + '</span>';
  const tbs = el('tb-sub'); if(tbs) tbs.textContent = FY;

  if (id==='dashboard')   loadDash();
  if (id==='reports')     loadReports();
  if (id==='approvals')   loadApprovals();
  if (id==='rules')       loadRules();
  if (id==='synclog')     loadSyncLog();
  if (id==='observations')loadObs();
  if (id==='auditlog')    loadAuditLog();
  if (id==='users')       loadUsers();
  if (id==='settings')    loadSettings();
}

// ── DASHBOARD ────────────────────────────────────────────────
async function loadDash() {
  el('dash-content').innerHTML = loading();
  try {
    const [files, approvals, obs, logs] = await Promise.all([
      api('files', {fy:FY}),
      api('approvals', {fy:FY}),
      api('obs', {fy:FY}),
      api('log', {limit:'5'}),
    ]);
    const openObs = (obs.observations||[]).filter(o=>o.status!=='resolved').length;
    el('dash-content').innerHTML = `
      <div class="greeting">Good day, ${USER.name} 👋</div>
      <div class="greeting-sub">ZeroLag · ${FY}</div>
      <div class="stats-grid">
        ${statCard((files.files||[]).length, 'Total files', '#4736FE', '📂', 'reports')}
        ${statCard((approvals.approvals||[]).length, 'Approvals captured', '#00875F', '✉️', 'approvals')}
        ${statCard(openObs, 'Open observations', '#C07000', '🔍', 'observations')}
        ${statCard((files.files||[]).filter(f=>f.source==='manual').length, 'Manual uploads', '#7C3AED', '📎', 'upload')}
      </div>
      <div class="dash-grid">
        <div>
          <div class="card">
            <div class="card-head"><div class="card-title">Recent activity</div></div>
            <div class="card-body" id="dash-activity">
              ${(logs.events||[]).slice(0,5).map(e=>`
                <div class="act-item">
                  <div class="act-icon" style="background:${actBg(e.type)}">${actIcon(e.type)}</div>
                  <div style="flex:1;min-width:0">
                    <div class="act-title">${e.action}</div>
                    <div class="act-meta">${e.user_name||e.user_email||'System'} · ${fmtDate(e.ts)}</div>
                  </div>
                </div>`).join('') || '<div class="empty-sm">No activity yet</div>'}
            </div>
          </div>
        </div>
        <div>
          <div class="card">
            <div class="card-head"><div class="card-title">System health</div></div>
            <div class="card-body">
              <div class="hrow"><div class="hdot g"></div><span>Portal</span><span class="hval">hraudit.github.io/ZeroLag</span></div>
              <div class="hrow"><div class="hdot g"></div><span>Database</span><span class="hval">Google Sheets</span></div>
              <div class="hrow"><div class="hdot g"></div><span>Storage</span><span class="hval">Google Drive</span></div>
              <div class="hrow"><div class="hdot g"></div><span>Active rules</span><span class="hval">Checking…</span></div>
            </div>
          </div>
        </div>
      </div>`;
  } catch(e) {
    el('dash-content').innerHTML = errBox(e.message);
  }
}

// ── REPORTS ──────────────────────────────────────────────────
let RDATA = [], rCat = 'all', rFilters = {sub:'',src:'',q:'',sort:'date-desc'};

async function loadReports() {
  el('r-sections').innerHTML = loading();
  try {
    const res = await api('files', {fy:FY});
    RDATA = res.files || [];
    updateRTabs();
    renderReports();
  } catch(e) {
    el('r-sections').innerHTML = errBox(e.message);
  }
}

function updateRTabs() {
  const cats = {'01_Darwinbox':0,'02_Medical':0,'03_Compliance':0,'04_Approvals':0,'05_Manual':0};
  RDATA.forEach(f => { if(cats[f.category]!==undefined) cats[f.category]++; });
  const tabMap = {'01_Darwinbox':'rs-01','02_Medical':'rs-02','03_Compliance':'rs-03','04_Approvals':'rs-04','05_Manual':'rs-05'};
  const allEl = el('rs-all'); if(allEl) allEl.querySelector('.r-stat-n').textContent = RDATA.length;
  Object.entries(tabMap).forEach(([cat,id])=>{
    const e = el(id); if(e) e.querySelector('.r-stat-n').textContent = cats[cat]||0;
  });
  const ct = el('r-count-txt'); if(ct) ct.textContent = RDATA.length + ' files · ' + FY;
}

function renderReports() {
  const filtered = RDATA.filter(f => {
    if (rCat !== 'all' && f.category !== rCat) return false;
    if (rFilters.sub && f.sub_folder !== rFilters.sub) return false;
    if (rFilters.src && f.source !== rFilters.src) return false;
    if (rFilters.q && !f.name.toLowerCase().includes(rFilters.q) && !f.doc_name.toLowerCase().includes(rFilters.q)) return false;
    return true;
  }).sort((a,b) => {
    const [col,dir] = (rFilters.sort||'date-desc').split('-');
    const av = col==='date' ? new Date(a.uploaded_at).getTime() : a.name.toLowerCase();
    const bv = col==='date' ? new Date(b.uploaded_at).getTime() : b.name.toLowerCase();
    const r = av < bv ? -1 : av > bv ? 1 : 0;
    return dir==='asc' ? r : -r;
  });

  el('r-result-txt').textContent = filtered.length + ' file' + (filtered.length!==1?'s':'');
  
  if (!filtered.length) {
    el('r-sections').innerHTML = `<div class="empty-state"><div class="empty-icon">📂</div><div class="empty-title">No files match your filters</div><div class="empty-sub">Try clearing filters or upload a file</div></div>`;
    return;
  }

  const RCAT = {
    '01_Darwinbox':{l:'Darwinbox',i:'📊',c:'#4736FE',bg:'rgba(71,54,254,0.08)'},
    '02_Medical':{l:'Medical',i:'🏥',c:'#0078A0',bg:'rgba(0,120,160,0.08)'},
    '03_Compliance':{l:'Compliance',i:'🔏',c:'#C07000',bg:'rgba(192,112,0,0.08)'},
    '04_Approvals':{l:'Approvals',i:'✉️',c:'#00875F',bg:'rgba(0,135,95,0.08)'},
    '05_Manual':{l:'Documents',i:'📁',c:'#7C3AED',bg:'rgba(124,58,237,0.08)'},
  };
  const EXTI = {csv:'📋',xlsx:'📗',xls:'📗',pdf:'📕',png:'🖼',jpg:'🖼',jpeg:'🖼',docx:'📘',doc:'📘'};

  const groups = {};
  filtered.forEach(f => {
    const c = f.category||'Other';
    if(!groups[c]) groups[c]=[];
    groups[c].push(f);
  });

  el('r-sections').innerHTML = Object.entries(groups).map(([cat,files]) => {
    const cfg = RCAT[cat]||{l:cat,i:'📄',c:'#666',bg:'#f5f5f5'};
    return `<div class="r-section">
      <div class="r-sec-head" onclick="this.parentElement.classList.toggle('open')">
        <div class="r-sec-icon" style="background:${cfg.bg}">${cfg.i}</div>
        <div class="r-sec-name">${cfg.l}</div>
        <span class="r-sec-ct">${files.length} file${files.length!==1?'s':''}</span>
        <svg class="r-chev" width="13" height="13" viewBox="0 0 13 13" fill="none"><path d="M2.5 4.5l4 4 4-4" stroke="#9491B8" stroke-width="1.4" stroke-linecap="round"/></svg>
      </div>
      <div class="r-sec-body">
        ${files.map(f => {
          const ext = (f.name||'').split('.').pop().toLowerCase();
          return `<div class="r-file">
            <div class="r-file-icon" style="background:${cfg.bg}">${EXTI[ext]||'📄'}</div>
            <div class="r-file-info">
              <div class="r-file-name">${f.doc_name||f.name}</div>
              <div class="r-file-meta">${f.size||'—'} · ${f.sub_folder.replace(/_/g,' ')} · ${fmtDate(f.uploaded_at)}</div>
            </div>
            <span class="tag ${f.source==='manual'?'tag-purple':'tag-green'}">${f.source==='manual'?'manual':'auto'}</span>
            ${f.drive_url ? `<a href="${f.drive_url}" target="_blank" class="btn-sm btn-primary" onclick="logAction('download','${f.name}')">↗ Open</a>` : ''}
          </div>`;
        }).join('')}
      </div>
    </div>`;
  }).join('');

  // Auto-expand first section
  const first = el('r-sections').querySelector('.r-section');
  if (first) first.classList.add('open');
}

function setRCat(cat, el_) {
  rCat = cat;
  document.querySelectorAll('.r-stat').forEach(s=>s.classList.remove('on'));
  if(el_) el_.classList.add('on');
  renderReports();
}
function applyRF() {
  rFilters = {
    sub: val('rf-sub'), src: val('rf-source'),
    q: (val('rf-search')||'').toLowerCase(), sort: val('rf-sort')||'date-desc',
  };
  renderReports();
}
function clearRF() {
  ['rf-sub','rf-source','rf-search'].forEach(id=>{const e=el(id);if(e)e.value='';});
  const rs = el('rf-sort'); if(rs) rs.value='date-desc';
  rFilters = {sub:'',src:'',q:'',sort:'date-desc'};
  renderReports();
}

// ── APPROVALS ────────────────────────────────────────────────
async function loadApprovals() {
  el('ap-list').innerHTML = loading();
  try {
    const res = await api('approvals', {fy:FY});
    const items = res.approvals || [];
    const badge = el('nb-ap-ct');
    if (badge && items.length) { badge.textContent = items.length; badge.style.display=''; }
    
    if (!items.length) {
      el('ap-list').innerHTML = `<div class="empty-state"><div class="empty-icon">✉️</div><div class="empty-title">No approvals captured yet</div><div class="empty-sub">Approvals from Gaurav & Sahil where hraudit@cars24.com is in CC are auto-captured</div></div>`;
      return;
    }
    const COLORS = ['#4736FE','#00875F','#C07000','#0078A0','#7C3AED'];
    el('ap-list').innerHTML = items.map((a,i)=>{
      const col = COLORS[i%COLORS.length];
      const init = (a.from_name||a.from_email||'?').split(' ').map(w=>w[0]).join('').slice(0,2).toUpperCase();
      return `<div class="ap-card" onclick="openApDrawer(${i})">
        <div class="ap-card-inner">
          <div class="ap-av" style="background:${col}">${init}</div>
          <div class="ap-info">
            <div class="ap-subj">${a.subject}</div>
            <div class="ap-meta">${a.from_name||a.from_email} · ${fmtDate(a.received_at)}</div>
          </div>
          <div class="ap-tags">
            <span class="tag tag-blue">${a.type||'General'}</span>
            ${a.attachment_count>0?`<span class="tag tag-grey">📎 ${a.attachment_count}</span>`:''}
          </div>
        </div>
        <div class="ap-card-foot">
          ${a.email_file_id?`<a href="https://drive.google.com/file/d/${a.email_file_id}/view" target="_blank" class="btn-sm btn-red" onclick="event.stopPropagation()">📕 Email file</a>`:''}
          ${(a.attachment_ids||'').split(',').filter(Boolean).map(id=>`<a href="https://drive.google.com/file/d/${id}/view" target="_blank" class="btn-sm" onclick="event.stopPropagation()">📎 Attachment</a>`).join('')}
        </div>
      </div>`;
    }).join('');
    window._AP_DATA = items;
  } catch(e) {
    el('ap-list').innerHTML = errBox(e.message);
  }
}

function openApDrawer(i) {
  const a = window._AP_DATA[i];
  if (!a) return;
  el('dr-title').textContent = a.subject;
  el('dr-meta').textContent = (a.from_name||a.from_email) + ' · ' + fmtDate(a.received_at);
  el('dr-body').innerHTML = `
    <div class="info-grid">
      ${infoCell('From', a.from_name||a.from_email)}
      ${infoCell('Email', a.from_email)}
      ${infoCell('Date', fmtDate(a.received_at))}
      ${infoCell('Type', a.type||'General')}
      ${infoCell('Attachments', (a.attachment_count||0) + ' files')}
      ${infoCell('FY', a.fy)}
    </div>
    <div class="section-label">Actions</div>
    <div style="display:flex;flex-wrap:wrap;gap:8px">
      ${a.email_file_id?`<a href="https://drive.google.com/file/d/${a.email_file_id}/view" target="_blank" class="btn btn-red">📕 View email file</a>`:''}
      ${(a.attachment_ids||'').split(',').filter(Boolean).map((id,i)=>`<a href="https://drive.google.com/file/d/${id}/view" target="_blank" class="btn">📎 Attachment ${i+1}</a>`).join('')}
    </div>`;
  el('dr-foot').innerHTML = `<button class="btn" onclick="closeDrawer()">Close</button>`;
  openDrawer();
}

// ── UPLOAD ───────────────────────────────────────────────────
let upFiles = [];

function handleFiles(files) {
  const allowed = ['pdf','png','jpg','jpeg','xlsx','xls','csv','doc','docx'];
  Array.from(files).forEach(f => {
    const ext = f.name.split('.').pop().toLowerCase();
    if (!allowed.includes(ext)) return;
    if (upFiles.find(x=>x.name===f.name)) return;
    upFiles.push(f);
  });
  renderFileList();
  updatePathPreview();
}
function dzOver(e) { e.preventDefault(); el('dz').classList.add('over'); }
function dzLeave() { el('dz').classList.remove('over'); }
function dzDrop(e) { e.preventDefault(); el('dz').classList.remove('over'); handleFiles(e.dataTransfer.files); }
function removeFile(i) { upFiles.splice(i,1); renderFileList(); }

function renderFileList() {
  const ICONS = {pdf:'📕',xlsx:'📗',xls:'📗',csv:'📋',png:'🖼',jpg:'🖼',jpeg:'🖼',docx:'📘',doc:'📘'};
  el('file-list').innerHTML = upFiles.map((f,i)=>{
    const ext = f.name.split('.').pop().toLowerCase();
    const sz = f.size>1048576?(f.size/1048576).toFixed(1)+' MB':Math.round(f.size/1024)+' KB';
    return `<div class="file-item">
      <span style="font-size:18px">${ICONS[ext]||'📄'}</span>
      <span class="file-item-name">${f.name}</span>
      <span class="file-item-size">${sz}</span>
      <button onclick="removeFile(${i})" class="file-item-rm">×</button>
    </div>`;
  }).join('');
}

const SUBCATS = {
  '01_Darwinbox':['Employee_Master','Leave_Reports','Attendance_Reports','Payroll_Reports','Separation_Reports','Onboarding_Reports','Perf_Reports','CompBen_Reports','Asset_Reports','Loan_Advances'],
  '02_Medical':['GMC_Claims','OPD_Reports','Premium_Paid','Policy_Docs'],
  '03_Compliance':['PF_ECR','ESIC','PT_Returns','Labour_Law'],
  '04_Approvals':['Approval_Records'],
  '05_Manual':['Policies','SOPs','Adhoc_Uploads'],
};

function updateSubCat() {
  const cat = val('up-cat');
  const sub = el('up-sub');
  if (!cat || !SUBCATS[cat]) { sub.innerHTML='<option value="">Select category first</option>'; sub.disabled=true; updatePathPreview(); return; }
  sub.disabled = false;
  sub.innerHTML = '<option value="">Select sub-category</option>' + SUBCATS[cat].map(s=>`<option value="${s}">${s.replace(/_/g,' ')}</option>`).join('');
  updatePathPreview();
}

function updatePathPreview() {
  const cat = val('up-cat'), sub = val('up-sub'), fy = val('up-fy')||FY;
  const catNames = {'01_Darwinbox':'Darwinbox','02_Medical':'Medical','03_Compliance':'Compliance','04_Approvals':'Approvals','05_Manual':'Documents'};
  el('pp-fy').textContent = fy;
  const pc = el('pp-cat'); if(pc){pc.textContent=cat?catNames[cat]+' /':'← category'; pc.style.color=cat?'var(--t1)':'var(--t4)';}
  const ps = el('pp-sub'); if(ps){ps.textContent=sub?sub.replace(/_/g,' ')+' /':'← sub-category'; ps.style.color=sub?'var(--t1)':'var(--t4)';}
  const pf = el('pp-file');
  if(pf){
    if(upFiles.length){const ts=new Date().toISOString().replace(/[-T:.Z]/g,'').slice(0,15);pf.textContent=upFiles[0].name.replace(/(\.[^.]+)$/,'_'+ts+'$1');pf.style.color='var(--brand)';}
    else{pf.textContent='filename_timestamp.ext';pf.style.color='var(--t4)';}
  }
}

async function submitUpload() {
  const name = val('up-name'), cat = val('up-cat'), sub = val('up-sub'), fy = val('up-fy')||FY;
  const st = el('up-status');
  if (!upFiles.length) { st.textContent='⚠ Select at least one file'; st.style.color='var(--amber)'; return; }
  if (!cat) { st.textContent='⚠ Select a category'; st.style.color='var(--amber)'; return; }
  if (!sub) { st.textContent='⚠ Select a sub-category'; st.style.color='var(--amber)'; return; }
  if (!name) { st.textContent='⚠ Enter a document name'; st.style.color='var(--amber)'; return; }

  const btn = el('up-submit');
  btn.disabled = true;
  btn.innerHTML = spinner() + ' Uploading…';
  st.textContent = 'Reading file…';
  st.style.color = 'var(--t3)';

  try {
    for (let i = 0; i < upFiles.length; i++) {
      const file = upFiles[i];
      st.textContent = `Uploading ${i+1}/${upFiles.length}: ${file.name}`;
      
      const b64 = await toBase64(file);
      const res = await apiPost('upload_file', {
        name: file.name, doc_name: name, category: cat, sub_folder: sub,
        fy, description: val('up-desc'), tags: val('up-tags'),
        base64: b64.split(',')[1], mime_type: file.type,
        email: USER ? USER.email : '', source: 'manual',
      });
      
      if (!res.ok) throw new Error(res.error || 'Upload failed');
    }
    
    st.textContent = `✓ ${upFiles.length} file${upFiles.length>1?'s':''} uploaded successfully`;
    st.style.color = 'var(--green)';
    btn.disabled = false;
    btn.innerHTML = '↑ Upload to Drive';
    clearFiles();
    resetUpload();
  } catch(e) {
    st.textContent = '✗ Error: ' + e.message;
    st.style.color = 'var(--red)';
    btn.disabled = false;
    btn.innerHTML = '↑ Upload to Drive';
  }
}

function toBase64(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = () => reject(new Error('File read failed'));
    reader.readAsDataURL(file);
  });
}

function resetUpload() {
  upFiles = [];
  ['up-name','up-desc','up-tags'].forEach(id=>{const e=el(id);if(e)e.value='';});
  const uc = el('up-cat'); if(uc) uc.value='';
  const us = el('up-sub'); if(us){us.innerHTML='<option value="">Select category first</option>';us.disabled=true;}
  renderFileList();
  updatePathPreview();
}
function clearFiles() { upFiles=[]; renderFileList(); updatePathPreview(); }

// ── OBSERVATIONS ─────────────────────────────────────────────
async function loadObs() {
  const content = el('obs-list');
  content.innerHTML = loading();
  try {
    const res = await api('obs', {fy:FY});
    const items = res.observations || [];
    const open = items.filter(o=>o.status!=='resolved').length;
    const badge = el('nb-obs-ct');
    if(badge){badge.textContent=open;badge.style.display=open?'':'none';}

    if (!items.length) {
      content.innerHTML = `<div class="empty-state"><div class="empty-icon">🔍</div><div class="empty-title">No observations yet</div><div class="empty-sub">Raise an observation using the button above</div></div>`;
      return;
    }
    const SC = {critical:'#FF4444',major:'#F5A623',minor:'#4736FE'};
    content.innerHTML = items.map((o,i)=>`
      <div class="obs-card" onclick="openObsDr(${i})">
        <div class="obs-sev-bar" style="background:${SC[o.severity]||'#999'}"></div>
        <div class="obs-body">
          <div class="obs-title">${o.title}</div>
          <div class="obs-meta">${o.raised_by} · ${fmtDate(o.raised_at)}</div>
          <div class="obs-tags">
            <span class="tag ${o.severity==='critical'?'tag-red':o.severity==='major'?'tag-amber':'tag-blue'}">${o.severity}</span>
            <span class="tag ${o.status==='resolved'?'tag-green':o.status==='inprogress'?'tag-amber':'tag-red'}">${o.status}</span>
            <span class="tag tag-grey">${o.category}</span>
          </div>
        </div>
      </div>`).join('');
    window._OBS_DATA = items;
  } catch(e) { content.innerHTML = errBox(e.message); }
}

function openObsDr(i) {
  const o = window._OBS_DATA[i];
  if (!o) return;
  el('dr-title').textContent = o.title;
  el('dr-meta').textContent = 'Raised by ' + o.raised_by + ' · ' + fmtDate(o.raised_at);
  el('dr-body').innerHTML = `
    <div class="info-grid">
      ${infoCell('Severity', `<span class="tag ${o.severity==='critical'?'tag-red':o.severity==='major'?'tag-amber':'tag-blue'}">${o.severity}</span>`)}
      ${infoCell('Status', `<span class="tag ${o.status==='resolved'?'tag-green':o.status==='inprogress'?'tag-amber':'tag-red'}">${o.status}</span>`)}
      ${infoCell('Category', o.category)}
      ${infoCell('FY', o.fy)}
    </div>
    <div class="section-label">Description</div>
    <div class="desc-box">${o.description}</div>
    ${o.status!=='resolved'?`
      <div class="section-label">Reply / Resolve</div>
      <textarea class="fi" id="obs-reply" rows="3" placeholder="Write a reply or resolution note…" style="margin-bottom:8px"></textarea>
      <select class="fi" id="obs-status">
        <option value="">Keep current status</option>
        <option value="inprogress">Mark In Progress</option>
        <option value="resolved">Mark Resolved ✓</option>
      </select>`:''}`;
  el('dr-foot').innerHTML = o.status!=='resolved'
    ? `<button class="btn btn-primary" onclick="submitObsReply('${o.obs_id}')">Send reply</button><button class="btn" onclick="closeDrawer()" style="margin-left:auto">Close</button>`
    : `<button class="btn" onclick="closeDrawer()">Close</button>`;
  openDrawer();
}

async function openObsModal() {
  el('modal-title').textContent = 'Raise observation';
  el('modal-body').innerHTML = `
    <div class="fg"><label class="fl">Title *</label><input class="fi" id="om-title" placeholder="Short description of the issue"></div>
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px">
      <div class="fg"><label class="fl">Severity *</label>
        <select class="fi" id="om-sev"><option value="">Select</option><option value="critical">🔴 Critical</option><option value="major">🟡 Major</option><option value="minor">🔵 Minor</option></select>
      </div>
      <div class="fg"><label class="fl">Category *</label>
        <select class="fi" id="om-cat"><option value="">Select</option><option>Darwinbox</option><option>Medical</option><option>Compliance</option><option>Documents</option><option>Approvals</option><option>Other</option></select>
      </div>
    </div>
    <div class="fg"><label class="fl">Description *</label><textarea class="fi" id="om-desc" rows="4" placeholder="What is missing or incorrect…"></textarea></div>
    <div class="fg"><label class="fl">Linked file (optional)</label><input class="fi" id="om-file" placeholder="e.g. Leave policy, March attendance"></div>`;
  el('modal-foot').innerHTML = `<button class="btn" onclick="closeModal()">Cancel</button><button class="btn btn-primary" onclick="saveObs()">Raise observation</button>`;
  openModal();
}

async function saveObs() {
  const title=val('om-title'),sev=val('om-sev'),cat=val('om-cat'),desc=val('om-desc');
  if(!title||!sev||!cat||!desc){alert('Fill all required fields');return;}
  try {
    const res = await apiPost('add_obs', {title,severity:sev,category:cat,description:desc,fy:FY,
      raised_by:USER.name,email:USER.email,linked_file:val('om-file')});
    if(!res.ok) throw new Error(res.error);
    closeModal(); loadObs();
  } catch(e) { alert('Error: '+e.message); }
}

async function submitObsReply(obsId) {
  const msg=val('obs-reply'), status=val('obs-status');
  if(!msg){alert('Write a reply first');return;}
  try {
    await apiPost('resolve_obs', {obs_id:obsId,note:msg,status,email:USER.email});
    closeDrawer(); loadObs();
  } catch(e) { alert('Error: '+e.message); }
}

// ── AUDIT LOG ────────────────────────────────────────────────
async function loadAuditLog() {
  el('al-body').innerHTML = loading();
  try {
    const res = await api('log', {limit:'200'});
    const items = res.events || [];
    if (!items.length) { el('al-body').innerHTML = `<div class="empty-state"><div class="empty-icon">🗂️</div><div class="empty-title">No events yet</div></div>`; return; }
    const ICONS = {login:'🔐',download:'⬇️',upload:'📎',sync:'🔄',observation:'🔍',settings:'⚙️',user:'👥',approval:'✉️'};
    const byDate = {};
    items.forEach(e => {
      const d = new Date(e.ts).toLocaleDateString('en-IN',{day:'2-digit',month:'short',year:'numeric'});
      if(!byDate[d]) byDate[d]=[];
      byDate[d].push(e);
    });
    el('al-body').innerHTML = Object.entries(byDate).map(([date,events])=>`
      <div class="al-date-header">${date}</div>
      ${events.map(e=>`
        <div class="al-row">
          <div class="al-icon">${ICONS[e.type]||'📋'}</div>
          <div class="al-info">
            <div class="al-action">${e.action}</div>
            <div class="al-detail">${e.detail}</div>
          </div>
          <div class="al-right">
            <div class="al-user">${e.user_name||e.user_email||'System'}</div>
            <div class="al-time">${new Date(e.ts).toLocaleTimeString('en-IN',{hour:'2-digit',minute:'2-digit'})}</div>
          </div>
        </div>`).join('')}`).join('');
  } catch(e) { el('al-body').innerHTML = errBox(e.message); }
}

// ── SYNC RULES ───────────────────────────────────────────────
let RULES_DATA = [];
async function loadRules() {
  el('rules-body').innerHTML = loading();
  try {
    const res = await api('rules');
    RULES_DATA = res.rules||[];
    const badge = el('nb-rules-ct');
    if(badge){const active=RULES_DATA.filter(r=>r.is_active==='TRUE').length;badge.textContent=active;badge.style.display=active?'':'none';}
    if(!RULES_DATA.length){
      el('rules-body').innerHTML=`<div class="empty-state"><div class="empty-icon">⚙️</div><div class="empty-title">No sync rules yet</div><div class="empty-sub">Add rules to auto-capture Darwinbox emails</div></div>`;
      return;
    }
    el('rules-body').innerHTML = RULES_DATA.map((r,i)=>`
      <div class="rule-card">
        <div class="rule-head">
          <span class="rule-id">${r.rule_id}</span>
          <div style="flex:1;min-width:0"><div class="rule-name">${r.rule_name}</div><div class="rule-kw">keyword: ${r.subject_exact}</div></div>
          <span class="tag ${r.is_active==='TRUE'?'tag-green':'tag-amber'}">${r.is_active==='TRUE'?'active':'paused'}</span>
          <label class="toggle"><input type="checkbox" ${r.is_active==='TRUE'?'checked':''} onchange="toggleRule('${r.rule_id}',this.checked)"><span class="toggle-sl"></span></label>
        </div>
        <div class="rule-foot">
          <span class="rule-path">${r.category.replace(/^\d+_/,'')} / ${r.sub_folder.replace(/_/g,' ')}</span>
          <button class="btn-sm" onclick="openRuleModal(${i})">✏ Edit</button>
        </div>
      </div>`).join('');
  } catch(e) { el('rules-body').innerHTML = errBox(e.message); }
}

async function toggleRule(ruleId, active) {
  try { await apiPost('toggle_rule',{rule_id:ruleId,active}); loadRules(); } catch(e){alert('Error: '+e.message);}
}

function openRuleModal(editIdx) {
  const r = editIdx !== undefined ? RULES_DATA[editIdx] : null;
  el('modal-title').textContent = r ? 'Edit rule' : 'Add sync rule';
  el('modal-body').innerHTML = `
    <div class="fg"><label class="fl">Rule name *</label><input class="fi" id="rm-name" value="${r?r.rule_name:''}" placeholder="e.g. Leave Reports"></div>
    <div class="fg"><label class="fl">Email subject keyword *</label><input class="fi" id="rm-kw" value="${r?r.subject_exact:''}" placeholder="e.g. Leave_Report_Monthly"></div>
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px">
      <div class="fg"><label class="fl">Category *</label>
        <select class="fi" id="rm-cat" onchange="updateRMSub()">
          <option value="">Select</option>
          <option value="01_Darwinbox" ${r&&r.category==='01_Darwinbox'?'selected':''}>Darwinbox</option>
          <option value="02_Medical" ${r&&r.category==='02_Medical'?'selected':''}>Medical</option>
          <option value="03_Compliance" ${r&&r.category==='03_Compliance'?'selected':''}>Compliance</option>
          <option value="05_Manual" ${r&&r.category==='05_Manual'?'selected':''}>Documents</option>
        </select>
      </div>
      <div class="fg"><label class="fl">Sub-category *</label>
        <select class="fi" id="rm-sub"><option value="${r?r.sub_folder:''}">${r?r.sub_folder.replace(/_/g,' '):'Select category first'}</option></select>
      </div>
    </div>
    <div class="fg"><label class="fl">Notes</label><input class="fi" id="rm-notes" value="${r?r.notes||'':''}" placeholder="Optional notes"></div>`;
  el('modal-foot').innerHTML = `<button class="btn" onclick="closeModal()">Cancel</button><button class="btn btn-primary" onclick="saveRule('${r?r.rule_id:''}')">${r?'Save changes':'Add rule'}</button>`;
  if (r) setTimeout(()=>{ el('rm-cat').value=r.category; updateRMSub(r.sub_folder); },50);
  openModal();
}

function updateRMSub(sel) {
  const cat = val('rm-cat');
  const sub = el('rm-sub');
  if (!cat||!SUBCATS[cat]) { sub.innerHTML='<option>Select category first</option>'; sub.disabled=true; return; }
  sub.disabled = false;
  sub.innerHTML = '<option value="">Select</option>' + SUBCATS[cat].map(s=>`<option value="${s}" ${sel===s?'selected':''}>${s.replace(/_/g,' ')}</option>`).join('');
}

async function saveRule(editId) {
  const name=val('rm-name'),kw=val('rm-kw'),cat=val('rm-cat'),sub=val('rm-sub');
  if(!name||!kw||!cat||!sub){alert('Fill all required fields');return;}
  try {
    if(editId) { await apiPost('edit_rule',{rule_id:editId,rule_name:name,subject_exact:kw,category:cat,sub_folder:sub,notes:val('rm-notes')}); }
    else { await apiPost('add_rule',{rule_name:name,subject_exact:kw,category:cat,sub_folder:sub,notes:val('rm-notes')}); }
    closeModal(); loadRules();
  } catch(e) { alert('Error: '+e.message); }
}

// ── SYNC LOG ─────────────────────────────────────────────────
async function loadSyncLog() {
  el('sl-body').innerHTML = loading();
  try {
    const res = await api('synclog');
    const items = res.logs||[];
    if(!items.length){ el('sl-body').innerHTML=`<div class="empty-state"><div class="empty-icon">📋</div><div class="empty-title">No sync log entries yet</div></div>`; return; }
    el('sl-body').innerHTML = `<table class="tbl"><thead><tr><th>Time</th><th>Rule</th><th>File</th><th>Size</th><th>Status</th></tr></thead><tbody>
    ${items.map(l=>`<tr>
      <td class="tbl-mono">${fmtDate(l.synced_at)}</td>
      <td><span class="rule-id">${l.rule_id||'—'}</span></td>
      <td class="tbl-name">${l.file_name||'—'}</td>
      <td class="tbl-mono">${l.size||'—'}</td>
      <td><span class="tag ${l.status==='error'?'tag-red':l.status==='skipped'?'tag-amber':'tag-green'}">${l.status||'—'}</span></td>
    </tr>`).join('')}
    </tbody></table>`;
  } catch(e) { el('sl-body').innerHTML = errBox(e.message); }
}

// ── USERS ────────────────────────────────────────────────────
let USERS_DATA = [];
const PERMS = [
  {key:'can_view',label:'View reports',sub:'Browse and read files'},
  {key:'can_download',label:'Download files',sub:'Download from Drive'},
  {key:'can_upload',label:'Upload documents',sub:'Upload via portal'},
  {key:'can_config',label:'Manage rules & logs',sub:'Add/edit sync rules'},
  {key:'can_manage_users',label:'Manage users',sub:'Add/edit/deactivate users'},
  {key:'can_observe',label:'Raise observations',sub:'Create audit observations'},
  {key:'can_resolve',label:'Resolve observations',sub:'Reply and close observations'},
];

async function loadUsers() {
  el('users-body').innerHTML = loading();
  try {
    const res = await api('users');
    USERS_DATA = res.users||[];
    if(!USERS_DATA.length){ el('users-body').innerHTML=`<div class="empty-state"><div class="empty-icon">👥</div><div class="empty-title">No users yet</div></div>`; return; }
    el('users-body').innerHTML = `<div class="card" style="overflow:hidden"><table class="tbl"><thead><tr><th>User</th><th>Email</th><th>Role</th><th>Last login</th><th>Status</th><th></th></tr></thead><tbody>
    ${USERS_DATA.map((u,i)=>{
      const active = u.is_active==='TRUE'||u.is_active===true;
      const init = (u.name||'').split(' ').map(w=>w[0]).join('').slice(0,2).toUpperCase();
      return `<tr>
        <td><div style="display:flex;align-items:center;gap:9px"><div class="user-av-sm">${init}</div><span class="tbl-bold">${u.name}</span></div></td>
        <td class="tbl-mono">${u.email}</td>
        <td><span class="tag tag-blue">${u.role}</span></td>
        <td class="tbl-mono">${u.last_login?fmtDate(u.last_login):'Never'}</td>
        <td><span class="tag ${active?'tag-green':'tag-amber'}">${active?'active':'inactive'}</span></td>
        <td><button class="btn-sm" onclick="openUserModal(${i})">✏ Edit</button></td>
      </tr>`;
    }).join('')}
    </tbody></table></div>`;
  } catch(e) { el('users-body').innerHTML = errBox(e.message); }
}

function openUserModal(editIdx) {
  const u = editIdx !== undefined ? USERS_DATA[editIdx] : null;
  el('modal-title').textContent = u ? 'Edit — '+u.name : 'Add user';
  el('modal-body').innerHTML = `
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:16px">
      <div class="fg"><label class="fl">Full name *</label><input class="fi" id="um-name" value="${u?u.name:''}" placeholder="e.g. Ravi Mehta" ${u&&u.email==='vivek.naiwal@cars24.com'?'readonly':''}></div>
      <div class="fg"><label class="fl">Work email *</label><input class="fi" id="um-email" type="email" value="${u?u.email:''}" placeholder="name@cars24.com" ${u?'readonly':''}></div>
      <div class="fg"><label class="fl">Role</label>
        <select class="fi" id="um-role">
          <option value="Admin" ${u&&u.role==='Admin'?'selected':''}>Admin</option>
          <option value="Sub-admin" ${u&&u.role==='Sub-admin'?'selected':''}>Sub-admin</option>
          <option value="Auditor" ${u&&u.role==='Auditor'?'selected':''}>Auditor</option>
          <option value="Viewer" ${u&&u.role==='Viewer'?'selected':''}>Viewer</option>
        </select>
      </div>
      <div class="fg"><label class="fl">Status</label>
        <select class="fi" id="um-status">
          <option value="TRUE" ${(!u||u.is_active==='TRUE')?'selected':''}>Active</option>
          <option value="FALSE" ${u&&u.is_active==='FALSE'?'selected':''}>Inactive</option>
        </select>
      </div>
    </div>
    <div style="font-size:11px;font-weight:700;color:var(--t3);text-transform:uppercase;letter-spacing:0.08em;margin-bottom:9px">Permissions</div>
    <div class="perms-grid">
      ${PERMS.map(p=>{
        const checked = u ? (u[p.key]==='TRUE'||u[p.key]===true) : false;
        const disabled = u && u.email==='vivek.naiwal@cars24.com';
        return `<div class="perm-row">
          <div><div class="perm-label">${p.label}</div><div class="perm-sub">${p.sub}</div></div>
          <label class="toggle"><input type="checkbox" id="perm-${p.key}" ${checked?'checked':''} ${disabled?'checked disabled':''}><span class="toggle-sl"></span></label>
        </div>`;
      }).join('')}
    </div>
    <div style="display:flex;gap:6px;margin-top:10px">
      <button class="btn-sm" onclick="PERMS.forEach(p=>{const e=document.getElementById('perm-'+p.key);if(e&&!e.disabled)e.checked=true})">Select all</button>
      <button class="btn-sm" onclick="PERMS.forEach(p=>{const e=document.getElementById('perm-'+p.key);if(e&&!e.disabled)e.checked=false})">Clear all</button>
    </div>`;
  el('modal-foot').innerHTML = `<button class="btn" onclick="closeModal()">Cancel</button><button class="btn btn-primary" onclick="saveUser('${u?u.email:''}')">${u?'Save changes':'Add user'}</button>`;
  openModal();
}

async function saveUser(editEmail) {
  const name=val('um-name'),email=val('um-email');
  if(!name||!email){alert('Name and email required');return;}
  const body = {name,email,role:val('um-role'),is_active:val('um-status'),added_by:USER?USER.email:''};
  PERMS.forEach(p=>{const e=el('perm-'+p.key);body[p.key]=e?e.checked?'TRUE':'FALSE':'FALSE';});
  try {
    if(editEmail){await apiPost('edit_user',{...body,edited_by:USER?USER.email:''});}
    else{await apiPost('add_user',body);}
    closeModal(); loadUsers();
  } catch(e){alert('Error: '+e.message);}
}

// ── SETTINGS ─────────────────────────────────────────────────
async function loadSettings() {
  try {
    const res = await api('config');
    const cfg = res.config||{};
    const setOn = cfg.cookie_set_on;
    let days = 0;
    if(setOn) days = Math.max(0, 30 - Math.floor((Date.now()-new Date(setOn))/86400000));
    const pct = Math.min(100,(days/30)*100);
    const color = days>10?'var(--green)':days>5?'var(--amber)':'var(--red)';
    
    const fillEl = el('set-cookie-fill');
    const stEl = el('set-cookie-status');
    if(fillEl){fillEl.style.width=pct+'%';fillEl.style.background=color;}
    if(stEl){stEl.textContent=days>10?'✓ Cookie healthy — '+days+' days left':'⚠ Expires in '+days+' days';stEl.style.color=color;}
  } catch(e){}
}

async function saveCookie() {
  const v = val('set-cookie-input');
  if(!v){alert('Paste the cookie value first');return;}
  try {
    await apiPost('save_cookie',{value:v,email:USER?USER.email:''});
    alert('✓ Cookie saved! Syncs will resume.');
    el('set-cookie-input').value='';
    loadSettings();
  } catch(e){alert('Error: '+e.message);}
}

// ── MODAL / DRAWER ───────────────────────────────────────────
function openModal(){el('modal-overlay').classList.add('open');el('modal').classList.add('open');}
function closeModal(){el('modal-overlay').classList.remove('open');el('modal').classList.remove('open');}
function openDrawer(){el('drawer-overlay').classList.add('open');el('drawer').classList.add('open');}
function closeDrawer(){el('drawer-overlay').classList.remove('open');el('drawer').classList.remove('open');}

// ── HELPERS ──────────────────────────────────────────────────
const el = id => document.getElementById(id);
const val = id => { const e=el(id); return e?e.value:''; };
const toBool = v => v==='TRUE'||v===true;
const loading = () => '<div class="loading-state"><div class="spin"></div><span>Loading…</span></div>';
const spinner = () => '<span class="spin-inline"></span>';
const errBox = msg => `<div class="err-box"><strong>⚠ Error</strong><br><span class="mono">${msg}</span><br><small>Check browser console (F12) for details</small></div>`;
const infoCell = (l,v) => `<div><div class="info-label">${l}</div><div class="info-val">${v}</div></div>`;
const statCard = (n,l,c,icon,page) => `<div class="stat-card" onclick="nav('${page}',el('nb-${page}'))"><div class="stat-icon">${icon}</div><div class="stat-n" style="color:${c}">${n}</div><div class="stat-l">${l}</div></div>`;
const fmtDate = ts => { if(!ts) return '—'; try{return new Date(ts).toLocaleString('en-IN',{day:'2-digit',month:'short',year:'numeric',hour:'2-digit',minute:'2-digit'});}catch{return ts;}};
const actIcon = t => ({login:'🔐',download:'⬇️',upload:'📎',sync:'🔄',observation:'🔍',settings:'⚙️',user:'👥',approval:'✉️'})[t]||'📋';
const actBg = t => ({login:'rgba(0,135,95,0.1)',sync:'rgba(71,54,254,0.08)',upload:'rgba(71,54,254,0.08)',download:'rgba(0,120,160,0.1)',observation:'rgba(192,112,0,0.1)',user:'rgba(71,54,254,0.08)'})[t]||'var(--bg3)';

async function logAction(type, detail) {
  if(!USER) return;
  try { await apiPost('audit_log',{type,action:type==='download'?'File downloaded':type,detail,email:USER.email,name:USER.name}); } catch(e){}
}

function copyWebAppURL() {
  navigator.clipboard.writeText(API_URL).then(()=>alert('API URL copied!'));
}
function toggleCookieVis() {
  const i=el('set-cookie-input');const b=el('cookie-eye');
  i.type=i.type==='password'?'text':'password';
  b.textContent=i.type==='password'?'Show':'Hide';
}

// ── INIT ─────────────────────────────────────────────────────
window.addEventListener('DOMContentLoaded', () => {
  // Auto-fill detection
  setTimeout(()=>{ const em=el('em'); if(em&&em.value) zlCheck(em); },400);
  setTimeout(()=>{ const em=el('em'); if(em&&em.value) zlCheck(em); },1000);
});
