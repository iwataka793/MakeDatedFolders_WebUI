const $ = (id)=>document.getElementById(id);
const state = { mode: 'Range', lastItems: [] };
let _autoTimer = null;

function iso(d){ return d.toISOString().slice(0,10); }

function setMode(mode){
  state.mode = mode;
  $('modeRange').classList.toggle('active', mode==='Range');
  $('modeDays').classList.toggle('active', mode==='Days');
  $('endWrap').classList.toggle('hidden', mode==='Days');
  $('daysWrap').classList.toggle('hidden', mode!=='Days');
  scheduleAutoPreview();
}

function setStatus(t){
  const el = $('status');
  el.textContent = t;
  el.classList.toggle('error', String(t||'').startsWith('error:'));
}

function setHealth(t){ $('health').textContent = t; }
function setCfgPath(t){ $('cfgPath').textContent = t || '---'; }

function updateSummary(sum){
  if(!sum){ $('sumDays').textContent='days: -'; $('sumTotal').textContent='total: -'; $('sumCreate').textContent='create: -'; $('sumSkip').textContent='skip: -'; return; }
  $('sumDays').textContent = `days: ${sum.days ?? '-'}`;
  $('sumTotal').textContent = `total: ${sum.total}`;
  $('sumCreate').textContent = `create: ${sum.create}`;
  $('sumSkip').textContent = `skip: ${sum.skip}`;
}

function badge(action){
  if(!action){ return ''; }
  const cls = action==='Create' ? 'create' : 'skip';
  return `<span class="badge ${cls}">${action}</span>`;
}

function render(items){
  state.lastItems = items || [];
  const tb = $('resultTable').querySelector('tbody');
  tb.innerHTML = '';
  for(const it of state.lastItems){
    const tr = document.createElement('tr');
    const d = it.Date ? String(it.Date).slice(0,10) : '';
    const action = it.Action || (String(it.Result || '').startsWith('Created') ? 'Create' : (String(it.Result || '').startsWith('Skipped') ? 'Skip' : ''));
    const kind = it.Kind || (String(it.Result || '').includes('Year') ? 'Year' : (String(it.Result || '').includes('Month') ? 'Month' : 'Day'));
    const fullPath = it.FullPath ?? '';
    const openLink = fullPath ? `<a href="#" class="openlink" data-path="${encodeURIComponent(fullPath)}">開く</a>` : '';
    tr.innerHTML = `
      <td class="kind">${kind || ''}</td>
      <td>${d}</td>
      <td>${it.FolderName ?? ''}</td>
      <td>${badge(action)}</td>
      <td class="path">${fullPath} ${openLink}</td>
    `;
    tb.appendChild(tr);
  }
}

function payload(){
  const basePath = $('basePath').value.trim();
  const startDate = $('startDate').value;
  const endDate = $('endDate').value;
  const daysToMake = parseInt($('daysToMake').value,10);
  const foldersPerDay = parseInt($('foldersPerDay').value,10);
  const firstDayStartIndex = parseInt($('firstDayStartIndex').value,10);
  return { basePath, mode: state.mode, startDate, endDate, daysToMake, foldersPerDay, firstDayStartIndex };
}

function canAutoPreview(){
  const p = payload();
  if(!p.basePath) return false;
  if(!p.startDate) return false;
  if(p.mode==='Range' && !p.endDate) return false;
  if(p.mode==='Days' && (!Number.isFinite(p.daysToMake) || p.daysToMake<1)) return false;
  if(!Number.isFinite(p.foldersPerDay) || p.foldersPerDay<1) return false;
  if(!Number.isFinite(p.firstDayStartIndex) || p.firstDayStartIndex<1) return false;
  return true;
}

function scheduleAutoPreview(){
  clearTimeout(_autoTimer);
  // BasePath が入ったら「作成先を開く」を有効化
  $('btnOpenBase').disabled = !($('basePath').value || '').trim();
  _autoTimer = setTimeout(()=>{
    if(canAutoPreview()) onPreview();
  }, 450);
}

async function api(path, body){
  const res = await fetch(path, { method:'POST', headers:{'Content-Type':'application/json; charset=utf-8'}, body: JSON.stringify(body) });
  const j = await res.json().catch(()=>({ok:false, errors:['invalid json response']}));
  if(!res.ok || !j.ok){ throw new Error((j.errors && j.errors[0]) || `HTTP ${res.status}`); }
  return j;
}

async function loadConfig(){
  try{
    const res = await fetch('/api/config');
    const j = await res.json();
    if(j.ok){
      setCfgPath(j.configPath);
      if(j.config && j.config.DefaultBasePath && !$('basePath').value){ $('basePath').value = j.config.DefaultBasePath; }
      setStatus('config loaded');
      scheduleAutoPreview();
    } else {
      setStatus('config load failed');
    }
  } catch(e){ setStatus('config load error'); }
}

async function health(){
  try{
    const res = await fetch('/api/health');
    const j = await res.json();
    if(j.ok){ setHealth('OK'); }
    else { setHealth('NG'); }
  } catch(e){ setHealth('NG'); }
}

async function onPreview(){
  setStatus('preview...');
  try{
    const j = await api('/api/preview', payload());
    updateSummary(j.summary);
    render(j.items);
    setStatus('preview ready');
  } catch(e){
    updateSummary(null);
    render([]);
    setStatus('error: ' + e.message);
  }
}

async function onRun(){
  if(!confirm('フォルダを作成します。よろしいですか？')) return;
  setStatus('running...');
  try{
    const j = await api('/api/run', payload());
    const created = j.created ?? 0;
    const skipped = j.skipped ?? 0;
    setStatus(`done (created: ${created}, skipped: ${skipped})`);
    updateSummary({ days: '-', total: created + skipped, create: created, skip: skipped });
    render(j.items);
    // 作成先フォルダへ移動（任意）
    $('btnOpenBase').disabled = !($('basePath').value || '').trim();
  } catch(e){
    setStatus('error: ' + e.message);
  }
}

async function onOpenBase(){
  const basePath = ($('basePath').value || '').trim();
  if(!basePath){ setStatus('error: base path is empty'); return; }
  try{
    await api('/api/openFolder', { path: basePath });
    setStatus('opened');
  } catch(e){
    setStatus('error: ' + e.message);
  }
}

async function onBrowse(){
  setStatus('select folder...');
  try{
    const res = await fetch('/api/pickFolder', {
      method:'POST',
      headers:{'Content-Type':'application/json; charset=utf-8'},
      body: JSON.stringify({ initialPath: $('basePath').value.trim() })
    });
    const j = await res.json().catch(()=>({ok:false, errors:['invalid json response']}));
    if(!res.ok || !j.ok){
      if(j.canceled){ setStatus('canceled'); return; }
      throw new Error((j.errors && j.errors[0]) || `HTTP ${res.status}`);
    }
    if(j.path){ $('basePath').value = j.path; }
    setStatus('folder selected');
    scheduleAutoPreview();
  } catch(e){
    setStatus('error: ' + e.message);
  }
}

async function onSaveBase(){
  const basePath = $('basePath').value.trim();
  if(!basePath){
    setStatus('error: base path is empty');
    return;
  }
  setStatus('saving base path...');
  try{
    const j = await api('/api/config/basePath', { basePath });
    if(j.configPath){ setCfgPath(j.configPath); }
    setStatus('base path saved');
  } catch(e){
    setStatus('error: ' + e.message);
  }
}

async function keepAlive(){
  try{
    await fetch('/api/ping', { method:'POST', headers:{'Content-Type':'application/json; charset=utf-8'}, body:'{}', keepalive:true });
  } catch(e){}
}

let _closeSent = false;
function closeServer(){
  if(_closeSent) return;
  _closeSent = true;
  try{
    const blob = new Blob(['{}'], { type: 'application/json; charset=utf-8' });
    if(navigator.sendBeacon){
      navigator.sendBeacon('/api/close', blob);
    } else {
      fetch('/api/close', { method:'POST', headers:{'Content-Type':'application/json; charset=utf-8'}, body:'{}', keepalive:true }).catch(()=>{});
    }
  } catch(e){}
}

function wireDatePicker(inputId, buttonId){
  const input = $(inputId);
  const btn = $(buttonId);
  if(!input || !btn){ return; }
  const openPicker = ()=>{
    if(typeof input.showPicker === 'function'){
      input.showPicker();
    } else {
      input.focus();
      input.click();
    }
  };
  btn.addEventListener('click', (e)=>{
    e.preventDefault();
    openPicker();
  });
}

function init(){
  const now = new Date();
  $('startDate').value = iso(now);
  const end = new Date(now.getTime()); end.setDate(end.getDate()+2);
  $('endDate').value = iso(end);
  setMode('Range');

  // 初期状態: BasePath が空なら「作成先を開く」は無効
  $('btnOpenBase').disabled = !($('basePath').value || '').trim();

  $('modeRange').addEventListener('click', ()=>setMode('Range'));
  $('modeDays').addEventListener('click', ()=>setMode('Days'));
  $('btnPreview').addEventListener('click', onPreview);
  $('btnRun').addEventListener('click', onRun);
  $('btnBrowse').addEventListener('click', onBrowse);
  $('btnSaveBase').addEventListener('click', onSaveBase);
  $('btnOpenBase').addEventListener('click', onOpenBase);
  $('btnReload').addEventListener('click', ()=>location.reload());

  // 結果テーブルの「開く」リンク
  $('resultTable').addEventListener('click', (e)=>{
    const a = e.target && e.target.closest ? e.target.closest('a.openlink') : null;
    if(!a) return;
    e.preventDefault();
    const p = decodeURIComponent(a.getAttribute('data-path') || '');
    if(!p) return;
    api('/api/openFolder', { path: p }).catch(err=>setStatus('error: ' + err.message));
  });

  // 入力変更で自動プレビュー（フォルダが選ばれている時だけ）
  for(const id of ['basePath','startDate','endDate','daysToMake','foldersPerDay','firstDayStartIndex']){
    const el = $(id);
    el.addEventListener('input', scheduleAutoPreview);
    el.addEventListener('change', scheduleAutoPreview);
  }

  wireDatePicker('startDate', 'startDateBtn');
  wireDatePicker('endDate', 'endDateBtn');

  loadConfig();
  health();
  keepAlive();
  setInterval(health, 5000);
  setInterval(keepAlive, 5000);

  // ウィンドウを閉じたらサーバも終了（コンソールを自動で閉じるため）
  window.addEventListener('pagehide', closeServer);
  window.addEventListener('beforeunload', closeServer);
}

init();
