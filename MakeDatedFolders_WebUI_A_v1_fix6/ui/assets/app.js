const $ = (id)=>document.getElementById(id);
const state = {
  mode: 'Range',
  lastItems: [],
  previewSeq: 0,
  previewController: null,
  isPreviewing: false,
  isDirty: false
};
let _closeRequested = false;
let _reloadRequested = false;
let _shutdownRequested = false;
const _skipCloseKey = 'skipCloseOnce';
const PREVIEW_TIMEOUT_MS = 30000;

function iso(d){ return d.toISOString().slice(0,10); }

function setMode(mode){
  state.mode = mode;
  $('modeRange').classList.toggle('active', mode==='Range');
  $('modeDays').classList.toggle('active', mode==='Days');
  $('endWrap').classList.toggle('hidden', mode==='Days');
  $('daysWrap').classList.toggle('hidden', mode!=='Days');
  markDirty('mode');
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
    tr.innerHTML = `
      <td class="kind">${kind || ''}</td>
      <td>${d}</td>
      <td>${it.FolderName ?? ''}</td>
      <td>${badge(action)}</td>
      <td class="path">${fullPath}</td>
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

function validatePayload(){
  const errors = [];
  const p = payload();
  if(!p.basePath) errors.push('作成先パスを入力してください');
  if(!p.startDate) errors.push('開始日を入力してください');
  if(p.mode==='Range' && !p.endDate) errors.push('終了日を入力してください');
  if(p.mode==='Days' && (!Number.isFinite(p.daysToMake) || p.daysToMake<1)) errors.push('日数は1以上にしてください');
  if(!Number.isFinite(p.foldersPerDay) || p.foldersPerDay<1) errors.push('最大番号は1以上にしてください');
  if(!Number.isFinite(p.firstDayStartIndex) || p.firstDayStartIndex<1) errors.push('初日の開始番号は1以上にしてください');
  return errors;
}

function setPreviewBusy(busy){
  state.isPreviewing = busy;
  $('btnPreview').disabled = busy;
  $('btnPreview').classList.toggle('ghost', busy);
}

function markDirty(){
  state.isDirty = true;
  if(!state.isPreviewing){
    setStatus('変更されました。プレビューを押してください。');
  }
}

async function api(path, body, opts = {}){
  const res = await fetch(path, { method:'POST', headers:{'Content-Type':'application/json; charset=utf-8'}, body: JSON.stringify(body), signal: opts.signal });
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
      if(j.config && j.config.DefaultBasePath){ markDirty(); }
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
  const errors = validatePayload();
  if(errors.length){
    setStatus('error: ' + errors[0]);
    return;
  }
  const seq = ++state.previewSeq;
  if(state.previewController){ state.previewController.abort(); }
  const controller = new AbortController();
  state.previewController = controller;
  setPreviewBusy(true);
  setStatus('プレビュー中...');
  let timedOut = false;
  const timeoutId = setTimeout(()=>{
    timedOut = true;
    controller.abort();
  }, PREVIEW_TIMEOUT_MS);
  try{
    const j = await api('/api/preview', payload(), { signal: controller.signal });
    if(seq !== state.previewSeq){ return; }
    updateSummary(j.summary);
    render(j.items);
    state.isDirty = false;
    setStatus('プレビュー完了');
  } catch(e){
    if(seq !== state.previewSeq){ return; }
    if(e.name === 'AbortError'){
      setStatus(timedOut ? 'プレビューがタイムアウトしました。' : 'プレビューをキャンセルしました。');
      return;
    }
    updateSummary(null);
    render([]);
    setStatus('error: ' + e.message);
  } finally {
    clearTimeout(timeoutId);
    if(seq === state.previewSeq){
      setPreviewBusy(false);
    }
  }
}

async function onRun(){
  setStatus('running...');
  try{
    const j = await api('/api/run', payload());
    const created = j.created ?? 0;
    const skipped = j.skipped ?? 0;
    setStatus(`done (created: ${created}, skipped: ${skipped})`);
    updateSummary({ days: '-', total: created + skipped, create: created, skip: skipped });
    render(j.items);
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
    markDirty();
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

function shouldSkipClose(){
  try{
    if (sessionStorage.getItem(_skipCloseKey)){
      sessionStorage.removeItem(_skipCloseKey);
      return true;
    }
  } catch(e){}
  return false;
}

function requestClose(){
  if (_closeRequested || _reloadRequested || shouldSkipClose()) { return; }
  _closeRequested = true;
  try{
    const payload = JSON.stringify({ ts: new Date().toISOString() });
    if (navigator.sendBeacon){
      const blob = new Blob([payload], { type: 'application/json' });
      navigator.sendBeacon('/api/close', blob);
    } else {
      fetch('/api/close', {
        method:'POST',
        headers:{'Content-Type':'application/json; charset=utf-8'},
        body: payload,
        keepalive: true
      });
    }
  } catch(e){}
}

async function requestShutdown(){
  if (_shutdownRequested) { return; }
  _shutdownRequested = true;
  _closeRequested = true;
  try{
    const payload = JSON.stringify({ ts: new Date().toISOString() });
    await fetch('/api/shutdown', {
      method:'POST',
      headers:{'Content-Type':'application/json; charset=utf-8'},
      body: payload,
      keepalive: true
    });
  } catch(e){}
  try { window.close(); } catch(e){}
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

  $('modeRange').addEventListener('click', ()=>setMode('Range'));
  $('modeDays').addEventListener('click', ()=>setMode('Days'));
  $('btnPreview').addEventListener('click', onPreview);
  $('btnRun').addEventListener('click', onRun);
  $('btnBrowse').addEventListener('click', onBrowse);
  $('btnSaveBase').addEventListener('click', onSaveBase);
  $('btnReload').addEventListener('click', ()=>{
    _reloadRequested = true;
    try { sessionStorage.setItem(_skipCloseKey, '1'); } catch(e){}
    location.reload();
  });
  $('btnShutdown').addEventListener('click', requestShutdown);

  // 入力変更はプレビュー未実行の案内のみ
  for(const id of ['basePath','startDate','endDate','daysToMake','foldersPerDay','firstDayStartIndex']){
    const el = $(id);
    el.addEventListener('input', markDirty);
    el.addEventListener('change', markDirty);
  }

  wireDatePicker('startDate', 'startDateBtn');
  wireDatePicker('endDate', 'endDateBtn');

  loadConfig();
  health();
  keepAlive();
  setInterval(health, 5000);
  setInterval(keepAlive, 5000);
  window.addEventListener('beforeunload', requestClose);
  window.addEventListener('pagehide', requestClose);

}

init();
