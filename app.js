'use strict';
const VERSION = '1.3.0';
const CUTOFF_DAY = 25;
const HEB_MONTHS = ['ינואר','פברואר','מרץ','אפריל','מאי','יוני','יולי','אוגוסט','ספטמבר','אוקטובר','נובמבר','דצמבר'];

const ASHAL_NAME = 'א.ש.ל';
const SETTINGS_VERSION = 2;
// categories = expense columns of the accounting Excel form (replaced from the uploaded file itself)
const FORM_CATEGORIES = ['ארוחות עובדים', 'ארוחות עסקיות', 'דלק', 'ציוד רכב', 'ציוד סלולארי', 'נסיעות וחניונים', 'משרדיות', 'כיבוד', 'שונות'];
const DEFAULT_SETTINGS = {
  v: SETTINGS_VERSION,
  name: '', role: '', company: '', manager: '', acctEmail: '', acctName: 'הנהלת חשבונות', notes: '',
  categories: [...FORM_CATEGORIES.map(name => ({name, ashal: false})), {name: ASHAL_NAME, ashal: true}],
  // A.SH.L policy: up to 9h ₪55, up to 12h ₪100, above 12h ₪150 (reported in ProperTime, not in the Excel form)
  ashalTiers: [{upTo: 9, amount: 55}, {upTo: 12, amount: 100}, {upTo: null, amount: 150}],
  sentReports: {} // periodKey -> {sentAt, total, count}
};

/* ---------------- IndexedDB ---------------- */
const DB = {
  db: null,
  open() {
    return new Promise((res, rej) => {
      const r = indexedDB.open('expenses-app', 1);
      r.onupgradeneeded = () => {
        const d = r.result;
        d.createObjectStore('expenses', {keyPath: 'id'});
        d.createObjectStore('files', {keyPath: 'id'});
        d.createObjectStore('kv', {keyPath: 'k'});
      };
      r.onsuccess = () => { this.db = r.result; res(); };
      r.onerror = () => rej(r.error);
    });
  },
  tx(store, mode, fn) {
    return new Promise((res, rej) => {
      const t = this.db.transaction(store, mode);
      const s = t.objectStore(store);
      let out;
      const req = fn(s);
      if (req) req.onsuccess = () => { out = req.result; };
      t.oncomplete = () => res(out);
      t.onerror = () => rej(t.error);
      t.onabort = () => rej(t.error);
    });
  },
  all(store) { return this.tx(store, 'readonly', s => s.getAll()); },
  get(store, key) { return this.tx(store, 'readonly', s => s.get(key)); },
  put(store, val) { return this.tx(store, 'readwrite', s => s.put(val)); },
  del(store, key) { return this.tx(store, 'readwrite', s => s.delete(key)); },
  clear(store) { return this.tx(store, 'readwrite', s => s.clear()); }
};

/* ---------------- helpers ---------------- */
const $ = id => document.getElementById(id);
const uid = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
const money = n => '₪' + (Number(n) || 0).toLocaleString('he-IL', {minimumFractionDigits: 2, maximumFractionDigits: 2});
const num2 = n => (Number(n) || 0).toLocaleString('he-IL', {minimumFractionDigits: 2, maximumFractionDigits: 2});
const pad = n => String(n).padStart(2, '0');
const todayISO = () => { const d = new Date(); return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`; };
const fmtDate = iso => { if (!iso) return ''; const [y, m, d] = iso.split('-'); return `${d}/${m}/${y}`; };
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({'&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'}[c]));

function periodKeyFor(iso) {
  let [y, m, d] = iso.split('-').map(Number);
  if (d > CUTOFF_DAY) { m++; if (m > 12) { m = 1; y++; } }
  return `${y}-${pad(m)}`;
}
function periodRange(key) {
  const [y, m] = key.split('-').map(Number);
  let py = y, pm = m - 1; if (pm < 1) { pm = 12; py--; }
  return {start: `${py}-${pad(pm)}-${pad(CUTOFF_DAY + 1)}`, end: `${y}-${pad(m)}-${pad(CUTOFF_DAY)}`};
}
const periodLabel = key => { const [y, m] = key.split('-').map(Number); return `${HEB_MONTHS[m - 1]} ${y}`; };
const currentPeriod = () => periodKeyFor(todayISO());

function toast(msg, ms = 2600) {
  const t = $('toast'); t.textContent = msg; t.classList.add('on');
  clearTimeout(toast._t); toast._t = setTimeout(() => t.classList.remove('on'), ms);
}
function spin(on, txt) { $('spinTxt').textContent = txt || 'מכין…'; $('spin').classList.toggle('on', on); }

const urlCache = new Map();
async function fileURL(id) {
  if (urlCache.has(id)) return urlCache.get(id);
  const f = await DB.get('files', id); if (!f) return '';
  const u = URL.createObjectURL(f.blob); urlCache.set(id, u); return u;
}
function downloadBlob(blob, name) {
  const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name;
  document.body.appendChild(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 4000);
}

/* ---------------- state ---------------- */
let S = null;      // settings
let EXP = [];      // expenses
async function loadSettings() {
  const r = await DB.get('kv', 'settings');
  S = Object.assign(structuredClone(DEFAULT_SETTINGS), r ? r.v : {});
  if ((S.v || 1) < SETTINGS_VERSION) { // v1 had free categories + hour tiers; move to the Excel form's columns and the policy rates
    const d = structuredClone(DEFAULT_SETTINGS);
    S.categories = d.categories; S.ashalTiers = d.ashalTiers; S.v = SETTINGS_VERSION;
    delete S.empId; delete S.dept; delete S.ashalMode; delete S.ashalRate;
    await saveSettingsObj();
  }
}
async function saveSettingsObj() { await DB.put('kv', {k: 'settings', v: S}); }
async function loadExpenses() { EXP = (await DB.all('expenses')).sort((a, b) => b.date.localeCompare(a.date) || b.createdAt - a.createdAt); }
const catByName = n => S.categories.find(c => c.name === n);
const isAshal = e => !!(catByName(e.category)?.ashal || e.ashal);

function ashalAmount(hours) {
  hours = Number(hours) || 0; if (!(hours > 0)) return 0;
  const tiers = [...S.ashalTiers].sort((a, b) => (a.upTo ?? Infinity) - (b.upTo ?? Infinity));
  const t = tiers.find(t => hours <= (t.upTo ?? Infinity)) || tiers[tiers.length - 1];
  return t ? Number(t.amount) || 0 : 0;
}

/* ---------------- navigation ---------------- */
function show(v) {
  document.querySelectorAll('.view').forEach(x => x.classList.toggle('on', x.id === 'view-' + v));
  document.querySelectorAll('nav button').forEach(b => b.classList.toggle('on', b.dataset.v === v));
  $('fab').classList.toggle('hidden', v !== 'expenses');
  if (v === 'report') renderReport();
  if (v === 'settings') renderSettings();
  if (v === 'expenses') renderExpenses();
  window.scrollTo(0, 0);
}
async function goTab(v) {
  if ($('view-settings').classList.contains('on')) { readSettingsForm(); await saveSettingsObj(); }
  const onHome = $('view-expenses').classList.contains('on');
  if (v === 'expenses') { if (!closeLayer('tab')) show('expenses'); return; }
  if (onHome) pushLayer('tab', () => show('expenses'));
  show(v);
}
document.querySelectorAll('nav button').forEach(b => b.onclick = () => goTab(b.dataset.v));

/* ---------------- phone back button: one step back, confirm before leaving ---------------- */
const LAYERS = [];
function pushLayer(id, hide) { LAYERS.push({id, hide}); history.pushState({layer: id, n: LAYERS.length}, ''); }
function closeLayer(id) { // UI-initiated close: let history drive it so back-stack stays in sync
  const top = LAYERS[LAYERS.length - 1];
  if (!top || top.id !== id) return false;
  history.back(); return true;
}
let EXIT_ARMED = false;
window.addEventListener('popstate', () => {
  const top = LAYERS.pop();
  if (top) { top.hide(); return; }
  if (EXIT_ARMED) return;
  // at the root: we are now on the base entry; another back would leave the app
  EXIT_ARMED = true; $('exitDlg').classList.add('on');
});
$('exitStay').onclick = () => { $('exitDlg').classList.remove('on'); EXIT_ARMED = false; history.pushState({guard: 1}, ''); };
$('exitGo').onclick = () => {
  $('exitDlg').classList.remove('on');
  try { window.close(); } catch {}
  setTimeout(() => toast('לחץ "חזור" פעם נוספת כדי לצאת', 3000), 150);
  setTimeout(() => { if (EXIT_ARMED) { EXIT_ARMED = false; history.pushState({guard: 1}, ''); } }, 3200);
};

/* ---------------- expenses list ---------------- */
function periodsList() {
  const set = new Set(EXP.map(e => e.period)); set.add(currentPeriod());
  return [...set].sort().reverse();
}
function fillPeriodSelect(sel, value) {
  sel.innerHTML = periodsList().map(p => {
    const sent = S.sentReports[p] ? ' ✓ נשלח' : '';
    return `<option value="${p}">${periodLabel(p)}${p === currentPeriod() ? ' (נוכחי)' : ''}${sent}</option>`;
  }).join('');
  sel.value = value && periodsList().includes(value) ? value : currentPeriod();
}

async function renderExpenses() {
  const cp = currentPeriod(), r = periodRange(cp);
  $('hdrPeriod').textContent = 'דוח ' + periodLabel(cp);
  const cur = EXP.filter(e => e.period === cp);
  $('curRange').textContent = `${fmtDate(r.start)} – ${fmtDate(r.end)}`;
  $('curCount').textContent = `${cur.length} הוצאות`;
  $('curTotal').textContent = money(cur.reduce((s, e) => s + Number(e.amount), 0));

  // reminder
  const d = new Date().getDate();
  let rem = '';
  if (S.sentReports[cp]) rem = `<div class="banner ok">✓ הדוח של ${periodLabel(cp)} נשלח. הוצאות חדשות ייכנסו לדוח הבא.</div>`;
  else if (cur.length && d >= 20 && d <= CUTOFF_DAY) rem = `<div class="banner">⏰ ${CUTOFF_DAY - d === 0 ? 'היום' : 'בעוד ' + (CUTOFF_DAY - d) + ' ימים'} ה-${CUTOFF_DAY} – זמן להגיש את דוח ${periodLabel(cp)}. <a href="#" id="goRep">לדוח ←</a></div>`;
  const prev = periodsList().filter(p => p < cp && !S.sentReports[p] && EXP.some(e => e.period === p));
  if (prev.length) rem += `<div class="banner">יש דוח שלא סומן כנשלח: ${prev.map(periodLabel).join(', ')}</div>`;
  $('remind').innerHTML = rem;
  const g = $('goRep'); if (g) g.onclick = ev => { ev.preventDefault(); goTab('report'); };

  const sel = $('listPeriod');
  fillPeriodSelect(sel, listKey); listKey = sel.value;
  const list = EXP.filter(e => e.period === sel.value);
  const box = $('expList');
  if (!list.length) { box.innerHTML = `<div class="empty">אין הוצאות בתקופה זו.<br>לחץ + כדי להוסיף.</div>`; return; }
  box.innerHTML = list.map(e => {
    const ash = isAshal(e);
    const tag = ash ? `<span class="tag">${e.hours || 0} ש'</span>` : (e.receipts?.length ? '' : '<span class="tag warn">חסרה קבלה</span>');
    return `<div class="item" data-id="${e.id}">
      <div class="thumb" data-f="${!ash && e.receipts?.[0] ? e.receipts[0] : ''}">${ash ? '🕒' : (e.receipts?.length ? '' : '🧾')}</div>
      <div class="mid"><div class="t">${esc(e.category)}${tag}${periodKeyFor(e.date) !== e.period ? '<span class="tag warn">מאוחרת</span>' : ''}</div><div class="muted">${fmtDate(e.date)} · ${esc(e.description || '')}</div></div>
      <div class="amt">${money(e.amount)}</div></div>`;
  }).join('');
  box.querySelectorAll('.item').forEach(el => el.onclick = () => openForm(EXP.find(x => x.id === el.dataset.id)));
  for (const th of box.querySelectorAll('.thumb[data-f]')) {
    const id = th.dataset.f; if (!id) continue;
    const f = await DB.get('files', id);
    if (f && f.type.startsWith('image/')) th.style.backgroundImage = `url(${await fileURL(id)})`;
    else if (f) th.textContent = 'PDF';
  }
}
let listKey = null;
$('listPeriod').onchange = () => { listKey = $('listPeriod').value; renderExpenses(); };

/* ---------------- expense form ---------------- */
let F = null; // editing state {id, receipts:[{id,blob,type,isNew}] , removed:[]}
function fillCatSelect(sel, value) {
  const names = S.categories.map(c => c.name);
  if (value && !names.includes(value)) names.push(value);
  sel.innerHTML = names.map(n => `<option>${esc(n)}</option>`).join('');
  if (value) sel.value = value;
}
async function openForm(e) {
  F = {id: e?.id || null, receipts: [], removed: []};
  $('fTitle').textContent = e ? 'עריכת הוצאה' : 'הוצאה חדשה';
  fillCatSelect($('fCat'), e?.category || S.categories[0]?.name);
  setFormDate(e?.date || todayISO(), e?.period);
  $('fAmount').value = e && !isAshal(e) ? e.amount : '';
  $('fHours').value = e?.hours || '';
  $('fDesc').value = e?.description || '';
  $('fDelete').classList.toggle('hidden', !e);
  if (e?.receipts) for (const rid of e.receipts) { const f = await DB.get('files', rid); if (f) F.receipts.push({id: rid, blob: f.blob, type: f.type, orig: f.orig || (f.type.startsWith('image/') ? f.blob : null), edit: normEdit(f.edit)}); }
  if (e && S.sentReports[e.period]) toast('שים לב: ההוצאה שייכת לדוח שכבר נשלח');
  $('fScanMode').value = S.scanMode || 'color';
  onCatChange(); renderFormReceipts();
  $('sheet').classList.add('on'); pushLayer('form', hideForm);
}
function hideForm() { $('sheet').classList.remove('on'); $('calDlg').classList.remove('on'); F = null; }
function closeForm() { if (!closeLayer('form')) hideForm(); }
$('fClose').onclick = closeForm;
$('sheet').onclick = ev => { if (ev.target.id === 'sheet') closeForm(); };
$('fab').onclick = () => openForm(null);

/* date (dd/mm/yyyy, no future) + which report the expense goes to */
function setFormDate(iso, keepPeriod) { F.date = iso; $('fDateBtn').textContent = fmtDate(iso); updPeriodChoice(keepPeriod); }
function updPeriodChoice(keep) {
  const nat = periodKeyFor(F.date), cur = currentPeriod(), wrap = $('fPeriodWrap');
  if (nat >= cur && (!keep || keep === nat)) { wrap.classList.add('hidden'); return; }
  const opts = [[cur, `דוח נוכחי – ${periodLabel(cur)} (הוצאה מאוחרת)`]];
  if (nat < cur && !S.sentReports[nat]) opts.push([nat, `דוח ${periodLabel(nat)} (לפי התאריך)`]);
  if (keep && !opts.some(o => o[0] === keep)) opts.push([keep, `דוח ${periodLabel(keep)}${S.sentReports[keep] ? ' – נשלח' : ''}`]);
  $('fPeriod').innerHTML = opts.map(o => `<option value="${o[0]}">${esc(o[1])}</option>`).join('');
  $('fPeriod').value = keep && opts.some(o => o[0] === keep) ? keep : cur;
  wrap.classList.remove('hidden');
}
const CAL = {y: 0, m: 0};
function openCal() {
  [CAL.y, CAL.m] = F.date.split('-').map(Number);
  renderCal(); $('calDlg').classList.add('on'); pushLayer('cal', () => $('calDlg').classList.remove('on'));
}
function renderCal() {
  const t = todayISO(), [ty, tm] = t.split('-').map(Number);
  $('calTitle').textContent = `${HEB_MONTHS[CAL.m - 1]} ${CAL.y}`;
  $('calNext').disabled = CAL.y > ty || (CAL.y === ty && CAL.m >= tm);
  const first = new Date(CAL.y, CAL.m - 1, 1).getDay(), days = new Date(CAL.y, CAL.m, 0).getDate();
  let h = ['א', 'ב', 'ג', 'ד', 'ה', 'ו', 'ש'].map(d => `<div class="cdh">${d}</div>`).join('');
  h += '<div></div>'.repeat(first);
  for (let d = 1; d <= days; d++) {
    const iso = `${CAL.y}-${pad(CAL.m)}-${pad(d)}`;
    h += `<button class="cd${iso === F.date ? ' sel' : ''}${iso === t ? ' today' : ''}" data-d="${iso}" ${iso > t ? 'disabled' : ''}>${d}</button>`;
  }
  $('calGrid').innerHTML = h;
  $('calGrid').querySelectorAll('.cd').forEach(b => b.onclick = () => { setFormDate(b.dataset.d); closeLayer('cal'); });
}
$('fDateBtn').onclick = openCal;
$('calPrev').onclick = () => { if (--CAL.m < 1) { CAL.m = 12; CAL.y--; } renderCal(); };
$('calNext').onclick = () => { if (++CAL.m > 12) { CAL.m = 1; CAL.y++; } renderCal(); };
$('calToday').onclick = () => { setFormDate(todayISO()); closeLayer('cal'); };
$('calYest').onclick = () => { const d = new Date(Date.now() - 864e5); setFormDate(`${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`); closeLayer('cal'); };
$('calDlg').onclick = ev => { if (ev.target.id === 'calDlg') closeLayer('cal'); };

function onCatChange() {
  const ash = !!catByName($('fCat').value)?.ashal;
  $('fAshal').classList.toggle('hidden', !ash);
  $('fAmountWrap').classList.toggle('hidden', ash);
  $('fRcptWrap').classList.toggle('hidden', ash);
  if (ash) updAshal();
  $('fDesc').placeholder = $('fCat').value === 'ארוחות עסקיות' ? 'עם מי הייתה הפגישה ובאיזה נושא / פרויקט' : 'לדוגמה: פז – נסיעה לאתר';
}
function updAshal() {
  const a = ashalAmount($('fHours').value);
  $('fAshalAmt').value = num2(a);
  $('fAshalNote').textContent = 'א.ש.ל – ללא קבלה. לא נכנס לדוח האקסל; מרוכז בלשונית "דוח" לדיווח בפרופר טיים.';
}
$('fCat').onchange = onCatChange;
$('fHours').oninput = updAshal;

$('btnCam').onclick = () => $('camInput').click();
$('btnPick').onclick = () => $('pickInput').click();
$('camInput').onchange = ev => addReceiptFiles(ev.target.files, ev.target);
$('pickInput').onchange = ev => addReceiptFiles(ev.target.files, ev.target);

async function addReceiptFiles(fileList, input) {
  const files = [...fileList];
  spin(true, 'מעבד תמונה…');
  const mode = $('fScanMode').value; S.scanMode = mode; saveSettingsObj();
  try {
    for (const f of files) {
      if (f.type === 'application/pdf') F.receipts.push({id: uid(), blob: f, type: 'application/pdf', isNew: true});
      else if (f.type.startsWith('image/') || /\.(heic|heif)$/i.test(f.name)) {
        const orig = await normalizeOrig(f);
        const edit = {rot: 0, angle: 0, crop: null, mode};
        if (mode !== 'none') { const a = ImgProc.auto(await blobCanvas(orig, 1200)); if (a) Object.assign(edit, a); }
        const blob = await renderReceipt(orig, edit);
        F.receipts.push({id: uid(), blob, orig, edit, type: 'image/jpeg', isNew: true});
      }
    }
  } catch (err) { console.error(err); toast('לא הצלחתי לעבד את הקובץ'); }
  input.value = ''; spin(false); renderFormReceipts();
  const last = F.receipts[F.receipts.length - 1];
  if (last && last.orig && last.isNew && files.length === 1) openViewer(F.receipts.length - 1); // let the user check / fix the crop right away
}

function normEdit(e) { e = e || {}; return {rot: e.rot || 0, angle: e.angle || 0, crop: e.crop || null, mode: e.mode || 'none'}; }
async function loadBitmap(blob) {
  try { return await createImageBitmap(blob, {imageOrientation: 'from-image'}); }
  catch {
    return await new Promise((res, rej) => { const i = new Image(); i.onload = () => res(i); i.onerror = rej; i.src = URL.createObjectURL(blob); });
  }
}
async function blobCanvas(blob, max = 2400) {
  const bmp = await loadBitmap(blob);
  const k = Math.min(1, max / Math.max(bmp.width, bmp.height));
  const c = ImgProc.mkCanvas(Math.round(bmp.width * k), Math.round(bmp.height * k));
  const x = c.getContext('2d'); x.imageSmoothingQuality = 'high'; x.drawImage(bmp, 0, 0, c.width, c.height);
  return c;
}
const canvasBlob = (c, type = 'image/jpeg', q = 0.85) => new Promise(res => c.toBlob(res, type, q));
async function normalizeOrig(file) { return canvasBlob(await blobCanvas(file, 2400), 'image/jpeg', 0.9); } // upright, size-capped original kept for re-editing
async function renderReceipt(orig, edit) {
  let c = await blobCanvas(orig, 2400);
  c = ImgProc.rotateAny(c, (edit.rot || 0) + (edit.angle || 0));
  c = ImgProc.crop(c, edit.crop);
  const k = Math.min(1, 1800 / Math.max(c.width, c.height));
  if (k < 1) { const s = ImgProc.mkCanvas(Math.round(c.width * k), Math.round(c.height * k)); const x = s.getContext('2d'); x.imageSmoothingQuality = 'high'; x.drawImage(c, 0, 0, s.width, s.height); c = s; }
  c = ImgProc.enhance(c, edit.mode);
  return canvasBlob(c, 'image/jpeg', 0.82);
}

function renderFormReceipts() {
  const box = $('fReceipts');
  box.innerHTML = F.receipts.map((r, i) => `<div class="rc" data-i="${i}">${r.type === 'application/pdf' ? 'PDF' : ''}
     <button class="del" data-i="${i}">✕</button></div>`).join('');
  box.querySelectorAll('.rc').forEach(el => {
    const r = F.receipts[el.dataset.i];
    if (r.type.startsWith('image/')) { el.style.backgroundImage = `url(${URL.createObjectURL(r.blob)})`; el.onclick = () => openViewer(+el.dataset.i); }
  });
  box.querySelectorAll('.del').forEach(b => b.onclick = ev => {
    ev.stopPropagation();
    const r = F.receipts.splice(b.dataset.i, 1)[0]; if (!r.isNew) F.removed.push(r.id); renderFormReceipts();
  });
}

/* ---------------- receipt viewer: pinch-zoom, filters, rotate, crop & align ---------------- */
const V = {i: -1, readonly: false, s: 1, tx: 0, ty: 0, fit: null, ptrs: new Map(), start: null, lastTap: 0};
function vRec() { return V.readonly ? V.ro : F.receipts[V.i]; }
async function openViewer(i, ro) {
  V.readonly = !!ro; V.ro = ro || null; V.i = i;
  $('viewer').classList.add('on'); pushLayer('viewer', hideViewer);
  $('vTools').classList.toggle('hidden', V.readonly);
  $('vModes').classList.toggle('hidden', V.readonly || !vRec().orig);
  $('vCropBar').classList.add('hidden'); $('vCrop').classList.add('hidden'); $('vStage').classList.remove('hidden');
  syncModes(); await vShow();
}
function hideViewer() { $('viewer').classList.remove('on'); if (!V.readonly && F) renderFormReceipts(); }
function closeViewer() { if (!closeLayer('viewer')) hideViewer(); }
function syncModes() {
  const r = vRec(); const m = r?.edit?.mode || 'none';
  document.querySelectorAll('#vModes button').forEach(b => b.classList.toggle('on', b.dataset.m === m));
  $('vCropBtn').disabled = $('vRotBtn').disabled = !r?.orig;
}
async function vShow() {
  const img = $('vImg'); const r = vRec();
  await new Promise(res => { img.onload = res; img.onerror = res; img.src = URL.createObjectURL(r.blob); });
  vFit();
}
function vFit() {
  const st = $('vStage').getBoundingClientRect(), img = $('vImg');
  const k = Math.min(st.width / img.naturalWidth, st.height / img.naturalHeight);
  const w = img.naturalWidth * k, h = img.naturalHeight * k;
  V.fit = {w, h, l: (st.width - w) / 2, t: (st.height - h) / 2};
  Object.assign(img.style, {width: w + 'px', height: h + 'px', left: V.fit.l + 'px', top: V.fit.t + 'px'});
  V.s = 1; V.tx = 0; V.ty = 0; vApply();
}
function vClamp() {
  if (V.s <= 1) { V.s = 1; V.tx = 0; V.ty = 0; return; }
  const st = $('vStage').getBoundingClientRect(), f = V.fit;
  const w = f.w * V.s, h = f.h * V.s;
  const minX = Math.min(0, st.width - w) - f.l, maxX = Math.max(0, st.width - w) - f.l;
  const minY = Math.min(0, st.height - h) - f.t, maxY = Math.max(0, st.height - h) - f.t;
  if (w <= st.width) V.tx = (st.width - w) / 2 - f.l; else V.tx = Math.min(maxX, Math.max(minX, V.tx));
  if (h <= st.height) V.ty = (st.height - h) / 2 - f.t; else V.ty = Math.min(maxY, Math.max(minY, V.ty));
}
function vApply() { $('vImg').style.transform = `translate(${V.tx}px,${V.ty}px) scale(${V.s})`; }
function vZoomAt(mx, my, s) {
  const f = V.fit, lx = (mx - f.l - V.tx) / V.s, ly = (my - f.t - V.ty) / V.s;
  V.s = Math.min(6, Math.max(1, s)); V.tx = mx - f.l - V.s * lx; V.ty = my - f.t - V.s * ly; vClamp(); vApply();
}
(function viewerGestures() {
  const st = $('vStage');
  const local = e => { const r = st.getBoundingClientRect(); return [e.clientX - r.left, e.clientY - r.top]; };
  st.addEventListener('pointerdown', e => {
    st.setPointerCapture(e.pointerId); V.ptrs.set(e.pointerId, local(e));
    const p = [...V.ptrs.values()];
    if (p.length === 2) V.start = {d: Math.hypot(p[0][0] - p[1][0], p[0][1] - p[1][1]), s: V.s, m: [(p[0][0] + p[1][0]) / 2, (p[0][1] + p[1][1]) / 2], tx: V.tx, ty: V.ty};
    else if (p.length === 1) {
      const now = Date.now();
      if (now - V.lastTap < 300) { vZoomAt(p[0][0], p[0][1], V.s > 1.2 ? 1 : 2.5); V.lastTap = 0; } else V.lastTap = now;
      V.start = {pan: p[0], tx: V.tx, ty: V.ty};
    }
  });
  st.addEventListener('pointermove', e => {
    if (!V.ptrs.has(e.pointerId)) return; V.ptrs.set(e.pointerId, local(e));
    const p = [...V.ptrs.values()];
    if (p.length >= 2 && V.start?.d) {
      const d = Math.hypot(p[0][0] - p[1][0], p[0][1] - p[1][1]), m = [(p[0][0] + p[1][0]) / 2, (p[0][1] + p[1][1]) / 2];
      const f = V.fit, s0 = V.start.s, lx = (V.start.m[0] - f.l - V.start.tx) / s0, ly = (V.start.m[1] - f.t - V.start.ty) / s0;
      V.s = Math.min(6, Math.max(1, s0 * d / V.start.d)); V.tx = m[0] - f.l - V.s * lx; V.ty = m[1] - f.t - V.s * ly; vClamp(); vApply();
    } else if (p.length === 1 && V.start?.pan && V.s > 1) {
      V.tx = V.start.tx + p[0][0] - V.start.pan[0]; V.ty = V.start.ty + p[0][1] - V.start.pan[1]; vClamp(); vApply();
    }
  });
  const up = e => { V.ptrs.delete(e.pointerId); const p = [...V.ptrs.values()]; V.start = p.length === 1 ? {pan: p[0], tx: V.tx, ty: V.ty} : null; };
  st.addEventListener('pointerup', up); st.addEventListener('pointercancel', up);
  st.addEventListener('wheel', e => { e.preventDefault(); const [x, y] = local(e); vZoomAt(x, y, V.s * (e.deltaY < 0 ? 1.2 : 1 / 1.2)); }, {passive: false});
  window.addEventListener('resize', () => { if ($('viewer').classList.contains('on') && !$('vStage').classList.contains('hidden')) vFit(); });
})();
async function vReprocess() {
  const r = vRec(); spin(true, 'מעבד…');
  try {
    r.blob = await renderReceipt(r.orig, r.edit);
    if (!r.isNew) { F.removed.push(r.id); r.id = uid(); r.isNew = true; }
    await vShow();
  } catch (err) { console.error(err); toast('העיבוד נכשל'); }
  spin(false);
}
document.querySelectorAll('#vModes button').forEach(b => b.onclick = async () => {
  const r = vRec(); if (!r.orig || r.edit.mode === b.dataset.m) return;
  r.edit.mode = b.dataset.m; syncModes(); await vReprocess();
});
$('vRotBtn').onclick = async () => { const r = vRec(); r.edit.rot = ((r.edit.rot || 0) + 90) % 360; await vReprocess(); };
$('vDelBtn').onclick = () => { const r = F.receipts.splice(V.i, 1)[0]; if (r && !r.isNew) F.removed.push(r.id); closeViewer(); };
$('vDone').onclick = closeViewer;

/* crop (right angles only) + straighten by rotating the photo */
const C = {base: null, prev: null, crop: null, rot: 0, angle: 0, scale: 1, off: [0, 0], drag: null};
async function openCrop() {
  const r = vRec(); spin(true, 'טוען…');
  C.base = await blobCanvas(r.orig, 1400); spin(false);
  C.rot = r.edit.rot || 0; C.angle = r.edit.angle || 0;
  C.crop = r.edit.crop ? {...r.edit.crop} : {x: 0, y: 0, w: 1, h: 1};
  $('vCrop').classList.remove('hidden'); $('vStage').classList.add('hidden');
  $('vTools').classList.add('hidden'); $('vModes').classList.add('hidden'); $('vCropBar').classList.remove('hidden');
  $('cAngle').value = C.angle; $('cAngleVal').textContent = C.angle.toFixed(1) + '°';
  pushLayer('crop', hideCrop);
  cropPreview();
}
function cropPreview() {
  C.prev = ImgProc.rotateAny(C.base, C.rot + C.angle);
  const area = $('vCrop').getBoundingClientRect(), pad = 24;
  C.scale = Math.min((area.width - 2 * pad) / C.prev.width, (area.height - 2 * pad) / C.prev.height);
  const w = C.prev.width * C.scale, h = C.prev.height * C.scale;
  C.off = [(area.width - w) / 2, (area.height - h) / 2];
  const cv = $('cBase'); cv.width = Math.round(w * devicePixelRatio); cv.height = Math.round(h * devicePixelRatio);
  Object.assign(cv.style, {width: w + 'px', height: h + 'px', left: C.off[0] + 'px', top: C.off[1] + 'px'});
  const x = cv.getContext('2d'); x.imageSmoothingQuality = 'high'; x.drawImage(C.prev, 0, 0, cv.width, cv.height);
  $('cSvg').setAttribute('viewBox', `0 0 ${area.width} ${area.height}`);
  drawCrop();
}
function cropRectPx() {
  const w = C.prev.width * C.scale, h = C.prev.height * C.scale;
  return {x: C.off[0] + C.crop.x * w, y: C.off[1] + C.crop.y * h, w: C.crop.w * w, h: C.crop.h * h};
}
function drawCrop(active = -1) {
  const r = cropRectPx(), area = $('vCrop').getBoundingClientRect();
  const pts = [[r.x, r.y], [r.x + r.w, r.y], [r.x + r.w, r.y + r.h], [r.x, r.y + r.h]];
  const grid = [1, 2].map(i => `<line x1="${r.x + r.w * i / 3}" y1="${r.y}" x2="${r.x + r.w * i / 3}" y2="${r.y + r.h}"/><line x1="${r.x}" y1="${r.y + r.h * i / 3}" x2="${r.x + r.w}" y2="${r.y + r.h * i / 3}"/>`).join('');
  $('cSvg').innerHTML = `<path d="M0 0H${area.width}V${area.height}H0Z M${r.x} ${r.y}H${r.x + r.w}V${r.y + r.h}H${r.x}Z" fill="rgba(0,0,0,.55)" fill-rule="evenodd"/>
    <g stroke="rgba(255,255,255,.35)" stroke-width="1">${grid}</g>
    <rect x="${r.x}" y="${r.y}" width="${r.w}" height="${r.h}" fill="none" stroke="#4fc3f7" stroke-width="2"/>
    ${pts.map((p, i) => `<circle cx="${p[0]}" cy="${p[1]}" r="${i === active ? 16 : 13}" fill="rgba(79,195,247,.35)" stroke="#fff" stroke-width="3"/>`).join('')}`;
}
function loupe(pt) {
  const lp = $('cLoupe'); if (!pt) { lp.classList.add('hidden'); return; }
  lp.classList.remove('hidden');
  const S2 = 120; lp.width = lp.height = S2 * devicePixelRatio;
  const sx = (pt[0] - C.off[0]) / C.scale, sy = (pt[1] - C.off[1]) / C.scale, src = S2 / (C.scale * 3);
  const x2 = lp.getContext('2d'); x2.fillStyle = '#000'; x2.fillRect(0, 0, lp.width, lp.height);
  x2.drawImage(C.prev, sx - src / 2, sy - src / 2, src, src, 0, 0, lp.width, lp.height);
  x2.strokeStyle = '#4fc3f7'; x2.lineWidth = 2 * devicePixelRatio;
  x2.beginPath(); x2.moveTo(lp.width / 2, 0); x2.lineTo(lp.width / 2, lp.height); x2.moveTo(0, lp.height / 2); x2.lineTo(lp.width, lp.height / 2); x2.stroke();
  const area = $('vCrop').getBoundingClientRect();
  lp.style.left = (pt[0] < area.width / 2 ? area.width - S2 - 12 : 12) + 'px';
}
(function cropGestures() {
  const svg = $('cSvg');
  const loc = e => { const r = $('vCrop').getBoundingClientRect(); return [e.clientX - r.left, e.clientY - r.top]; };
  svg.addEventListener('pointerdown', e => {
    const [x, y] = loc(e), r = cropRectPx();
    const pts = [[r.x, r.y], [r.x + r.w, r.y], [r.x + r.w, r.y + r.h], [r.x, r.y + r.h]];
    let best = -1, bd = 44;
    pts.forEach((p, i) => { const d = Math.hypot(p[0] - x, p[1] - y); if (d < bd) { bd = d; best = i; } });
    if (best >= 0) C.drag = {corner: best};
    else if (x > r.x && x < r.x + r.w && y > r.y && y < r.y + r.h) C.drag = {move: [x, y], start: {...C.crop}};
    else return;
    svg.setPointerCapture(e.pointerId); drawCrop(best); if (best >= 0) loupe(pts[best]);
  });
  svg.addEventListener('pointermove', e => {
    if (!C.drag) return;
    const [x, y] = loc(e), W = C.prev.width * C.scale, H = C.prev.height * C.scale, MIN = 0.05;
    const nx = Math.min(1, Math.max(0, (x - C.off[0]) / W)), ny = Math.min(1, Math.max(0, (y - C.off[1]) / H));
    const c = C.crop;
    if (C.drag.move) {
      const s = C.drag.start;
      c.x = Math.min(1 - s.w, Math.max(0, s.x + (x - C.drag.move[0]) / W));
      c.y = Math.min(1 - s.h, Math.max(0, s.y + (y - C.drag.move[1]) / H));
      drawCrop(); return;
    }
    let x0 = c.x, y0 = c.y, x1 = c.x + c.w, y1 = c.y + c.h;
    const k = C.drag.corner;
    if (k === 0 || k === 3) x0 = Math.min(nx, x1 - MIN); else x1 = Math.max(nx, x0 + MIN);
    if (k === 0 || k === 1) y0 = Math.min(ny, y1 - MIN); else y1 = Math.max(ny, y0 + MIN);
    Object.assign(c, {x: x0, y: y0, w: x1 - x0, h: y1 - y0});
    drawCrop(k); const r = cropRectPx(); loupe([[r.x, r.y], [r.x + r.w, r.y], [r.x + r.w, r.y + r.h], [r.x, r.y + r.h]][k]);
  });
  const up = () => { C.drag = null; drawCrop(); loupe(null); };
  svg.addEventListener('pointerup', up); svg.addEventListener('pointercancel', up);
})();
$('cAngle').oninput = () => { C.angle = Number($('cAngle').value); $('cAngleVal').textContent = C.angle.toFixed(1) + '°'; cropPreview(); };
$('cAngleMinus').onclick = () => { $('cAngle').value = Math.max(-45, C.angle - 0.5); $('cAngle').oninput(); };
$('cAnglePlus').onclick = () => { $('cAngle').value = Math.min(45, C.angle + 0.5); $('cAngle').oninput(); };
$('cRot90').onclick = () => { C.rot = (C.rot + 90) % 360; C.crop = {x: 0, y: 0, w: 1, h: 1}; cropPreview(); };
$('cAuto').onclick = () => {
  const a = ImgProc.auto(ImgProc.rotate(C.base, C.rot)); // detect on the 90°-oriented photo
  if (!a) return toast('לא זוהה מסמך – כוון ידנית');
  C.angle = a.angle; C.crop = a.crop; $('cAngle').value = C.angle; $('cAngleVal').textContent = C.angle.toFixed(1) + '°'; cropPreview();
};
$('cFull').onclick = () => { C.crop = {x: 0, y: 0, w: 1, h: 1}; C.angle = 0; $('cAngle').value = 0; $('cAngleVal').textContent = '0.0°'; cropPreview(); };
function hideCrop() {
  $('vCrop').classList.add('hidden'); $('vStage').classList.remove('hidden'); $('vCropBar').classList.add('hidden'); loupe(null);
  $('vTools').classList.remove('hidden'); $('vModes').classList.toggle('hidden', !vRec()?.orig); vFit();
}
$('vCropBtn').onclick = openCrop;
$('cCancel').onclick = () => closeLayer('crop');
$('cOk').onclick = async () => {
  const r = vRec(), full = C.crop.x < 0.005 && C.crop.y < 0.005 && C.crop.w > 0.99 && C.crop.h > 0.99;
  r.edit.rot = C.rot; r.edit.angle = C.angle; r.edit.crop = full ? null : {...C.crop};
  closeLayer('crop'); await vReprocess();
};

$('fSave').onclick = async () => {
  const cat = $('fCat').value; const ash = !!catByName(cat)?.ashal;
  const date = F.date;
  if (!date) return toast('יש לבחור תאריך');
  let amount, hours = null;
  if (ash) {
    hours = Number($('fHours').value);
    if (!(hours > 0)) return toast('יש להזין מספר שעות');
    amount = ashalAmount(hours);
  } else {
    amount = Number($('fAmount').value);
    if (!(amount > 0)) return toast('יש להזין סכום');
  }
  const old = F.id ? EXP.find(e => e.id === F.id) : null;
  const rec = {
    id: F.id || uid(), date, period: $('fPeriodWrap').classList.contains('hidden') ? periodKeyFor(date) : $('fPeriod').value, category: cat, ashal: ash,
    amount: Math.round(amount * 100) / 100, hours, description: $('fDesc').value.trim(),
    receipts: ash ? [] : F.receipts.map(r => r.id), createdAt: old?.createdAt || Date.now(), updatedAt: Date.now()
  };
  for (const r of F.receipts) if (r.isNew && !ash) await DB.put('files', {id: r.id, blob: r.blob, type: r.type, orig: r.orig || null, edit: r.edit || null});
  const removed = ash ? [...F.removed, ...F.receipts.filter(r => !r.isNew).map(r => r.id)] : F.removed;
  for (const id of removed) await DB.del('files', id);
  await DB.put('expenses', rec);
  await loadExpenses();
  const msg = rec.period !== currentPeriod() ? `נשמר – בדוח ${periodLabel(rec.period)}` : (!ash && !rec.receipts.length ? 'נשמר (ללא קבלה)' : 'נשמר ✓');
  closeForm(); listKey = rec.period; renderExpenses(); toast(msg);
};
$('fDelete').onclick = async () => {
  if (!confirm('למחוק את ההוצאה?')) return;
  const e = EXP.find(x => x.id === F.id);
  for (const id of e.receipts || []) await DB.del('files', id);
  await DB.del('expenses', e.id); await loadExpenses(); closeForm(); renderExpenses(); toast('נמחק');
};

/* ---------------- report ---------------- */
function reportItems(key) { // items that go into the Excel form (A.SH.L is reported separately in ProperTime)
  return EXP.filter(e => e.period === key && !isAshal(e)).sort((a, b) => a.date.localeCompare(b.date) || a.createdAt - b.createdAt)
    .map((e, i) => ({...e, index: i + 1}));
}
function ashalItems(key) {
  return EXP.filter(e => e.period === key && isAshal(e)).sort((a, b) => a.date.localeCompare(b.date));
}
function byCategory(items) {
  const m = new Map();
  items.forEach(e => { const c = m.get(e.category) || {name: e.category, count: 0, sum: 0}; c.count++; c.sum += Number(e.amount); m.set(e.category, c); });
  return [...m.values()];
}
function reportChecks(items) {
  const out = [];
  if (!items.length) out.push(['bad', 'אין הוצאות לדוח בתקופה זו']);
  const miss = items.filter(e => !e.receipts?.length);
  if (miss.length) out.push(['warn', `${miss.length} הוצאות ללא קבלה: ${miss.map(e => '#' + e.index).join(', ')}`]);
  const bm = items.filter(e => e.category === 'ארוחות עסקיות' && !e.description);
  if (bm.length) out.push(['warn', `ארוחה עסקית בלי פרטי פגישה (עם מי ובאיזה פרויקט): ${bm.map(e => '#' + e.index).join(', ')}`]);
  if (!S.name) out.push(['bad', 'חסר שם עובד (הגדרות)']);
  if (!S.acctEmail) out.push(['bad', 'חסר מייל הנהלת חשבונות (הגדרות)']);
  return out;
}
function ashalHtml(key) {
  const a = ashalItems(key); if (!a.length) return '';
  const tot = a.reduce((s, e) => s + Number(e.amount), 0);
  return `<div class="card"><h2>א.ש.ל – לדיווח בפרופר טיים</h2>
    <div class="muted" style="margin-bottom:8px">לפי הנוהל א.ש.ל מדווח במערכת השעות ולא נכנס לדוח ההוצאות.</div>
    <table class="rep"><thead><tr><th>תאריך</th><th>שעות</th><th>תיאור</th><th>סכום</th></tr></thead><tbody>
    ${a.map(e => `<tr><td>${fmtDate(e.date)}</td><td>${e.hours}</td><td>${esc(e.description)}</td><td>${num2(e.amount)}</td></tr>`).join('')}
    </tbody><tfoot><tr><td colspan="3">סה"כ ${a.length} ימים</td><td>${num2(tot)}</td></tr></tfoot></table></div>`;
}
let repKey = null;
async function renderReport() {
  fillPeriodSelect($('repPeriod'), repKey || currentPeriod());
  repKey = $('repPeriod').value;
  const items = reportItems(repKey), r = periodRange(repKey);
  const sent = S.sentReports[repKey];
  $('repStatus').innerHTML = `<div class="muted">${fmtDate(r.start)} – ${fmtDate(r.end)} · הגשה עד ${fmtDate(r.end)}</div>` +
    (sent ? `<div class="banner ok" style="margin:8px 0 0">✓ נשלח ב-${new Date(sent.sentAt).toLocaleDateString('he-IL')} (${money(sent.total)})
      <a href="#" id="unsend">ביטול סימון</a></div>` : '');
  const un = $('unsend'); if (un) un.onclick = async ev => { ev.preventDefault(); delete S.sentReports[repKey]; await saveSettingsObj(); renderReport(); };
  const sig = await DB.get('files', 'signature');
  const tpl = await DB.get('files', 'template');
  const checks = reportChecks(items);
  if (!sig) checks.push(['bad', 'חסרה חתימה (הגדרות)']);
  if (!tpl) checks.push(['bad', 'חסר קובץ האקסל של הנהלת החשבונות (הגדרות ← תבנית הדוח)']);
  $('repChecks').innerHTML = checks.map(([lvl, t]) => `<div class="banner">${lvl === 'bad' ? '⛔' : '⚠️'} ${esc(t)}</div>`).join('');
  const total = items.reduce((s, e) => s + Number(e.amount), 0);
  $('repSummary').innerHTML = `<h2>דוח הוצאות – ${periodLabel(repKey)}</h2>` +
    (items.length ? `<table class="rep"><thead><tr><th>סוג הוצאה</th><th>פריטים</th><th>סכום</th></tr></thead><tbody>` +
      byCategory(items).map(c => `<tr><td>${esc(c.name)}</td><td>${c.count}</td><td>${num2(c.sum)}</td></tr>`).join('') +
      `</tbody><tfoot><tr><td>סה"כ</td><td>${items.length}</td><td>${num2(total)}</td></tr></tfoot></table>` : '<div class="empty">אין הוצאות לדוח</div>');
  $('repAshal').innerHTML = ashalHtml(repKey);
  $('btnReview').disabled = checks.some(c => c[0] === 'bad');
}
$('repPeriod').onchange = () => { repKey = $('repPeriod').value; renderReport(); };

async function openReview() {
  const items = reportItems(repKey), total = items.reduce((s, e) => s + Number(e.amount), 0);
  $('rvTitle').textContent = 'עיון בדוח – ' + periodLabel(repKey);
  const r = periodRange(repKey);
  const body = $('rvBody');
  body.innerHTML = `
  <div class="card">
    <div style="font-weight:700;font-size:17px">דוח הוצאות עובדים – ${periodLabel(repKey)}</div>
    <div class="muted">${esc(S.name)}${S.role ? ' · ' + esc(S.role) : ''}${S.company ? ' · ' + esc(S.company) : ''}${S.manager ? ' · מנהל מאשר: ' + esc(S.manager) : ''}</div>
    <div class="muted">${fmtDate(r.start)} – ${fmtDate(r.end)}</div>
  </div>
  <div class="card" style="overflow-x:auto">
    <table class="rep"><thead><tr><th># חשבונית</th><th>תאריך</th><th>סוג</th><th>תיאור</th><th>סכום</th></tr></thead><tbody>
    ${items.map(e => `<tr><td>${e.receipts?.length ? e.index : '<b style="color:var(--bad)">חסרה</b>'}</td><td>${fmtDate(e.date)}${periodKeyFor(e.date) !== e.period ? ' <span class="tag warn">מאוחרת</span>' : ''}</td><td>${esc(e.category)}</td><td>${esc(e.description)}</td><td>${num2(e.amount)}</td></tr>`).join('')}
    </tbody><tfoot><tr><td colspan="4">סה"כ</td><td>${num2(total)}</td></tr></tfoot></table>
    <div class="row" style="margin-top:6px">
      <div><label>מקדמות (₪, אם יש)</label><input id="rvAdv" type="number" inputmode="decimal" placeholder="0"></div>
      <div><label>הערות לדוח</label><input id="rvNotes" value="${esc(S.notes)}"></div>
    </div>
  </div>
  <div class="card"><h2>קבלות מצורפות</h2><div class="gallery" id="rvGal"></div></div>
  ${ashalHtml(repKey)}
  <div class="card">
    <button class="btn sec block" id="rvDraft" style="margin-top:0">📊 פתיחת טיוטת Excel (ללא חתימה)</button>
    <label style="display:flex;align-items:center;color:var(--ink);font-size:15px;margin-top:14px">
      <input type="checkbox" id="rvOk">עיינתי בדוח ואני מאשר/ת את נכונותו</label>
    <button class="btn ok block" id="rvSign" disabled>✍️ אישור, חתימה והכנה לשליחה</button>
  </div>`;
  const gal = $('rvGal'); const cells = [];
  for (const e of items) for (const id of e.receipts || []) {
    const f = await DB.get('files', id); if (!f) continue;
    cells.push(f.type.startsWith('image/') ? `<div data-f="${id}" style="background-image:url(${await fileURL(id)})"><span>#${e.index}</span></div>` : `<div>PDF<span>#${e.index}</span></div>`);
  }
  gal.innerHTML = cells.join('') || '<div class="muted" style="grid-column:1/-1">אין קבלות</div>';
  gal.querySelectorAll('[data-f]').forEach(el => el.onclick = async () => { const f = await DB.get('files', el.dataset.f); if (f) openViewer(-1, {blob: f.blob}); });
  $('rvOk').onchange = () => $('rvSign').disabled = !$('rvOk').checked;
  $('rvDraft').onclick = async () => {
    spin(true, 'מכין טיוטה…');
    try { const b = await buildXlsx(repKey, false); downloadBlob(b, fileNames(repKey).draft); }
    catch (err) { console.error(err); alert('שגיאה ביצירת הדוח: ' + err.message); }
    spin(false);
  };
  $('rvSign').onclick = () => signAndPrepare();
  if (!$('review').classList.contains('on')) pushLayer('review', hideReview);
  $('review').classList.add('on'); $('review').scrollTop = 0;
}
$('btnReview').onclick = openReview;
function hideReview() { $('review').classList.remove('on'); renderReport(); }
$('rvClose').onclick = () => { if (!closeLayer('review')) hideReview(); };

function fileNames(key) {
  const n = (S.name || 'עובד').replace(/[\\/:*?"<>|]/g, '').trim().replace(/\s+/g, '_');
  return {xlsx: `דוח_הוצאות_${n}_${key}.xlsx`, draft: `טיוטה_דוח_הוצאות_${key}.xlsx`, pdf: `קבלות_${n}_${key}.pdf`};
}

/* ----- XLSX ----- */
const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
async function getTemplate() {
  const t = await DB.get('files', 'template');
  if (!t) throw new Error('חסר קובץ האקסל של הנהלת החשבונות (הגדרות)');
  return await t.blob.arrayBuffer();
}
async function sigPng() {
  const sig = await DB.get('files', 'signature'); if (!sig) throw new Error('חסרה חתימה');
  const bmp = await loadBitmap(sig.blob);
  return {buffer: new Uint8Array(await sig.blob.arrayBuffer()), width: bmp.width, height: bmp.height};
}
async function buildXlsx(key, signed) {
  const items = reportItems(key);
  const [y, m] = key.split('-').map(Number);
  const utc = iso => { const [a, b, c] = iso.split('-').map(Number); return new Date(Date.UTC(a, b - 1, c)); };
  const adv = $('rvAdv') ? Number($('rvAdv').value) || 0 : 0;
  const notes = $('rvNotes') ? $('rvNotes').value.trim() : S.notes;
  const data = {
    name: S.name, role: S.role, company: S.company, manager: S.manager, notes, advances: adv,
    monthDate: new Date(Date.UTC(y, m - 1, 1)),
    items: items.map(e => ({date: utc(e.date), description: e.description || '', category: e.category, amount: Number(e.amount), invoice: e.receipts?.length ? e.index : null}))
  };
  const {buffer} = await ExpenseXlsx.fill(ExcelJS, PizZip, await getTemplate(), data, signed ? await sigPng() : null);
  return new Blob([buffer], {type: XLSX_MIME});
}

/* ----- receipts PDF ----- */
async function buildReceiptsPdf(key) {
  const {PDFDocument, StandardFonts, rgb} = PDFLib;
  const items = reportItems(key).filter(e => e.receipts?.length);
  if (!items.length) return null;
  const pdf = await PDFDocument.create();
  const font = await pdf.embedFont(StandardFonts.HelveticaBold);
  const A4 = [595.28, 841.89], M = 28;
  const label = (page, e, n, cnt) => {
    const txt = `#${e.index}   ${e.date}   ${Number(e.amount).toFixed(2)} ILS${cnt > 1 ? `   (${n}/${cnt})` : ''}`;
    const {width, height} = page.getSize();
    page.drawRectangle({x: 0, y: height - 22, width, height: 22, color: rgb(0.12, 0.31, 0.47), opacity: 0.9});
    page.drawText(txt, {x: 10, y: height - 16, size: 11, font, color: rgb(1, 1, 1)});
  };
  for (const e of items) {
    for (let n = 0; n < e.receipts.length; n++) {
      const f = await DB.get('files', e.receipts[n]); if (!f) continue;
      const buf = await f.blob.arrayBuffer();
      if (f.type === 'application/pdf') {
        try {
          const src = await PDFDocument.load(buf, {ignoreEncryption: true});
          const pages = await pdf.copyPages(src, src.getPageIndices());
          pages.forEach(p => { pdf.addPage(p); label(p, e, n + 1, e.receipts.length); });
        } catch (err) { console.warn('pdf receipt failed', err); }
      } else {
        const img = f.type === 'image/png' ? await pdf.embedPng(buf) : await pdf.embedJpg(buf);
        const page = pdf.addPage(A4);
        const maxW = A4[0] - 2 * M, maxH = A4[1] - 2 * M - 26;
        const k = Math.min(maxW / img.width, maxH / img.height);
        const w = img.width * k, h = img.height * k;
        page.drawImage(img, {x: (A4[0] - w) / 2, y: M + (maxH - h) / 2, width: w, height: h});
        label(page, e, n + 1, e.receipts.length);
      }
    }
  }
  return new Blob([await pdf.save()], {type: 'application/pdf'});
}

/* ----- sign & send ----- */
let OUT = null;
async function signAndPrepare() {
  spin(true, 'חותם ומכין קבצים…');
  try {
    const names = fileNames(repKey);
    const xlsx = await buildXlsx(repKey, true);
    const pdf = await buildReceiptsPdf(repKey);
    OUT = {key: repKey, files: [new File([xlsx], names.xlsx, {type: XLSX_MIME})]};
    if (pdf) OUT.files.push(new File([pdf], names.pdf, {type: 'application/pdf'}));
  } catch (err) { console.error(err); spin(false); alert('שגיאה: ' + err.message); return; }
  spin(false);
  renderSend();
}
function mailText(key) {
  const items = reportItems(key), total = items.reduce((s, e) => s + Number(e.amount), 0);
  const subject = `דוח הוצאות אישיות – ${S.name} – ${periodLabel(key)}`;
  const body = `שלום ${S.acctName || ''},\n\nמצורף דוח ההוצאות האישיות שלי לחודש ${periodLabel(key)} (חתום), יחד עם צילומי הקבלות והחשבוניות.\n` +
    `סה"כ: ${money(total)} (${items.length} פריטים).\n\nתודה,\n${S.name}`;
  return {subject, body};
}
function renderSend() {
  const {subject, body} = mailText(OUT.key);
  const canShare = !!(navigator.canShare && navigator.canShare({files: OUT.files}));
  const sizeKB = OUT.files.reduce((s, f) => s + f.size, 0) / 1024;
  $('rvTitle').textContent = 'שליחה להנהלת חשבונות';
  $('rvBody').innerHTML = `
  <div class="banner ok">✓ הדוח נחתם. קבצים מוכנים (${sizeKB > 1024 ? (sizeKB / 1024).toFixed(1) + ' MB' : Math.round(sizeKB) + ' KB'}):<br>
    ${OUT.files.map(f => '• <bdi>' + esc(f.name) + '</bdi>').join('<br>')}</div>
  <div class="card">
    <div class="muted">נמען</div>
    <div style="display:flex;gap:8px;align-items:center;margin:4px 0 10px"><b dir="ltr" style="flex:1">${esc(S.acctEmail)}</b>
      <button class="btn sec" id="cpEmail" style="padding:8px 12px">העתק</button></div>
    <div class="muted">נושא</div><div style="margin:4px 0 10px">${esc(subject)}</div>
    <ol class="steps">
      ${canShare ? `<li>לחץ "שליחה" ובחר Outlook / Gmail</li><li>המייל נפתח עם הקבצים מצורפים – הדבק את כתובת הנמען (הועתקה)</li><li>שלח, וחזור לכאן לסמן "נשלח"</li>`
                 : `<li>הורד את הקבצים</li><li>פתח מייל חדש (הכפתור ממלא נמען, נושא וגוף)</li><li>צרף את הקבצים ושלח, וחזור לכאן לסמן "נשלח"</li>`}
    </ol>
    ${canShare ? `<button class="btn block" id="doShare">📤 שליחה</button>` : ''}
    <button class="btn ${canShare ? 'sec' : ''} block" id="doDownload">⬇️ הורדת הקבצים</button>
    <a class="btn sec block" id="doMailto" href="mailto:${encodeURIComponent(S.acctEmail)}?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}">✉️ פתיחת מייל (נמען + נושא)</a>
  </div>
  <button class="btn ok block" id="markSent">✓ סימון הדוח כנשלח</button>`;
  $('cpEmail').onclick = async () => { try { await navigator.clipboard.writeText(S.acctEmail); toast('הכתובת הועתקה'); } catch { toast('העתקה נכשלה'); } };
  const sh = $('doShare');
  if (sh) sh.onclick = async () => {
    try { await navigator.clipboard.writeText(S.acctEmail); } catch {}
    try { await navigator.share({files: OUT.files, title: subject, text: body}); toast('הכתובת הועתקה – הדבק בשדה "אל"', 4000); }
    catch (err) { if (err.name !== 'AbortError') toast('השיתוף נכשל – השתמש בהורדה'); }
  };
  $('doDownload').onclick = () => OUT.files.forEach((f, i) => setTimeout(() => downloadBlob(f, f.name), i * 600));
  $('markSent').onclick = async () => {
    const items = reportItems(OUT.key);
    S.sentReports[OUT.key] = {sentAt: Date.now(), total: items.reduce((s, e) => s + Number(e.amount), 0), count: items.length};
    await saveSettingsObj(); if (!closeLayer('review')) hideReview(); toast('הדוח סומן כנשלח ✓');
  };
}

/* ---------------- settings ---------------- */
async function renderSettings() {
  $('sName').value = S.name; $('sRole').value = S.role; $('sManager').value = S.manager; $('sCompany').value = S.company;
  $('sAcctEmail').value = S.acctEmail; $('sAcctName').value = S.acctName; $('sNotes').value = S.notes;
  renderCats(); renderTiers();
  await renderSigTpl();
  try {
    const est = await navigator.storage.estimate(); const p = await navigator.storage.persisted?.();
    $('storageInfo').textContent = `נתונים שמורים בטלפון בלבד · ${(est.usage / 1048576).toFixed(1)} MB${p ? ' · אחסון קבוע ✓' : ''}. מומלץ לייצא גיבוי פעם בחודש.`;
  } catch { $('storageInfo').textContent = 'הנתונים שמורים בטלפון בלבד. מומלץ לייצא גיבוי פעם בחודש.'; }
  $('ver').textContent = VERSION;
}
async function renderSigTpl() {
  const sig = await DB.get('files', 'signature');
  $('sigPreview').innerHTML = sig ? `<img src="${URL.createObjectURL(sig.blob)}">` : '<span class="muted">אין חתימה</span>';
  const t = await DB.get('files', 'template');
  $('tplStatus').innerHTML = t ? `✓ ${esc(t.name || 'template.xlsx')}` : '<b style="color:var(--bad)">לא הועלה קובץ.</b> העלה את קובץ האקסל של הנהלת החשבונות (דוח החזר הוצאות).';
  $('btnTplReset').classList.toggle('hidden', !t);
}
function readSettingsForm() {
  S.name = $('sName').value.trim(); S.role = $('sRole').value.trim(); S.manager = $('sManager').value.trim(); S.company = $('sCompany').value.trim();
  S.acctEmail = $('sAcctEmail').value.trim(); S.acctName = $('sAcctName').value.trim(); S.notes = $('sNotes').value.trim();
  const rows = [...document.querySelectorAll('#tierList .catrow')];
  if (rows.length) S.ashalTiers = rows.map(r => { const [a, b] = r.querySelectorAll('input'); return {upTo: a.value === '' ? null : Number(a.value), amount: Number(b.value) || 0}; });
}
function renderCats() {
  $('catList').innerHTML = S.categories.map(c => `<span class="tag" style="font-size:13px;padding:4px 10px;margin:3px">${esc(c.name)}${c.ashal ? ' (פרופר טיים)' : ''}</span>`).join('');
}
function renderTiers() {
  $('tierList').innerHTML = S.ashalTiers.map(t => `<div class="catrow"><span class="muted" style="white-space:nowrap">עד</span>
    <input type="number" inputmode="decimal" value="${t.upTo ?? ''}" placeholder="ללא הגבלה"><span class="muted" style="white-space:nowrap">שעות ←</span>
    <input type="number" inputmode="decimal" value="${t.amount || ''}" placeholder="₪"></div>`).join('');
}
let _asT;
$('view-settings').addEventListener('input', ev => {
  if (ev.target.type === 'file') return;
  clearTimeout(_asT); _asT = setTimeout(() => { readSettingsForm(); saveSettingsObj(); }, 500);
});
$('view-settings').addEventListener('change', ev => { if (ev.target.type !== 'file') { readSettingsForm(); saveSettingsObj(); } });
$('btnAshalReset').onclick = () => { S.ashalTiers = structuredClone(DEFAULT_SETTINGS.ashalTiers); renderTiers(); saveSettingsObj(); toast('תעריפי הנוהל שוחזרו'); };
$('btnSaveSettings').onclick = async () => {
  readSettingsForm();
  if (S.acctEmail && !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(S.acctEmail)) return toast('כתובת מייל לא תקינה');
  await saveSettingsObj(); toast('ההגדרות נשמרו ✓');
};

// template (accounting's Excel form)
$('btnTplUpload').onclick = () => $('tplFile').click();
$('tplFile').onchange = async ev => {
  const f = ev.target.files[0]; ev.target.value = ''; if (!f) return;
  spin(true, 'בודק את הקובץ…');
  try {
    const L = await ExpenseXlsx.inspect(ExcelJS, await f.arrayBuffer());
    await DB.put('files', {id: 'template', blob: f, type: XLSX_MIME, name: f.name});
    readSettingsForm();
    S.categories = [...L.categories.map(c => ({name: c.name, ashal: false})), {name: ASHAL_NAME, ashal: true}];
    await saveSettingsObj(); renderCats(); renderSigTpl();
    const rows = L.lastData - L.firstData + 1;
    toast(`✓ הקובץ נשמר: ${L.categories.length} סוגי הוצאות, ${rows} שורות${L.signature ? ', מקום לחתימה' : ' – ⚠️ לא נמצא מקום לחתימה'}`, 5000);
  } catch (err) { console.error(err); alert('הקובץ לא זוהה כטופס דוח ההוצאות: ' + err.message); }
  spin(false);
};
$('btnTplReset').onclick = async () => { if (!confirm('להסיר את קובץ האקסל?')) return; await DB.del('files', 'template'); renderSigTpl(); };

// signature: upload (with background removal) or draw
$('btnSigUpload').onclick = () => $('sigFile').click();
$('sigFile').onchange = async ev => {
  const f = ev.target.files[0]; ev.target.value = ''; if (!f) return;
  spin(true, 'מעבד חתימה…');
  try { const png = await cleanSignature(f); await DB.put('files', {id: 'signature', blob: png, type: 'image/png'}); renderSigTpl(); toast('החתימה נשמרה'); }
  catch (err) { console.error(err); toast('לא הצלחתי לעבד את התמונה'); }
  spin(false);
};
async function cleanSignature(blob) {
  const bmp = await loadBitmap(blob);
  const k = Math.min(1, 1400 / Math.max(bmp.width, bmp.height));
  const W = Math.round(bmp.width * k), H = Math.round(bmp.height * k);
  const c = document.createElement('canvas'); c.width = W; c.height = H;
  const x = c.getContext('2d'); x.drawImage(bmp, 0, 0, W, H);
  const im = x.getImageData(0, 0, W, H), d = im.data;
  let minX = W, minY = H, maxX = -1, maxY = -1;
  for (let y = 0; y < H; y++) for (let i = 0; i < W; i++) {
    const p = (y * W + i) * 4;
    if (d[p + 3] < 20) continue;
    const g = d[p] * 0.299 + d[p + 1] * 0.587 + d[p + 2] * 0.114;
    if (g > 170) { d[p + 3] = 0; continue; }
    d[p + 3] = Math.min(255, (170 - g) * 3); d[p] = d[p + 1] = 20; d[p + 2] = 90; // dark-blue ink
    if (i < minX) minX = i; if (i > maxX) maxX = i; if (y < minY) minY = y; if (y > maxY) maxY = y;
  }
  x.putImageData(im, 0, 0);
  if (maxX < 0) throw new Error('empty');
  const m = 10; minX = Math.max(0, minX - m); minY = Math.max(0, minY - m); maxX = Math.min(W - 1, maxX + m); maxY = Math.min(H - 1, maxY + m);
  const o = document.createElement('canvas'); o.width = maxX - minX + 1; o.height = maxY - minY + 1;
  o.getContext('2d').drawImage(c, minX, minY, o.width, o.height, 0, 0, o.width, o.height);
  return await new Promise(res => o.toBlob(res, 'image/png'));
}
(function sigPad() {
  const cv = $('sigpad'); let ctx, drawing = false, dirty = false;
  const setup = () => {
    const r = cv.getBoundingClientRect(), dpr = window.devicePixelRatio || 1;
    cv.width = r.width * dpr; cv.height = r.height * dpr; ctx = cv.getContext('2d'); ctx.scale(dpr, dpr);
    ctx.lineWidth = 2.6; ctx.lineCap = 'round'; ctx.lineJoin = 'round'; ctx.strokeStyle = '#14145a'; dirty = false;
  };
  const pos = e => { const r = cv.getBoundingClientRect(); return [e.clientX - r.left, e.clientY - r.top]; };
  cv.onpointerdown = e => { drawing = true; dirty = true; cv.setPointerCapture(e.pointerId); ctx.beginPath(); ctx.moveTo(...pos(e)); };
  cv.onpointermove = e => { if (!drawing) return; ctx.lineTo(...pos(e)); ctx.stroke(); };
  cv.onpointerup = cv.onpointercancel = () => { drawing = false; };
  $('btnSigDraw').onclick = () => { $('sigSheet').classList.add('on'); pushLayer('sig', () => $('sigSheet').classList.remove('on')); requestAnimationFrame(setup); };
  $('sigClose').onclick = () => { if (!closeLayer('sig')) $('sigSheet').classList.remove('on'); };
  $('sigClear').onclick = setup;
  $('sigSave').onclick = async () => {
    if (!dirty) return toast('יש לחתום קודם');
    const raw = await new Promise(res => cv.toBlob(res, 'image/png'));
    const png = await cleanSignature(raw);
    await DB.put('files', {id: 'signature', blob: png, type: 'image/png'});
    if (!closeLayer('sig')) $('sigSheet').classList.remove('on'); renderSigTpl(); toast('החתימה נשמרה');
  };
})();

// backup / restore
const b64 = buf => { let s = ''; const a = new Uint8Array(buf); for (let i = 0; i < a.length; i += 0x8000) s += String.fromCharCode.apply(null, a.subarray(i, i + 0x8000)); return btoa(s); };
const unb64 = s => { const bin = atob(s), a = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) a[i] = bin.charCodeAt(i); return a; };
$('btnBackup').onclick = async () => {
  spin(true, 'מכין גיבוי…');
  try {
    readSettingsForm();
    const files = [];
    for (const f of await DB.all('files')) files.push({id: f.id, type: f.type, name: f.name, data: b64(await f.blob.arrayBuffer()), orig: f.orig ? b64(await f.orig.arrayBuffer()) : null, edit: f.edit || null});
    const data = {app: 'expenses-app', version: VERSION, exportedAt: new Date().toISOString(), settings: S, expenses: await DB.all('expenses'), files};
    const blob = new Blob([JSON.stringify(data)], {type: 'application/json'});
    const name = `גיבוי_הוצאות_${todayISO()}.json`;
    const file = new File([blob], name, {type: 'application/json'});
    spin(false);
    if (navigator.canShare && navigator.canShare({files: [file]})) {
      try { await navigator.share({files: [file], title: name}); return; } catch (e) { if (e.name === 'AbortError') return; }
    }
    downloadBlob(blob, name);
  } catch (err) { spin(false); console.error(err); toast('הגיבוי נכשל'); }
};
$('btnRestore').onclick = () => $('restoreFile').click();
$('restoreFile').onchange = async ev => {
  const f = ev.target.files[0]; ev.target.value = ''; if (!f) return;
  if (!confirm('השחזור יחליף את כל הנתונים הקיימים בטלפון. להמשיך?')) return;
  spin(true, 'משחזר…');
  try {
    const data = JSON.parse(await f.text());
    if (data.app !== 'expenses-app') throw new Error('קובץ לא מתאים');
    await DB.clear('expenses'); await DB.clear('files');
    for (const e of data.expenses) await DB.put('expenses', e);
    for (const x of data.files) await DB.put('files', {id: x.id, type: x.type, name: x.name, blob: new Blob([unb64(x.data)], {type: x.type}), orig: x.orig ? new Blob([unb64(x.orig)], {type: 'image/jpeg'}) : null, edit: x.edit || null});
    S = Object.assign(structuredClone(DEFAULT_SETTINGS), data.settings); await saveSettingsObj();
    urlCache.clear(); await loadExpenses(); renderSettings(); toast(`שוחזרו ${data.expenses.length} הוצאות`);
  } catch (err) { console.error(err); alert('השחזור נכשל: ' + err.message); }
  spin(false);
};
$('btnIcs').onclick = () => {
  const now = new Date(); let y = now.getFullYear(), m = now.getMonth() + 1; if (now.getDate() > 24) { m++; if (m > 12) { m = 1; y++; } }
  const ics = ['BEGIN:VCALENDAR', 'VERSION:2.0', 'PRODID:-//expenses-app//HE', 'BEGIN:VEVENT', `UID:expenses-reminder-${Date.now()}@expenses-app`,
    `DTSTAMP:${now.toISOString().replace(/[-:]/g, '').split('.')[0]}Z`, `DTSTART;TZID=Asia/Jerusalem:${y}${pad(m)}24T090000`,
    'DURATION:PT15M', 'RRULE:FREQ=MONTHLY;BYMONTHDAY=24', 'SUMMARY:הגשת דוח הוצאות אישיות (עד ה-25)',
    'BEGIN:VALARM', 'ACTION:DISPLAY', 'DESCRIPTION:דוח הוצאות', 'TRIGGER:PT0M', 'END:VALARM', 'END:VEVENT', 'END:VCALENDAR'].join('\r\n');
  downloadBlob(new Blob([ics], {type: 'text/calendar'}), 'expenses-reminder.ics');
};

/* ---------------- boot ---------------- */
(async function boot() {
  try {
    history.replaceState({base: 1}, ''); history.pushState({guard: 1}, '');
    await DB.open(); await loadSettings(); await loadExpenses();
    try { await navigator.storage?.persist?.(); } catch {}
    renderExpenses();
    if (!S.name) setTimeout(() => toast('ברוך הבא! התחל בהגדרות: שם, מייל הנה"ח וחתימה', 4500), 400);
  } catch (err) { console.error(err); document.body.insertAdjacentHTML('afterbegin', `<div class="banner">שגיאה בטעינת האחסון: ${esc(err.message)}</div>`); }
  if ('serviceWorker' in navigator && location.protocol === 'https:') navigator.serviceWorker.register('sw.js').catch(() => {});
})();
