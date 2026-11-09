'use strict';
const VERSION = '1.0.0';
const CUTOFF_DAY = 25;
const HEB_MONTHS = ['ינואר','פברואר','מרץ','אפריל','מאי','יוני','יולי','אוגוסט','ספטמבר','אוקטובר','נובמבר','דצמבר'];

const DEFAULT_SETTINGS = {
  name: '', empId: '', dept: '', company: '', acctEmail: '', acctName: 'הנהלת חשבונות', notes: '',
  categories: [
    {name: 'דלק', ashal: false}, {name: 'חניה', ashal: false}, {name: 'כבישי אגרה', ashal: false},
    {name: 'תחבורה ציבורית / מונית', ashal: false}, {name: 'ארוחות וכיבוד', ashal: false},
    {name: 'ציוד וחומרים', ashal: false}, {name: 'טלפון ותקשורת', ashal: false},
    {name: 'לינה', ashal: false}, {name: 'א.ש.ל', ashal: true}, {name: 'אחר', ashal: false}
  ],
  ashalMode: 'tiers', ashalRate: 0,
  ashalTiers: [{from: 4, amount: 0}, {from: 8, amount: 0}],
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
}
async function saveSettingsObj() { await DB.put('kv', {k: 'settings', v: S}); }
async function loadExpenses() { EXP = (await DB.all('expenses')).sort((a, b) => b.date.localeCompare(a.date) || b.createdAt - a.createdAt); }
const catByName = n => S.categories.find(c => c.name === n);
const isAshal = e => !!(catByName(e.category)?.ashal || e.ashal);

function ashalAmount(hours) {
  hours = Number(hours) || 0;
  if (S.ashalMode === 'hourly') return Math.round(hours * (Number(S.ashalRate) || 0) * 100) / 100;
  let amt = 0;
  [...S.ashalTiers].sort((a, b) => a.from - b.from).forEach(t => { if (hours >= Number(t.from)) amt = Number(t.amount) || 0; });
  return amt;
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
document.querySelectorAll('nav button').forEach(b => b.onclick = async () => {
  if ($('view-settings').classList.contains('on')) { readSettingsForm(); await saveSettingsObj(); }
  show(b.dataset.v);
});

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
  const g = $('goRep'); if (g) g.onclick = ev => { ev.preventDefault(); show('report'); };

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
      <div class="mid"><div class="t">${esc(e.category)}${tag}</div><div class="muted">${fmtDate(e.date)} · ${esc(e.description || '')}</div></div>
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
  $('fDate').value = e?.date || todayISO();
  $('fAmount').value = e && !isAshal(e) ? e.amount : '';
  $('fHours').value = e?.hours || '';
  $('fDesc').value = e?.description || '';
  $('fDelete').classList.toggle('hidden', !e);
  if (e?.receipts) for (const rid of e.receipts) { const f = await DB.get('files', rid); if (f) F.receipts.push({id: rid, blob: f.blob, type: f.type}); }
  if (e && S.sentReports[e.period]) toast('שים לב: ההוצאה שייכת לדוח שכבר נשלח');
  onCatChange(); renderFormReceipts();
  $('sheet').classList.add('on');
}
function closeForm() { $('sheet').classList.remove('on'); F = null; }
$('fClose').onclick = closeForm;
$('sheet').onclick = ev => { if (ev.target.id === 'sheet') closeForm(); };
$('fab').onclick = () => openForm(null);

function onCatChange() {
  const ash = !!catByName($('fCat').value)?.ashal;
  $('fAshal').classList.toggle('hidden', !ash);
  $('fAmountWrap').classList.toggle('hidden', ash);
  $('fRcptWrap').classList.toggle('hidden', ash);
  if (ash) updAshal();
}
function updAshal() {
  const a = ashalAmount($('fHours').value);
  $('fAshalAmt').value = num2(a);
  const noRate = S.ashalMode === 'hourly' ? !(Number(S.ashalRate) > 0) : !S.ashalTiers.some(t => Number(t.amount) > 0);
  $('fAshalNote').textContent = noRate ? '⚠️ לא הוגדר תעריף א.ש.ל – יש להגדיר בהגדרות.' : 'א.ש.ל – לא נדרשת קבלה.';
}
$('fCat').onchange = onCatChange;
$('fHours').oninput = updAshal;

$('btnCam').onclick = () => $('camInput').click();
$('btnPick').onclick = () => $('pickInput').click();
$('camInput').onchange = ev => addReceiptFiles(ev.target.files, ev.target);
$('pickInput').onchange = ev => addReceiptFiles(ev.target.files, ev.target);

async function addReceiptFiles(files, input) {
  spin(true, 'מעבד תמונה…');
  try {
    for (const f of files) {
      if (f.type === 'application/pdf') F.receipts.push({id: uid(), blob: f, type: 'application/pdf', isNew: true});
      else if (f.type.startsWith('image/') || /\.(heic|heif)$/i.test(f.name)) {
        const blob = await processImage(f, $('fScan').checked);
        F.receipts.push({id: uid(), blob, type: 'image/jpeg', isNew: true});
      }
    }
  } catch (err) { console.error(err); toast('לא הצלחתי לעבד את הקובץ'); }
  input.value = ''; spin(false); renderFormReceipts();
}

async function loadBitmap(blob) {
  try { return await createImageBitmap(blob, {imageOrientation: 'from-image'}); }
  catch {
    return await new Promise((res, rej) => { const i = new Image(); i.onload = () => res(i); i.onerror = rej; i.src = URL.createObjectURL(blob); });
  }
}
async function processImage(file, scan, rotate = 0) {
  const bmp = await loadBitmap(file);
  const MAX = 1800;
  let w = bmp.width, h = bmp.height; const k = Math.min(1, MAX / Math.max(w, h)); w = Math.round(w * k); h = Math.round(h * k);
  const rot = rotate % 360; const swap = rot === 90 || rot === 270;
  const c = document.createElement('canvas'); c.width = swap ? h : w; c.height = swap ? w : h;
  const x = c.getContext('2d');
  x.translate(c.width / 2, c.height / 2); x.rotate(rot * Math.PI / 180); x.drawImage(bmp, -w / 2, -h / 2, w, h);
  if (scan) {
    const im = x.getImageData(0, 0, c.width, c.height), d = im.data;
    const hist = new Uint32Array(256);
    for (let i = 0; i < d.length; i += 4) { const g = (d[i] * 0.299 + d[i + 1] * 0.587 + d[i + 2] * 0.114) | 0; d[i] = g; hist[g]++; }
    const total = d.length / 4; let lo = 0, hi = 255, acc = 0;
    for (let i = 0; i < 256; i++) { acc += hist[i]; if (acc > total * 0.02) { lo = i; break; } }
    acc = 0; for (let i = 255; i >= 0; i--) { acc += hist[i]; if (acc > total * 0.15) { hi = i; break; } }
    if (hi - lo < 30) { lo = 0; hi = 255; }
    const span = hi - lo;
    for (let i = 0; i < d.length; i += 4) {
      let v = (d[i] - lo) * 255 / span; v = v < 0 ? 0 : v > 255 ? 255 : v;
      v = 255 * Math.pow(v / 255, 1.4); // darken ink a bit
      d[i] = d[i + 1] = d[i + 2] = v;
    }
    x.putImageData(im, 0, 0);
  }
  return await new Promise(res => c.toBlob(res, 'image/jpeg', 0.8));
}

function renderFormReceipts() {
  const box = $('fReceipts');
  box.innerHTML = F.receipts.map((r, i) => `<div class="rc" data-i="${i}">${r.type === 'application/pdf' ? 'PDF' : ''}
     <button class="del" data-i="${i}">✕</button>${r.type.startsWith('image/') ? `<button class="rot" data-i="${i}">⟳</button>` : ''}</div>`).join('');
  box.querySelectorAll('.rc').forEach(el => {
    const r = F.receipts[el.dataset.i];
    if (r.type.startsWith('image/')) el.style.backgroundImage = `url(${URL.createObjectURL(r.blob)})`;
  });
  box.querySelectorAll('.del').forEach(b => b.onclick = () => {
    const r = F.receipts.splice(b.dataset.i, 1)[0]; if (!r.isNew) F.removed.push(r.id); renderFormReceipts();
  });
  box.querySelectorAll('.rot').forEach(b => b.onclick = async () => {
    const r = F.receipts[b.dataset.i];
    r.blob = await processImage(r.blob, false, 90);
    if (!r.isNew) { F.removed.push(r.id); r.id = uid(); r.isNew = true; }
    renderFormReceipts();
  });
}

$('fSave').onclick = async () => {
  const cat = $('fCat').value; const ash = !!catByName(cat)?.ashal;
  const date = $('fDate').value;
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
    id: F.id || uid(), date, period: periodKeyFor(date), category: cat, ashal: ash,
    amount: Math.round(amount * 100) / 100, hours, description: $('fDesc').value.trim(),
    receipts: ash ? [] : F.receipts.map(r => r.id), createdAt: old?.createdAt || Date.now(), updatedAt: Date.now()
  };
  for (const r of F.receipts) if (r.isNew && !ash) await DB.put('files', {id: r.id, blob: r.blob, type: r.type});
  const removed = ash ? [...F.removed, ...F.receipts.filter(r => !r.isNew).map(r => r.id)] : F.removed;
  for (const id of removed) await DB.del('files', id);
  await DB.put('expenses', rec);
  await loadExpenses();
  const msg = rec.period !== currentPeriod() && date > periodRange(currentPeriod()).end
    ? `נשמר – ייכנס לדוח ${periodLabel(rec.period)}` : (!ash && !rec.receipts.length ? 'נשמר (ללא קבלה)' : 'נשמר ✓');
  closeForm(); listKey = rec.period; renderExpenses(); toast(msg);
};
$('fDelete').onclick = async () => {
  if (!confirm('למחוק את ההוצאה?')) return;
  const e = EXP.find(x => x.id === F.id);
  for (const id of e.receipts || []) await DB.del('files', id);
  await DB.del('expenses', e.id); await loadExpenses(); closeForm(); renderExpenses(); toast('נמחק');
};

/* ---------------- report ---------------- */
function reportItems(key) {
  return EXP.filter(e => e.period === key).sort((a, b) => a.date.localeCompare(b.date) || a.createdAt - b.createdAt)
    .map((e, i) => ({...e, index: i + 1}));
}
function byCategory(items) {
  const m = new Map();
  items.forEach(e => { const c = m.get(e.category) || {name: e.category, count: 0, sum: 0}; c.count++; c.sum += Number(e.amount); m.set(e.category, c); });
  return [...m.values()];
}
function reportChecks(items) {
  const out = [];
  if (!items.length) out.push(['bad', 'אין הוצאות בתקופה זו']);
  const miss = items.filter(e => !isAshal(e) && !e.receipts?.length);
  if (miss.length) out.push(['warn', `${miss.length} הוצאות ללא קבלה: ${miss.map(e => '#' + e.index).join(', ')}`]);
  if (items.some(e => isAshal(e) && !(e.amount > 0))) out.push(['warn', 'יש א.ש.ל בסכום 0 – בדוק תעריף בהגדרות']);
  if (!S.name) out.push(['bad', 'חסר שם עובד (הגדרות)']);
  if (!S.acctEmail) out.push(['bad', 'חסר מייל הנהלת חשבונות (הגדרות)']);
  return out;
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
  const checks = reportChecks(items); if (!sig) checks.push(['bad', 'חסרה חתימה (הגדרות)']);
  $('repChecks').innerHTML = checks.map(([lvl, t]) => `<div class="banner ${lvl === 'ok' ? 'ok' : ''}">${lvl === 'bad' ? '⛔' : '⚠️'} ${esc(t)}</div>`).join('');
  const total = items.reduce((s, e) => s + Number(e.amount), 0);
  $('repSummary').innerHTML = `<h2>סיכום – ${periodLabel(repKey)}</h2>` +
    (items.length ? `<table class="rep"><thead><tr><th>סוג הוצאה</th><th>פריטים</th><th>סכום</th></tr></thead><tbody>` +
      byCategory(items).map(c => `<tr><td>${esc(c.name)}</td><td>${c.count}</td><td>${num2(c.sum)}</td></tr>`).join('') +
      `</tbody><tfoot><tr><td>סה"כ</td><td>${items.length}</td><td>${num2(total)}</td></tr></tfoot></table>` : '<div class="empty">אין הוצאות</div>');
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
    <div class="muted">${esc(S.company)}</div>
    <div style="font-weight:700;font-size:17px">דוח הוצאות אישיות – ${periodLabel(repKey)}</div>
    <div class="muted">${fmtDate(r.start)} – ${fmtDate(r.end)} · ${esc(S.name)}${S.empId ? ' · עובד ' + esc(S.empId) : ''}</div>
  </div>
  <div class="card" style="overflow-x:auto">
    <table class="rep"><thead><tr><th>#</th><th>תאריך</th><th>סוג</th><th>תיאור</th><th>קבלה</th><th>סכום</th></tr></thead><tbody>
    ${items.map(e => `<tr><td>${e.index}</td><td>${fmtDate(e.date)}</td><td>${esc(e.category)}</td><td>${esc(e.description)}${isAshal(e) ? ` (${e.hours} ש')` : ''}</td>
      <td>${isAshal(e) ? '—' : (e.receipts?.length ? 'מצורפת' : '<b style="color:var(--bad)">חסרה</b>')}</td><td>${num2(e.amount)}</td></tr>`).join('')}
    </tbody><tfoot><tr><td colspan="5">סה"כ</td><td>${num2(total)}</td></tr></tfoot></table>
  </div>
  <div class="card"><h2>קבלות מצורפות</h2><div class="gallery" id="rvGal"></div></div>
  <div class="card">
    <button class="btn sec block" id="rvDraft" style="margin-top:0">📄 פתיחת טיוטת Word (ללא חתימה)</button>
    <label style="display:flex;align-items:center;color:var(--ink);font-size:15px;margin-top:14px">
      <input type="checkbox" id="rvOk">עיינתי בדוח ואני מאשר/ת את נכונותו</label>
    <button class="btn ok block" id="rvSign" disabled>✍️ אישור, חתימה והכנה לשליחה</button>
  </div>`;
  const gal = $('rvGal'); const cells = [];
  for (const e of items) for (const id of e.receipts || []) {
    const f = await DB.get('files', id); if (!f) continue;
    cells.push(f.type.startsWith('image/') ? `<div style="background-image:url(${await fileURL(id)})"><span>#${e.index}</span></div>` : `<div>PDF<span>#${e.index}</span></div>`);
  }
  gal.innerHTML = cells.join('') || '<div class="muted" style="grid-column:1/-1">אין קבלות</div>';
  $('rvOk').onchange = () => $('rvSign').disabled = !$('rvOk').checked;
  $('rvDraft').onclick = async () => {
    spin(true, 'מכין טיוטה…');
    try { const b = await buildDocx(repKey, false); downloadBlob(b, fileNames(repKey).draft); }
    catch (err) { console.error(err); alert('שגיאה ביצירת הדוח: ' + err.message); }
    spin(false);
  };
  $('rvSign').onclick = () => signAndPrepare();
  $('review').classList.add('on'); $('review').scrollTop = 0;
}
$('btnReview').onclick = openReview;
$('rvClose').onclick = () => { $('review').classList.remove('on'); renderReport(); };

function fileNames(key) {
  const n = (S.name || 'עובד').replace(/[\\/:*?"<>|]/g, '').trim().replace(/\s+/g, '_');
  return {docx: `דוח_הוצאות_${n}_${key}.docx`, draft: `טיוטה_דוח_הוצאות_${key}.docx`, pdf: `קבלות_${n}_${key}.pdf`};
}

/* ----- DOCX ----- */
async function getTemplate() {
  const t = await DB.get('files', 'template');
  if (t) return await t.blob.arrayBuffer();
  const r = await fetch('template-default.docx'); return await r.arrayBuffer();
}
async function buildDocx(key, signed) {
  const items = reportItems(key), r = periodRange(key), total = items.reduce((s, e) => s + Number(e.amount), 0);
  const zip = new PizZip(await getTemplate());
  const doc = new window.docxtemplater(zip, {paragraphLoop: true, linebreaks: true, nullGetter: () => ''});
  const today = fmtDate(todayISO());
  doc.render({
    companyName: S.company, employeeName: S.name, employeeId: S.empId, department: S.dept, notes: S.notes,
    periodLabel: periodLabel(key), periodStart: fmtDate(r.start), periodEnd: fmtDate(r.end), reportDate: today,
    signDate: signed ? today : '', total: num2(total), itemCount: String(items.length),
    items: items.map(e => ({
      index: String(e.index), date: fmtDate(e.date), category: e.category,
      description: (e.description || '') + (isAshal(e) ? ` (${e.hours} שעות)` : ''),
      receipt: isAshal(e) ? 'לא נדרש' : (e.receipts?.length ? 'מצורפת' : 'חסרה'), amount: num2(e.amount)
    })),
    byCategory: byCategory(items).map(c => ({name: c.name, count: String(c.count), sum: num2(c.sum)}))
  });
  const out = doc.getZip();
  if (signed) {
    const sig = await DB.get('files', 'signature');
    if (!sig) throw new Error('חסרה חתימה');
    const ok = await injectSignature(out, sig.blob);
    if (!ok) throw new Error('לא נמצא מקום לחתימה בתבנית (תמונה עם Alt Text = SIGNATURE)');
  }
  return out.generate({type: 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', compression: 'DEFLATE'});
}
async function injectSignature(zip, sigBlob) {
  const path = 'word/document.xml';
  let xml = zip.file(path).asText();
  const m = /<wp:docPr\b[^>]*(?:descr|title|name)="SIGNATURE"[^>]*>/i.exec(xml);
  if (!m) return false;
  const dStart = xml.lastIndexOf('<w:drawing>', m.index), dEnd = xml.indexOf('</w:drawing>', m.index);
  if (dStart < 0 || dEnd < 0) return false;
  let seg = xml.slice(dStart, dEnd);
  const em = /r:embed="([^"]+)"/.exec(seg); if (!em) return false;
  const rels = zip.file('word/_rels/document.xml.rels').asText();
  const tm = new RegExp(`<Relationship[^>]*Id="${em[1]}"[^>]*>`).exec(rels); if (!tm) return false;
  let target = /Target="([^"]+)"/.exec(tm[0])[1];
  const mediaPath = target.startsWith('/') ? target.slice(1) : 'word/' + target.replace(/^\.\//, '');
  // convert signature to PNG of correct aspect
  const bmp = await loadBitmap(sigBlob);
  const c = document.createElement('canvas'); c.width = bmp.width; c.height = bmp.height; c.getContext('2d').drawImage(bmp, 0, 0);
  const png = await new Promise(res => c.toBlob(res, 'image/png'));
  const pngBuf = new Uint8Array(await png.arrayBuffer());
  if (/\.png$/i.test(mediaPath)) zip.file(mediaPath, pngBuf, {binary: true});
  else { // non-png placeholder: add new png media and repoint relationship
    const newT = 'media/signature_' + Date.now() + '.png';
    zip.file('word/' + newT, pngBuf, {binary: true});
    zip.file('word/_rels/document.xml.rels', rels.replace(tm[0], tm[0].replace(/Target="[^"]+"/, `Target="${newT}"`)));
    let ct = zip.file('[Content_Types].xml').asText();
    if (!/Extension="png"/i.test(ct)) ct = ct.replace('<Types ', '<Types ').replace(/(<Types[^>]*>)/, '$1<Default Extension="png" ContentType="image/png"/>');
    zip.file('[Content_Types].xml', ct);
  }
  // keep width, fix height to signature aspect ratio
  const ex = /<wp:extent cx="(\d+)" cy="(\d+)"\/>/.exec(seg);
  if (ex) {
    const cx = Number(ex[1]), cy = Math.round(cx * bmp.height / bmp.width);
    seg = seg.replace(/<wp:extent cx="\d+" cy="\d+"\/>/, `<wp:extent cx="${cx}" cy="${cy}"/>`)
             .replace(/<a:ext cx="\d+" cy="\d+"\/>/, `<a:ext cx="${cx}" cy="${cy}"/>`);
    xml = xml.slice(0, dStart) + seg + xml.slice(dEnd);
    zip.file(path, xml);
  }
  return true;
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
    const docx = await buildDocx(repKey, true);
    const pdf = await buildReceiptsPdf(repKey);
    OUT = {key: repKey, files: [new File([docx], names.docx, {type: docx.type})]};
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
    await saveSettingsObj(); $('review').classList.remove('on'); show('report'); toast('הדוח סומן כנשלח ✓');
  };
}

/* ---------------- settings ---------------- */
async function renderSettings() {
  $('sName').value = S.name; $('sEmpId').value = S.empId; $('sDept').value = S.dept; $('sCompany').value = S.company;
  $('sAcctEmail').value = S.acctEmail; $('sAcctName').value = S.acctName; $('sNotes').value = S.notes;
  $('sAshalMode').value = S.ashalMode; $('sAshalRate').value = S.ashalRate || '';
  renderCats(); renderTiers(); onAshalMode();
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
  $('tplStatus').textContent = t ? `תבנית מותאמת: ${t.name || 'template.docx'}` : 'תבנית ברירת מחדל (עד שתעלה את התבנית של הנה"ח)';
}
function readSettingsForm() {
  S.name = $('sName').value.trim(); S.empId = $('sEmpId').value.trim(); S.dept = $('sDept').value.trim(); S.company = $('sCompany').value.trim();
  S.acctEmail = $('sAcctEmail').value.trim(); S.acctName = $('sAcctName').value.trim(); S.notes = $('sNotes').value.trim();
  S.ashalMode = $('sAshalMode').value; S.ashalRate = Number($('sAshalRate').value) || 0;
  S.categories = [...document.querySelectorAll('#catList .catrow')].map(r => ({name: r.querySelector('input[type=text]').value.trim(), ashal: r.querySelector('input[type=checkbox]').checked})).filter(c => c.name);
  S.ashalTiers = [...document.querySelectorAll('#tierList .catrow')].map(r => { const [a, b] = r.querySelectorAll('input'); return {from: Number(a.value) || 0, amount: Number(b.value) || 0}; });
}
function renderCats() {
  $('catList').innerHTML = S.categories.map((c, i) => `<div class="catrow"><input type="text" value="${esc(c.name)}">
    <label style="margin:0;display:flex;align-items:center;white-space:nowrap;font-size:12px"><input type="checkbox" ${c.ashal ? 'checked' : ''}>א.ש.ל</label>
    <button class="sm" data-i="${i}">✕</button></div>`).join('');
  $('catList').querySelectorAll('.sm').forEach(b => b.onclick = () => { readSettingsForm(); S.categories.splice(b.dataset.i, 1); renderCats(); saveSettingsObj(); });
}
function renderTiers() {
  $('tierList').innerHTML = S.ashalTiers.map((t, i) => `<div class="catrow"><input type="number" inputmode="decimal" value="${t.from}" placeholder="משעה">
    <input type="number" inputmode="decimal" value="${t.amount || ''}" placeholder="₪"><button class="sm" data-i="${i}">✕</button></div>`).join('');
  $('tierList').querySelectorAll('.sm').forEach(b => b.onclick = () => { readSettingsForm(); S.ashalTiers.splice(b.dataset.i, 1); renderTiers(); saveSettingsObj(); });
}
function onAshalMode() { const h = $('sAshalMode').value === 'hourly'; $('ashalHourly').classList.toggle('hidden', !h); $('ashalTiers').classList.toggle('hidden', h); }
$('sAshalMode').onchange = onAshalMode;
let _asT;
$('view-settings').addEventListener('input', ev => {
  if (ev.target.type === 'file') return;
  clearTimeout(_asT); _asT = setTimeout(() => { readSettingsForm(); saveSettingsObj(); }, 500);
});
$('view-settings').addEventListener('change', ev => { if (ev.target.type !== 'file') { readSettingsForm(); saveSettingsObj(); } });
$('btnAddCat').onclick = () => { readSettingsForm(); S.categories.push({name: '', ashal: false}); renderCats(); $('catList').lastChild.querySelector('input').focus(); };
$('btnAddTier').onclick = () => { readSettingsForm(); S.ashalTiers.push({from: 0, amount: 0}); renderTiers(); };
$('btnSaveSettings').onclick = async () => {
  readSettingsForm();
  if (S.acctEmail && !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(S.acctEmail)) return toast('כתובת מייל לא תקינה');
  await saveSettingsObj(); toast('ההגדרות נשמרו ✓');
};

// template
$('btnTplUpload').onclick = () => $('tplFile').click();
$('tplFile').onchange = async ev => {
  const f = ev.target.files[0]; ev.target.value = ''; if (!f) return;
  try {
    const zip = new PizZip(await f.arrayBuffer());
    const xml = zip.file('word/document.xml').asText();
    new window.docxtemplater(zip, {paragraphLoop: true, linebreaks: true}); // validates tags
    await DB.put('files', {id: 'template', blob: f, type: f.type, name: f.name});
    const hasSig = /(descr|title|name)="SIGNATURE"/i.test(xml);
    const hasTags = /\{[#\/]?\w+\}/.test(xml.replace(/<[^>]+>/g, ''));
    renderSigTpl(); toast(`התבנית נשמרה${hasTags ? '' : ' – ⚠️ לא נמצאו שדות {…}'}${hasSig ? '' : ' – ⚠️ אין מקום לחתימה'}`, 5000);
  } catch (err) { console.error(err); alert('התבנית לא תקינה: ' + (err.properties?.errors?.map(e => e.properties.explanation).join('\n') || err.message)); }
};
$('btnTplReset').onclick = async () => { await DB.del('files', 'template'); renderSigTpl(); toast('חזרה לתבנית ברירת המחדל'); };

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
  $('btnSigDraw').onclick = () => { $('sigSheet').classList.add('on'); requestAnimationFrame(setup); };
  $('sigClose').onclick = () => $('sigSheet').classList.remove('on');
  $('sigClear').onclick = setup;
  $('sigSave').onclick = async () => {
    if (!dirty) return toast('יש לחתום קודם');
    const raw = await new Promise(res => cv.toBlob(res, 'image/png'));
    const png = await cleanSignature(raw);
    await DB.put('files', {id: 'signature', blob: png, type: 'image/png'});
    $('sigSheet').classList.remove('on'); renderSigTpl(); toast('החתימה נשמרה');
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
    for (const f of await DB.all('files')) files.push({id: f.id, type: f.type, name: f.name, data: b64(await f.blob.arrayBuffer())});
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
    for (const x of data.files) await DB.put('files', {id: x.id, type: x.type, name: x.name, blob: new Blob([unb64(x.data)], {type: x.type})});
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
    await DB.open(); await loadSettings(); await loadExpenses();
    try { await navigator.storage?.persist?.(); } catch {}
    renderExpenses();
    if (!S.name) setTimeout(() => toast('ברוך הבא! התחל בהגדרות: שם, מייל הנה"ח וחתימה', 4500), 400);
  } catch (err) { console.error(err); document.body.insertAdjacentHTML('afterbegin', `<div class="banner">שגיאה בטעינת האחסון: ${esc(err.message)}</div>`); }
  if ('serviceWorker' in navigator && location.protocol === 'https:') navigator.serviceWorker.register('sw.js').catch(() => {});
})();
