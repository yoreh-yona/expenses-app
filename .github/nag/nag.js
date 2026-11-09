// Hourly nag: Web Push to the phone from the 25th at START_HOUR (Israel time) every INTERVAL hours until the report is marked sent.
// Inputs come from repository variables written by the app: NAG_CONFIG {sub, vapidPublic, vapidPrivate, interval, startHour, endHour}
// and NAG_STATE {pending: ["YYYY-MM", ...]} (periods that have expenses and were not marked sent).
const fs = require('fs');
const MONTHS = ['ינואר','פברואר','מרץ','אפריל','מאי','יוני','יולי','אוגוסט','ספטמבר','אוקטובר','נובמבר','דצמבר'];
const label = k => { const [y, m] = k.split('-').map(Number); return `${MONTHS[m - 1]} ${y}`; };
const parse = s => { try { return JSON.parse(s || 'null'); } catch { return null; } };

function israelNow(d = new Date()) {
  const p = Object.fromEntries(new Intl.DateTimeFormat('en-GB', {timeZone: 'Asia/Jerusalem', year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', hourCycle: 'h23'})
    .formatToParts(d).map(x => [x.type, x.value]));
  return {y: +p.year, m: +p.month, d: +p.day, h: +p.hour, key: `${p.year}-${p.month}-${p.day}`};
}

function decide(cfg, st, now, test) {
  if (!cfg || !cfg.sub) return {send: false, why: 'not configured'};
  const start = cfg.startHour ?? 9, end = cfg.endHour ?? 21, every = Math.max(1, cfg.interval ?? 2);
  const today = `${now.y}-${String(now.m).padStart(2, '0')}-${String(now.d).padStart(2, '0')}`;
  const overdue = (st?.pending || []).filter(p => today > `${p}-25` || (today === `${p}-25` && now.h >= start)).sort();
  if (test) return {send: true, overdue, why: 'test'};
  if (!overdue.length) return {send: false, why: 'nothing due'};
  if (now.h < start || now.h > end) return {send: false, why: 'quiet hours'};
  if ((now.h - start) % every) return {send: false, why: 'not this hour'};
  return {send: true, overdue};
}

function message(overdue, now) {
  if (!overdue.length) return {title: 'דוח הוצאות – בדיקה', body: 'ההתראות עובדות ✓'};
  const cur = overdue[overdue.length - 1];
  const isToday = now.d === 25 && `${now.y}-${String(now.m).padStart(2, '0')}` === cur && overdue.length === 1;
  return isToday
    ? {title: `היום ה-25: דוח הוצאות ${label(cur)}`, body: 'הדוח עוד לא נשלח להנהלת החשבונות. לחץ כדי לעבור לדוח.'}
    : {title: `דוח הוצאות ${overdue.map(label).join(', ')} – באיחור`, body: 'הדוח עוד לא נשלח. לחץ כדי לעבור לדוח.'};
}

async function main() {
  const cfg = parse(process.env.NAG_CONFIG), st = parse(process.env.NAG_STATE), test = process.env.TEST === 'true';
  const now = israelNow(process.env.NOW ? new Date(process.env.NOW) : new Date());
  const r = decide(cfg, st, now, test);
  console.log(`Israel ${now.key} ${now.h}:00 →`, r.send ? 'SEND' : 'skip', r.why || '', JSON.stringify(r.overdue || []));
  if (process.argv.includes('--decide')) { if (process.env.GITHUB_OUTPUT) fs.appendFileSync(process.env.GITHUB_OUTPUT, `send=${r.send}\n`); return; }
  if (!r.send) return;
  const webpush = require('web-push');
  webpush.setVapidDetails('https://yoreh-yona.github.io/expenses-app/', cfg.vapidPublic, cfg.vapidPrivate);
  const msg = {...message(r.overdue, now), url: './#report'};
  try {
    const res = await webpush.sendNotification(cfg.sub, JSON.stringify(msg), {TTL: 3 * 3600, urgency: 'high'});
    console.log('sent', res.statusCode);
  } catch (e) {
    console.error('push failed', e.statusCode, e.body);
    if (e.statusCode === 404 || e.statusCode === 410) console.error('Subscription expired – re-enable notifications in the app (Settings → Notifications).');
    process.exit(1);
  }
}
module.exports = {decide, message, israelNow};
if (require.main === module) main();
