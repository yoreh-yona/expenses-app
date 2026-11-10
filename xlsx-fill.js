/* Fills the accounting department's Excel expense form.
 * Layout is detected from the template itself (labels + header row), so a revised form with
 * the same labels keeps working. Shared by the app (browser) and the test harness (node). */
(function (root) {
  'use strict';
  const norm = v => String(v ?? '').replace(/[\s ]+/g, ' ').trim();
  const stripCode = s => norm(s).replace(/\(\s*\d+\s*\)/g, '').replace(/\s+/g, ' ').trim();
  const cellText = c => {
    const v = c.value;
    if (v && typeof v === 'object' && v.richText) return v.richText.map(t => t.text).join('');
    if (v && typeof v === 'object' && 'formula' in v) return '';
    return v == null ? '' : String(v);
  };
  const colLetter = n => { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = (n - m - 1) / 26; } return s; };

  function findLabel(ws, test) {
    let hit = null;
    ws.eachRow({includeEmpty: false}, (row, r) => row.eachCell({includeEmpty: false}, (c, col) => {
      if (!hit && test(norm(cellText(c)))) hit = {row: r, col};
    }));
    return hit;
  }

  /** Analyse a template worksheet; returns layout or throws with a Hebrew message. */
  function analyse(ws) {
    const hdr = findLabel(ws, t => t === 'תאריך');
    if (!hdr) throw new Error('לא נמצאה שורת כותרת עם "תאריך"');
    const L = {headerRow: hdr.row, dateCol: hdr.col, descCol: hdr.col + 1, categories: [], totalCol: null, invoiceCol: null};
    const row = ws.getRow(hdr.row);
    const maxCol = Math.max(ws.columnCount, 30);
    for (let c = hdr.col + 2; c <= maxCol; c++) {
      const cell = row.getCell(c);
      if (cell.isMerged && cell.master.address !== cell.address) continue;
      const t = norm(cellText(cell)); if (!t) continue;
      if (t.includes('סך הכל')) { L.totalCol = c; continue; }
      if (t.includes('חשבונית')) { L.invoiceCol = c; continue; }
      if (!L.totalCol) L.categories.push({name: stripCode(t), header: t, col: c});
    }
    if (!L.totalCol || !L.categories.length) throw new Error('לא זוהו עמודות סוגי ההוצאות / סך הכל');
    // data rows: from header+1 until the totals row (first category column holds SUM over a range)
    let r = hdr.row + 1;
    while (r < hdr.row + 500) {
      const v = ws.getRow(r).getCell(L.categories[0].col).value;
      if (v && typeof v === 'object' && v.formula && /:/.test(v.formula)) break;
      r++;
    }
    L.firstData = hdr.row + 1; L.lastData = r - 1; L.totalsRow = r;
    const lab = (txt, exact) => findLabel(ws, t => exact ? t === txt : t.startsWith(txt));
    const val = h => h ? {row: h.row, col: h.col + 1} : null;
    L.fields = {
      month: val(lab('דוח לחודש')), name: val(lab('שם', true)), role: val(lab('תפקיד')),
      company: val(lab('חברה', true)), manager: val(lab('מנהל מאשר')), notes: val(lab('הערות')),
      advances: val(lab('מקדמות'))
    };
    L.signature = val(lab('חתימת עובד'));
    L.subtotal = val(lab('סיכום ביניים'));
    L.grandTotal = (() => { // 'סך הכל' label below the table (not the header)
      let hit = null;
      ws.eachRow((rw, rr) => { if (rr > L.totalsRow) rw.eachCell((c, cc) => { if (!hit && norm(cellText(c)) === 'סך הכל') hit = {row: rr, col: cc + 1}; }); });
      return hit;
    })();
    return L;
  }

  /**
   * data: {name, role, company, manager, notes, monthDate:Date(UTC), items:[{date:Date(UTC), description, category, amount, invoice}]}
   * sig: optional {buffer (Uint8Array/ArrayBuffer of PNG), width, height}
   */
  async function fill(ExcelJS, PizZip, templateBuf, data, sig) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(templateBuf);
    const ws = wb.worksheets[0];
    let L = analyse(ws);
    const items = data.items;
    const capacity = L.lastData - L.firstData + 1;
    if (items.length > capacity) {
      const extra = items.length - capacity;
      ws.duplicateRow(L.lastData, extra, true);
      L = analyse(ws);
    }
    const setF = (rc, v) => { if (rc && v != null && v !== '') ws.getRow(rc.row).getCell(rc.col).value = v; };
    setF(L.fields.month, data.monthDate);
    if (L.fields.month) ws.getRow(L.fields.month.row).getCell(L.fields.month.col).numFmt = 'mm/yyyy';
    setF(L.fields.name, data.name); setF(L.fields.role, data.role);
    setF(L.fields.company, data.company); setF(L.fields.manager, data.manager);
    setF(L.fields.notes, data.notes);

    const catCol = name => {
      const n = stripCode(name);
      return (L.categories.find(c => c.name === n) || L.categories.find(c => c.name.includes('שונות')) || L.categories[L.categories.length - 1]).col;
    };
    const first = colLetter(L.categories[0].col), last = colLetter(L.categories[L.categories.length - 1].col);
    const colSums = {};
    let grand = 0;
    for (let r = L.firstData; r <= L.lastData; r++) {
      const row = ws.getRow(r);
      const it = items[r - L.firstData];
      // clear previous values in the data area (keep styles)
      row.getCell(L.dateCol).value = null; row.getCell(L.descCol).value = null;
      L.categories.forEach(c => { row.getCell(c.col).value = null; });
      if (L.invoiceCol) row.getCell(L.invoiceCol).value = null;
      let rowSum = 0;
      if (it) {
        row.getCell(L.dateCol).value = it.date; row.getCell(L.dateCol).numFmt = 'dd/mm/yyyy';
        row.getCell(L.descCol).value = it.description;
        const cc = catCol(it.category);
        row.getCell(cc).value = it.amount;
        colSums[cc] = (colSums[cc] || 0) + it.amount;
        rowSum = it.amount; grand += it.amount;
        if (L.invoiceCol && it.invoice != null) row.getCell(L.invoiceCol).value = it.invoice;
      }
      row.getCell(L.totalCol).value = {formula: `SUM(${first}${r}:${last}${r})`, result: rowSum};
    }
    const tr = ws.getRow(L.totalsRow);
    L.categories.forEach(c => {
      const cl = colLetter(c.col);
      tr.getCell(c.col).value = {formula: `SUM(${cl}${L.firstData}:${cl}${L.lastData})`, result: colSums[c.col] || 0};
    });
    const tc = colLetter(L.totalCol);
    if (L.subtotal) ws.getRow(L.subtotal.row).getCell(L.subtotal.col).value = {formula: `SUM(${tc}${L.firstData}:${tc}${L.lastData})`, result: grand};
    if (L.grandTotal && L.subtotal && L.fields.advances) {
      const adv = Number(data.advances) || 0;
      if (adv) setF(L.fields.advances, adv);
      ws.getRow(L.grandTotal.row).getCell(L.grandTotal.col).value =
        {formula: `(${tc}${L.subtotal.row}-${colLetter(L.fields.advances.col)}${L.fields.advances.row})`, result: grand - adv};
    }
    if (sig && L.signature) {
      const id = wb.addImage({buffer: sig.buffer, extension: 'png'});
      const r = L.signature.row, c = L.signature.col, cell = ws.getRow(r).getCell(c);
      const colPx = k => Math.round((ws.getColumn(k).width ?? 8.43) * 7 + 5), rowPx = k => Math.round((ws.getRow(k).height ?? 15) * 4 / 3);
      let c2 = c; // the signature cell is usually merged across a few columns
      while (c2 < c + 10) { const nx = ws.getRow(r).getCell(c2 + 1); if (nx.isMerged && nx.master.address === cell.address) c2++; else break; }
      let boxW = 0; for (let k = c; k <= c2; k++) boxW += colPx(k);
      const aspect = sig.width / sig.height;
      let h = 50, w = h * aspect;
      if (w > boxW * 0.92) { w = boxW * 0.92; h = w / aspect; }
      // centred over the signature line, crossing it slightly like a hand signature
      let col = c, xo = (boxW - w) / 2;
      while (xo >= colPx(col) && col < c2) { xo -= colPx(col); col++; }
      let row = r, yo = rowPx(r) + 12 - h;
      while (yo < 0 && row > 1) { row--; yo += rowPx(row); }
      while (yo >= rowPx(row)) { yo -= rowPx(row); row++; }
      ws.addImage(id, {tl: {nativeCol: col - 1, nativeColOff: Math.round(xo * 9525), nativeRow: row - 1, nativeRowOff: Math.round(yo * 9525)},
        ext: {width: Math.round(w), height: Math.round(h)}, editAs: 'oneCell'});
    }
    wb.calcProperties = Object.assign(wb.calcProperties || {}, {fullCalcOnLoad: true});
    let out = await wb.xlsx.writeBuffer();
    // ExcelJS drops a custom indexed-colour palette; copy it back from the template so fills keep their colours
    try {
      const src = new PizZip(templateBuf).file('xl/styles.xml').asText();
      const m = /<colors>[\s\S]*?<\/colors>/.exec(src);
      if (m) {
        const z = new PizZip(out); let st = z.file('xl/styles.xml').asText();
        if (!/<colors>/.test(st)) {
          st = /<extLst/.test(st) ? st.replace(/<extLst/, m[0] + '<extLst') : st.replace('</styleSheet>', m[0] + '</styleSheet>');
          z.file('xl/styles.xml', st);
          out = z.generate({type: 'uint8array', compression: 'DEFLATE'});
        }
      }
    } catch (e) { console.warn('palette restore failed', e); }
    return {buffer: out, layout: L};
  }

  async function inspect(ExcelJS, templateBuf) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(templateBuf);
    return analyse(wb.worksheets[0]);
  }

  root.ExpenseXlsx = {fill, inspect, stripCode};
})(typeof window !== 'undefined' ? window : globalThis);

/* Print-preview: render the first worksheet (cells, merges, fills, borders, fonts, images) as an SVG image. */
(function (root) {
  'use strict';
  const DEF_PAL = ['000000','FFFFFF','FF0000','00FF00','0000FF','FFFF00','FF00FF','00FFFF','000000','FFFFFF','FF0000','00FF00','0000FF','FFFF00','FF00FF','00FFFF',
    '800000','008000','000080','808000','800080','008080','C0C0C0','808080','9999FF','993366','FFFFCC','CCFFFF','660066','FF8080','0066CC','CCCCFF',
    '000080','FF00FF','FFFF00','00FFFF','800080','800000','008080','0000FF','00CCFF','CCFFFF','CCFFCC','FFFF99','99CCFF','FF99CC','CC99FF','FFCC99',
    '3366FF','33CCCC','99CC00','FFCC00','FF9900','FF6600','666699','969696','003366','339966','003300','333300','993300','993366','333399','333333'];
  const esc = s => String(s).replace(/[&<>"]/g, c => ({'&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;'}[c]));
  const toB64 = u8 => { let s = ''; for (let i = 0; i < u8.length; i += 0x8000) s += String.fromCharCode.apply(null, u8.subarray(i, i + 0x8000)); return btoa(s); };

  function palettes(PizZip, buf) {
    const pal = DEF_PAL.slice(), theme = [];
    try {
      const z = new PizZip(buf);
      const st = z.file('xl/styles.xml')?.asText() || '';
      const ic = /<indexedColors>([\s\S]*?)<\/indexedColors>/.exec(st);
      if (ic) [...ic[1].matchAll(/rgb="([0-9A-Fa-f]{8})"/g)].forEach((m, i) => { pal[i] = m[1].slice(2); });
      const th = z.file('xl/theme/theme1.xml')?.asText() || '';
      const cs = /<a:clrScheme[^>]*>([\s\S]*?)<\/a:clrScheme>/.exec(th);
      if (cs) {
        const get = n => { const m = new RegExp(`<a:${n}>[\\s\\S]*?(?:srgbClr val="([0-9A-Fa-f]{6})"|lastClr="([0-9A-Fa-f]{6})")`).exec(cs[1]); return m ? (m[1] || m[2]) : '000000'; };
        ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'].forEach(n => theme.push(get(n)));
      }
    } catch {}
    return {pal, theme};
  }
  function color(c, P, fallback) {
    if (!c) return fallback;
    let hex = null;
    if (c.argb) hex = c.argb.length === 8 ? c.argb.slice(2) : c.argb;
    else if (c.indexed != null) { if (c.indexed === 64 || c.indexed === 65) return fallback; hex = P.pal[c.indexed]; }
    else if (c.theme != null) hex = P.theme[c.theme] || null;
    if (!hex) return fallback;
    let r = parseInt(hex.slice(0, 2), 16), g = parseInt(hex.slice(2, 4), 16), b = parseInt(hex.slice(4, 6), 16);
    const t = c.tint || 0;
    const f = v => Math.round(t > 0 ? v + (255 - v) * t : v * (1 + t));
    if (t) { r = f(r); g = f(g); b = f(b); }
    return `rgb(${r},${g},${b})`;
  }
  const BORDER = {thin: '1px solid', hair: '1px solid', medium: '2px solid', thick: '3px solid', double: '3px double', dotted: '1px dotted', dashed: '1px dashed',
    mediumDashed: '2px dashed', dashDot: '1px dashed', mediumDashDot: '2px dashed', dashDotDot: '1px dotted', slantDashDot: '2px dashed'};
  const pad = n => String(n).padStart(2, '0');

  function fmtValue(v, numFmt) {
    if (v && typeof v === 'object' && !(v instanceof Date)) {
      if (v.richText) return v.richText.map(t => t.text).join('');
      if ('result' in v || 'formula' in v) v = v.result;
      else if (v.text) return v.text;
      else if (v.error) return v.error;
    }
    if (v == null || v === '') return '';
    const f = String(numFmt || '');
    if (v instanceof Date) {
      const d = v.getUTCDate(), m = v.getUTCMonth() + 1, y = v.getUTCFullYear();
      if (/d/i.test(f.replace(/\[[^\]]*\]/g, '')) ) return `${pad(d)}/${pad(m)}/${y}`;
      return `${pad(m)}/${y}`;
    }
    if (typeof v === 'number') {
      const dec = /0\.(0+)/.exec(f), n = dec ? dec[1].length : (Number.isInteger(v) ? 0 : 2);
      const s = Math.abs(v).toLocaleString('en-US', {minimumFractionDigits: n, maximumFractionDigits: n, useGrouping: /#,##/.test(f) || Math.abs(v) >= 1000});
      if (v === 0 && /;\s*"?-?"?\s*;/.test(f) && /-/.test(f.split(';')[2] || '')) return '-';
      return (v < 0 ? '-' : '') + s;
    }
    return String(v);
  }
  const colToPx = w => Math.round((w == null ? 8.43 : w) * 7 + 5);
  const rowToPx = h => Math.round((h == null ? 15 : h) * 4 / 3);

  async function toSvg(ExcelJS, PizZip, buf) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buf);
    const ws = wb.worksheets[0], P = palettes(PizZip, buf);
    const rtl = !!(ws.views && ws.views[0] && ws.views[0].rightToLeft);
    // range: print area if set, else used range
    let r0 = 1, c0 = 1, r1 = ws.rowCount, c1 = ws.columnCount;
    const pa = ws.pageSetup && ws.pageSetup.printArea;
    const m = pa && /\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)/.exec(pa.split(',')[0]);
    const colNum = s => s.split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
    if (m) { c0 = colNum(m[1]); r0 = +m[2]; c1 = colNum(m[3]); r1 = +m[4]; }
    const cols = [], rows = [];
    for (let c = c0; c <= c1; c++) { const col = ws.getColumn(c); cols.push(col.hidden ? 0 : colToPx(col.width)); }
    for (let r = r0; r <= r1; r++) { const row = ws.getRow(r); rows.push(row.hidden ? 0 : rowToPx(row.height)); }
    const W = cols.reduce((a, b) => a + b, 0), H = rows.reduce((a, b) => a + b, 0);
    // merges
    const span = {}, covered = new Set();
    Object.values(ws._merges || {}).forEach(mg => {
      const {top, left, bottom, right} = mg.model || mg;
      span[`${top}:${left}`] = {rs: bottom - top + 1, cs: right - left + 1};
      for (let r = top; r <= bottom; r++) for (let c = left; c <= right; c++) if (r !== top || c !== left) covered.add(`${r}:${c}`);
    });
    let html = `<table style="border-collapse:collapse;table-layout:fixed;width:${W}px;direction:${rtl ? 'rtl' : 'ltr'};font-family:Arial,'Segoe UI',sans-serif">`;
    html += '<colgroup>' + cols.map(w => `<col style="width:${w}px"/>`).join('') + '</colgroup>';
    for (let r = r0; r <= r1; r++) {
      const row = ws.getRow(r);
      html += `<tr style="height:${rows[r - r0]}px">`;
      for (let c = c0; c <= c1; c++) {
        if (covered.has(`${r}:${c}`)) continue;
        const cell = row.getCell(c), st = cell.style || {}, sp = span[`${r}:${c}`];
        const css = ['padding:0 3px', 'overflow:hidden', 'box-sizing:border-box'];
        const fill = st.fill && st.fill.type === 'pattern' && st.fill.pattern !== 'none' ? color(st.fill.fgColor, P, null) : null;
        css.push('background:' + (fill || '#fff'));
        const f = st.font || {};
        css.push(`font-size:${Math.round((f.size || 11) * 4 / 3)}px`, `color:${color(f.color, P, '#000')}`);
        if (f.bold) css.push('font-weight:bold'); if (f.italic) css.push('font-style:italic'); if (f.underline) css.push('text-decoration:underline');
        const b = st.border || {}, side = (k, cssSide) => { const s = b[k] && BORDER[b[k].style]; if (s) css.push(`border-${cssSide}:${s} ${color(b[k].color, P, '#000')}`); };
        side('top', 'top'); side('bottom', 'bottom'); side('left', rtl ? 'right' : 'left'); side('right', rtl ? 'left' : 'right');
        const al = st.alignment || {}, v = cell.value;
        const isNum = typeof v === 'number' || v instanceof Date || (v && typeof v === 'object' && typeof v.result === 'number');
        let h = al.horizontal || (isNum ? 'right' : 'general');
        if (h === 'general') h = rtl ? 'right' : 'left';
        if (h === 'centerContinuous') h = 'center';
        css.push('text-align:' + (h === 'fill' || h === 'justify' || h === 'distributed' ? 'center' : h));
        css.push('vertical-align:' + (al.vertical === 'top' ? 'top' : al.vertical === 'middle' || al.vertical === 'center' ? 'middle' : 'bottom'));
        css.push(al.wrapText ? 'white-space:normal;word-break:break-word;line-height:1.15' : 'white-space:nowrap');
        const txt = fmtValue(v, st.numFmt);
        html += `<td${sp ? ` rowspan="${sp.rs}" colspan="${sp.cs}"` : ''} style="${css.join(';')}">${esc(txt)}</td>`;
      }
      html += '</tr>';
    }
    html += '</table>';
    // images (e.g. the signature)
    let imgs = '';
    const colX = x => { const i = Math.floor(x) - (c0 - 1); let px = 0; for (let k = 0; k < Math.min(i, cols.length); k++) px += cols[k]; return px + (x - Math.floor(x)) * (cols[i] || 0); };
    const rowY = y => { const i = Math.floor(y) - (r0 - 1); let px = 0; for (let k = 0; k < Math.min(i, rows.length); k++) px += rows[k]; return px + (y - Math.floor(y)) * (rows[i] || 0); };
    for (const im of ws.getImages()) {
      const media = wb.getImage(Number(im.imageId)); if (!media || !media.buffer) continue;
      const tl = im.range.tl, ext = im.range.ext || {};
      const sumTo = (arr, n) => { let p = 0; for (let k = 0; k < Math.min(n, arr.length); k++) p += arr[k]; return p; };
      const x = tl.nativeCol != null ? sumTo(cols, tl.nativeCol - (c0 - 1)) + (tl.nativeColOff || 0) / 9525 : colX(tl.col);
      const y = tl.nativeRow != null ? sumTo(rows, tl.nativeRow - (r0 - 1)) + (tl.nativeRowOff || 0) / 9525 : rowY(tl.row);
      const w = ext.width || 120, h = ext.height || 40;
      const left = rtl ? W - x - w : x;
      imgs += `<img src="data:image/${media.extension || 'png'};base64,${toB64(new Uint8Array(media.buffer))}" style="position:absolute;left:${left}px;top:${y}px;width:${w}px;height:${h}px"/>`;
    }
    const M = 24, SW = W + 2 * M, SH = H + 2 * M;
    const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${SW}" height="${SH}" viewBox="0 0 ${SW} ${SH}">
<rect width="100%" height="100%" fill="#fff"/>
<foreignObject x="${M}" y="${M}" width="${W}" height="${H}"><div xmlns="http://www.w3.org/1999/xhtml" style="position:relative;width:${W}px;height:${H}px">${html}${imgs}</div></foreignObject></svg>`;
    return {svg, width: SW, height: SH};
  }
  root.ExpenseXlsx = Object.assign(root.ExpenseXlsx || {}, {toSvg});
})(typeof window !== 'undefined' ? window : globalThis);
