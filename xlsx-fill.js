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
      const h = 40, w = Math.min(160, Math.round(h * sig.width / sig.height));
      // sit on the signature line: start half a row above the value cell
      ws.addImage(id, {tl: {col: L.signature.col - 1 + 0.15, row: L.signature.row - 1 - 0.55}, ext: {width: w, height: h}, editAs: 'oneCell'});
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
