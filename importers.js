/* Roster Convert — extra input formats.
   PDF (Roster Report table printed to PDF) and iCalendar (.ics) are turned into the
   same rows/columns as the Excel export, so the rest of the app treats them identically.
   Everything runs locally; pdf.js is loaded from this site only when a PDF is opened. */
'use strict';

/** Bump on every release: added to file URLs so browsers fetch the new version instead of a cached one. */
const APP_VERSION = '2026.11.01-23';

/* ------------------------------------------------------------------ *
 *  PDF → array-of-arrays (same layout as the Excel Roster Report)
 * ------------------------------------------------------------------ */
let pdfjsReady = null;
function loadPdfJs(){
  return pdfjsReady ||= new Promise((resolve, reject) => {
    const s = document.createElement('script');
    s.src = 'vendor/pdf.min.js?v='+APP_VERSION;
    s.onload = () => { window.pdfjsLib.GlobalWorkerOptions.workerSrc = 'vendor/pdf.worker.min.js?v='+APP_VERSION; resolve(window.pdfjsLib); };
    s.onerror = () => { pdfjsReady = null; reject(new Error('PDF reader failed to load')); };
    document.head.appendChild(s);
  });
}

/** Text items of every page in top-down page coordinates (handles rotated pages). */
async function pdfItems(buf){
  const pdfjs = await loadPdfJs();
  const doc = await pdfjs.getDocument({data: new Uint8Array(buf), isEvalSupported: false, disableFontFace: true}).promise;
  const pages = [];
  for(let n = 1; n <= doc.numPages; n++){
    const page = await doc.getPage(n);
    const vp = page.getViewport({scale: 1});
    const tc = await page.getTextContent();
    pages.push(tc.items.filter(i => i.str && i.str.trim()).map(i => {
      const t = pdfjs.Util.transform(vp.transform, i.transform);
      const rotated = Math.abs(t[1]) > Math.abs(t[0]);          // text running vertically after transform
      return {s: i.str.trim(), x: t[4], y: t[5], w: rotated ? i.width : i.width * Math.hypot(t[0], t[1]) / Math.hypot(i.transform[0], i.transform[1])};
    }));
  }
  return pages;
}

/** Rebuild the Roster Report table from positioned PDF text. */
function pdfItemsToAoa(pages){
  const aoa = [];
  let headerDone = false;
  for(const items of pages){
    const dateHdr = items.find(i => i.s === 'Date' && items.some(j => /^Pairing/.test(j.s) && Math.abs(j.y - i.y) < 8));
    if(!dateHdr) continue;
    const hy = dateHdr.y;
    // Header labels are stacked on up to two lines around the "Date" baseline: group by horizontal overlap.
    const hdr = items.filter(i => Math.abs(i.y - hy) <= 9).sort((a, b) => a.x - b.x);
    const cols = [];
    for(const it of hdr){
      const c = cols.find(c => it.x < c.x1 + 1 && it.x + it.w > c.x0 - 1);
      if(c){ c.items.push(it); c.x0 = Math.min(c.x0, it.x); c.x1 = Math.max(c.x1, it.x + it.w); }
      else cols.push({x0: it.x, x1: it.x + it.w, items: [it]});
    }
    cols.forEach(c => { c.label = c.items.sort((a, b) => a.y - b.y || a.x - b.x).map(i => i.s).join(' '); c.cx = (c.x0 + c.x1) / 2; });
    cols.sort((a, b) => a.cx - b.cx);
    const colOf = it => { const cx = it.x + it.w / 2; let best = 0; cols.forEach((c, k) => { if(Math.abs(c.cx - cx) < Math.abs(cols[best].cx - cx)) best = k; }); return best; };
    const iDate = cols.findIndex(c => c.label === 'Date');
    const iDep = cols.findIndex(c => /^Dep/.test(c.label));

    if(!headerDone){
      // Report title / period / "NAME | STAFF | FLEET | BASE | RANK" line above the table.
      aoa.push(items.filter(i => i.y < hy - 9).sort((a, b) => a.y - b.y || a.x - b.x).map(i => i.s));
      aoa.push(cols.map(c => c.label));
      headerDone = true;
    }

    const data = items.filter(i => i.y > hy + 9).map(i => Object.assign(i, {col: colOf(i)}));
    // Each record has one main baseline: where its Date and/or Dep Stn sits.
    const anchors = [];
    data.filter(i => i.col === iDate || i.col === iDep).forEach(i => { if(!anchors.some(a => Math.abs(a - i.y) < 1.5)) anchors.push(i.y); });
    anchors.sort((a, b) => a - b);
    if(!anchors.length) continue;
    const onLine = i => anchors.find(a => Math.abs(a - i.y) < 1.5);
    const nearest = y => anchors.reduce((b, a) => Math.abs(a - y) < Math.abs(b - y) ? a : b, anchors[0]);
    const rows = new Map(anchors.map(a => [a, cols.map(() => [])]));
    for(const it of data){
      let a = onLine(it);
      if(a === undefined){
        // Wrapped cell text (e.g. hotel over 3 lines): follow a same-column piece that sits on a baseline.
        const buddy = data.filter(j => j.col === it.col && onLine(j) !== undefined && Math.abs(j.y - it.y) <= 13)
                          .sort((p, q) => Math.abs(p.y - it.y) - Math.abs(q.y - it.y))[0];
        a = buddy ? onLine(buddy) : nearest(it.y);
      }
      rows.get(a)[it.col].push(it);
    }
    anchors.forEach(a => aoa.push(rows.get(a).map(cell => cell.sort((p, q) => p.y - q.y || p.x - q.x).map(i => i.s).join(' '))));
  }
  if(!headerDone) throw new Error('This PDF doesn\'t look like a crew Roster Report (no "Date" / "Pairing/Activity" table found).');
  return aoa;
}

async function pdfToWorkbook(buf){
  const aoa = pdfItemsToAoa(await pdfItems(buf));
  return {SheetNames: ['PDF'], Sheets: {PDF: XLSX.utils.aoa_to_sheet(aoa)}};
}

/* ------------------------------------------------------------------ *
 *  iCalendar (.ics) → array-of-arrays
 *  Expected events (crew roster calendar export):
 *    all-day   "AL - KUL"                       → leave / off day
 *    timed     "Report time: 22:25 / Debrief…"  → duty period (report → debrief)
 *    timed     "MH 020 KUL - CDG"               → flight (dep → arr)
 *  Times without "Z" are treated like the Excel export (home-base clock).
 * ------------------------------------------------------------------ */
function parseIcsEvents(text){
  const lines = text.replace(/\r\n?/g, '\n').replace(/\n[ \t]/g, '').split('\n');
  const events = []; let ev = null;
  for(const line of lines){
    if(line === 'BEGIN:VEVENT'){ ev = {}; continue; }
    if(line === 'END:VEVENT'){ if(ev) events.push(ev); ev = null; continue; }
    if(!ev) continue;
    const m = /^([A-Z-]+)((?:;[^:]*)?):(.*)$/.exec(line);
    if(!m) continue;
    ev[m[1]] = {params: m[2], value: m[3].replace(/\\n/gi, ' ').replace(/\\([,;\\])/g, '$1').trim()};
  }
  return events;
}

/** "20260607T222500" / "20260601" → {date: Date(local midnight), time: 'HH:MM'|'' , utc: bool} */
function icsTime(p){
  if(!p) return null;
  const m = /^(\d{4})(\d{2})(\d{2})(?:T(\d{2})(\d{2})\d{0,2}(Z?))?$/.exec(p.value);
  if(!m) return null;
  return {date: new Date(+m[1], +m[2]-1, +m[3]), time: m[4] ? m[4]+':'+m[5] : '', utc: m[6] === 'Z', allDay: !m[4]};
}

function icsToAoa(text){
  const events = parseIcsEvents(text);
  if(!events.length) throw new Error('No events found in this calendar file.');
  const at = t => t.date.getTime() + (t.time ? (+t.time.slice(0, 2)) * 3600e3 + (+t.time.slice(3)) * 60e3 : 0);
  const rel = (t, anchor) => { if(!t || !t.time) return ''; const n = dayDiff(anchor, t.date); return t.time + (n > 0 ? '(+'+n+')' : ''); };
  const dur = (a, b) => { const mins = Math.round((at(b) - at(a)) / 60000); return mins > 0 ? minToHHMM(mins) : ''; };

  const allDay = [], duties = [], flights = [], other = [];
  let anyUtc = false;
  for(const e of events){
    const st = icsTime(e.DTSTART), en = icsTime(e.DTEND);
    const sum = (e.SUMMARY && e.SUMMARY.value) || '', desc = (e.DESCRIPTION && e.DESCRIPTION.value) || '';
    if(!st) continue;
    if(st.utc) anyUtc = true;
    let fm;
    if(st.allDay){
      const [code, stn] = sum.split(/\s+-\s+/);
      const days = en && en.allDay ? Math.max(1, dayDiff(st.date, en.date)) : 1;
      for(let d = 0; d < days; d++) allDay.push({date: addDays(st.date, d), code: (code || '').trim(), stn: (stn || '').trim(), desc});
    } else if(/report\s*time/i.test(sum)){
      const pm = /\(?\s*([^\s()]+)\s*\(/.exec(desc);
      duties.push({st, en: en || st, pairing: pm ? pm[1] : ''});
    } else if((fm = /^([A-Z0-9]{2})\s*(\d{1,4}[A-Z]?)\s+([A-Z]{3,4})\s*-\s*([A-Z]{3,4})/.exec(sum))){
      flights.push({st, en: en || st, item: fm[1]+fm[2], dep: fm[3], arr: fm[4]});
    } else {
      const [code, stn] = sum.split(/\s+-\s+/);
      other.push({st, en: en || st, code: (code || sum).trim().split(/\s+/)[0], stn: (stn || '').trim()});
    }
  }

  const rows = [];          // [sortKey, cells[]]
  const used = new Set();
  duties.sort((a, b) => at(a.st) - at(b.st));
  for(const d of duties){
    const legs = flights.filter(f => !used.has(f) && at(f.st) >= at(d.st) && at(f.st) <= at(d.en)).sort((a, b) => at(a.st) - at(b.st));
    legs.forEach(f => used.add(f));
    const anchor = d.st.date;
    if(!legs.length){ rows.push([at(d.st), [fmtDate(anchor), d.pairing, rel(d.st, anchor), d.pairing, '', '', '', rel(d.en, anchor), '', dur(d.st, d.en)]]); continue; }
    legs.forEach((f, k) => rows.push([at(d.st) + k, [
      k === 0 ? fmtDate(anchor) : '', k === 0 ? d.pairing : '', k === 0 ? rel(d.st, anchor) : '', f.item, '',
      f.dep+' '+rel(f.st, anchor), f.arr+' '+rel(f.en, anchor), k === legs.length - 1 ? rel(d.en, anchor) : '',
      dur(f.st, f.en), k === 0 ? dur(d.st, d.en) : '']]));
  }
  for(const f of flights.filter(f => !used.has(f))){
    const anchor = f.st.date;
    rows.push([at(f.st), [fmtDate(anchor), '', '', f.item, '', f.dep+' '+rel(f.st, anchor), f.arr+' '+rel(f.en, anchor), '', dur(f.st, f.en), '']]);
  }
  for(const o of other){
    const anchor = o.st.date, stn = o.stn;
    rows.push([at(o.st), [fmtDate(anchor), o.code, rel(o.st, anchor), o.code, '', stn ? stn+' '+rel(o.st, anchor) : '', stn ? stn+' '+rel(o.en, anchor) : '', rel(o.en, anchor), '', dur(o.st, o.en)]]);
  }
  for(const a of allDay){
    const stn = a.stn;
    rows.push([at({date: a.date, time: ''}), [fmtDate(a.date), a.code, '00:00', '', '', stn ? stn+' 00:00' : '', stn ? stn+' 23:59' : '', '23:59', '', '0:00']]);
  }
  rows.sort((a, b) => a[0] - b[0]);
  const tz = anyUtc ? ' (UTC)' : '';
  return {
    aoa: [['Date', 'Pairing/Activity', 'Duty Report', 'Item', 'Work Type', 'Dep Stn / Dep Time'+tz, 'Arr. Stn / Arr. Time'+tz, 'Duty Debrief', 'Flying Hrs', 'Duty Hrs']]
           .concat(rows.map(r => r[1])),
    descriptions: Object.fromEntries(allDay.filter(a => a.code && a.desc).map(a => [a.code, a.desc])),
  };
}

function icsToWorkbook(text){
  const {aoa, descriptions} = icsToAoa(text);
  return {wb: {SheetNames: ['Calendar'], Sheets: {Calendar: XLSX.utils.aoa_to_sheet(aoa)}}, descriptions};
}

/* ------------------------------------------------------------------ *
 *  Saved web page (.html) of the Roster Report (JasperReports HTML)
 *  Parsed as inert data with DOMParser: no scripts run, nothing loads.
 * ------------------------------------------------------------------ */
/** Lay out a <table> as a grid that honours colspan/rowspan: [{cells:[{col, span, text}]}]. */
function tableGrid(table){
  const taken = [];          // taken[row][col] = true when filled by a rowspan from above
  const rows = [];
  [...table.rows].forEach((tr, r) => {
    const cells = []; let c = 0;
    taken[r] = taken[r] || [];
    for(const td of tr.cells){
      while(taken[r][c]) c++;
      const span = Math.max(1, +td.colSpan || 1), rspan = Math.max(1, +td.rowSpan || 1);
      for(let dr = 0; dr < rspan; dr++){ taken[r+dr] = taken[r+dr] || []; for(let dc = 0; dc < span; dc++) taken[r+dr][c+dc] = true; }
      const text = (td.innerText !== undefined ? td.innerText : td.textContent).replace(/ /g, ' ').replace(/\s+/g, ' ').trim();
      cells.push({col: c, span, text});
      c += span;
    }
    rows.push(cells);
  });
  return rows;
}

function htmlToAoa(text){
  const doc = new DOMParser().parseFromString(text, 'text/html');
  const aoa = [];
  let labels = null;                       // header labels in output order (from the first table)
  for(const table of doc.querySelectorAll('table')){
    const grid = tableGrid(table);
    const hi = grid.findIndex(cells => cells.some(c => c.text === 'Date') && cells.some(c => /^Pairing/.test(c.text)));
    if(hi < 0) continue;
    const hdr = grid[hi].filter(c => c.text);
    const labelAt = col => { const h = hdr.find(h => col >= h.col && col < h.col + h.span); return h ? h.text : null; };
    if(!labels){
      // Lines above the header: title, period, "NAME | STAFF | FLEET | BASE | RANK", totals.
      grid.slice(0, hi).forEach(cells => { const t = cells.map(c => c.text).filter(Boolean); if(t.length) aoa.push(t); });
      labels = hdr.map(h => h.text);
      aoa.push(labels);
    }
    for(const cells of grid.slice(hi + 1)){
      if(!cells.some(c => c.text)) continue;
      const row = labels.map(() => '');
      cells.forEach(c => { const l = c.text && labelAt(c.col); const k = l ? labels.indexOf(l) : -1; if(k >= 0) row[k] = row[k] ? row[k]+' '+c.text : c.text; });
      if(row.some(Boolean)) aoa.push(row);
    }
  }
  if(!labels) throw new Error('This web page doesn\'t contain a crew Roster Report table (no "Date" / "Pairing/Activity" header found).');
  return aoa;
}

function htmlToWorkbook(text){
  return {SheetNames: ['Web page'], Sheets: {'Web page': XLSX.utils.aoa_to_sheet(htmlToAoa(text))}};
}
