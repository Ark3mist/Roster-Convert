/* Roster Convert — xls/xlsx roster → daily roster table → PDF.
   Everything runs client-side; no file is ever sent anywhere. */
'use strict';

/* ------------------------------------------------------------------ *
 *  State
 * ------------------------------------------------------------------ */
const S = {
  file: null,          // File
  wb: null,            // SheetJS workbook
  sheet: null,         // active sheet name
  mode: 'roster',      // 'roster' | 'raw'
  rosterOK: false,     // active sheet looks like a crew roster export
  model: null,         // {kind, title, subtitle, columns, rows, stats, codes}
};

const $ = id => document.getElementById(id);
const store = {
  get(k, fb){ try{ const v = localStorage.getItem('rosterConvert.'+k); return v ? JSON.parse(v) : fb; }catch(e){ return fb; } },
  set(k, v){ try{ localStorage.setItem('rosterConvert.'+k, JSON.stringify(v)); }catch(e){} },
};

/* ------------------------------------------------------------------ *
 *  Logging
 * ------------------------------------------------------------------ */
function log(msg, cls){ const d = document.createElement('div'); d.className = cls || ''; d.textContent = msg; $('log').appendChild(d); }
function clearLog(){ $('log').innerHTML = ''; }
function setBadge(text, ok){ const b = $('fmtBadge'); b.textContent = text; b.className = 'badge ' + (ok ? 'on' : 'bad'); }

/* ------------------------------------------------------------------ *
 *  Date / time helpers
 * ------------------------------------------------------------------ */
const MON = {jan:0,feb:1,mar:2,apr:3,may:4,jun:5,jul:6,aug:7,sep:8,oct:9,nov:10,dec:11};
const MON_UP = ['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC'];
const DOW = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];

const ymd = d => d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0')+'-'+String(d.getDate()).padStart(2,'0');
const addDays = (d, n) => { const r = new Date(d); r.setDate(r.getDate()+n); return r; };
const dayDiff = (a, b) => Math.round((Date.UTC(b.getFullYear(),b.getMonth(),b.getDate()) - Date.UTC(a.getFullYear(),a.getMonth(),a.getDate())) / 864e5);
const fmtDate = d => String(d.getDate()).padStart(2,'0')+'-'+MON_UP[d.getMonth()]+'-'+d.getFullYear();

/** Parse a date cell. Accepts Excel serials, Date objects and common text forms. */
function parseDateCell(raw, text){
  if(typeof raw === 'number' && raw > 20000 && raw < 80000){
    const p = XLSX.SSF.parse_date_code(raw);
    if(p) return new Date(p.y, p.m-1, p.d);
  }
  if(raw instanceof Date && !isNaN(raw)) return new Date(raw.getFullYear(), raw.getMonth(), raw.getDate());
  const s = String(text || raw || '').trim();
  if(!s) return null;
  let m;
  // 01-Nov-2026, 1 Nov 2026, 01Nov26, Sat 01-Nov-2026
  if((m = /(\d{1,2})[-\s/]?([A-Za-z]{3})[A-Za-z]*[-\s/]?(\d{2,4})/.exec(s))){
    const mo = MON[m[2].toLowerCase()];
    if(mo !== undefined){ let y = +m[3]; if(y < 100) y += 2000; return new Date(y, mo, +m[1]); }
  }
  // 2026-11-01
  if((m = /^(\d{4})-(\d{1,2})-(\d{1,2})/.exec(s))) return new Date(+m[1], +m[2]-1, +m[3]);
  // 01/11/2026 (day first — crew rosters outside the US use D/M/Y)
  if((m = /^(\d{1,2})[/.](\d{1,2})[/.](\d{2,4})$/.exec(s))){ let y = +m[3]; if(y < 100) y += 2000; return new Date(y, +m[2]-1, +m[1]); }
  return null;
}

/** "06:30", "6:30:00", "06:30(+1)" → {time:'06:30', roll:1} */
function parseTime(t){
  const s = String(t || '').trim();
  if(!s) return {time:'', roll:0};
  const m = /^(\d{1,2}):(\d{2})(?::\d{2})?\s*(?:\(\s*\+\s*(\d)\s*\))?$/.exec(s);
  if(!m) return {time: s.replace(/\(\s*\+\s*\d\s*\)/, '').trim(), roll: /\(\s*\+\s*(\d)/.test(s) ? +RegExp.$1 : 0};
  return {time: m[1].padStart(2,'0')+':'+m[2], roll: m[3] ? +m[3] : 0};
}

/** "KUL 08:30(+1)" → ['KUL', '08:30(+1)'] */
function splitStnTime(s){
  s = String(s || '').trim();
  if(!s) return ['', ''];
  const m = /^([A-Z]{3,4})\s*(\d{1,2}:\d{2}(?::\d{2})?\s*(?:\(\s*\+\s*\d\s*\))?)$/.exec(s);
  return m ? [m[1], m[2]] : [s, ''];
}

/** Duration "12:35" / "1:05:00" / 1.5 → minutes */
function durToMin(s){
  if(s === '' || s == null) return 0;
  const m = /^(\d{1,3}):(\d{2})/.exec(String(s).trim());
  if(m) return (+m[1])*60 + (+m[2]);
  const n = parseFloat(s);
  return isFinite(n) && /^\d+(\.\d+)?$/.test(String(s).trim()) ? Math.round(n*60) : 0;
}
const minToHHMM = n => Math.floor(n/60)+':'+String(n%60).padStart(2,'0');

/* ------------------------------------------------------------------ *
 *  Codes / legend
 * ------------------------------------------------------------------ */
const DEFAULT_LEGEND = {
  DO:'Day Off', D:'Day Off', OFF:'Day Off', AL:'Annual Leave', EL:'Emergency Leave',
  MCH:'Medical Leave', MC:'Medical Leave', SL:'Sick Leave', LO:'Layover',
  OP:'Operating', DH:'Deadhead', DHD:'Deadhead', PAX:'Positioning as passenger',
  SB:'Standby', SBY:'Standby', RES:'Reserve', TRG:'Training', SIM:'Simulator',
};
const isOffCode = c => /^(D|DO\d*|OFF|RDO|AL|EL|MCH|MC|SL|LV)$/i.test(String(c || '').trim());

function legend(){ return Object.assign({}, DEFAULT_LEGEND, store.get('legend', {})); }
function describe(code){
  const L = legend();
  if(L[code]) return L[code];
  if(/^DO\d+$/.test(code)) return 'Day Off';
  if(/^S\d?[A-Z]{2,4}$/.test(code) || /^SB/.test(code)) return 'Standby';
  if(/RCNT|RECUR|TRNG|TRAIN|SEP|CRM/i.test(code)) return 'Training / Recurrent';
  return '';
}

/* ------------------------------------------------------------------ *
 *  Sheet reading
 * ------------------------------------------------------------------ */
function sheetMatrices(ws){
  const raw  = XLSX.utils.sheet_to_json(ws, {header:1, raw:true,  defval:''});
  const text = XLSX.utils.sheet_to_json(ws, {header:1, raw:false, defval:''});
  return {raw, text};
}

const norm = s => String(s == null ? '' : s).replace(/\s+/g, ' ').trim();
const normKey = s => norm(s).toLowerCase().replace(/[^a-z0-9/]/g, '');

/** Locate the crew-roster header row ("Date" + "Pairing/Activity"). */
function findRosterHeader(text){
  for(let i = 0; i < Math.min(text.length, 40); i++){
    const keys = text[i].map(normKey);
    if(keys.includes('date') && keys.some(k => k.startsWith('pairing') || k === 'activity' || k === 'duty' )) return i;
  }
  return -1;
}

/** Crew metadata cell like "NAME | STAFF ID | FLEET | BASE | RANK" in the first rows. */
function findCrewMeta(text, headerIdx){
  for(let i = 0; i < Math.min(headerIdx < 0 ? 6 : headerIdx, 8); i++){
    for(const c of text[i] || []){
      const s = norm(c);
      if(s.includes('|')){
        const parts = s.split('|').map(norm).filter(Boolean);
        if(parts.length >= 2) return {name: parts[0], details: parts.slice(1)};
      }
    }
  }
  return null;
}

/* ------------------------------------------------------------------ *
 *  Crew roster parser
 * ------------------------------------------------------------------ */
function parseRosterRows(ws){
  const {raw, text} = sheetMatrices(ws);
  const h = findRosterHeader(text);
  if(h < 0) throw new Error('No roster header row found (expected "Date" and "Pairing/Activity" columns).');
  const header = text[h].map(normKey);
  const col = (...names) => {
    for(const n of names){ const k = normKey(n); const i = header.findIndex(x => x === k); if(i >= 0) return i; }
    for(const n of names){ const k = normKey(n); const i = header.findIndex(x => x && x.startsWith(k)); if(i >= 0) return i; }
    return -1;
  };
  const idx = {
    date: col('Date'), pairing: col('Pairing/Activity', 'Pairing', 'Activity'),
    report: col('Duty Report', 'Report'), item: col('Item', 'Flight'), rank: col('Overridden Rank', 'Rank'),
    work: col('Work Type'), dep: col('Dep Stn', 'Dep. Stn', 'Departure'), arr: col('Arr. Stn', 'Arr Stn', 'Arrival'),
    debrief: col('Duty Debrief', 'Debrief'), fly: col('Flying Hrs', 'Block Hrs', 'Block'), duty: col('Duty Hrs'),
    sdc: col('SDC', 'Duty Code'), ac: col('A/C Type', 'AC Type', 'Aircraft'), hotel: col('Hotel'),
  };

  const rows = [];
  let lastDate = null;
  for(let i = h+1; i < text.length; i++){
    const t = text[i] || [], r = raw[i] || [];
    if(t.every(c => norm(c) === '')) continue;
    const get = j => j < 0 ? '' : norm(t[j]);
    if(normKey(get(idx.date)) === 'date') continue;              // repeated header on page breaks
    const own = idx.date >= 0 ? parseDateCell(r[idx.date], t[idx.date]) : null;
    if(own) lastDate = own;
    if(!lastDate) continue;
    const [depStn, depTime] = splitStnTime(get(idx.dep));
    const [arrStn, arrTime] = splitStnTime(get(idx.arr));
    rows.push({
      ownDate: own, pairing: get(idx.pairing), report: get(idx.report), item: get(idx.item),
      rank: get(idx.rank), work: get(idx.work), depStn, depTime, arrStn, arrTime,
      debrief: get(idx.debrief), fly: get(idx.fly), duty: get(idx.duty), sdc: get(idx.sdc),
      ac: get(idx.ac), hotel: get(idx.hotel),
    });
  }
  if(!rows.length) throw new Error('Roster header found, but no data rows below it.');
  return {rows, meta: findCrewMeta(text, h)};
}

function rowsToLegs(rows){
  const legs = [];
  let date = null, pairing = '';
  for(const r of rows){
    if(r.ownDate) date = r.ownDate;
    if(r.pairing) pairing = r.pairing;
    const dep = parseTime(r.depTime), arr = parseTime(r.arrTime), rep = parseTime(r.report), deb = parseTime(r.debrief);
    const isFlight = /^[A-Z0-9]{2}\d{1,4}[A-Z]?$/.test(r.item) && !!(r.depStn && r.arrStn) ||
                     (r.depStn && r.arrStn && r.depStn !== r.arrStn && /^[A-Z]{3,4}$/.test(r.depStn) && /^[A-Z]{3,4}$/.test(r.arrStn));
    legs.push({
      isFlight, pairing,
      item: r.item || r.pairing || '',
      depDate: addDays(date, dep.roll), depStn: r.depStn, depTime: dep.time,
      arrDate: addDays(date, arr.roll), arrStn: r.arrStn, arrTime: arr.time,
      repDate: rep.time ? addDays(date, rep.roll) : null, repTime: rep.time,
      debDate: deb.time ? addDays(date, deb.roll) : null, debTime: deb.time,
      rank: r.rank, work: r.work, fly: r.fly, duty: r.duty, sdc: r.sdc, ac: r.ac, hotel: r.hotel,
    });
  }
  return legs;
}

/** Most common departure station of flights = home base. */
function guessHome(legs, meta){
  const count = {};
  legs.filter(l => l.isFlight && l.depStn).forEach(l => { count[l.depStn] = (count[l.depStn] || 0) + 1; });
  if(meta){ for(const d of meta.details){ if(count[d.toUpperCase()]) return d.toUpperCase(); } }
  let best = '', n = 0;
  for(const k in count){ if(count[k] > n){ best = k; n = count[k]; } }
  return best;
}

const ROSTER_COLUMNS = [
  {key:'date',  label:'Date'},
  {key:'dow',   label:'Day'},
  {key:'start', label:'Duty Start'},
  {key:'item',  label:'Item'},
  {key:'rank',  label:'Rank'},
  {key:'dep',   label:'Dep / Start'},
  {key:'arr',   label:'Arr / End'},
  {key:'end',   label:'Duty End'},
  {key:'work',  label:'Work Type'},
  {key:'block', label:'Block Hrs', num:true},
  {key:'duty',  label:'Duty Hrs',  num:true},
  {key:'code',  label:'Duty Code'},
  {key:'ac',    label:'A/C'},
  {key:'hotel', label:'Hotel'},
];

function buildRosterModel(ws){
  const {rows, meta} = parseRosterRows(ws);
  const legs = rowsToLegs(rows);
  const home = guessHome(legs, meta);
  const days = {};
  const dayOf = d => (days[ymd(d)] ||= {date: d, rows: []});
  const rel = (time, date, rowDate) => {             // show "(+1)" when time falls on a later day than its row
    if(!time) return '';
    const n = date ? dayDiff(rowDate, date) : 0;
    return n > 0 ? time+'(+'+n+')' : n < 0 ? time+'('+n+')' : time;
  };
  const st = (stn, time) => [stn, time].filter(Boolean).join(' ');

  let minD = null, maxD = null;
  const span = d => { if(!d) return; if(!minD || d < minD) minD = d; if(!maxD || d > maxD) maxD = d; };

  legs.forEach((l, i) => {
    [l.depDate, l.arrDate, l.repDate, l.debDate].forEach(span);
    const base = {rank:l.rank, work:l.work, code:l.sdc, ac:l.ac, hotel:l.hotel};

    if(!l.isFlight){
      const d = l.depDate;
      dayOf(d).rows.push(Object.assign({}, base, {
        start: rel(l.repTime, l.repDate, d), item: l.item, dep: st(l.depStn, rel(l.depTime, l.depDate, d)),
        arr: st(l.arrStn, rel(l.arrTime, l.arrDate, d)), end: rel(l.debTime, l.debDate, d), block:'', duty: l.duty,
      }));
      return;
    }

    const sameDay = ymd(l.depDate) === ymd(l.arrDate);
    // report the evening before a just-after-midnight departure
    if(l.repDate && dayDiff(l.repDate, l.depDate) > 0){
      dayOf(l.repDate).rows.push({start: l.repTime, item: l.item, rank:'', dep:'', arr:'', end:'', work:'', block:'', duty:'', code:'', ac:'', hotel:''});
    }
    const showRep = l.repDate && ymd(l.repDate) === ymd(l.depDate);
    dayOf(l.depDate).rows.push(Object.assign({}, base, {
      start: showRep ? l.repTime : '', item: l.item,
      dep: st(l.depStn, l.depTime),
      arr: sameDay ? st(l.arrStn, l.arrTime) : '',
      end: sameDay ? rel(l.debTime, l.debDate, l.depDate) : '',
      block: l.fly, duty: l.duty,
    }));
    if(!sameDay){
      dayOf(l.arrDate).rows.push({start:'', item:'', rank:'', dep:'', arr: st(l.arrStn, l.arrTime),
        end: rel(l.debTime, l.debDate, l.arrDate), work:'', block:'', duty:'', code:'', ac:'', hotel: ''});
    }

  });
  if(!minD) throw new Error('Could not work out the roster dates.');

  // Second pass: empty days between an outstation arrival and the next flight are layovers.
  const flights = legs.filter(l => l.isFlight);
  flights.forEach((l, i) => {
    const next = flights[i+1];
    if(!next || !l.arrStn || l.arrStn === home || (next.depStn && next.depStn !== l.arrStn)) return;
    for(let c = addDays(l.arrDate, 1); c < next.depDate; c = addDays(c, 1)){
      const d = dayOf(c);
      if(d.rows.length) continue;
      const showRep = next.repDate && ymd(next.repDate) === ymd(c);
      d.rows.push({start: showRep ? next.repTime : '', item:'LO', rank:'', dep: l.arrStn, arr:'', end:'', work:'', block:'', duty:'', code:'', ac:'', hotel: next.hotel || l.hotel});
    }
  });

  // Model rows: every calendar day in range, first row carries the date.
  const out = [];
  let blockMin = 0, dutyMin = 0, sectors = 0, offBase = 0, offAway = 0, workDays = 0;
  const codes = new Set();
  for(let c = new Date(minD); c <= maxD; c = addDays(c, 1)){
    const d = days[ymd(c)] || {date: c, rows: []};
    const rs = d.rows.length ? d.rows : [{}];
    const weekend = c.getDay() === 0 || c.getDay() === 6;
    rs.forEach((r, ri) => {
      const cls = [];
      if(ri === 0) cls.push('daystart'); else cls.push('leg2');
      if(ri === 0 && isOffCode(r.item)) cls.push('off');
      else if(ri === 0 && r.item === 'LO') cls.push('layover');
      if(weekend) cls.push('weekend');
      out.push({cls, cells: {
        date: ri === 0 ? fmtDate(c) : '', dow: ri === 0 ? DOW[c.getDay()] : '',
        start: r.start || '', item: r.item || '', rank: r.rank || '', dep: r.dep || '', arr: r.arr || '',
        end: r.end || '', work: r.work || '', block: r.block || '', duty: r.duty || '', code: r.code || '',
        ac: r.ac || '', hotel: r.hotel || '',
      }});
      blockMin += durToMin(r.block); dutyMin += durToMin(r.duty);
      if(r.block) sectors++;
      if(r.item) codes.add(r.item);
      if(r.work) codes.add(r.work);
      String(r.code || '').split(/[,;\s]+/).filter(Boolean).forEach(x => codes.add(x));
    });
    const first = rs[0];
    if(rs.length === 1 && isOffCode(first.item)){
      const stn = (first.dep || '').split(' ')[0];
      if(stn && home && stn !== home) offAway++; else offBase++;
    } else if(rs.length === 1 && first.item === 'LO') offAway++;
    else if(d.rows.some(r => !isOffCode(r.item) && r.item !== 'LO' && (r.item || r.start || r.arr || r.end))) workDays++;
  }

  const dayCount = dayDiff(minD, maxD) + 1;
  return {
    kind: 'roster',
    title: meta ? meta.name : 'Daily Roster',
    subtitle: [meta ? meta.details.join(' · ') : '', fmtDate(minD)+' to '+fmtDate(maxD)].filter(Boolean).join(' · '),
    columns: ROSTER_COLUMNS,
    rows: out,
    stats: [
      ['Days', String(dayCount)],
      ['Block Hours', minToHHMM(blockMin)],
      ['Duty Hours', minToHHMM(dutyMin)],
      ['Sectors', String(sectors)],
      ['Duty Days', String(workDays)],
      ['Off at Base'+(home ? ' ('+home+')' : ''), String(offBase)],
      ['Off / Layover Away', String(offAway)],
    ],
    codes: [...codes].filter(Boolean).sort(),
  };
}

/* ------------------------------------------------------------------ *
 *  Generic sheet → table (any spreadsheet)
 * ------------------------------------------------------------------ */
function buildRawModel(ws, sheetName){
  const {text} = sheetMatrices(ws);
  let rows = text.map(r => r.map(norm));
  rows = rows.filter(r => r.some(c => c !== ''));
  if(!rows.length) throw new Error('This sheet is empty.');
  const width = Math.max(...rows.map(r => r.length));
  const used = [];
  for(let j = 0; j < width; j++) if(rows.some(r => (r[j] || '') !== '')) used.push(j);
  // Header = first row that fills at least half the used columns.
  let h = rows.findIndex(r => used.filter(j => r[j]).length >= Math.max(1, Math.ceil(used.length/2)));
  if(h < 0) h = 0;
  const pre = rows.slice(0, h).map(r => r.filter(Boolean).join('  ')).filter(Boolean);
  const columns = used.map((j, n) => {
    const label = rows[h][j] || ('Column '+(n+1));
    const vals = rows.slice(h+1).map(r => r[j] || '').filter(Boolean);
    const num = vals.length > 0 && vals.every(v => /^-?[\d,.:]+%?$/.test(v));
    return {key: 'c'+j, label, num};
  });
  const body = rows.slice(h+1).map(r => {
    const cells = {}; used.forEach(j => { cells['c'+j] = r[j] || ''; });
    return {cls: [], cells};
  });
  return {kind: 'raw', title: sheetName || 'Sheet', subtitle: pre.join('   '), columns, rows: body, stats: [], codes: []};
}

/* ------------------------------------------------------------------ *
 *  Column visibility
 * ------------------------------------------------------------------ */
function hiddenCols(){ return new Set(store.get('hidden.'+S.model.kind, [])); }
function visibleColumns(){
  const hid = hiddenCols();
  return S.model.columns.filter(c => !hid.has(c.key) && S.model.rows.some(r => r.cells[c.key]));
}
function renderColumnsPanel(){
  const hid = hiddenCols();
  const list = $('columnsList'); list.innerHTML = '';
  S.model.columns.forEach(c => {
    const empty = !S.model.rows.some(r => r.cells[c.key]);
    const lab = document.createElement('label');
    const cb = document.createElement('input');
    cb.type = 'checkbox'; cb.checked = !hid.has(c.key); cb.disabled = empty;
    cb.onchange = () => {
      const h = hiddenCols();
      cb.checked ? h.delete(c.key) : h.add(c.key);
      store.set('hidden.'+S.model.kind, [...h]);
      renderTable();
    };
    lab.append(cb, document.createTextNode(c.label + (empty ? ' (empty)' : '')));
    list.appendChild(lab);
  });
}

/* ------------------------------------------------------------------ *
 *  Rendering
 * ------------------------------------------------------------------ */
function renderTable(){
  const m = S.model, cols = visibleColumns();
  const thead = $('outHead'), tbody = $('outBody');
  thead.innerHTML = ''; tbody.innerHTML = '';
  const hr = document.createElement('tr');
  cols.forEach(c => { const th = document.createElement('th'); th.textContent = c.label; hr.appendChild(th); });
  thead.appendChild(hr);
  const frag = document.createDocumentFragment();
  m.rows.forEach(r => {
    const tr = document.createElement('tr');
    if(r.cls.length) tr.className = r.cls.join(' ');
    cols.forEach(c => {
      const td = document.createElement('td');
      td.textContent = r.cells[c.key] || '';
      if(c.num) td.className = 'num';
      if(c.key === 'date') td.className = 'date';
      tr.appendChild(td);
    });
    frag.appendChild(tr);
  });
  tbody.appendChild(frag);
}

function renderLegend(){
  const body = $('legendBody'); body.innerHTML = '';
  const m = S.model;
  if(m.kind !== 'roster'){ $('legendCard').hidden = true; return; }
  const extra = store.get('legend', {});
  const codes = new Set(m.codes.filter(c => describe(c)));
  Object.keys(extra).forEach(c => { if(m.codes.includes(c)) codes.add(c); });
  [...codes].sort().forEach(c => addLegendRow(c, describe(c)));
  $('legendCard').hidden = false;
}
function addLegendRow(code = '', desc = ''){
  const row = document.createElement('div');
  row.className = 'legendRow';
  const ci = Object.assign(document.createElement('input'), {className:'code', value: code, placeholder:'CODE'});
  const di = Object.assign(document.createElement('input'), {className:'desc', value: desc, placeholder:'Description'});
  const del = Object.assign(document.createElement('button'), {className:'ghost', textContent:'✕', title:'Remove'});
  let prevCode = code;
  const save = () => {
    const L = store.get('legend', {});
    const c = ci.value.trim().toUpperCase();
    if(prevCode && prevCode !== c) delete L[prevCode];
    if(c) L[c] = di.value.trim();
    prevCode = c;
    store.set('legend', L);
  };
  ci.onchange = save; di.onchange = save;
  del.onclick = () => { const L = store.get('legend', {}); if(prevCode) L[prevCode] = ''; store.set('legend', L); row.remove(); };
  row.append(ci, di, del);
  $('legendBody').appendChild(row);
}
function legendEntries(){
  return [...document.querySelectorAll('#legendBody .legendRow')]
    .map(r => [r.querySelector('.code').value.trim(), r.querySelector('.desc').value.trim()])
    .filter(([c, d]) => c && d);
}

function render(){
  const m = S.model;
  $('titleInput').value = m.title;
  $('rangeLbl').textContent = m.subtitle;
  renderTable();
  renderColumnsPanel();
  $('statsBody').innerHTML = '';
  m.stats.forEach(([k, v]) => {
    const d = document.createElement('div'); const b = document.createElement('b');
    d.textContent = k; b.textContent = v; d.appendChild(b); $('statsBody').appendChild(d);
  });
  $('statsCard').hidden = !m.stats.length;
  renderLegend();
  $('outCard').hidden = false;
  $('toolbar').hidden = false;
}

/* ------------------------------------------------------------------ *
 *  File handling
 * ------------------------------------------------------------------ */
async function handleFile(file){
  clearLog();
  S.file = file;
  $('status').hidden = false;
  $('fileNameLbl').textContent = file.name;
  try{
    const buf = await file.arrayBuffer();
    S.wb = XLSX.read(buf, {type:'array', cellDates:false, cellNF:false});
  }catch(err){
    setBadge('Unreadable', false);
    log('Could not read this file as a spreadsheet: '+err.message, 'err');
    return;
  }
  // Prefer the first sheet that looks like a crew roster.
  const names = S.wb.SheetNames;
  S.sheet = names.find(n => findRosterHeader(sheetMatrices(S.wb.Sheets[n]).text) >= 0) || names[0];
  const sel = $('sheetSelect'); sel.innerHTML = '';
  names.forEach(n => sel.add(new Option(n, n, false, n === S.sheet)));
  $('sheetPick').hidden = names.length < 2;
  loadSheet();
}

function loadSheet(preferredMode){
  clearLog();
  const ws = S.wb.Sheets[S.sheet];
  S.rosterOK = findRosterHeader(sheetMatrices(ws).text) >= 0;
  S.mode = preferredMode || (S.rosterOK ? 'roster' : 'raw');
  $('modePick').hidden = !S.rosterOK;
  $('modeSelect').value = S.mode;
  try{
    if(S.mode === 'roster'){
      S.model = buildRosterModel(ws);
      setBadge('Crew roster', true);
      log('Parsed '+S.model.stats[0][1]+' days. Check the table, then press Download PDF.', 'ok');
    } else {
      S.model = buildRawModel(ws, S.sheet);
      setBadge(S.rosterOK ? 'Raw sheet' : 'Spreadsheet (as-is)', true);
      log('Loaded '+S.model.rows.length+' rows × '+S.model.columns.length+' columns.'+(S.rosterOK ? '' : ' Not a recognised crew roster, so it is converted as-is.'), 'ok');
    }
    render();
  }catch(err){
    console.error(err);
    if(S.mode === 'roster'){
      log('Roster parse failed ('+err.message+') — showing the raw sheet instead.', 'warn');
      return loadSheet('raw');
    }
    setBadge('Error', false);
    log(err.message, 'err');
  }
}

/* ------------------------------------------------------------------ *
 *  Export: PDF / print / CSV
 * ------------------------------------------------------------------ */
const pdfSafe = s => String(s || '').replace(/[·•]/g, '-').replace(/[–—]/g, '-').replace(/[^\x20-\x7E\xA0-\xFF]/g, '');
const fileBase = () => {
  const t = ($('titleInput').value || 'roster').trim();
  const range = S.model.kind === 'roster' && S.model.rows.length ? '_'+S.model.rows[0].cells.date : '';
  return (t + range).replace(/[^\w\-]+/g, '_').replace(/_+/g, '_').replace(/^_|_$/g, '') || 'roster';
};

function exportPDF(){
  const {jsPDF} = window.jspdf;
  const cols = visibleColumns();
  const m = S.model;
  const paper = $('paperSelect').value;
  let orient = $('orientSelect').value;
  if(orient === 'auto'){
    // Estimate table width at 7.5pt and pick the orientation that avoids squeezing.
    const chars = cols.reduce((a, c) => a + Math.max(c.label.length * 0.8, ...m.rows.map(r => (r.cells[c.key] || '').length)) + 2, 0);
    const mmNeeded = chars * 1.55;
    orient = mmNeeded > (paper === 'letter' ? 196 : 190) ? 'landscape' : 'portrait';
  }
  const doc = new jsPDF({orientation: orient, unit: 'mm', format: paper});
  const W = doc.internal.pageSize.getWidth(), M = 10;
  const title = pdfSafe($('titleInput').value || m.title);

  doc.setFont('helvetica', 'bold'); doc.setFontSize(14); doc.setTextColor(17);
  doc.text(title, M, 14);
  doc.setFont('helvetica', 'normal'); doc.setFontSize(8.5); doc.setTextColor(90);
  const sub = doc.splitTextToSize(pdfSafe(m.subtitle), W - 2*M);
  doc.text(sub, M, 19);
  let y = 19 + sub.length * 3.6 + 1;

  const fills = {off:[236,247,214], layover:[225,240,252]};
  const texts = {off:[47,74,0], layover:[13,74,122]};
  doc.autoTable({
    startY: y,
    margin: {left: M, right: M, top: 12, bottom: 12},
    head: [cols.map(c => pdfSafe(c.label))],
    body: m.rows.map(r => cols.map(c => pdfSafe(r.cells[c.key]))),
    theme: 'grid',
    styles: {font:'helvetica', fontSize: cols.length > 12 ? 6.5 : 7.5, cellPadding: 1.2, lineColor: [210,210,210], lineWidth: 0.1, textColor: 20, overflow: 'linebreak', valign: 'top'},
    headStyles: {fillColor: [27,30,34], textColor: [198,255,61], fontStyle: 'bold'},
    columnStyles: Object.fromEntries(cols.map((c, i) => [i, c.num ? {halign:'right'} : {}])),
    didParseCell: data => {
      if(data.section !== 'body') return;
      const row = m.rows[data.row.index];
      const k = row.cls.includes('off') ? 'off' : row.cls.includes('layover') ? 'layover' : null;
      if(k){ data.cell.styles.fillColor = fills[k]; data.cell.styles.textColor = texts[k]; }
      if(cols[data.column.index].key === 'date' && row.cls.includes('weekend')) data.cell.styles.textColor = [163,95,0];
      if(row.cls.includes('daystart') && data.row.index > 0) data.cell.styles.lineWidth = {top: 0.35, bottom: 0.1, left: 0.1, right: 0.1};
    },
  });
  y = doc.lastAutoTable.finalY + 6;

  const smallTable = (heading, rows) => {
    if(!rows.length) return;
    if(y > doc.internal.pageSize.getHeight() - 30){ doc.addPage(); y = 14; }
    doc.setFont('helvetica', 'bold'); doc.setFontSize(10); doc.setTextColor(17);
    doc.text(heading, M, y); y += 2;
    doc.autoTable({
      startY: y, margin: {left: M, right: M}, tableWidth: Math.min(120, W - 2*M),
      body: rows.map(r => r.map(pdfSafe)), theme: 'plain',
      styles: {fontSize: 8, cellPadding: 1}, columnStyles: {0: {fontStyle: 'bold', cellWidth: 45}},
    });
    y = doc.lastAutoTable.finalY + 6;
  };
  smallTable('Statistics', m.stats);
  smallTable('Code Description', legendEntries());

  const pages = doc.getNumberOfPages();
  const stamp = 'Generated '+fmtDate(new Date())+' with Roster Convert';
  for(let p = 1; p <= pages; p++){
    doc.setPage(p);
    doc.setFont('helvetica', 'normal'); doc.setFontSize(7); doc.setTextColor(140);
    const H = doc.internal.pageSize.getHeight();
    doc.text(stamp, M, H - 5);
    doc.text('Page '+p+' of '+pages, W - M, H - 5, {align: 'right'});
  }
  doc.save(fileBase()+'.pdf');
}

function exportCSV(){
  const cols = visibleColumns();
  const esc = v => /[",\n]/.test(v) ? '"'+v.replace(/"/g, '""')+'"' : v;
  const lines = [cols.map(c => esc(c.label)).join(',')]
    .concat(S.model.rows.map(r => cols.map(c => esc(r.cells[c.key] || '')).join(',')));
  const blob = new Blob(['﻿'+lines.join('\r\n')], {type: 'text/csv;charset=utf-8'});
  const a = Object.assign(document.createElement('a'), {href: URL.createObjectURL(blob), download: fileBase()+'.csv'});
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(a.href), 1000);
}

/* ------------------------------------------------------------------ *
 *  Wiring
 * ------------------------------------------------------------------ */
(function init(){
  const drop = $('drop'), input = $('fileInput');
  $('chooseBtn').onclick = () => input.click();
  input.onchange = () => { if(input.files.length) handleFile(input.files[0]); input.value = ''; };
  ['dragenter', 'dragover'].forEach(e => drop.addEventListener(e, ev => { ev.preventDefault(); drop.classList.add('drag'); }));
  ['dragleave', 'drop'].forEach(e => drop.addEventListener(e, ev => { ev.preventDefault(); drop.classList.remove('drag'); }));
  drop.addEventListener('drop', ev => { const f = ev.dataTransfer.files[0]; if(f) handleFile(f); });
  // Allow dropping anywhere on the page.
  window.addEventListener('dragover', ev => ev.preventDefault());
  window.addEventListener('drop', ev => { ev.preventDefault(); if(ev.target.closest && ev.target.closest('#drop')) return; const f = ev.dataTransfer.files[0]; if(f) handleFile(f); });

  $('sheetSelect').onchange = e => { S.sheet = e.target.value; loadSheet(); };
  $('modeSelect').onchange = e => loadSheet(e.target.value);
  $('pdfBtn').onclick = () => { try{ exportPDF(); }catch(err){ console.error(err); log('PDF export failed: '+err.message, 'err'); } };
  $('printBtn').onclick = () => window.print();
  $('csvBtn').onclick = exportCSV;
  $('columnsBtn').onclick = () => { $('columnsPanel').hidden = !$('columnsPanel').hidden; };
  $('addCodeBtn').onclick = () => addLegendRow();
  ['orientSelect', 'paperSelect'].forEach(id => {
    const el = $(id); el.value = store.get(id, el.value);
    el.onchange = () => store.set(id, el.value);
  });
  if(!window.XLSX || !window.jspdf) log('A library failed to load — make sure the vendor/ folder was deployed with index.html.', 'err');
})();
