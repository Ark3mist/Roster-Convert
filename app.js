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
  warnings: [],
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
const MONTH_FULL = ['January','February','March','April','May','June','July','August','September','October','November','December'];
/** "SHAFIQ FAKHRI BIN OTH" → "Shafiq Fakhri Bin Oth" */
const titleCase = s => String(s || '').toLowerCase().replace(/(^|[\s\-'/(])([a-z])/g, (_, a, b) => a + b.toUpperCase());

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
  MCH:'Medical Leave', MC:'Medical Leave', MC1:'Medical Leave', SL:'Sick Leave', LO:'Layover',
  OP:'Operating', DH:'Deadhead', DHD:'Deadhead', PAX:'Positioning as passenger',
  SB:'Standby', SBY:'Standby', RES:'Reserve', TRG:'Training', SIM:'Simulator',
};
// Off / leave codes, with optional number suffix (DO1, MC1, AL2…).
const isOffCode = c => /^(D|DO|OFF|RDO|AL|EL|MCH|MC|SL|LV)\d*$/i.test(String(c || '').trim());

function legend(){ return Object.assign({}, DEFAULT_LEGEND, store.get('legend', {})); }
function describe(code){
  const L = legend();
  if(L[code]) return L[code];
  if(S.icsDescriptions && S.icsDescriptions[code]) return titleCase(S.icsDescriptions[code]);
  if(/^DO\d+$/.test(code)) return 'Day Off';
  if(/^S\d?[A-Z]{2,4}$/.test(code) || /^SB/.test(code)) return 'Standby';
  if(/RCNT|RECUR|TRNG|TRAIN|SEP|CRM/i.test(code)) return 'Training / Recurrent';
  return '';
}

/* ------------------------------------------------------------------ *
 *  Public holidays
 *  Built-in: Malaysia national + Federal Territory (KL) holidays.
 *  Fixed-date holidays are exact; lunar/religious ones are the expected
 *  dates and can shift by a day once officially announced — the list is
 *  editable in the app (add / remove, saved on this device).
 * ------------------------------------------------------------------ */
const BUILTIN_HOLIDAYS = {
  // 2026
  '2026-01-01':"New Year's Day", '2026-02-01':'Thaipusam / Federal Territory Day', '2026-02-02':'Replacement holiday',
  '2026-02-17':'Chinese New Year', '2026-02-18':'Chinese New Year (2nd day)', '2026-03-07':'Nuzul Al-Quran',
  '2026-03-21':'Hari Raya Aidilfitri', '2026-03-22':'Hari Raya Aidilfitri (2nd day)', '2026-03-23':'Replacement holiday',
  '2026-05-01':'Labour Day', '2026-05-27':'Hari Raya Haji', '2026-05-31':'Wesak Day',
  '2026-06-01':"Agong's Birthday", '2026-06-17':'Awal Muharram', '2026-08-25':'Maulidur Rasul',
  '2026-08-31':'National Day', '2026-09-16':'Malaysia Day', '2026-11-08':'Deepavali', '2026-11-09':'Replacement holiday',
  '2026-12-25':'Christmas Day',
  // 2027
  '2027-01-01':"New Year's Day", '2027-01-22':'Thaipusam', '2027-02-01':'Federal Territory Day',
  '2027-02-06':'Chinese New Year', '2027-02-07':'Chinese New Year (2nd day)', '2027-02-08':'Replacement holiday',
  '2027-02-25':'Nuzul Al-Quran', '2027-03-10':'Hari Raya Aidilfitri', '2027-03-11':'Hari Raya Aidilfitri (2nd day)',
  '2027-05-01':'Labour Day', '2027-05-17':'Hari Raya Haji', '2027-05-20':'Wesak Day',
  '2027-06-07':"Agong's Birthday", '2027-06-06':'Awal Muharram', '2027-08-15':'Maulidur Rasul',
  '2027-08-16':'Replacement holiday', '2027-08-31':'National Day', '2027-09-16':'Malaysia Day',
  '2027-10-28':'Deepavali', '2027-12-25':'Christmas Day',
};
function holidays(){
  const added = store.get('holidays.added', {}), removed = new Set(store.get('holidays.removed', []));
  const all = Object.assign({}, BUILTIN_HOLIDAYS, added);
  removed.forEach(k => { if(!(k in added)) delete all[k]; });
  return all;
}
const holidayName = key => key ? holidays()[key] || '' : '';

/* ------------------------------------------------------------------ *
 *  Time zones (Local ⇄ UTC)
 *  Station → IANA zone; offsets (incl. daylight saving) come from the
 *  browser's own tz database for the actual date, so they stay correct.
 * ------------------------------------------------------------------ */
const STATION_TZ = (() => {
  const z = {}, put = (zone, list) => list.split(' ').forEach(c => { z[c] = zone; });
  put('Asia/Kuala_Lumpur', 'KUL SZB PEN LGK KBR KUA TGG IPH AOR JHB MKZ TOD RDN SZB');
  put('Asia/Kuching', 'BKI KCH MYY SBW TWU SDK LBU MZV LDU MKM BTU LMN LWY BBN TMG GSA KUD');
  put('Asia/Singapore', 'SIN XSP');
  put('Asia/Brunei', 'BWN');
  put('Asia/Jakarta', 'CGK HLP KNO PDG PKU BTH PLM BDO SRG JOG YIA SOC SUB TNJ PNK DJB BTJ');
  put('Asia/Makassar', 'DPS UPG BPN LOP MDC BDJ KOE');
  put('Asia/Jayapura', 'DJJ');
  put('Asia/Bangkok', 'BKK DMK HKT CNX KBV USM HDY CEI UTH');
  put('Asia/Ho_Chi_Minh', 'SGN HAN DAD PQC CXR HPH VCA');
  put('Asia/Phnom_Penh', 'PNH REP KTI KOS');
  put('Asia/Vientiane', 'VTE LPQ');
  put('Asia/Yangon', 'RGN MDL');
  put('Asia/Manila', 'MNL CEB CRK DVO ILO KLO PPS');
  put('Asia/Hong_Kong', 'HKG'); put('Asia/Macau', 'MFM'); put('Asia/Taipei', 'TPE KHH TSA');
  put('Asia/Shanghai', 'PEK PKX PVG SHA CAN SZX XMN KMG CTU TFU CSX HGH NKG WUH CKG XIY TAO FOC NNG HAK SYX');
  put('Asia/Seoul', 'ICN GMP PUS CJU');
  put('Asia/Tokyo', 'NRT HND KIX ITM NGO FUK CTS OKA');
  put('Asia/Kolkata', 'DEL BOM MAA BLR HYD COK TRV TRZ AMD CCU GOI GOX ATQ IXE CJB');
  put('Asia/Colombo', 'CMB'); put('Indian/Maldives', 'MLE'); put('Asia/Dhaka', 'DAC CGP');
  put('Asia/Kathmandu', 'KTM'); put('Asia/Karachi', 'KHI LHE ISB');
  put('Asia/Dubai', 'DXB AUH SHJ'); put('Asia/Qatar', 'DOH'); put('Asia/Riyadh', 'JED MED RUH DMM');
  put('Asia/Muscat', 'MCT'); put('Asia/Bahrain', 'BAH'); put('Asia/Kuwait', 'KWI'); put('Asia/Tehran', 'IKA');
  put('Europe/Istanbul', 'IST SAW'); put('Europe/London', 'LHR LGW MAN'); put('Europe/Paris', 'CDG ORY');
  put('Europe/Amsterdam', 'AMS'); put('Europe/Berlin', 'FRA MUC'); put('Europe/Rome', 'FCO');
  put('Europe/Zurich', 'ZRH'); put('Europe/Madrid', 'MAD');
  put('Australia/Sydney', 'SYD'); put('Australia/Melbourne', 'MEL'); put('Australia/Brisbane', 'BNE OOL');
  put('Australia/Perth', 'PER'); put('Australia/Adelaide', 'ADL'); put('Australia/Darwin', 'DRW');
  put('Pacific/Auckland', 'AKL CHC');
  put('Africa/Johannesburg', 'JNB'); put('Africa/Nairobi', 'NBO'); put('America/Los_Angeles', 'LAX SFO');
  return z;
})();

const tzFmtCache = {};
/** Offset (minutes east of UTC) of `zone` at the instant `utcMs`. */
function tzOffset(zone, utcMs){
  utcMs = Math.floor(utcMs / 60000) * 60000;          // whole minutes, so seconds can't skew the result
  const f = tzFmtCache[zone] ||= new Intl.DateTimeFormat('en-US', {timeZone: zone, hourCycle: 'h23',
    year: 'numeric', month: 'numeric', day: 'numeric', hour: 'numeric', minute: 'numeric'});
  const p = Object.fromEntries(f.formatToParts(new Date(utcMs)).map(x => [x.type, x.value]));
  return Math.round((Date.UTC(+p.year, p.month-1, +p.day, +p.hour % 24, +p.minute) - utcMs) / 60000);
}

/** Re-express a wall-clock (date, "HH:MM") in zone `fromZ` as wall-clock time in zone `toZ`. */
function shiftTime(date, time, fromZ, toZ){
  const m = /^(\d{2}):(\d{2})$/.exec(time || '');
  if(!date || !m || !fromZ || !toZ || fromZ === toZ) return [date, time];
  const wall = Date.UTC(date.getFullYear(), date.getMonth(), date.getDate(), +m[1], +m[2]);
  let off = tzOffset(fromZ, wall); off = tzOffset(fromZ, wall - off*60000);   // second pass settles DST edges
  const instant = wall - off*60000;
  const d = new Date(instant + tzOffset(toZ, instant)*60000);
  return [new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate()),
          String(d.getUTCHours()).padStart(2,'0')+':'+String(d.getUTCMinutes()).padStart(2,'0')];
}

/**
 * Convert every leg's times between clocks:
 *   'base'  = home-base time for every station (how the crew Roster Report exports times)
 *   'local' = each station's own local time
 *   'utc'   = UTC
 * Report times belong to the departure station, debrief times to the arrival station.
 * Off/leave days are left on their date.
 */
function convertLegs(legs, from, to, home){
  const unknown = new Set();
  const baseZ = STATION_TZ[home] || 'Asia/Kuala_Lumpur';
  const stnZ = stn => { const z = STATION_TZ[(stn || '').toUpperCase()]; if(stn && !z && /^[A-Z]{3}$/.test(stn)) unknown.add(stn); return z || baseZ; };
  const clock = (mode, stationZone) => mode === 'utc' ? 'UTC' : mode === 'base' ? baseZ : stationZone;
  for(const l of legs){
    // Days off / leave are calendar days at home base, whatever clock is chosen for duties.
    const target = isOffCode(l.item) ? 'base' : to;
    if(from === target) continue;
    const depZ = stnZ(l.depStn), arrZ = stnZ(l.arrStn || l.depStn);
    [l.depDate, l.depTime] = shiftTime(l.depDate, l.depTime, clock(from, depZ), clock(target, depZ));
    [l.arrDate, l.arrTime] = shiftTime(l.arrDate, l.arrTime, clock(from, arrZ), clock(target, arrZ));
    if(l.repTime) [l.repDate, l.repTime] = shiftTime(l.repDate, l.repTime, clock(from, depZ), clock(target, depZ));
    if(l.debTime) [l.debDate, l.debTime] = shiftTime(l.debDate, l.debTime, clock(from, arrZ), clock(target, arrZ));
  }
  return [...unknown];
}
/** UTC wall-clock of home-base midnight (start) or 23:59 (end) around the roster's dates, e.g. KUL → "16:00". */
function utcMidnightOf(home, legs, end){
  const z = STATION_TZ[home] || 'Asia/Kuala_Lumpur';
  const d = (legs.find(l => l.depDate) || {}).depDate || new Date();
  const [, t] = shiftTime(d, end ? '23:59' : '00:00', z, 'UTC');
  return t;
}
/**
 * Some Roster Report exports (e.g. the portal's HTML) give times in UTC without saying so.
 * Tell-tale: days off run 00:00–23:59 home-base time, so in a UTC export they start at
 * home-base midnight in UTC (16:00 for KUL) instead of 00:00.
 */
function detectUtcExport(legs, home){
  const offs = legs.filter(l => isOffCode(l.item) && l.depTime);
  if(!offs.length) return 'base';
  const utcStart = utcMidnightOf(home, legs);
  if(utcStart === '00:00') return 'base';
  const utcLike = offs.filter(l => l.depTime === utcStart).length;
  return utcLike > offs.length / 2 ? 'utc' : 'base';
}
/**
 * Station-local exports give each airport's own local time. Tell-tale: on a flight between
 * airports in different time zones, arrival − departure ≠ block time, and the gap equals the
 * zone difference (MH727 KUL 22:35 → CGK 23:45 = 1:10 for a 2:10 block, CGK being −1h).
 * In home-base or UTC exports both ends use one clock, so arrival − departure = block.
 * Returns {same, local}: vote counts over flights that can tell the two apart.
 */
function clockVotes(legs, home){
  const baseZ = STATION_TZ[home] || 'Asia/Kuala_Lumpur';
  const zoneOf = stn => STATION_TZ[(stn || '').toUpperCase()] || baseZ;
  const at = (d, t) => { const m = /^(\d{2}):(\d{2})$/.exec(t || ''); return d && m ? Date.UTC(d.getFullYear(), d.getMonth(), d.getDate(), +m[1], +m[2]) : null; };
  let same = 0, local = 0;
  for(const l of legs){
    if(!l.isFlight || isOffCode(l.item)) continue;
    const dep = at(l.depDate, l.depTime), arr = at(l.arrDate, l.arrTime), block = durToMin(l.fly);
    if(dep == null || arr == null || !block) continue;
    const delta = tzOffset(zoneOf(l.arrStn), dep) - tzOffset(zoneOf(l.depStn), dep);   // minutes
    if(!delta) continue;                                   // same zone at both ends: can't tell
    const elapsed = Math.round((arr - dep) / 60000);
    if(Math.abs(elapsed - block) <= 5) same++;
    else if(Math.abs(elapsed - (block + delta)) <= 5) local++;
    // anything else (e.g. credited hours on augmented sectors) gives no vote
  }
  return {same, local};
}

/** Work out the file's clock when its header doesn't say: station local, UTC or home base. */
function detectClock(legs, home){
  const v = clockVotes(legs, home);
  if(v.local > v.same) return {clock: 'local', why: 'flight times differ from block hours by the airport time-zone difference'};
  if(detectUtcExport(legs, home) === 'utc') return {clock: 'utc', why: 'days off run '+utcMidnightOf(home, legs)+'–'+utcMidnightOf(home, legs, true)+' instead of 00:00–23:59'};
  return {clock: 'base', why: ''};
}

const utcOffsetLabel = zone => { const m = tzOffset(zone, Date.now()); return 'UTC'+(m >= 0 ? '+' : '-')+Math.floor(Math.abs(m)/60)+(Math.abs(m)%60 ? ':'+String(Math.abs(m)%60).padStart(2,'0') : ''); };
const TIME_LABEL = {local: 'All times station local (LT)', utc: 'All times UTC',
  base: home => 'All times '+(home || 'home base')+' ('+utcOffsetLabel(STATION_TZ[home] || 'Asia/Kuala_Lumpur')+')'};
const timeLabel = (mode, home) => mode === 'base' ? TIME_LABEL.base(home) : TIME_LABEL[mode];

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
  // Which clock the export uses. The crew Roster Report gives every station's times in
  // home-base time (e.g. LHR departures in KUL time) — that's the default. A header marked
  // "UTC"/"GMT"/"(Z)" means UTC; "(LT)"/"Local" means each station's own local time.
  const depHead = idx.dep >= 0 ? String(text[h][idx.dep]) : '';
  const srcTime = /\b(UTC|GMT)\b|\(Z\)/i.test(depHead) ? 'utc' : /\(LT\)|\blocal\b/i.test(depHead) ? 'local' : 'base';
  return {rows, meta: findCrewMeta(text, h), srcTime};
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
];

/** Day off / leave: just the code. Times (00:00–23:59) are noise; the station is kept only when away from base. */
function offCells(r, home){
  const stn = String(r.dep || '').split(' ')[0];
  return {start: '', dep: stn && home && stn !== home ? stn : '', arr: '', end: '', duty: '', block: ''};
}

function buildRosterModel(ws){
  const parsed = parseRosterRows(ws);
  const {rows, meta} = parsed;
  const legs = rowsToLegs(rows);
  const home = guessHome(legs, meta);
  // File clock: chosen by the user, or worked out (header label, then days-off pattern).
  const srcChoice = $('srcSelect').value;
  let srcTime = srcChoice !== 'auto' ? srcChoice : parsed.srcTime;
  if(srcChoice === 'auto' && parsed.srcTime === 'base'){
    const d = detectClock(legs, home);
    srcTime = d.clock;
    if(d.why) S.warnings.push('This file\'s times are '+(d.clock === 'utc' ? 'in UTC' : 'already in each airport\'s local time')+' (detected: '+d.why+'). They have been read accordingly. If that\'s wrong, change "File times".');
  }
  const timeMode = $('timeSelect').value;
  const unknownStations = convertLegs(legs, srcTime, timeMode, home);
  if(unknownStations.length) S.warnings.push('No time zone known for '+unknownStations.join(', ')+' — times there were converted using the home base ('+home+') zone. Tell the developer to add them.');
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
      // Flagged so it can't be overlooked: the classic missed report (evening before an after-midnight departure).
      const n = dayDiff(l.repDate, l.depDate);
      dayOf(l.repDate).rows.push({start: l.repTime, item: l.item, rank:'', dep: '> dep '+(n === 1 ? 'next day' : '+'+n+' days'),
        arr:'', end:'', work:'', block:'', duty:'', code:'', ac:'', hotel:'', earlyRpt: true});
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

  // Always show the full roster month(s), even when the export starts after the 1st
  // or ends before month end. The roster month = month holding most of the dated days;
  // earlier/later data (e.g. a trip spilling into the next month) is never cut off.
  const perMonth = {};
  Object.values(days).forEach(d => { const k = d.date.getFullYear()*12 + d.date.getMonth(); perMonth[k] = (perMonth[k] || 0) + 1; });
  const mainKey = +Object.keys(perMonth).sort((a, b) => perMonth[b] - perMonth[a] || a - b)[0];
  const monthStart = new Date(Math.floor(mainKey/12), mainKey % 12, 1);
  const monthEnd = new Date(Math.floor(mainKey/12), mainKey % 12 + 1, 0);
  if(monthStart < minD) minD = monthStart;
  if(monthEnd > maxD) maxD = monthEnd;

  // Model rows: every calendar day in range, first row carries the date.
  const out = [];
  let blockMin = 0, dutyMin = 0, sectors = 0, offBase = 0, offAway = 0, workDays = 0;
  const codes = new Set();
  for(let c = new Date(minD); c <= maxD; c = addDays(c, 1)){
    const d = days[ymd(c)] || {date: c, rows: []};
    const rs = d.rows.length ? d.rows : [{}];
    const weekend = c.getDay() === 0 || c.getDay() === 6;
    const isDutyRow = r => !isOffCode(r.item) && r.item !== 'LO' && !!(r.item || r.start || r.dep || r.arr || r.end);
    rs.forEach((r, ri) => {
      const cls = [];
      if(ri === 0) cls.push('daystart'); else cls.push('leg2');
      if(ri === 0 && isOffCode(r.item)) cls.push('off');
      else if(ri === 0 && r.item === 'LO') cls.push('layover');
      else if(isDutyRow(r)) cls.push('duty');
      if(weekend) cls.push('weekend');
      if(r.earlyRpt) cls.push('earlyrpt');
      out.push({cls, ymd: ri === 0 ? ymd(c) : '', cells: Object.assign({
        date: ri === 0 ? fmtDate(c) : '', dow: ri === 0 ? DOW[c.getDay()] : '',
        start: r.start || '', item: r.item || '', rank: r.rank || '', dep: r.dep || '', arr: r.arr || '',
        end: r.end || '', work: r.work || '', block: r.block || '', duty: r.duty || '', code: r.code || '',
        ac: r.ac || '', hotel: r.hotel || '',
      }, isOffCode(r.item) ? offCells(r, home) : {})});
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
    title: meta ? titleCase(meta.name) : 'Daily Roster',
    staffId: meta ? (meta.details.find(d => /^\d{4,}$/.test(d)) || '') : '',
    monthLabel: MONTH_FULL[mainKey % 12]+' '+Math.floor(mainKey/12),
    subtitle: [meta ? meta.details.join(' · ') : '', fmtDate(minD)+' to '+fmtDate(maxD), timeLabel(timeMode, home)].filter(Boolean).join(' · '),
    timeMode, home, srcTime,
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
  const H = holidays();
  m.rows.forEach(r => {
    const tr = document.createElement('tr');
    const ph = r.ymd && H[r.ymd];
    tr.className = r.cls.concat(ph ? ['holiday'] : []).join(' ');
    cols.forEach(c => {
      const td = document.createElement('td');
      td.textContent = r.cells[c.key] || '';
      if(c.num) td.className = 'num';
      if(c.key === 'date'){ td.className = 'date'; if(ph) td.title = ph; }
      if(c.key === 'start' && !r.cls.includes('off')){
        td.className = 'start';
        if(td.textContent){ const sp = document.createElement('span'); sp.className = 'rpt'; sp.textContent = td.textContent; td.textContent = ''; td.appendChild(sp); }
      }
      if(c.key === 'dep' && r.cls.includes('earlyrpt')) td.className = 'earlynote';
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
/** Holidays falling inside the loaded roster's date range. */
function holidaysInRange(){
  const keys = S.model.rows.map(r => r.ymd).filter(Boolean);
  if(!keys.length) return [];
  const H = holidays(), lo = keys[0], hi = keys[keys.length-1];
  return Object.keys(H).filter(k => k >= lo && k <= hi).sort().map(k => [k, H[k]]);
}
function renderHolidays(){
  const card = $('holidayCard');
  if(S.model.kind !== 'roster'){ card.hidden = true; return; }
  const body = $('holidayBody'); body.innerHTML = '';
  const list = holidaysInRange();
  if(!list.length){
    const p = document.createElement('div'); p.className = 'legendRow empty'; p.textContent = 'No public holidays in this period.'; body.appendChild(p);
  }
  list.forEach(([k, name]) => {
    const row = document.createElement('div'); row.className = 'legendRow';
    const d = document.createElement('span'); d.className = 'phDate'; d.textContent = fmtDate(new Date(k+'T00:00:00')) + ' ' + DOW[new Date(k+'T00:00:00').getDay()];
    const n = document.createElement('span'); n.className = 'phName'; n.textContent = name;
    const del = Object.assign(document.createElement('button'), {className:'ghost', textContent:'✕', title:'Not a holiday — remove'});
    del.onclick = () => {
      const added = store.get('holidays.added', {}); delete added[k]; store.set('holidays.added', added);
      const removed = new Set(store.get('holidays.removed', [])); removed.add(k); store.set('holidays.removed', [...removed]);
      renderTable(); renderHolidays();
    };
    row.append(d, n, del); body.appendChild(row);
  });
  card.hidden = false;
}
function addHoliday(){
  const k = $('phDateInput').value, name = $('phNameInput').value.trim() || 'Public Holiday';
  if(!/^\d{4}-\d{2}-\d{2}$/.test(k)){ $('phDateInput').focus(); return; }
  const added = store.get('holidays.added', {}); added[k] = name; store.set('holidays.added', added);
  const removed = store.get('holidays.removed', []).filter(x => x !== k); store.set('holidays.removed', removed);
  $('phNameInput').value = '';
  renderTable(); renderHolidays();
}

function legendEntries(){
  return [...document.querySelectorAll('#legendBody .legendRow')]
    .map(r => [r.querySelector('.code').value.trim(), r.querySelector('.desc').value.trim()])
    .filter(([c, d]) => c && d);
}

function render(){
  const m = S.model;
  // A name typed into the title box is remembered per staff number (exports often truncate names).
  $('titleInput').value = (m.staffId && store.get('names', {})[m.staffId]) || m.title;
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
  renderHolidays();
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
  const ext = (file.name.split('.').pop() || '').toLowerCase();
  S.source = ext === 'pdf' ? 'pdf' : ext === 'ics' ? 'ics' : /^html?$/.test(ext) ? 'html' : 'sheet';
  S.icsDescriptions = {};
  try{
    const buf = await file.arrayBuffer();
    if(S.source === 'pdf'){
      setBadge('Reading PDF…', true);
      S.wb = await pdfToWorkbook(buf);
    } else if(S.source === 'html'){
      S.wb = htmlToWorkbook(new TextDecoder('utf-8').decode(buf));
    } else if(S.source === 'ics'){
      const r = icsToWorkbook(new TextDecoder('utf-8').decode(buf));
      S.wb = r.wb; S.icsDescriptions = r.descriptions;
    } else {
      S.wb = XLSX.read(buf, {type:'array', cellDates:false, cellNF:false});
    }
  }catch(err){
    console.error(err);
    setBadge('Unreadable', false);
    log('Could not read this file: '+err.message, 'err');
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
  $('timePick').hidden = S.mode !== 'roster';
  $('srcPick').hidden = S.mode !== 'roster';
  S.warnings = [];
  try{
    if(S.mode === 'roster'){
      S.model = buildRosterModel(ws);
      setBadge(S.source === 'pdf' ? 'Crew roster · PDF' : S.source === 'ics' ? 'Crew roster · calendar' : S.source === 'html' ? 'Crew roster · web page' : 'Crew roster', true);
      const srcName = {base: (S.model.home || 'home base')+' time', utc: 'UTC', local: 'station local time'}[S.model.srcTime];
      log('Parsed '+S.model.stats[0][1]+' days. File times read as '+srcName+'; showing '+timeLabel(S.model.timeMode, S.model.home).replace('All times ', '')+'. Check the table, then press Download PDF.', 'ok');
      if(S.source === 'ics') log('Calendar files don\'t include aircraft, rank, work type or duty codes, and Block Hrs here are the scheduled flight times (credited hours for augmented crew can be lower). Use the Excel export for exact figures.', 'warn');
      S.warnings.forEach(w => log(w, 'warn'));
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
const pdfSafe = s => String(s || '').replace(/[•]/g, '·').replace(/[–—]/g, '-').replace(/[^\x20-\x7E\xA0-\xFF]/g, '');
/** File name: "Name, Staff number, Month Year" for rosters (e.g. "Shafiq Fakhri Othman, 2106537, November 2026"). */
const fileBase = () => {
  const name = ($('titleInput').value || '').trim();
  const parts = S.model.kind === 'roster' ? [name, S.model.staffId, S.model.monthLabel] : [name || 'Sheet'];
  return parts.filter(Boolean).join(', ').replace(/[\\/:*?"<>|\r\n\t]+/g, ' ').replace(/\s+/g, ' ').trim() || 'Roster';
};

/** Register DM Mono in a jsPDF doc; falls back to Courier if the font file is missing. */
function pdfFont(doc){
  const F = window.DM_MONO_TTF;
  if(!F) return {regular: ['courier', 'normal'], medium: ['courier', 'bold']};
  doc.addFileToVFS('DMMono-Regular.ttf', F.regular); doc.addFont('DMMono-Regular.ttf', 'DMMono', 'normal');
  doc.addFileToVFS('DMMono-Medium.ttf', F.medium);   doc.addFont('DMMono-Medium.ttf', 'DMMono', 'bold');
  return {regular: ['DMMono', 'normal'], medium: ['DMMono', 'bold']};
}

/** Build the PDF at a given scale (1 = normal). Smaller scales shrink text and spacing. */
function buildPDF(scale = 1){
  const {jsPDF} = window.jspdf;
  const cols = visibleColumns();
  const m = S.model;
  const k = scale;
  const paper = $('paperSelect').value;
  const fontSize = (cols.length > 14 ? 6 : cols.length > 11 ? 6.8 : 7.5) * k;
  let orient = $('orientSelect').value;
  if(orient === 'auto'){
    // Monospace width estimate at normal size; long free-text columns (hotel etc.) wrap, so cap them.
    const chW = (fontSize / k) * 0.6 * 0.3528;
    const chars = cols.reduce((a, c) => a + Math.min(16, Math.max(5, ...m.rows.map(r => (r.cells[c.key] || '').length))) + 1.6, 0);
    orient = chars * chW > (paper === 'letter' ? 200 : 194) ? 'landscape' : 'portrait';
  }
  const doc = new jsPDF({orientation: orient, unit: 'mm', format: paper});
  const F = pdfFont(doc);
  const W = doc.internal.pageSize.getWidth(), PH = doc.internal.pageSize.getHeight(), M = 10;
  const INK = [17,17,17], DIM = [110,116,122];
  // Font colour only — no background highlights anywhere.
  const texts = {duty:[29,78,216], off:[63,120,0], layover:[109,40,217]};
  const PINK = [219,39,119], ORANGE = [204,102,0];
  const HILITE = [255,226,92];          // highlighter yellow behind report times (the only highlight)
  const H = holidays();
  const padX = 1.4 * k, padY = 1.05 * k;

  // Title line: name in medium weight, details in grey beside / under it.
  let y = 9 + 5 * k;
  const title = pdfSafe($('titleInput').value || m.title);
  doc.setFont(...F.medium); doc.setFontSize(12 * Math.max(k, 0.8)); doc.setTextColor(...INK);
  doc.text(title, M, y);
  const tw = doc.getTextWidth(title);
  doc.setFont(...F.regular); doc.setFontSize(8 * Math.max(k, 0.8)); doc.setTextColor(...DIM);
  const subText = pdfSafe(m.subtitle);
  if(tw + 4 + doc.getTextWidth(subText) <= W - 2*M) doc.text(subText, M + tw + 4, y);
  else { const sub = doc.splitTextToSize(subText, W - 2*M); y += 4.2; doc.text(sub, M, y); y += (sub.length - 1) * 3.2; }
  y += 5 * k;

  if(m.kind === 'roster'){
    const key = [[texts.duty, 'Duty'], [texts.off, 'Off / Leave'], [texts.layover, 'Layover'], [PINK, 'Public holiday'], [ORANGE, 'Weekend'],
                 [INK, timeLabel(m.timeMode, m.home).replace('All times', 'Times:')]];
    let x = M;
    doc.setFont(...F.regular); doc.setFontSize(6.8 * Math.max(k, 0.85));
    key.forEach(([color, label]) => {
      doc.setTextColor(...color); doc.text(label, x, y);
      x += doc.getTextWidth(label) + 6;
    });
    // Report time sample: same highlighter mark as in the table.
    const lbl = 'Report time', lw = doc.getTextWidth(lbl), fh = doc.getFontSize() * 0.3528;
    doc.setFillColor(...HILITE); doc.roundedRect(x - 0.9, y - fh * 0.95, lw + 1.8, fh * 1.3, 0.7, 0.7, 'F');
    doc.setTextColor(...INK); doc.text(lbl, x, y);
    y += 3.2 * k;
  }

  doc.autoTable({
    startY: y,
    margin: {left: M, right: M, top: 10, bottom: 10},
    head: [cols.map(c => pdfSafe(c.label))],
    body: m.rows.map(r => cols.map(c => pdfSafe(r.cells[c.key]))),
    theme: 'plain',
    styles: {font: F.regular[0], fontStyle: F.regular[1], fontSize, cellPadding: {top: padY, bottom: padY, left: padX, right: padX},
             textColor: INK, overflow: 'linebreak', valign: 'top', lineWidth: 0},
    headStyles: {font: F.medium[0], fontStyle: F.medium[1], fillColor: false, textColor: [85,90,96], valign: 'bottom',
                 lineWidth: {bottom: 0.25}, lineColor: [190,195,200]},
    // Short columns never wrap (keeps "KUL 19:00" on one line); only long free text (hotel…) wraps.
    columnStyles: Object.fromEntries(cols.map((c, i) => {
      const long = m.rows.some(r => (r.cells[c.key] || '').length > 14);
      return [i, Object.assign({cellWidth: long ? 'auto' : 'wrap'}, c.num ? {halign:'right'} : {})];
    })),
    didParseCell: data => {
      if(data.section !== 'body') return;
      const row = m.rows[data.row.index], st = data.cell.styles;
      const c = ['off', 'layover', 'duty'].find(x => row.cls.includes(x));
      if(c) st.textColor = texts[c];
      const key = cols[data.column.index].key;
      if(key === 'date' && row.cls.includes('weekend')) st.textColor = ORANGE;
      if(key === 'date' && row.ymd && H[row.ymd]){ st.textColor = PINK; st.font = F.medium[0]; st.fontStyle = F.medium[1]; }
      // Report time: bold, black and one size up so it stands out from the blue duty times.
      // Report time: bold black on a highlighter mark (drawn in willDrawCell).
      if(key === 'start' && data.cell.raw && !row.cls.includes('off')){ st.font = F.medium[0]; st.fontStyle = F.medium[1]; st.textColor = INK; data.cell.rpt = true; }
      if(key === 'dep' && row.cls.includes('earlyrpt')){ st.font = F.medium[0]; st.fontStyle = F.medium[1]; st.textColor = INK; }
      // No lines or fills: a little extra space above each new day keeps days apart.
      if(row.cls.includes('daystart') && data.row.index > 0) st.cellPadding = {top: padY * 1.8, bottom: padY, left: padX, right: padX};
    },
    willDrawCell: data => {
      if(data.section !== 'body' || !data.cell.rpt) return;
      const c = data.cell, fs = c.styles.fontSize, txt = c.text.join(' ');
      doc.setFont(F.medium[0], F.medium[1]); doc.setFontSize(fs);
      const w = doc.getTextWidth(txt), h = fs * 0.3528 * 1.15, padH = 0.9 * k;
      doc.setFillColor(...HILITE);
      doc.roundedRect(c.x + c.padding('left') - padH, c.y + c.padding('top') - 0.35 * k, w + 2 * padH, h + 0.7 * k, 0.8 * k, 0.8 * k, 'F');
    },
  });
  y = doc.lastAutoTable.finalY + 6 * k;

  // Bottom block: Statistics | Code Description | Public Holidays side by side, to keep it on one page.
  const blocks = [];
  if(m.stats.length) blocks.push(['Statistics', m.stats, null]);
  const legend = legendEntries();
  if(legend.length) blocks.push(['Code Description', legend, null]);
  if(m.kind === 'roster'){
    const ph = holidaysInRange().map(([d, n]) => { const dt = new Date(d+'T00:00:00'); return [fmtDate(dt)+' '+DOW[dt.getDay()], n]; });
    if(ph.length) blocks.push(['Public Holidays', ph, PINK]);
  }
  if(blocks.length){
    const gap = 6, colW = (W - 2*M - gap * (blocks.length - 1)) / blocks.length;
    let bottom = y;
    blocks.forEach(([heading, rows, firstColor], i) => {
      const x = M + i * (colW + gap);
      doc.setFont(...F.medium); doc.setFontSize(8.5 * Math.max(k, 0.8)); doc.setTextColor(...INK);
      doc.text(heading, x, y + 3 * k);
      doc.autoTable({
        startY: y + 4.5 * k, margin: {left: x, right: W - x - colW, top: 10, bottom: 10}, tableWidth: colW,
        body: rows.map(r => r.map(pdfSafe)), theme: 'plain',
        styles: {font: F.regular[0], fontStyle: F.regular[1], fontSize: 7 * k, cellPadding: {top: 0.6 * k, bottom: 0.6 * k, left: 0, right: 3},
                 textColor: INK, overflow: 'linebreak'},
        columnStyles: {0: {cellWidth: 'wrap', textColor: firstColor || DIM}, 1: heading === 'Statistics' ? {font: F.medium[0], fontStyle: F.medium[1], halign: 'right'} : {}},
      });
      bottom = Math.max(bottom, doc.lastAutoTable.finalY);
    });
    y = bottom;
  }

  const pages = doc.getNumberOfPages();
  for(let p = 1; p <= pages; p++){
    doc.setPage(p);
    doc.setFont(...F.regular); doc.setFontSize(6); doc.setTextColor(150);
    doc.text('Generated '+fmtDate(new Date()), M, PH - 4.5);
    if(pages > 1) doc.text(p+' / '+pages, W - M, PH - 4.5, {align: 'right'});
  }
  return doc;
}

/** Shrink step by step until the whole roster fits on one page (stops at a still-readable size). */
function buildOnePagePDF(){
  let doc = null;
  for(let k = 1; k >= 0.6; k = Math.round((k - 0.04) * 100) / 100){
    doc = buildPDF(k);
    if(doc.getNumberOfPages() === 1) return {doc, fit: true, scale: k};
  }
  return {doc, fit: false, scale: 0.6};
}

function exportPDF(){
  const {doc, fit} = buildOnePagePDF();
  if(!fit) log('This roster is too long to fit on one page at a readable size, so it continues onto page 2. Try Landscape, or hide some columns.', 'warn');
  doc.save(fileBase()+'.pdf');
}

/** Print the same one-page PDF (opens it in a new tab with the print dialog). */
function printPDF(){
  const {doc} = buildOnePagePDF();
  doc.setProperties({title: fileBase()});
  doc.autoPrint();
  const url = doc.output('bloburl');
  const w = window.open(url, '_blank');
  if(!w){ // pop-up blocked: fall back to the browser's own page printing
    const t = document.title; document.title = fileBase(); window.print(); document.title = t;
  }
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
  // The browser's "Save as PDF" uses the page title as the file name.
  $('printBtn').onclick = () => { try{ printPDF(); }catch(err){ console.error(err); log('Print failed: '+err.message, 'err'); } };
  $('titleInput').addEventListener('change', () => {
    const m = S.model; if(!m || !m.staffId) return;
    const names = store.get('names', {}), v = $('titleInput').value.trim();
    if(v && v !== m.title) names[m.staffId] = v; else delete names[m.staffId];
    store.set('names', names);
  });
  $('csvBtn').onclick = exportCSV;
  $('columnsBtn').onclick = () => { $('columnsPanel').hidden = !$('columnsPanel').hidden; };
  $('addCodeBtn').onclick = () => addLegendRow();
  $('clearSettingsBtn').onclick = () => {
    if(!confirm('Clear all settings saved in this browser (hidden columns, code descriptions, holiday edits, time mode, page size and saved names)?')) return;
    try{ Object.keys(localStorage).filter(k => k.startsWith('rosterConvert.')).forEach(k => localStorage.removeItem(k)); }catch(e){}
    location.reload();
  };
  $('addHolidayBtn').onclick = addHoliday;
  $('timeSelect').value = store.get('timeSelect', 'local');
  $('appVersion').textContent = APP_VERSION;
  $('srcSelect').onchange = () => { if(S.wb && S.mode === 'roster') loadSheet('roster'); };
  $('timeSelect').onchange = e => { store.set('timeSelect', e.target.value); if(S.wb && S.mode === 'roster') loadSheet('roster'); };
  ['orientSelect', 'paperSelect'].forEach(id => {
    const el = $(id); el.value = store.get(id, el.value);
    el.onchange = () => store.set(id, el.value);
  });
  if(!window.XLSX || !window.jspdf) log('A library failed to load — make sure the vendor/ folder was deployed with index.html.', 'err');
})();
