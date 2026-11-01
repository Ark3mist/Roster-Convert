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
  AD:'Available Day', SB:'Standby', SBY:'Standby', RES:'Reserve', TRG:'Training', SIM:'Simulator',
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
/** One name per clock, used everywhere (menus, messages, PDF). */
const CLOCK_NAME = {base: 'Home base KUL (UTC+8)', local: 'Station local (LT)', utc: 'UTC'};
const TIME_LABEL = {local: 'All times Station local (LT)', utc: 'All times UTC',
  base: home => 'All times Home base '+(home || 'KUL')+' ('+utcOffsetLabel(STATION_TZ[home] || 'Asia/Kuala_Lumpur')+')'};
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
  {key:'work',  label:'Work Type', center:true},
  {key:'block', label:'Block Hrs', center:true},
  {key:'duty',  label:'Duty Hrs',  center:true},
  {key:'code',  label:'Duty Code', center:true},
  {key:'ac',    label:'A/C',       center:true},
  {key:'hotel', label:'Hotel'},
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
  if(srcChoice === 'auto' && parsed.srcTime !== 'utc'){
    // Evidence from the flights decides; a "(LT)" header is only a tie-breaker when the flights can't tell.
    const v = clockVotes(legs, home);
    const d = detectClock(legs, home);
    srcTime = (parsed.srcTime === 'local' && !v.same && !v.local) ? 'local' : d.clock;
    if(srcTime !== 'base') S.warnings.push('This file\'s times are '+(srcTime === 'utc' ? 'in UTC' : 'already in each airport\'s local time')+' (detected: '+(d.why || 'column header says LT')+'). They have been converted accordingly.');
  }
  // (hidden) keeps the detected clock for reference.
  $('srcSelect').options[0].textContent = 'Auto-detect' + (srcChoice === 'auto' ? ' → ' + CLOCK_NAME[srcTime] : '');
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
      // Arrival stays on the flight's own line, marked (+1) when it lands the next day
      // (e.g. Home base: MH727 KUL 22:35 → CGK 00:45(+1); Station local: → CGK 23:45).
      arr: st(l.arrStn, rel(l.arrTime, l.arrDate, l.depDate)),
      end: rel(l.debTime, l.debDate, l.depDate),
      block: l.fly, duty: l.duty,
    }));
    // Landed back at base after midnight: the arrival day keeps a short arrival/debrief line, so it is
    // part of the duty (not shown as an available day). At an outstation the next day is a layover instead.
    if(ymd(l.arrDate) !== ymd(l.depDate) && l.arrStn && l.arrStn === home){
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
    // Days at the outstation before the next departure — including the arrival day itself when the
    // flight landed after midnight (its arrival is shown on the previous day's line with (+1)).
    const firstAway = ymd(l.arrDate) !== ymd(l.depDate) ? l.arrDate : addDays(l.arrDate, 1);
    for(let c = new Date(firstAway); c < next.depDate; c = addDays(c, 1)){
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
  let blockMin = 0, dutyMin = 0, sectors = 0, offBase = 0, offAway = 0, workDays = 0, availDays = 0;
  const codes = new Set();
  const stnOf = v => { const m = /^([A-Z]{3,4})\b/.exec(String(v || '')); return m ? m[1] : ''; };
  let lastStn = home;                      // where the crew member is: last arrival / station seen
  for(let c = new Date(minD); c <= maxD; c = addDays(c, 1)){
    const d = days[ymd(c)] || {date: c, rows: []};
    // No code at all for the day → AD (available day) at the last station from the previous duty.
    const rs = d.rows.length ? d.rows : [{item: 'AD', dep: lastStn, avail: true}];
    d.rows.forEach(r => { lastStn = stnOf(r.arr) || stnOf(r.dep) || lastStn; });
    const weekend = c.getDay() === 0 || c.getDay() === 6;
    const isDutyRow = r => !r.avail && !isOffCode(r.item) && r.item !== 'LO' && !!(r.item || r.start || r.dep || r.arr || r.end);
    rs.forEach((r, ri) => {
      const cls = [];
      if(ri === 0) cls.push('daystart'); else cls.push('leg2');
      if(ri === 0 && isOffCode(r.item)) cls.push('off');
      else if(ri === 0 && r.item === 'LO') cls.push('layover');
      else if(r.avail) cls.push('avail');
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
    if(first.avail) availDays++;
    else if(rs.length === 1 && isOffCode(first.item)){
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
      ['Available Days (AD)', String(availDays)],
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
// Duty Code starts hidden for new users; once the user ticks/unticks columns, their choice is kept.
const DEFAULT_HIDDEN = {roster: ['code'], raw: []};
// Opt-in columns stay hidden for everyone, including users whose column choices were saved before
// the column existed, until they tick it. Kept in a separate "shown" list for that reason.
const OPT_IN = {roster: ['hotel'], raw: []};
function shownOptIn(){ return new Set(store.get('shown.'+S.model.kind, [])); }
function hiddenCols(){
  const v = store.get('hidden.'+S.model.kind, null), shown = shownOptIn();
  const hid = new Set(v === null ? DEFAULT_HIDDEN[S.model.kind] || [] : v);
  (OPT_IN[S.model.kind] || []).forEach(k => shown.has(k) ? hid.delete(k) : hid.add(k));
  return hid;
}
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
      const kind = S.model.kind;
      if((OPT_IN[kind] || []).includes(c.key)){
        const sh = shownOptIn(); cb.checked ? sh.add(c.key) : sh.delete(c.key);
        store.set('shown.'+kind, [...sh]);
      } else {
        const h = hiddenCols(); (OPT_IN[kind] || []).forEach(k => h.delete(k));
        cb.checked ? h.delete(c.key) : h.add(c.key);
        store.set('hidden.'+kind, [...h]);
      }
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
  cols.forEach(c => { const th = document.createElement('th'); th.textContent = c.label; if(c.center) th.className = 'center'; hr.appendChild(th); });
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
      if(c.center) td.className = 'center';
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
  $('srcSelect').value = 'auto';              // a manual override applies to one file only
  $('privacyHead').classList.remove('blink');   // stop drawing attention once a roster is open
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
  S.warnings = [];
  try{
    if(S.mode === 'roster'){
      S.model = buildRosterModel(ws);
      setBadge(S.source === 'pdf' ? 'Crew roster · PDF' : S.source === 'ics' ? 'Crew roster · calendar' : S.source === 'html' ? 'Crew roster · web page' : 'Crew roster', true);
      log('Parsed '+S.model.stats[0][1]+' days. File detected as '+CLOCK_NAME[S.model.srcTime]+'; showing '+CLOCK_NAME[S.model.timeMode]+'. Check the table, then press Download PDF.', 'ok');
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
function pdfFont(doc, T){
  // Theme font (Minecraft / Roblox style) when its file has been loaded; otherwise DM Mono.
  const tf = T && T.font, TF = tf && window[tf.global];
  if(TF){
    doc.addFileToVFS(tf.family+'-Regular.ttf', TF.regular); doc.addFont(tf.family+'-Regular.ttf', tf.family, 'normal');
    doc.addFileToVFS(tf.family+'-Medium.ttf', TF.medium);   doc.addFont(tf.family+'-Medium.ttf', tf.family, 'bold');
    return {regular: [tf.family, 'normal'], medium: [tf.family, 'bold']};
  }
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
  const F = pdfFont(doc, currentTheme());
  const W = doc.internal.pageSize.getWidth(), PH = doc.internal.pageSize.getHeight(), M = 10;
  const T = currentTheme();
  const INK = T.ink, DIM = T.dim;
  // Page background for dark/tinted themes (painted once per page, before anything else).
  const painted = new Set();
  const paintBg = () => { const n = doc.internal.getCurrentPageInfo().pageNumber; if(!T.bg || painted.has(n)) return; painted.add(n); doc.setFillColor(...T.bg); doc.rect(0, 0, W, PH, 'F'); };
  paintBg();
  // Font colour only — no background highlights anywhere.
  const texts = {duty: T.duty, off: T.off, layover: T.layover};
  const PINK = T.ph, ORANGE = T.weekend;
  const HILITE = T.hilite;
  const pageBg = T.bg || [255,255,255];
  const rowLines = $('rowLinesChk').checked;
  const ROW_LINE = T.rule.map((c, i) => Math.round(c * 0.45 + pageBg[i] * 0.55));   // faint hairline colour              // highlighter behind report times (the only highlight)
  const H = holidays();
  const padX = 1.4 * k, padY = 1.05 * k;

  // Title line: name in medium weight, details in grey beside / under it.
  let y = 9 + 5 * k;
  const title = pdfSafe($('titleInput').value || m.title);
  doc.setFont(...F.medium); doc.setFontSize(12 * Math.max(k, 0.8)); doc.setTextColor(...INK);
  // Theme badge (if any) just before the name, sized to the title's cap height.
  const capH = doc.getFontSize() * 0.3528 * 0.8, bsz = BADGES[$('themeSelect').value] ? capH * 1.35 : 0;
  if(bsz) drawBadgePDF(doc, $('themeSelect').value, M, y - capH - (bsz - capH) / 2, bsz);
  const tx = M + (bsz ? bsz + 2 : 0);
  doc.setTextColor(...INK);
  doc.text(title, tx, y);
  const tw = doc.getTextWidth(title) + (tx - M);
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
    doc.setTextColor(...T.hiliteInk); doc.text(lbl, x, y);
    y += 3.2 * k;
  }

  // 'wrap' sizes every column to its content and never limits the table to the page, so long free
  // text (hotel names) would run off the right edge. Measure the columns: short ones keep their
  // natural width, long ones share what is left of the page width and wrap inside it.
  const isLong = c => m.rows.some(r => (r.cells[c.key] || '').length > 14);
  doc.setFontSize(fontSize);
  const natW = cols.map(c => {
    doc.setFont(...F.medium); let w = doc.getTextWidth(pdfSafe(c.label));
    doc.setFont(...F.regular); m.rows.forEach(r => { w = Math.max(w, doc.getTextWidth(pdfSafe(r.cells[c.key]))); });
    return w * 1.06 + 2 * padX + 0.6;      // a little slack for bold report times and rounding
  });
  const longIdx = cols.map((c, i) => isLong(c) ? i : -1).filter(i => i >= 0);
  const room = W - 2*M - natW.reduce((a, w, i) => longIdx.includes(i) ? a : a + w, 0);
  const longW = {};
  longIdx.slice().sort((a, b) => natW[a] - natW[b]).forEach((i, n, arr) => {
    const share = Math.max(14 * k, (room - Object.values(longW).reduce((a, w) => a + w, 0)) / (arr.length - n));
    longW[i] = Math.min(natW[i], share);
  });
  doc.autoTable({
    startY: y,
    margin: {left: M, right: M, top: 10, bottom: 10},
    tableWidth: 'wrap',                     // compact: columns only as wide as their content
    head: [cols.map(c => pdfSafe(c.label))],
    body: m.rows.map(r => cols.map(c => pdfSafe(r.cells[c.key]))),
    theme: 'plain',
    styles: {font: F.regular[0], fontStyle: F.regular[1], fontSize, cellPadding: {top: padY, bottom: padY, left: padX, right: padX},
             textColor: INK, overflow: 'linebreak', valign: 'top', lineWidth: 0},
    headStyles: {font: F.medium[0], fontStyle: F.medium[1], fillColor: false, textColor: T.head, valign: 'bottom',
                 lineWidth: {bottom: 0.25}, lineColor: T.rule},
    willDrawPage: paintBg,
    // Short columns never wrap (keeps "KUL 19:00" on one line); only long free text (hotel…) wraps.
    columnStyles: Object.fromEntries(cols.map((c, i) => {
      return [i, Object.assign({cellWidth: i in longW ? longW[i] : 'wrap'}, c.center ? {halign:'center'} : c.num ? {halign:'right'} : {})];
    })),
    didParseCell: data => {
      if(data.section === 'head' && cols[data.column.index].center) data.cell.styles.halign = 'center';
      if(data.section !== 'body') return;
      const row = m.rows[data.row.index], st = data.cell.styles;
      const c = ['off', 'layover', 'duty'].find(x => row.cls.includes(x));
      if(c) st.textColor = texts[c];
      const key = cols[data.column.index].key;
      if(key === 'date' && row.cls.includes('weekend')) st.textColor = ORANGE;
      if(key === 'date' && row.ymd && H[row.ymd]){ st.textColor = PINK; st.font = F.medium[0]; st.fontStyle = F.medium[1]; }
      // Report time: bold, black and one size up so it stands out from the blue duty times.
      // Report time: bold black on a highlighter mark (drawn in willDrawCell).
      if(key === 'start' && data.cell.raw && !row.cls.includes('off')){ st.font = F.medium[0]; st.fontStyle = F.medium[1]; st.textColor = T.hiliteInk; data.cell.rpt = true; }
      if(key === 'dep' && row.cls.includes('earlyrpt')){ st.font = F.medium[0]; st.fontStyle = F.medium[1]; st.textColor = INK; }
      if(rowLines){
        // A very thin line under every row (soft tone of the theme's rule colour).
        st.lineWidth = {top: 0, right: 0, left: 0, bottom: 0.1}; st.lineColor = ROW_LINE;
      } else if(row.cls.includes('daystart') && data.row.index > 0){
        // No lines: a little extra space above each new day keeps days apart.
        st.cellPadding = {top: padY * 1.8, bottom: padY, left: padX, right: padX};
      }
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

  // Bottom block (compact): Statistics | Code | PH side by side, each only as wide as its content.
  const blocks = [];
  if(m.stats.length){
    // Statistics in two label/value pairs per line to halve the height.
    const half = Math.ceil(m.stats.length / 2), rows = [];
    for(let i = 0; i < half; i++) rows.push([...m.stats[i], ...(m.stats[i + half] || ['', ''])]);
    blocks.push({title: 'Statistics', rows, stats: true});
  }
  const legend = legendEntries();
  if(legend.length) blocks.push({title: 'Code', rows: legend});
  if(m.kind === 'roster'){
    const ph = holidaysInRange().map(([d, n]) => { const dt = new Date(d+'T00:00:00'); return [String(dt.getDate()).padStart(2,'0')+' '+MON_UP[dt.getMonth()]+' '+DOW[dt.getDay()], n]; });
    if(ph.length) blocks.push({title: 'PH', rows: ph, firstColor: PINK});
  }
  if(blocks.length){
    const gap = 5 * k, fs = 6.6 * k;
    let x = M, bottom = y;
    doc.setFont(...F.medium); doc.setFontSize(7.5 * Math.max(k, 0.8)); doc.setTextColor(...INK);
    for(const b of blocks){
      if(x > W - M - 20) break;
      doc.setFont(...F.medium); doc.setFontSize(7.5 * Math.max(k, 0.8)); doc.setTextColor(...INK);
      doc.text(b.title, x, y + 2.6 * k);
      const valueCol = {font: F.medium[0], fontStyle: F.medium[1], halign: 'right', textColor: INK, cellPadding: {top: 0.35 * k, bottom: 0.35 * k, left: 0, right: 5 * k}};
      doc.autoTable({
        startY: y + 3.6 * k, margin: {left: x, right: 0, top: 10, bottom: 10}, tableWidth: 'wrap', willDrawPage: paintBg,
        body: b.rows.map(r => r.map(pdfSafe)), theme: 'plain',
        styles: {font: F.regular[0], fontStyle: F.regular[1], fontSize: fs, cellPadding: {top: 0.35 * k, bottom: 0.35 * k, left: 0, right: 2.5 * k},
                 textColor: INK, overflow: 'visible', cellWidth: 'wrap'},
        columnStyles: b.stats ? {0: {textColor: DIM}, 1: valueCol, 2: {textColor: DIM}, 3: valueCol}
                              : {0: {textColor: b.firstColor || DIM}},
      });
      const t = doc.lastAutoTable;
      bottom = Math.max(bottom, t.finalY);
      x += t.columns.reduce((a, c) => a + c.width, 0) + gap;
    }
    y = bottom;
  }

  const footerLine = pdfSafe('Generated '+fmtDate(new Date())+' / '+(($('titleInput').value || m.title).trim())+' ('+deviceName()+') ~by LipasTerbang');
  const pages = doc.getNumberOfPages();
  for(let p = 1; p <= pages; p++){
    doc.setPage(p);
    doc.setFont(...F.regular); doc.setFontSize(6); doc.setTextColor(...T.footer);
    doc.text(footerLine, M, PH - 4.5);
    if(pages > 1) doc.text(p+' / '+pages, W - M, PH - 4.5, {align: 'right'});
  }
  return doc;
}

/* ------------------------------------------------------------------ *
 *  Colour themes (PDF, print and the on-screen roster)
 * ------------------------------------------------------------------ */
const THEMES = {
  default:   {name: 'Default',         bg: null,           ink: [17,17,17],  dim: [110,116,122], head: [85,90,96],   rule: [190,195,200],
              duty: [29,78,216],  off: [63,120,0],   layover: [109,40,217], ph: [219,39,119], weekend: [204,102,0],  hilite: [255,226,92], hiliteInk: [17,17,17], footer: [150,150,150]},
  bw:        {name: 'Black & White',   bg: null,           ink: [0,0,0],     dim: [90,90,90],    head: [60,60,60],   rule: [120,120,120],
              duty: [0,0,0],      off: [95,95,95],   layover: [0,0,0],      ph: [0,0,0],      weekend: [95,95,95],   hilite: [212,212,212], hiliteInk: [0,0,0], footer: [130,130,130]},
  dark:      {name: 'Dark',            bg: [17,19,22],     ink: [240,242,245], dim: [160,166,173], head: [160,166,173], rule: [70,75,80],
              duty: [125,178,255], off: [198,255,61], layover: [196,181,253], ph: [255,140,198], weekend: [255,180,84], hilite: [255,226,92], hiliteInk: [17,17,17], footer: [120,125,130]},
  minecraft: {name: 'Minecraft',       font: {family: 'PixelifySans', global: 'FONT_PIXELIFY', src: 'vendor/fonts/pixelify.js', css: "'Pixelify Sans', monospace"},
              bg: [244,236,214],  ink: [60,44,30],  dim: [125,105,85],  head: [90,70,50],   rule: [134,96,67],
              duty: [0,120,120],  off: [76,135,40],  layover: [31,62,158],  ph: [170,0,0],    weekend: [134,96,67],  hilite: [252,219,5], hiliteInk: [40,30,20], footer: [150,130,110]},
  roblox:    {name: 'Roblox',          font: {family: 'Montserrat', global: 'FONT_MONTSERRAT', src: 'vendor/fonts/montserrat.js', css: "'Montserrat', sans-serif"},
              bg: null,           ink: [27,42,53],  dim: [99,95,98],    head: [99,95,98],   rule: [163,162,165],
              duty: [13,105,172], off: [75,151,75],  layover: [107,50,124], ph: [196,40,28],  weekend: [218,133,65], hilite: [245,205,48], hiliteInk: [27,42,53], footer: [163,162,165]},
  // F1: black/red with F1 timing-screen colours (purple = fastest lap, green = personal best, yellow flag).
  f1:        {name: 'F1',              font: {family: 'TitilliumWeb', global: 'FONT_TITILLIUM', src: 'vendor/fonts/titillium.js', css: "'Titillium Web', sans-serif"},
              bg: null,           ink: [21,21,30],  dim: [105,105,115], head: [21,21,30],   rule: [225,6,0],
              duty: [21,21,30],   off: [0,150,70],   layover: [130,40,200], ph: [225,6,0],    weekend: [255,128,0],  hilite: [255,214,0], hiliteInk: [21,21,30], footer: [148,148,152]},
  // MotoGP: asphalt grey page, MotoGP red rule, factory-team blue duties, flag colours.
  motogp:    {name: 'MotoGP',          font: {family: 'BarlowCondensed', global: 'FONT_BARLOWCOND', src: 'vendor/fonts/barlowcond.js', css: "'Barlow Condensed', sans-serif"},
              bg: [240,240,238],  ink: [17,17,17],  dim: [100,100,100], head: [40,40,40],   rule: [215,24,42],
              duty: [0,62,165],   off: [0,140,60],   layover: [115,35,150], ph: [215,24,42],  weekend: [235,100,0],  hilite: [255,214,0], hiliteInk: [17,17,17], footer: [140,140,140]},
};
/* Original 8×8 pixel-art badges for the themed palettes (not the brands' logos, which are trademarks). */
const BADGES = {
  f1: {colors: {K: [21,21,30], W: [255,255,255], R: [225,6,0]}, rows: [        // chequered flag
    'RKWKWKW.', 'RWKWKWK.', 'RKWKWKW.', 'RWKWKWK.', 'R.......', 'R.......', 'R.......', 'R.......']},
  motogp: {colors: {R: [215,24,42], W: [255,255,255], D: [30,34,44], K: [60,60,60]}, rows: [   // racing helmet
    '..RRRR..', '.RRWWRR.', 'RRWRRRRR', 'RDDDDDRR', 'RDDDDDRR', 'RRRRRRRR', '.RRRRRW.', '..KKKK..']},
  minecraft: {colors: {G: [93,158,63], g: [126,190,84], B: [134,96,67], b: [100,70,48]}, rows: [  // grass block
    'GgGGgGGg', 'GGgGGGgG', 'gGbGGbGG', 'BBBbBBBB', 'BbBBBBbB', 'BBBBbBBB', 'BBbBBBBb', 'bBBBBbBB']},
  roblox: {colors: {Y: [245,205,48], K: [27,42,53]}, rows: [                   // blocky smiley head
    'YYYYYYYY', 'YYYYYYYY', 'YKKYYKKY', 'YKKYYKKY', 'YYYYYYYY', 'YKYYYYKY', 'YYKKKKYY', 'YYYYYYYY']},
};
/** Draw a badge into a jsPDF doc at (x, y) with the given size in mm. */
function drawBadgePDF(doc, id, x, y, size){
  const b = BADGES[id]; if(!b) return 0;
  const px = size / 8;
  b.rows.forEach((row, r) => [...row].forEach((ch, c) => {
    const col = b.colors[ch]; if(!col) return;
    doc.setFillColor(...col); doc.rect(x + c * px, y + r * px, px + 0.02, px + 0.02, 'F');
  }));
  return size;
}
/** Same badge as an inline SVG element for the page. */
function badgeSVG(id, sizePx){
  const b = BADGES[id]; if(!b) return null;
  const NS = 'http://www.w3.org/2000/svg', svg = document.createElementNS(NS, 'svg');
  svg.setAttribute('viewBox', '0 0 8 8'); svg.setAttribute('width', sizePx); svg.setAttribute('height', sizePx);
  svg.setAttribute('shape-rendering', 'crispEdges'); svg.setAttribute('aria-hidden', 'true'); svg.classList.add('themeBadge');
  b.rows.forEach((row, r) => [...row].forEach((ch, c) => {
    const col = b.colors[ch]; if(!col) return;
    const rect = document.createElementNS(NS, 'rect');
    rect.setAttribute('x', c); rect.setAttribute('y', r); rect.setAttribute('width', 1); rect.setAttribute('height', 1);
    rect.setAttribute('fill', 'rgb('+col.join(',')+')'); svg.appendChild(rect);
  }));
  return svg;
}

const currentTheme = () => THEMES[$('themeSelect') ? $('themeSelect').value : 'default'] || THEMES.default;

/** Mirror the PDF theme on the on-screen roster (default keeps the app's own dark styling). */
function applyScreenTheme(){
  const id = $('themeSelect').value, T = THEMES[id] || THEMES.default, area = $('printArea');
  const vars = ['--panel','--panel2','--text','--dim','--line','--duty-text','--off-text','--lo-text','--ph-text','--warn','--rpt-bg','--rpt-text'];
  vars.forEach(v => area.style.removeProperty(v));
  area.style.fontFamily = T.font ? T.font.css : '';
  const slot = $('themeLogo'); slot.innerHTML = '';
  const svg = badgeSVG(id, 18); if(svg) slot.appendChild(svg);
  slot.hidden = !svg;
  loadThemeFont(T).catch(() => {});
  if(id === 'default') return;
  const rgb = a => 'rgb('+a.join(',')+')';
  const bg = T.bg || [255,255,255], bg2 = bg.map(c => Math.round(c * 0.94));
  const set = {'--panel': bg, '--panel2': bg2, '--text': T.ink, '--dim': T.dim, '--line': T.rule, '--duty-text': T.duty, '--off-text': T.off,
               '--lo-text': T.layover, '--ph-text': T.ph, '--warn': T.weekend, '--rpt-bg': T.hilite, '--rpt-text': T.hiliteInk};
  Object.entries(set).forEach(([k, v]) => area.style.setProperty(k, rgb(v)));
}

/**
 * Device type and browser, e.g. "iPhone · Safari", "Android phone · Chrome", "Windows PC · Edge".
 * Browsers don't reveal a device's own name (e.g. "Shafiq's iPhone"), so this is the closest available.
 * Worked out locally from the user-agent; nothing is sent anywhere.
 */
function deviceName(){
  const ua = navigator.userAgent || '';
  const touchMac = /Macintosh/.test(ua) && navigator.maxTouchPoints > 1;          // iPadOS reports as a Mac
  const device = /iPhone/.test(ua) ? 'iPhone' : (/iPad/.test(ua) || touchMac) ? 'iPad'
    : /Android/.test(ua) ? (/Mobile/.test(ua) ? 'Android phone' : 'Android tablet')
    : /Windows/.test(ua) ? 'Windows PC' : /Macintosh|Mac OS X/.test(ua) ? 'Mac' : /CrOS/.test(ua) ? 'Chromebook'
    : /Linux/.test(ua) ? 'Linux PC' : 'Device';
  const browser = /Edg\//.test(ua) ? 'Edge' : /OPR\/|Opera/.test(ua) ? 'Opera' : /SamsungBrowser/.test(ua) ? 'Samsung Internet'
    : /Firefox|FxiOS/.test(ua) ? 'Firefox' : /CriOS|Chrome\//.test(ua) ? 'Chrome' : /Safari\//.test(ua) ? 'Safari' : '';
  return browser ? device+' · '+browser : device;
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

/** Load a theme's PDF font file on demand (only Minecraft / Roblox need one). */
const themeFontLoads = {};
function loadThemeFont(T){
  const tf = T && T.font;
  if(!tf || window[tf.global]) return Promise.resolve();
  return themeFontLoads[tf.global] ||= new Promise((resolve, reject) => {
    const s = document.createElement('script');
    s.src = tf.src+'?v='+APP_VERSION;
    s.onload = resolve;
    s.onerror = () => { delete themeFontLoads[tf.global]; reject(new Error('Theme font failed to load')); };
    document.head.appendChild(s);
  });
}

async function exportPDF(){
  await loadThemeFont(currentTheme()).catch(() => log('Theme font could not be loaded, using the default font.', 'warn'));
  const {doc, fit} = buildOnePagePDF();
  if(!fit) log('This roster is too long to fit on one page at a readable size, so it continues onto page 2. Try Landscape, or hide some columns.', 'warn');
  doc.save(fileBase()+'.pdf');
}

/** Print the same one-page PDF (opens it in a new tab with the print dialog). */
async function printPDF(){
  await loadThemeFont(currentTheme()).catch(() => {});
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
  $('pdfBtn').onclick = () => exportPDF().catch(err => { console.error(err); log('PDF export failed: '+err.message, 'err'); });
  // The browser's "Save as PDF" uses the page title as the file name.
  $('printBtn').onclick = () => printPDF().catch(err => { console.error(err); log('Print failed: '+err.message, 'err'); });
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
  $('themeSelect').value = store.get('themeSelect', 'default');
  if(!THEMES[$('themeSelect').value]) $('themeSelect').value = 'default';
  $('themeSelect').onchange = e => { store.set('themeSelect', e.target.value); applyScreenTheme(); };
  applyScreenTheme();
  // Row lines option (on by default), remembered on this device; mirrored on screen.
  $('rowLinesChk').checked = store.get('rowLines', true);
  const syncRowLines = () => $('outTable').classList.toggle('noLines', !$('rowLinesChk').checked);
  $('rowLinesChk').onchange = () => { store.set('rowLines', $('rowLinesChk').checked); syncRowLines(); };
  syncRowLines();
  ['orientSelect', 'paperSelect'].forEach(id => {
    const el = $(id); el.value = store.get(id, el.value);
    el.onchange = () => store.set(id, el.value);
  });
  if(!window.XLSX || !window.jspdf) log('A library failed to load — make sure the vendor/ folder was deployed with index.html.', 'err');
})();
