/**
 * Roster Bank: Google Apps Script receiver for Roster Convert.
 *
 * Saves PDFs that users chose to share into a private "Roster Bank" folder in
 * the Google Drive of whoever deploys this script. One sub-folder is created
 * per person ("Name, Staff number"). Only the Drive owner can see the files.
 *
 * Deploy: see bank/README.md.
 */

const FOLDER_NAME = 'Roster Bank';
const MAX_BYTES = 5 * 1024 * 1024;     // reject anything over 5 MB
const MAX_PER_HOUR = 60;               // crude spam guard across all senders

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) throw new Error('Empty request');
    const d = JSON.parse(e.postData.contents);
    if (d.kind !== 'roster-pdf' || d.consent !== true) throw new Error('Bad request');

    rateLimit_();

    const bytes = Utilities.base64Decode(String(d.pdf || ''));
    if (!bytes.length || bytes.length > MAX_BYTES) throw new Error('File missing or too large');
    if (bytes[0] !== 37 || bytes[1] !== 80 || bytes[2] !== 68 || bytes[3] !== 70) throw new Error('Not a PDF'); // "%PDF"

    const fileName = clean_(d.fileName, 150) || 'Roster';
    const person = [clean_(d.name, 80), clean_(d.staffId, 20)].filter(String).join(', ') || 'Unknown';

    const root = folder_(DriveApp.getRootFolder(), FOLDER_NAME);
    const dir = folder_(root, person);

    // Re-sharing the same month replaces the older copy instead of piling up duplicates.
    const old = dir.getFilesByName(fileName + '.pdf');
    while (old.hasNext()) old.next().setTrashed(true);

    const file = dir.createFile(Utilities.newBlob(bytes, 'application/pdf', fileName + '.pdf'));
    file.setDescription([
      'Shared via Roster Convert',
      'Name: ' + clean_(d.name, 80),
      'Staff no: ' + clean_(d.staffId, 20),
      'Month: ' + clean_(d.month, 30),
      'Times: ' + clean_(d.timeMode, 10),
      'Source file: ' + clean_(d.sourceFile, 150),
      'Shared at: ' + new Date().toISOString(),
    ].join('\n'));

    return json_({ ok: true, file: file.getName() });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

/** Visiting the URL in a browser just confirms the bank is running (and shows nothing private). */
function doGet() {
  return json_({ ok: true, service: 'Roster Bank' });
}

function folder_(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

function clean_(s, max) {
  return String(s == null ? '' : s).replace(/[\\/:*?"<>|\r\n\t]+/g, ' ').replace(/\s+/g, ' ').trim().slice(0, max);
}

function rateLimit_() {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const cache = CacheService.getScriptCache();
    const key = 'uploads-' + Math.floor(Date.now() / 3600000);
    const n = Number(cache.get(key) || 0);
    if (n >= MAX_PER_HOUR) throw new Error('Too many uploads right now, try again later');
    cache.put(key, String(n + 1), 3700);
  } finally {
    lock.releaseLock();
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
