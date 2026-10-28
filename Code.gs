/**
 * PENGURUSAN PEMBIMBING RAKAN SEBAYA (PRS) SK JAWI – VERSI DALAM TALIAN
 * "Bantu Diri, Untuk Bantu Orang Lain" · Memperkasakan Kepimpinan PRS
 *
 * Pembangun Inovasi : SAMSOR BIN MOHAMMED
 *                     Guru Bimbingan dan Kaunseling, SK Jawi
 * Bantuan lesen     : samsormohamed@gmail.com · 013-580 1515 (Telegram)
 *
 * Pangkalan data : Google Sheet ini (tab PRS, Takwim, Perjumpaan, KPI, Tetapan)
 * Gambar aktiviti: folder Google Drive "Gambar Aktiviti PRS SK Jawi"
 * PIN asal       : 1234 (tukar di tab Tetapan dalam apps)
 */

const NAMA_APP = 'Pengurusan PRS SK Jawi';
const PIN_ASAL = '1234';
const NAMA_FOLDER = 'Gambar Aktiviti PRS SK Jawi';

const JADUAL = {
  members:  { nama: 'PRS',        hdr: ['ID', 'Nama PRS', 'Kelas', 'Jawatan'] },
  takwim:   { nama: 'Takwim',     hdr: ['ID', 'Tarikh', 'Aktiviti', 'Tempat', 'Sasaran', 'Status'] },
  meetings: { nama: 'Perjumpaan', hdr: ['ID', 'Tarikh', 'Masa', 'Tajuk', 'Tempat', 'Ringkasan / Catatan', 'Hadir (ID PRS)', 'Bil. Hadir', 'Gambar (ID Drive)'] },
  kpi:      { nama: 'KPI',        hdr: ['ID', 'KPI', 'Sasaran', 'Unit', 'Sumber', 'Pencapaian Semasa'] },
  meta:     { nama: 'Tetapan',    hdr: ['Kunci', 'Nilai'] }
};

/* ================= WEB APP ================= */
function doGet() {
  ensure_();
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle(NAMA_APP)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Jalankan SEKALI dari editor untuk beri kebenaran dan sediakan tab Sheet + folder gambar. */
function setup() {
  ensure_();
  Logger.log('Siap. Sheet: ' + ss_().getUrl());
  Logger.log('Folder gambar: ' + folder_().getUrl());
  Logger.log('PIN semasa: ' + PropertiesService.getScriptProperties().getProperty('PIN'));
}

/* ================= FUNGSI DIPANGGIL OLEH APPS ================= */
function muatData(pin) {
  semak_(pin);
  const meta = {};
  baca_('meta').forEach(r => { if (r[0]) meta[String(r[0])] = String(r[1]); });
  return {
    members: baca_('members').filter(r => r[0] && r[1]).map(r => ({ id: s_(r[0]), nama: s_(r[1]), kelas: s_(r[2]), jawatan: s_(r[3]) || 'Ahli' })),
    takwim: baca_('takwim').filter(r => r[0]).map(r => ({ id: s_(r[0]), tarikh: d_(r[1]), aktiviti: s_(r[2]), tempat: s_(r[3]), sasaran: s_(r[4]), status: s_(r[5]) || 'Dirancang' })),
    meetings: baca_('meetings').filter(r => r[0]).map(r => ({
      id: s_(r[0]), tarikh: d_(r[1]), masa: s_(r[2]), tajuk: s_(r[3]), tempat: s_(r[4]), catatan: s_(r[5]),
      hadir: senarai_(r[6]), photos: senarai_(r[8])
    })),
    kpi: baca_('kpi').filter(r => r[0]).map(r => ({ id: s_(r[0]), nama: s_(r[1]), sasaran: Number(r[2]) || 0, unit: s_(r[3]), sumber: s_(r[4]) || 'manual', semasa: Number(r[5]) || 0 })),
    meta: meta,
    baharu: Object.keys(meta).length === 0,
    sheetUrl: ss_().getUrl(),
    folderUrl: folder_().getUrl()
  };
}

/** Simpan jadual yang berubah sahaja. ch = { members:[...], takwim:[...], meetings:[...], kpi:[...], meta:{...} } */
function simpan(pin, ch) {
  semak_(pin);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (ch.members)  tulis_('members',  ch.members.map(m => [m.id, m.nama, m.kelas, m.jawatan]));
    if (ch.takwim)   tulis_('takwim',   ch.takwim.map(t => [t.id, t.tarikh, t.aktiviti, t.tempat, t.sasaran, t.status]));
    if (ch.meetings) tulis_('meetings', ch.meetings.map(m => [m.id, m.tarikh, m.masa, m.tajuk, m.tempat, m.catatan, (m.hadir || []).join(','), String((m.hadir || []).length), (m.photos || []).join(',')]));
    if (ch.kpi)      tulis_('kpi',      ch.kpi.map(k => [k.id, k.nama, String(k.sasaran), k.unit, k.sumber, String(k.semasa)]));
    if (ch.meta)     tulis_('meta',     Object.keys(ch.meta).map(k => [k, String(ch.meta[k] == null ? '' : ch.meta[k])]));
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  return new Date().toISOString();
}

function muatNaikGambar(pin, dataUrl, nama) {
  semak_(pin);
  const m = String(dataUrl).match(/^data:(image\/[\w+.-]+);base64,(.+)$/);
  if (!m) throw new Error('Format gambar tidak dikenali.');
  const blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], (nama || 'Gambar PRS').replace(/[\\/:*?"<>|]/g, '-') + '.jpg');
  const f = folder_().createFile(blob);
  try { f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) { /* domain menyekat perkongsian awam */ }
  return f.getId();
}

function padamGambar(pin, id) {
  semak_(pin);
  try { DriveApp.getFileById(id).setTrashed(true); } catch (e) {}
  return true;
}

function tukarPin(pin, pinBaru) {
  semak_(pin);
  pinBaru = String(pinBaru || '').trim();
  if (!/^\d{4,8}$/.test(pinBaru)) throw new Error('PIN baharu mesti 4 hingga 8 digit nombor.');
  PropertiesService.getScriptProperties().setProperty('PIN', pinBaru);
  return true;
}

/* ================= DALAMAN ================= */
function semak_(pin) {
  const p = PropertiesService.getScriptProperties().getProperty('PIN') || PIN_ASAL;
  if (String(pin || '').trim() !== p) throw new Error('PIN tidak sah.');
}

function ss_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('SHEET_ID');
  if (id) { try { return SpreadsheetApp.openById(id); } catch (e) {} }
  let s = null;
  try { s = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) {}
  if (!s) s = SpreadsheetApp.create('Data Pengurusan PRS SK Jawi');
  props.setProperty('SHEET_ID', s.getId());
  return s;
}

function folder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('FOLDER_ID');
  if (id) { try { const f = DriveApp.getFolderById(id); if (!f.isTrashed()) return f; } catch (e) {} }
  const f = DriveApp.createFolder(NAMA_FOLDER);
  props.setProperty('FOLDER_ID', f.getId());
  return f;
}

function sheet_(key) {
  const j = JADUAL[key], s = ss_();
  let sh = s.getSheetByName(j.nama);
  if (!sh) {
    sh = s.insertSheet(j.nama);
    sh.getRange(1, 1, 1, j.hdr.length).setValues([j.hdr]).setFontWeight('bold').setBackground('#4B2380').setFontColor('#FFFFFF');
    sh.setFrozenRows(1);
    sh.getRange(1, 1, sh.getMaxRows(), j.hdr.length).setNumberFormat('@');
    sh.setColumnWidth(1, 90);
  }
  return sh;
}

function ensure_() {
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('PIN')) props.setProperty('PIN', PIN_ASAL);
  Object.keys(JADUAL).forEach(sheet_);
  const s = ss_(), def = s.getSheetByName('Sheet1') || s.getSheetByName('Helaian1');
  if (def && def.getLastRow() === 0 && s.getSheets().length > 1) s.deleteSheet(def);
  folder_();
}

function baca_(key) {
  const sh = sheet_(key), n = sh.getLastRow() - 1, w = JADUAL[key].hdr.length;
  return n > 0 ? sh.getRange(2, 1, n, w).getValues() : [];
}

function tulis_(key, rows) {
  const sh = sheet_(key), w = JADUAL[key].hdr.length, last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, w).clearContent();
  if (rows.length) {
    if (sh.getMaxRows() < rows.length + 1) sh.insertRowsAfter(sh.getMaxRows(), rows.length + 1 - sh.getMaxRows());
    const r = sh.getRange(2, 1, rows.length, w);
    r.setNumberFormat('@');
    r.setValues(rows.map(x => x.map(v => v == null ? '' : String(v).slice(0, 49000))));
  }
}

function s_(v) { return v == null ? '' : String(v); }
function d_(v) { return v instanceof Date ? Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(v || ''); }
function senarai_(v) { return String(v || '').split(',').map(x => x.trim()).filter(String); }
