/**
 * =====================================================================
 *  e-PANITIA — Sistem Pengurusan Panitia Sekolah Rendah
 *  Backend Google Apps Script (Google Sheets + Google Drive)
 *  Versi 1.5.2
 * =====================================================================
 *
 *  CARA PASANG
 *  1. Buka Google Sheet baharu > Sambungan > Apps Script.
 *  2. Tampal keseluruhan fail ini ke dalam Code.gs.
 *  3. (Pilihan A — hos dalam Apps Script) Tambah fail HTML bernama "index"
 *     dan tampal kandungan index.html. Tiada URL perlu diisi.
 *     (Pilihan B — hos luar seperti Netlify/GitHub Pages) Isi API_URL
 *     dalam index.html dengan URL Web App.
 *  4. Pilih fungsi setupSistem > Run. Benarkan akses (Sheets, Drive, trigger).
 *     Fungsi ini MENJANA SEMUA JADUAL secara automatik, mencipta folder
 *     Drive induk, memasukkan 12 panitia asas dan memasang trigger harian.
 *  5. Deploy > New deployment > Web app
 *       Execute as: Me
 *       Who has access: Anyone  (atau "Anyone within <domain>" jika hos
 *       dalam Apps Script dan semua guru guna akaun domain yang sama)
 *  6. Pengguna PERTAMA yang mendaftar akan dijadikan Super Admin.
 *
 *  Nota: setiap kali kod dikemas kini, buat "Manage deployments > Edit >
 *  Version: New version" supaya URL Web App kekal sama.
 * =====================================================================
 */

const APP_VERSION = '1.5.2';
const KATA_LALUAN_LALAI = '12345678';
const SESI_HARI = 7;
const TONG_HARI = 30;
const MAKS_FAIL_MB = 10;

/* ---------- Skema jadual (dijana automatik) ---------- */
const SCHEMA = {
  Tetapan:  ['kunci', 'nilai', 'catatan'],
  Pengguna: ['id', 'nama', 'noKP', 'emel', 'telefon', 'jawatan', 'opsyen', 'kelulusan', 'gred', 'pengalaman',
             'peranan', 'fotoId', 'jadualId', 'jadualMime', 'hash', 'salt', 'status', 'dicipta', 'dikemaskini', 'wajibTukar', 'namaJadual'],
  Panitia:  ['id', 'nama', 'kod', 'ketuaId', 'setiausahaId', 'folderId', 'subFolder', 'status', 'dicipta'],
  Keahlian: ['id', 'panitiaId', 'userId', 'tahun', 'status', 'dicipta'],
  Carta:    ['id', 'panitiaId', 'userId', 'tahap', 'susunan', 'jawatan'],
  Fail:     ['id', 'panitiaId', 'folderNo', 'tajuk', 'penerangan', 'jenis', 'url', 'driveId', 'mime', 'saiz',
             'tugasanId', 'pemilikId', 'pemilikNama', 'dicipta', 'status', 'dipadamPada', 'dipadamOleh'],
  Tugasan:  ['id', 'panitiaId', 'tajuk', 'folderNo', 'tarikhAkhir', 'diciptaOleh', 'dicipta'],
  Pemantau: ['id', 'panitiaId', 'kod', 'hash', 'salt', 'label', 'tamat', 'diciptaOleh', 'dicipta', 'status'],
  Sesi:     ['token', 'userId', 'tamat', 'dicipta'],
  Log:      ['id', 'masa', 'userId', 'nama', 'panitiaId', 'folderNo', 'tindakan', 'butiran']
};

const FOLDER_NAMA = [
  '01 Profil & Organisasi', '02 Perancangan & Takwim', '03 Mesyuarat Panitia',
  '04 Dokumen Kurikulum & PdP', '05 PBD & Pentaksiran', '06 Analisis & Pencapaian',
  '07 Intervensi & Program', '08 Pengurusan Guru', '09 Kewangan, Aset & Inventori',
  '10 Dokumentasi & Evidens'
];

const PANITIA_ASAS = [
  ['Bahasa Melayu', 'BM'], ['Bahasa Inggeris', 'BI'], ['Matematik', 'MT'], ['Sains', 'SN'],
  ['Pendidikan Islam', 'PI'], ['Pendidikan Moral', 'PM'], ['Sejarah', 'SJ'],
  ['Reka Bentuk dan Teknologi', 'RBT'], ['Pendidikan Jasmani & Kesihatan', 'PJK'],
  ['Pendidikan Seni Visual', 'PSV'], ['Pendidikan Muzik', 'MZ'], ['Bahasa Arab', 'BA']
];

const TETAPAN_ASAS = [
  ['namaSekolah', 'Nama Sekolah Anda', 'Nama sekolah dipapar dalam aplikasi dan laporan'],
  ['kodSekolah', '', 'Kod sekolah (pilihan)'],
  ['gbId', '', 'ID pengguna Guru Besar (Tahap 1 carta)'],
  ['gbNama', '', 'Nama Guru Besar jika belum berdaftar'],
  ['pkId', '', 'ID Penolong Kanan Pentadbiran (Tahap 2)'],
  ['pkNama', '', 'Nama PK Pentadbiran jika belum berdaftar'],
  ['sukId', '', 'ID Setiausaha Kurikulum (Tahap 2)'],
  ['sukNama', '', 'Nama SU Kurikulum jika belum berdaftar'],
  ['gbFoto', '', 'ID foto Guru Besar (jika tidak dipautkan kepada pengguna)'],
  ['pkFoto', '', 'ID foto PK Pentadbiran (jika tidak dipautkan kepada pengguna)'],
  ['sukFoto', '', 'ID foto SU Kurikulum (jika tidak dipautkan kepada pengguna)'],
  ['jadualCsvUrl', 'https://docs.google.com/spreadsheets/d/e/2PACX-1vTi4sLUWGblVsM6gyTxstwXtt8n57CgKRF4w3YwnsQ49HvOWyZEdlH3ljmjvIlysfPbLMAasSvV5NE8/pub?gid=635713887&single=true&output=csv', 'Pautan CSV jadual waktu (Google Sheet > File > Share > Publish to web > CSV)']
];

/* =====================================================================
 *  PEMASANGAN & MENU
 * ===================================================================== */
function setupSistem() {
  PropertiesService.getScriptProperties().deleteProperty('SCHEMA_VER');
  try { CacheService.getScriptCache().remove('SV'); } catch (e) { /* */ }
  _P = null;
  ensureSchema_();
  const induk = folderInduk_();
  if (!rows_('Panitia').length) {
    PANITIA_ASAS.forEach(([n, k]) => insert_('Panitia', { id: uid_(), nama: n, kod: k, status: 'AKTIF', dicipta: now_() }));
  }
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'bersihkanTong') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('bersihkanTong').timeBased().everyDays(1).atHour(2).create();
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === 'cacheBerubah') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('cacheBerubah').forSpreadsheet(getSS_()).onChange().create();
  ScriptApp.newTrigger('cacheBerubah').forSpreadsheet(getSS_()).onEdit().create();
  selesai_();
  Logger.log('✅ Sistem siap.');
  Logger.log('Spreadsheet : ' + getSS_().getUrl());
  Logger.log('Folder Drive: ' + induk.getUrl());
}

/**
 * DIAGNOSTIK: jalankan dari editor Apps Script jika muat naik foto/fail gagal.
 * Fungsi ini meminta kebenaran Google Drive (jika belum) dan menguji
 * cipta folder + fail. Lihat keputusan di Execution log.
 */
function ujiDrive() {
  const induk = folderInduk_();
  Logger.log('✅ Folder induk: ' + induk.getName() + ' | ' + induk.getUrl());
  const foto = folderSistem_('_FOTO_PROFIL');
  Logger.log('✅ Folder foto: ' + foto.getUrl());
  const f = foto.createFile(Utilities.newBlob('ujian', 'text/plain', 'ujian_drive.txt'));
  Logger.log('✅ Fail ujian berjaya dicipta: ' + f.getId());
  f.setTrashed(true);
  selesai_();
  Logger.log('✅ Google Drive berfungsi. Sekarang buat Manage deployments > Edit > New version.');
}

/**
 * DIAGNOSTIK JADUAL: jalankan dari editor Apps Script.
 * Meminta kebenaran "Connect to an external service" (jika belum) dan menguji bacaan CSV jadual.
 */
function ujiJadual() {
  ensureSchema_();
  const url = String(tetapan_().jadualCsvUrl || '');
  Logger.log('Pautan CSV: ' + (url || '(kosong)'));
  const idx = jadualIndeks_(true);
  selesai_();
  if (!idx) { Logger.log('❗ Pautan CSV jadual belum ditetapkan dalam jadual Tetapan.'); return; }
  const n = Object.keys(idx.slot).length;
  const baris = Object.keys(idx.slot).reduce((t, k) => t + idx.slot[k].length, 0);
  Logger.log('✅ Jadual dibaca: ' + n + ' guru, ' + baris + ' slot.');
  const guru = rows_('Pengguna').filter(u => u.status !== 'NYAHAKTIF');
  const tiada = guru.filter(u => !padanJadual_(u, idx)).map(u => u.nama);
  Logger.log('✅ ' + (guru.length - tiada.length) + '/' + guru.length + ' pengguna sistem dipadankan dengan jadual.');
  if (tiada.length) Logger.log('⚠️ Belum dipadankan (guru boleh pilih nama sendiri di Profil):\n' + tiada.join('\n'));
  Logger.log('Seterusnya: Manage deployments > Edit > Version: New version > Deploy.');
}

function onOpen() {
  try {
    SpreadsheetApp.getUi().createMenu('⚙️ e-Panitia')
      .addItem('Jana / baiki jadual', 'setupSistem')
      .addItem('Import guru dari tab Data Guru', 'importDataGuru')
      .addItem('Kosongkan cache', 'kosongkanCache')
      .addItem('Uji bacaan jadual waktu', 'ujiJadual')
      .addItem('Bersihkan tong sampah sekarang', 'bersihkanTong')
      .addToUi();
  } catch (e) { /* bukan konteks UI */ }
}

/**
 * IMPORT GURU PUKAL daripada tab "Data Guru" (dalam spreadsheet yang sama).
 * Lajur dikesan melalui tajuk baris pertama: Emel/Email, Nama, No KP/NoKP/IC.
 * Semua guru didaftarkan dengan kata laluan lalai 12345678 dan WAJIB
 * menukarnya semasa log masuk kali pertama.
 * Guru yang emel atau No. KP-nya sudah wujud akan dilangkau (tidak ditindih).
 * Selamat dijalankan berulang kali — hanya guru baharu ditambah.
 */
function importDataGuru() {
  ensureSchema_();
  const ss = getSS_();
  const sh = ss.getSheets().find(x => /^\s*data\s*guru\s*$/i.test(x.getName()));
  if (!sh) throw new Error('Tab "Data Guru" tidak dijumpai dalam spreadsheet ' + ss.getName());
  const v = sh.getDataRange().getValues();
  if (v.length < 2) throw new Error('Tab Data Guru kosong.');
  const head = v[0].map(h => String(h).toLowerCase().replace(/[\s._\-\/]/g, ''));
  const cari = pola => head.findIndex(h => pola.test(h));
  const cEmel = cari(/^(e?mel|email|emelrasmi|emeldelima|emelmoe)/);
  const cNama = cari(/^nama/);
  const cKP = cari(/^(no)?(kp|ic|mykad|kadpengenalan|nokadpengenalan)/);
  if (cEmel < 0 || cNama < 0 || cKP < 0) {
    throw new Error('Tajuk lajur tidak dikenali. Pastikan baris pertama mengandungi: Emel, Nama, No KP. Dijumpai: ' + v[0].join(' | '));
  }

  const ada = rows_('Pengguna');
  const emelAda = {}, kpAda = {};
  ada.forEach(u => { emelAda[u.emel] = 1; kpAda[u.noKP] = 1; });
  const pemilik = String(Session.getEffectiveUser().getEmail() || '').toLowerCase();
  let adaAdmin = ada.some(u => u.peranan === 'SUPER_ADMIN' && u.status !== 'NYAHAKTIF');

  const baru = [], langkau = [];
  for (let i = 1; i < v.length; i++) {
    const r = v[i];
    const emel = String(r[cEmel] || '').trim().toLowerCase();
    const nama = String(r[cNama] || '').trim().replace(/\s+/g, ' ');
    let kp = r[cKP];
    kp = (typeof kp === 'number' ? kp.toFixed(0) : String(kp || '')).replace(/\D/g, '');
    if (kp.length === 11) kp = '0' + kp; // sifar di hadapan hilang dalam Sheets
    const baris = 'Baris ' + (i + 1) + ' (' + (nama || emel || 'kosong') + ')';
    if (!emel && !nama && !kp) continue;
    if (!emelSah_(emel)) { langkau.push(baris + ': emel tidak sah'); continue; }
    if (nama.length < 3) { langkau.push(baris + ': nama kosong'); continue; }
    if (!/^\d{12}$/.test(kp)) { langkau.push(baris + ': No. KP bukan 12 digit'); continue; }
    if (emelAda[emel]) { langkau.push(baris + ': emel sudah berdaftar'); continue; }
    if (kpAda[kp]) { langkau.push(baris + ': No. KP sudah berdaftar'); continue; }
    const salt = uid_(16);
    let peranan = 'GURU';
    if (!adaAdmin && emel === pemilik) { peranan = 'SUPER_ADMIN'; adaAdmin = true; }
    baru.push({
      id: uid_(), nama, noKP: kp, emel, telefon: '', peranan, hash: hash_(KATA_LALUAN_LALAI, salt), salt,
      status: 'AKTIF', dicipta: now_(), dikemaskini: now_(), wajibTukar: 'YA'
    });
    emelAda[emel] = 1; kpAda[kp] = 1;
  }
  insertMany_('Pengguna', baru);
  if (baru.length) log_('Sistem', '', '', 'IMPORT GURU', baru.length + ' guru didaftarkan dari tab ' + sh.getName());

  Logger.log('✅ ' + baru.length + ' guru berjaya didaftarkan (kata laluan lalai: ' + KATA_LALUAN_LALAI + ').');
  if (langkau.length) Logger.log('⚠️ ' + langkau.length + ' baris dilangkau:\n' + langkau.join('\n'));
  if (!adaAdmin) Logger.log('❗ Tiada Super Admin. Tukar lajur "peranan" akaun anda kepada SUPER_ADMIN dalam jadual Pengguna.');
  else if (baru.some(u => u.peranan === 'SUPER_ADMIN')) Logger.log('👑 ' + pemilik + ' dijadikan Super Admin.');
  selesai_();
  try { SpreadsheetApp.getUi().alert('Import selesai: ' + baru.length + ' guru didaftarkan, ' + langkau.length + ' dilangkau.\nLihat butiran di Execution log.'); } catch (e) { /* dijalankan dari editor */ }
  return { didaftar: baru.length, dilangkau: langkau };
}

/* =====================================================================
 *  PINTU MASUK WEB APP
 * ===================================================================== */
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.api === 'ping') return json_({ ok: true, app: 'e-Panitia', versi: APP_VERSION });
  try {
    return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('e-Panitia')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (err) {
    return json_({ ok: true, app: 'e-Panitia', versi: APP_VERSION, nota: 'Fail HTML "index" tiada. Guna mod hos luar.' });
  }
}

function doPost(e) {
  let q = {};
  try { q = JSON.parse(e.postData.contents); } catch (err) { /* abaikan */ }
  return json_(handle_(q));
}

/** Dipanggil melalui google.script.run apabila HTML dihos dalam Apps Script */
function api(q) { return handle_(q); }

/** Ralat kunci (LockService) ditukar kepada mesej yang klien kenali sebagai boleh dicuba semula */
function mesejRalat_(err) {
  const m = String((err && err.message) || err);
  if (/lock|timed out|too many|service invoked too many|try again later/i.test(m)) return 'PELAYAN_SIBUK: Pelayan sedang sibuk. Perubahan akan dicuba semula.';
  return m;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

const AWAM = ['awam', 'daftar', 'login', 'pemantauData', 'imej', 'muatTurun'];

function handle_(q) {
  try {
    q = q || {};
    ensureSchema_();
    const fn = Object.prototype.hasOwnProperty.call(ACTIONS, q.action) ? ACTIONS[q.action] : null;
    if (!fn) throw new Error('Tindakan tidak dikenali: ' + q.action);
    const u = AWAM.indexOf(q.action) >= 0 ? null : auth_(q.token);
    const data = fn(q, u);
    selesai_();
    if (data && data.perluSync) {
      delete data.perluSync;
      data.sync = sync({ aktif: q.aktif }, userById_(u.id)); // keadaan terkini dalam jawapan yang sama
    }
    return { ok: true, data };
  } catch (err) {
    try { selesai_(); } catch (e) { /* */ }
    return { ok: false, error: mesejRalat_(err) };
  }
}

const ACTIONS = {
  awam, daftar, login, logout, init, kemaskiniProfil, tukarKataLaluan, muatNaikFoto, muatNaikJadual, lihatJadual,
  sync, batch, jadualGuru, setNamaJadual, muatNaikFotoPentadbir,
  imej, muatTurun, panitiaDetail, simpanCarta, lantikSU, sahkanAhli, calonAhli, tambahAhli, keluarkanAhli, muatNaikFail, tambahPautan, padamFail,
  pulihFail, padamKekal, tambahTugasan, padamTugasan, janaLaporan, eksportCartaPdf, janaPemantau, batalPemantau,
  pemantauData, senaraiPengguna, simpanPanitia, padamPanitia, lantikKP, kemaskiniPengguna, resetKataLaluan,
  simpanTetapan
};

/* =====================================================================
 *  LAPISAN DATA (Google Sheets) + MEMORY CACHE (CacheService)
 *
 *  Bacaan: memori permintaan (_C) → CacheService → Google Sheet.
 *  Setiap jadual ada "generasi" (G_<nama>) dalam CacheService.
 *  Setiap tulisan menukar generasi, jadi cache lama tidak akan dibaca lagi.
 *  Kunci (LockService) hanya diambil semasa tulisan pertama, jadi kerja
 *  Drive yang lambat (muat naik) tidak menyekat pengguna lain.
 * ===================================================================== */
const CACHE_TTL = 21600;      // 6 jam (maksimum CacheService)
const CACHE_CEBIS = 45000;    // saiz setiap cebisan (had CacheService 100KB)
const TANPA_CACHE = { Log: 1 };
let _SS = null, _P = null, _LOCK = null;
const _C = {}, _SH = {}, _KOTOR = {};

function props_() { if (!_P) _P = PropertiesService.getScriptProperties().getProperties(); return _P; }
function setProp_(k, v) { PropertiesService.getScriptProperties().setProperty(k, v); if (_P) _P[k] = v; }
let _G = null;
/** Versi (generasi) setiap jadual disimpan dalam CacheService — satu panggilan getAll, lebih pantas daripada Script Properties */
function gens_() {
  if (_G) return _G;
  const nama = Object.keys(SCHEMA);
  const m = cache_().getAll(nama.map(n => 'G_' + n));
  _G = {};
  const baru = {};
  nama.forEach(n => {
    if (m['G_' + n]) _G[n] = m['G_' + n];
    else { _G[n] = uid_(8); baru['G_' + n] = _G[n]; } // versi hilang → versi rawak baharu (cache lama tidak akan dibaca)
  });
  if (Object.keys(baru).length) cache_().putAll(baru, CACHE_TTL);
  return _G;
}
function gen_(name) { return gens_()[name] || '0'; }
function cache_() { return CacheService.getScriptCache(); }

function getSS_() {
  if (_SS) return _SS;
  const id = props_().SS_ID;
  if (id) { try { _SS = SpreadsheetApp.openById(id); return _SS; } catch (e) { /* hilang — cipta semula */ } }
  try { _SS = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) { _SS = null; }
  if (!_SS) _SS = SpreadsheetApp.create('Pangkalan Data e-Panitia');
  setProp_('SS_ID', _SS.getId());
  return _SS;
}
function sh_(name) {
  if (_SH[name]) return _SH[name];
  const sh = getSS_().getSheetByName(name);
  if (!sh) throw new Error('Jadual ' + name + ' tiada. Jalankan setupSistem().');
  _SH[name] = sh;
  return sh;
}

/* ---------- CacheService bercebis ---------- */
function cacheBaca_(kunci) {
  try {
    const c = cache_();
    const n = Number(c.get(kunci));
    if (!n) return null;
    const keys = [];
    for (let i = 0; i < n; i++) keys.push(kunci + '_' + i);
    const m = c.getAll(keys);
    let s = '';
    for (let i = 0; i < keys.length; i++) { if (m[keys[i]] == null) return null; s += m[keys[i]]; }
    return JSON.parse(s);
  } catch (e) { return null; }
}
function cacheTulis_(kunci, data, ttl) {
  try {
    const s = JSON.stringify(data);
    const o = {};
    let n = 0;
    for (let i = 0; i < s.length; i += CACHE_CEBIS) o[kunci + '_' + (n++)] = s.slice(i, i + CACHE_CEBIS);
    if (n > 150) return; // terlalu besar — baca terus dari Sheet
    o[kunci] = String(n);
    cache_().putAll(o, ttl || CACHE_TTL);
  } catch (e) { /* cache pilihan sahaja */ }
}

/* ---------- Kunci & generasi ---------- */
function kunci_() {
  if (_LOCK) return;
  _LOCK = LockService.getScriptLock();
  _LOCK.waitLock(25000);
  // Selepas memegang kunci, buang data yang dibaca sebelum kunci supaya tulisan berdasarkan data terkini
  _P = null;
  _G = null;
  Object.keys(_C).forEach(k => delete _C[k]);
}
function kotor_(name) {
  if (_KOTOR[name]) return;
  _KOTOR[name] = 1;
  gens_()[name] = uid_(8);
  cache_().put('G_' + name, _G[name], CACHE_TTL); // batalkan cache serta-merta
}
/** Dipanggil di hujung setiap permintaan: jana generasi baharu, isi semula cache, lepaskan kunci */
function selesai_() {
  flushLog_();
  const nama = Object.keys(_KOTOR);
  if (nama.length) {
    const baru = {};
    nama.forEach(n => { baru['G_' + n] = uid_(8); });
    cache_().putAll(baru, CACHE_TTL);
    nama.forEach(n => {
      gens_()[n] = baru['G_' + n];
      if (_C[n] && !TANPA_CACHE[n]) cacheTulis_('T_' + n + '_' + baru['G_' + n], { head: _C[n].head, rows: _C[n].rows });
      delete _KOTOR[n];
    });
  }
  if (_LOCK) { try { _LOCK.releaseLock(); } catch (e) { /* */ } _LOCK = null; }
}
/** Pencetus onChange/onEdit: sebarang suntingan manual dalam Sheet membatalkan semua cache */
function cacheBerubah() {
  const o = {};
  Object.keys(SCHEMA).forEach(n => { o['G_' + n] = uid_(8); });
  cache_().putAll(o, CACHE_TTL);
  cache_().removeAll(Object.keys(SCHEMA).map(n => 'H_' + n + '_' + APP_VERSION));
  _G = null;
}
function kosongkanCache() { cacheBerubah(); Logger.log('✅ Cache dikosongkan.'); }

/** Menjana semua jadual & lajur yang tiada. Selamat dijalankan berulang kali. */
function ensureSchema_() {
  if (cache_().get('SV') === APP_VERSION) return; // laluan pantas: tiada bacaan Script Properties
  if (props_().SCHEMA_VER === APP_VERSION) { cache_().put('SV', APP_VERSION, CACHE_TTL); return; }
  kunci_();
  if (props_().SCHEMA_VER === APP_VERSION) return;
  const ss = getSS_();
  Object.keys(SCHEMA).forEach(name => {
    const cols = SCHEMA[name];
    let sh = ss.getSheetByName(name);
    if (!sh) {
      sh = ss.insertSheet(name);
      sh.getRange(1, 1, 1, cols.length).setValues([cols]);
    } else {
      const lastCol = Math.max(sh.getLastColumn(), 1);
      const head = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
      if (head.join('') === '') {
        sh.getRange(1, 1, 1, cols.length).setValues([cols]);
      } else {
        const tiada = cols.filter(c => head.indexOf(c) < 0);
        if (tiada.length) sh.getRange(1, head.length + 1, 1, tiada.length).setValues([tiada]);
      }
    }
    const n = sh.getLastColumn();
    sh.getRange(1, 1, sh.getMaxRows(), n).setNumberFormat('@');
    sh.getRange(1, 1, 1, n).setFontWeight('bold').setBackground('#145C43').setFontColor('#FFFFFF');
    sh.setFrozenRows(1);
  });
  ['Sheet1', 'Helaian1'].forEach(nm => {
    const s = ss.getSheetByName(nm);
    if (s && s.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(s);
  });
  cacheBerubah();
  _P = null;
  Object.keys(_C).forEach(k => delete _C[k]);
  const tet = rows_('Tetapan');
  TETAPAN_ASAS.forEach(([k, v, c]) => { if (!tet.find(t => t.kunci === k)) insert_('Tetapan', { kunci: k, nilai: v, catatan: c }); });
  setProp_('SCHEMA_VER', APP_VERSION);
  cache_().put('SV', APP_VERSION, CACHE_TTL);
}

function bacaSheet_(name) {
  const values = sh_(name).getDataRange().getValues();
  const head = values.shift().map(String);
  const rows = values.map((r, i) => {
    const o = { _row: i + 2 };
    head.forEach((h, j) => {
      if (!h) return;
      const v = r[j];
      o[h] = v === null || v === undefined ? '' : (v instanceof Date ? v.toISOString() : String(v));
    });
    return o;
  }).filter(o => Object.keys(o).some(k => k !== '_row' && o[k] !== ''));
  return { head, rows };
}

/** fresh = true → abaikan semua cache dan baca terus dari Sheet */
function table_(name, fresh) {
  if (!fresh && _C[name]) return _C[name];
  const kunci = 'T_' + name + '_' + gen_(name);
  if (!fresh && !TANPA_CACHE[name]) {
    const c = cacheBaca_(kunci);
    if (c) { _C[name] = { head: c.head, rows: c.rows, dariCache: true }; return _C[name]; }
  }
  const d = bacaSheet_(name);
  _C[name] = { head: d.head, rows: d.rows, dariCache: false };
  if (!TANPA_CACHE[name]) cacheTulis_(kunci, d);
  return _C[name];
}

function rows_(name, fresh) { return table_(name, fresh).rows; }

function head_(name) {
  if (_C[name]) return _C[name].head;
  const k = 'H_' + name + '_' + APP_VERSION;
  const c = cache_().get(k);
  if (c) return JSON.parse(c);
  const sh = sh_(name);
  const h = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0].map(String);
  cache_().put(k, JSON.stringify(h), CACHE_TTL);
  return h;
}

function insert_(name, obj) { return insertMany_(name, [obj])[0]; }

function insertMany_(name, arr) {
  if (!arr.length) return arr;
  kunci_();
  kotor_(name);
  const head = head_(name);
  const sh = sh_(name);
  const start = sh.getLastRow() + 1;
  const perlu = start + arr.length - 1;
  if (perlu > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), perlu - sh.getMaxRows() + 50);
  const vals = arr.map(o => head.map(h => (o[h] === undefined || o[h] === null) ? '' : String(o[h])));
  sh.getRange(start, 1, arr.length, head.length).setNumberFormat('@').setValues(vals);
  const t = _C[name];
  if (t) arr.forEach((o, i) => { const r = { _row: start + i }; head.forEach((h, j) => { if (h) r[h] = vals[i][j]; }); t.rows.push(r); });
  return arr;
}

function update_(name, id, patch, key) {
  key = key || 'id';
  kunci_();
  kotor_(name);
  let t = table_(name);
  let r = t.rows.find(x => x[key] === String(id));
  // Semak baris dari cache masih sama (lindung daripada suntingan manual dalam Sheet)
  if (r && t.dariCache) {
    const kol = t.head.indexOf(key) + 1;
    if (String(sh_(name).getRange(r._row, kol).getValue()) !== String(id)) {
      t = table_(name, true);
      r = t.rows.find(x => x[key] === String(id));
    }
  }
  if (!r) throw new Error('Rekod tidak dijumpai.');
  Object.keys(patch).forEach(k => { r[k] = patch[k] === null || patch[k] === undefined ? '' : String(patch[k]); });
  const vals = t.head.map(h => (r[h] === undefined ? '' : String(r[h])));
  sh_(name).getRange(r._row, 1, 1, t.head.length).setValues([vals]);
  return r;
}

function deleteWhere_(name, pred) {
  kunci_();
  const t = table_(name, true);
  const sasaran = t.rows.filter(pred).map(r => r._row).sort((a, b) => b - a);
  if (!sasaran.length) return 0;
  if (sasaran.length > 15) { // banyak baris: tulis semula jadual sekali (deleteRow satu demi satu sangat lambat)
    const buang = {};
    sasaran.forEach(r => { buang[r] = 1; });
    tulisSemula_(name, t.rows.filter(r => !buang[r._row]));
    return sasaran.length;
  }
  kotor_(name);
  const sh = sh_(name);
  if (sasaran.length >= sh.getMaxRows() - 1) sh.insertRowsAfter(sh.getMaxRows(), 20);
  sasaran.forEach(r => sh.deleteRow(r));
  delete _C[name];
  return sasaran.length;
}

function tulisSemula_(name, rows) {
  kunci_();
  kotor_(name);
  const t = table_(name, true);
  const sh = sh_(name), n = t.head.length;
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, n).clearContent();
  if (rows.length) {
    if (rows.length + 1 > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), rows.length + 1 - sh.getMaxRows());
    sh.getRange(2, 1, rows.length, n).setNumberFormat('@')
      .setValues(rows.map(r => t.head.map(h => (r[h] === undefined || r[h] === null) ? '' : String(r[h]))));
  }
  delete _C[name];
}

/** Baca N baris terakhir Log sahaja (bukan keseluruhan jadual), dengan cache */
function logTerkini_(n) {
  const kunci = 'LT_' + gen_('Log');
  const c = cacheBaca_(kunci);
  if (c) return c;
  const sh = sh_('Log');
  const last = sh.getLastRow();
  if (last < 2) return [];
  const head = head_('Log');
  const mula = Math.max(2, last - n + 1);
  const out = sh.getRange(mula, 1, last - mula + 1, head.length).getValues().map(r => {
    const o = {};
    head.forEach((h, j) => { if (h) o[h] = r[j] instanceof Date ? r[j].toISOString() : String(r[j]); });
    return o;
  });
  cacheTulis_(kunci, out);
  return out;
}

/* =====================================================================
 *  UTILITI
 * ===================================================================== */
function uid_(n) { return Utilities.getUuid().replace(/-/g, '').slice(0, n || 12); }
function now_() { return new Date().toISOString(); }
function hash_(pw, salt) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + '::' + pw, Utilities.Charset.UTF_8)
    .map(b => ('0' + (b & 255).toString(16)).slice(-2)).join('');
}
function mesti_(c, msg) { if (!c) throw new Error(msg || 'Anda tiada kebenaran untuk tindakan ini.'); }
function bersih_(s, max) { return String(s === null || s === undefined ? '' : s).trim().slice(0, max || 500); }
function esc_(s) {
  return String(s || '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}
function emelSah_(e) { return /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(e); }
function tetapan_() { const o = {}; rows_('Tetapan').forEach(r => { o[r.kunci] = r.nilai; }); return o; }
function setTetapan_(k, v) {
  if (rows_('Tetapan').find(x => x.kunci === k)) update_('Tetapan', k, { nilai: v }, 'kunci');
  else insert_('Tetapan', { kunci: k, nilai: v, catatan: '' });
}
const _LOGBUF = [];
/** Log dikumpul dan ditulis sekali di hujung permintaan (satu tulisan Sheet) */
function flushLog_() { if (_LOGBUF.length) insertMany_('Log', _LOGBUF.splice(0)); }
function log_(u, panitiaId, folderNo, tindakan, butiran) {
  const nama = typeof u === 'string' ? u : (u ? u.nama : 'Sistem');
  _LOGBUF.push({
    id: uid_(), masa: now_(), userId: (u && u.id) || '', nama, panitiaId: panitiaId || '',
    folderNo: folderNo || '', tindakan, butiran: bersih_(butiran, 300)
  });
}
function userById_(id) { return id ? rows_('Pengguna').find(u => u.id === id) : null; }
function pub_(u, penuh) {
  if (!u) return null;
  const o = {};
  ['id', 'nama', 'emel', 'telefon', 'jawatan', 'opsyen', 'kelulusan', 'gred', 'pengalaman', 'peranan',
    'fotoId', 'jadualId', 'jadualMime', 'status'].forEach(k => { o[k] = u[k] || ''; });
  if (penuh) { o.noKP = u.noKP; o.wajibTukar = u.wajibTukar || ''; }
  return o;
}
function ringkas_(u) { return u ? { id: u.id, nama: u.nama, jawatan: u.jawatan, opsyen: u.opsyen, fotoId: u.fotoId } : null; }
function senaraiAktif_() { return rows_('Panitia').filter(p => p.status !== 'PADAM').sort((a, b) => a.nama.localeCompare(b.nama)); }
function panitia_(id) {
  const p = rows_('Panitia').find(x => x.id === id && x.status !== 'PADAM');
  mesti_(p, 'Panitia tidak dijumpai.');
  return p;
}
function failById_(id) {
  const f = rows_('Fail').find(x => x.id === id && x.status !== 'KEKAL_PADAM');
  mesti_(f, 'Fail tidak dijumpai.');
  return f;
}
function failPub_(f) {
  return {
    id: f.id, panitiaId: f.panitiaId, folderNo: Number(f.folderNo), tajuk: f.tajuk, penerangan: f.penerangan,
    jenis: f.jenis, url: f.url, driveId: f.driveId, mime: f.mime, saiz: Number(f.saiz) || 0, tugasanId: f.tugasanId,
    pemilikId: f.pemilikId, pemilikNama: f.pemilikNama, dicipta: f.dicipta, status: f.status,
    dipadamPada: f.dipadamPada, dipadamOleh: f.dipadamOleh
  };
}

/* ---------- Kawalan akses ---------- */
function auth_(token) {
  mesti_(token, 'SESI_TAMAT');
  const s = rows_('Sesi').find(r => r.token === token);
  mesti_(s && new Date(s.tamat).getTime() > Date.now(), 'SESI_TAMAT');
  const u = userById_(s.userId);
  mesti_(u && u.status !== 'NYAHAKTIF', 'SESI_TAMAT');
  return u;
}
function admin_(u) { mesti_(u && u.peranan === 'SUPER_ADMIN', 'Hanya Super Admin boleh melakukan tindakan ini.'); }

/** r: KP | SU | AHLI | MENUNGGU | null */
function akses_(u, p) {
  const admin = u.peranan === 'SUPER_ADMIN';
  let r = null;
  if (p.ketuaId && p.ketuaId === u.id) r = 'KP';
  else if (p.setiausahaId && p.setiausahaId === u.id) r = 'SU';
  else {
    const k = rows_('Keahlian').find(x => x.panitiaId === p.id && x.userId === u.id);
    if (k) r = k.status === 'SAH' ? 'AHLI' : 'MENUNGGU';
  }
  const lihat = admin || r === 'KP' || r === 'SU' || r === 'AHLI';
  return { r, admin, lihat, urus: admin || r === 'KP', su: admin || r === 'KP' || r === 'SU', kew: admin || r === 'KP' || r === 'SU' };
}

function ahliPanitia_(p) {
  const ids = [];
  rows_('Keahlian').forEach(k => { if (k.panitiaId === p.id && k.status === 'SAH' && ids.indexOf(k.userId) < 0) ids.push(k.userId); });
  [p.ketuaId, p.setiausahaId].forEach(id => { if (id && ids.indexOf(id) < 0) ids.push(id); });
  return ids.map(userById_).filter(u => u && u.status !== 'NYAHAKTIF');
}

function pentadbir_() {
  const t = tetapan_();
  const pilih = (id, nm, foto) => {
    const u = ringkas_(userById_(id));
    if (u) { if (!u.fotoId && foto) u.fotoId = foto; return u; }
    return nm ? { id: '', nama: nm, jawatan: '', fotoId: foto || '' } : null;
  };
  return { gb: pilih(t.gbId, t.gbNama, t.gbFoto), pk: pilih(t.pkId, t.pkNama, t.pkFoto), suk: pilih(t.sukId, t.sukNama, t.sukFoto) };
}

/* ---------- Google Drive ---------- */
function folderInduk_() {
  const id = props_().ROOT_ID;
  if (id) { try { const f = DriveApp.getFolderById(id); if (!f.isTrashed()) return f; } catch (e) { /* */ } }
  const f = DriveApp.createFolder('SISTEM PANITIA - ' + (tetapan_().namaSekolah || 'SEKOLAH'));
  setProp_('ROOT_ID', f.getId());
  return f;
}
/** ID folder sistem disimpan dalam Script Properties — elak carian Drive yang lambat */
function folderSistem_(nama, kongsi) {
  const k = 'F_' + nama;
  const id = props_()[k];
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) { /* cipta semula */ } }
  const r = folderInduk_();
  const it = r.getFoldersByName(nama);
  const f = it.hasNext() ? it.next() : r.createFolder(nama);
  if (kongsi) kongsi_(f);
  setProp_(k, f.getId());
  return f;
}
/**
 * Folder subjek dikongsi SEKALI ("anyone with link", view) semasa dicipta.
 * Fail di dalamnya mewarisi akses, jadi setiap muat naik tidak perlu setSharing (jimat ~1 saat).
 * Folder 09 (kewangan) tidak dikongsi.
 */
function petaFolder_(p) { try { return p.subFolder ? JSON.parse(p.subFolder) : {}; } catch (e) { return {}; } }
function folderPanitia_(p, no) {
  no = String(Number(no));
  let peta = petaFolder_(p);
  // Laluan pantas tanpa kunci: folder sudah wujud dan sudah dikongsi
  if (peta[no] && (no === '9' || peta['k' + no])) {
    try { return DriveApp.getFolderById(peta[no]); } catch (e) { /* folder hilang — cipta semula di bawah */ }
  }
  // Perlu ubah peta folder: ambil kunci dan baca rekod terkini supaya muat naik serentak
  // tidak mencipta folder berganda atau menindih peta folder lain.
  kunci_();
  p = rows_('Panitia').find(x => x.id === p.id) || p;
  peta = petaFolder_(p);
  if (peta[no]) {
    try {
      const f = DriveApp.getFolderById(peta[no]);
      if (no !== '9' && !peta['k' + no]) { kongsi_(f); peta['k' + no] = 1; update_('Panitia', p.id, { subFolder: JSON.stringify(peta) }); }
      return f;
    } catch (e) { /* folder hilang — cipta semula */ }
  }
  let utama = null;
  if (p.folderId) { try { utama = DriveApp.getFolderById(p.folderId); if (utama.isTrashed()) utama = null; } catch (e) { utama = null; } }
  if (!utama) {
    utama = folderInduk_().createFolder('PANITIA ' + p.nama.toUpperCase());
    peta = {};
    update_('Panitia', p.id, { folderId: utama.getId() });
  }
  const f = utama.createFolder(FOLDER_NAMA[Number(no) - 1]);
  peta[no] = f.getId();
  if (no !== '9') { kongsi_(f); peta['k' + no] = 1; }
  update_('Panitia', p.id, { subFolder: JSON.stringify(peta) });
  return f;
}
function simpanBlob_(folder, b64, mime, nama) {
  const data = String(b64 || '').replace(/^data:[^,]*,/, '');
  mesti_(data.length > 0, 'Data fail kosong. Pilih fail semula.');
  let bytes;
  try { bytes = Utilities.base64Decode(data); }
  catch (e) { throw new Error('Data fail rosak semasa dihantar. Cuba lagi.'); }
  try {
    return folder.createFile(Utilities.newBlob(bytes, mime || 'application/octet-stream', nama || 'fail'));
  } catch (e) {
    throw new Error('Gagal menyimpan ke Google Drive (' + e.message + '). Pemilik skrip perlu jalankan ujiDrive() dalam editor Apps Script.');
  }
}
function kongsi_(file) {
  try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); }
  catch (e) { try { file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW); } catch (e2) { /* polisi organisasi */ } }
}
function buangDrive_(id) { if (!id) return; try { DriveApp.getFileById(id).setTrashed(true); } catch (e) { /* */ } }
function saizB64_(b64) { return Math.floor(String(b64 || '').replace(/^data:[^,]*,/, '').length * 0.75); }

/* =====================================================================
 *  AKAUN & PROFIL
 * ===================================================================== */
function awam() {
  const t = tetapan_();
  return { namaSekolah: t.namaSekolah, versi: APP_VERSION, panitia: senaraiAktif_().map(p => ({ id: p.id, nama: p.nama, kod: p.kod })) };
}

function daftar(q) {
  kunci_(); // semakan emel/No. KP unik mesti dalam kunci
  const nama = bersih_(q.nama, 120);
  const noKP = String(q.noKP || '').replace(/\D/g, '');
  const emel = bersih_(q.emel, 120).toLowerCase();
  const tel = String(q.telefon || '').replace(/[^\d+]/g, '');
  mesti_(nama.length >= 3, 'Isi nama penuh.');
  mesti_(/^\d{12}$/.test(noKP), 'No. Kad Pengenalan mesti 12 digit.');
  mesti_(emelSah_(emel), 'Format emel tidak sah.');
  mesti_(tel.replace(/\D/g, '').length >= 9, 'Isi nombor telefon bimbit yang sah.');
  mesti_(String(q.kataLaluan || '').length >= 6, 'Kata laluan sekurang-kurangnya 6 aksara.');
  const semua = rows_('Pengguna');
  mesti_(!semua.find(u => u.emel === emel), 'Emel ini sudah didaftarkan. Sila log masuk.');
  mesti_(!semua.find(u => u.noKP === noKP), 'No. Kad Pengenalan ini sudah didaftarkan.');
  const salt = uid_(16);
  const u = {
    id: uid_(), nama, noKP, emel, telefon: tel, peranan: semua.length ? 'GURU' : 'SUPER_ADMIN',
    hash: hash_(String(q.kataLaluan), salt), salt, status: 'AKTIF', dicipta: now_(), dikemaskini: now_()
  };
  insert_('Pengguna', u);
  const disimpan = userById_(u.id);
  segerakSubjek_(disimpan, q.subjek);
  log_(disimpan, '', '', 'DAFTAR', 'Pendaftaran akaun baharu');
  return bukaSesi_(disimpan);
}

/** Had cubaan: selepas MAKS_CUBA kegagalan, kunci HAD_CUBA_MIN minit (CacheService) */
const MAKS_CUBA = 8, HAD_CUBA_MIN = 15;
function semakCuba_(kunci) {
  const n = Number(cache_().get('CUBA_' + kunci)) || 0;
  mesti_(n < MAKS_CUBA, 'Terlalu banyak cubaan gagal. Cuba lagi selepas ' + HAD_CUBA_MIN + ' minit.');
}
function gagalCuba_(kunci) {
  const k = 'CUBA_' + kunci;
  cache_().put(k, String((Number(cache_().get(k)) || 0) + 1), HAD_CUBA_MIN * 60);
}

function login(q) {
  const emel = bersih_(q.emel, 120).toLowerCase();
  semakCuba_('L_' + emel);
  const u = rows_('Pengguna').find(x => x.emel === emel);
  if (!(u && u.hash === hash_(String(q.kataLaluan || ''), u.salt))) {
    gagalCuba_('L_' + emel);
    throw new Error('Emel atau kata laluan tidak betul.');
  }
  cache_().remove('CUBA_L_' + emel);
  mesti_(u.status !== 'NYAHAKTIF', 'Akaun anda telah dinyahaktifkan. Hubungi pentadbir sistem.');
  return bukaSesi_(u);
}

function bukaSesi_(u) {
  const token = uid_(32);
  insert_('Sesi', { token, userId: u.id, tamat: new Date(Date.now() + SESI_HARI * 864e5).toISOString(), dicipta: now_() });
  const d = init({}, u);
  d.token = token;
  return d;
}

function logout(q) { deleteWhere_('Sesi', s => s.token === q.token); return true; }

function init(q, u) {
  const t = tetapan_();
  const aktif = senaraiAktif_();
  const fail = rows_('Fail').filter(f => f.status === 'AKTIF');
  const panitia = aktif.map(p => {
    const a = akses_(u, p);
    if (!a.lihat && a.r !== 'MENUNGGU') return null;
    const kp = userById_(p.ketuaId);
    const failP = fail.filter(f => f.panitiaId === p.id && (a.kew || f.folderNo !== '9'));
    const terkini = failP.reduce((m, f) => (f.dicipta > m ? f.dicipta : m), '');
    return {
      id: p.id, nama: p.nama, kod: p.kod, ketua: kp ? kp.nama : '', bilAhli: ahliPanitia_(p).length,
      bilFail: failP.length, terkini, peranan: a.r, admin: a.admin, lihat: a.lihat
    };
  }).filter(Boolean);
  return {
    user: pub_(u, true),
    subjek: rows_('Keahlian').filter(k => k.userId === u.id).map(k => ({ panitiaId: k.panitiaId, tahun: k.tahun, status: k.status })),
    panitia,
    semuaPanitia: aktif.map(p => ({ id: p.id, nama: p.nama, kod: p.kod })),
    tetapan: { namaSekolah: t.namaSekolah, kodSekolah: t.kodSekolah },
    versi: APP_VERSION,
    urlApp: urlApp_(),
    jadualAda: (() => { try { return !!padanJadual_(u, jadualIndeks_()); } catch (e) { return false; } })()
  };
}

function urlApp_() { try { return ScriptApp.getService().getUrl() || ''; } catch (e) { return ''; } }

function segerakSubjek_(u, subjek) {
  if (!Array.isArray(subjek)) return;
  const aktif = senaraiAktif_();
  const ids = aktif.map(p => p.id);
  const mahu = subjek.filter(s => s && ids.indexOf(s.panitiaId) >= 0);
  const ada = rows_('Keahlian').filter(k => k.userId === u.id);
  const buang = ada.filter(k => !mahu.find(s => s.panitiaId === k.panitiaId)).filter(k => {
    const p = aktif.find(x => x.id === k.panitiaId);
    return !p || (p.ketuaId !== u.id && p.setiausahaId !== u.id);
  }).map(k => k.id);
  if (buang.length) deleteWhere_('Keahlian', k => buang.indexOf(k.id) >= 0);
  mahu.forEach(s => {
    const tahun = bersih_(s.tahun, 60);
    const k = rows_('Keahlian').find(x => x.userId === u.id && x.panitiaId === s.panitiaId);
    if (k) { if (k.tahun !== tahun) update_('Keahlian', k.id, { tahun }); }
    else insert_('Keahlian', { id: uid_(), panitiaId: s.panitiaId, userId: u.id, tahun, status: u.peranan === 'SUPER_ADMIN' ? 'SAH' : 'MENUNGGU', dicipta: now_() });
  });
}

function kemaskiniProfil(q, u) {
  const d = q.profil || {};
  const patch = {};
  ['nama', 'jawatan', 'opsyen', 'kelulusan', 'gred', 'pengalaman'].forEach(k => { if (d[k] !== undefined) patch[k] = bersih_(d[k], 200); });
  if (patch.nama !== undefined) mesti_(patch.nama.length >= 3, 'Isi nama penuh.');
  if (d.telefon !== undefined) patch.telefon = String(d.telefon).replace(/[^\d+]/g, '');
  if (d.noKP !== undefined) {
    const n = String(d.noKP).replace(/\D/g, '');
    mesti_(/^\d{12}$/.test(n), 'No. Kad Pengenalan mesti 12 digit.');
    mesti_(!rows_('Pengguna').find(x => x.noKP === n && x.id !== u.id), 'No. Kad Pengenalan sudah digunakan.');
    patch.noKP = n;
  }
  if (d.emel !== undefined) {
    const e = bersih_(d.emel, 120).toLowerCase();
    mesti_(emelSah_(e), 'Format emel tidak sah.');
    mesti_(!rows_('Pengguna').find(x => x.emel === e && x.id !== u.id), 'Emel sudah digunakan oleh akaun lain.');
    patch.emel = e;
  }
  patch.dikemaskini = now_();
  update_('Pengguna', u.id, patch);
  segerakSubjek_(userById_(u.id), q.subjek);
  return _BATCH ? { ok: true } : init({}, userById_(u.id));
}

function tukarKataLaluan(q, u) {
  if (u.wajibTukar !== 'YA') mesti_(u.hash === hash_(String(q.lama || ''), u.salt), 'Kata laluan semasa tidak betul.');
  mesti_(String(q.baru || '').length >= 6, 'Kata laluan baharu sekurang-kurangnya 6 aksara.');
  mesti_(String(q.baru) !== KATA_LALUAN_LALAI, 'Kata laluan baharu tidak boleh sama dengan kata laluan lalai.');
  mesti_(u.hash !== hash_(String(q.baru), u.salt), 'Kata laluan baharu mesti berbeza daripada yang lama.');
  const salt = uid_(16);
  update_('Pengguna', u.id, { hash: hash_(String(q.baru), salt), salt, wajibTukar: '', dikemaskini: now_() });
  deleteWhere_('Sesi', s => s.userId === u.id && s.token !== q.token);
  return true;
}

function bolehUrusGuru_(u, sasaranId) {
  if (u.id === sasaranId || u.peranan === 'SUPER_ADMIN') return true;
  return senaraiAktif_().some(p => (p.ketuaId === u.id || p.setiausahaId === u.id) && ahliPanitia_(p).some(a => a.id === sasaranId));
}

function muatNaikFoto(q, u) {
  const sasaran = q.userId ? userById_(q.userId) : u;
  mesti_(sasaran, 'Pengguna tidak dijumpai.');
  mesti_(bolehUrusGuru_(u, sasaran.id));
  mesti_(saizB64_(q.data) < 3 * 1048576, 'Gambar terlalu besar.');
  const f = simpanBlob_(folderSistem_('_FOTO_PROFIL'), q.data, 'image/jpeg', 'foto_' + sasaran.id + '_' + Date.now() + '.jpg');
  const fotoLama = sasaran.fotoId;
  update_('Pengguna', sasaran.id, { fotoId: f.getId(), dikemaskini: now_() });
  const durl = 'data:image/jpeg;base64,' + String(q.data).replace(/^data:[^,]*,/, '');
  if (durl.length < 95000) try { cache_().put('IMG_' + f.getId(), durl, CACHE_TTL); } catch (e) { /* */ }
  buangDrive_(fotoLama);
  return { fotoId: f.getId() };
}

function muatNaikJadual(q, u) {
  const mime = String(q.mime || '');
  mesti_(/^(application\/pdf|image\/(jpeg|png|webp))$/.test(mime), 'Jadual waktu mesti PDF, JPEG atau PNG.');
  mesti_(saizB64_(q.data) <= MAKS_FAIL_MB * 1048576, 'Saiz fail melebihi ' + MAKS_FAIL_MB + 'MB.');
  const f = simpanBlob_(folderSistem_('_JADUAL_WAKTU', true), q.data, mime, 'Jadual_' + u.nama + '_' + Date.now());
  const lama = u.jadualId;
  update_('Pengguna', u.id, { jadualId: f.getId(), jadualMime: mime, dikemaskini: now_() });
  buangDrive_(lama);
  log_(u, '', '8', 'JADUAL', 'Jadual waktu peribadi dikemas kini');
  return { jadualId: f.getId(), jadualMime: mime };
}

function lihatJadual(q, u) {
  const s = userById_(q.userId || u.id);
  mesti_(s && s.jadualId, 'Jadual waktu belum dimuat naik.');
  mesti_(bolehUrusGuru_(u, s.id) || senaraiAktif_().some(p => akses_(u, p).lihat && ahliPanitia_(p).some(a => a.id === s.id)));
  let fd;
  try { fd = DriveApp.getFileById(s.jadualId); } catch (e) { throw new Error('Salinan jadual waktu tidak lagi wujud dalam Google Drive. Muat naik semula.'); }
  const b = fd.getBlob();
  return { nama: 'Jadual Waktu ' + s.nama, mime: b.getContentType(), driveId: s.jadualId, data: Utilities.base64Encode(b.getBytes()) };
}

/** Proksi gambar profil — elak masalah polisi perkongsian Drive organisasi */
function imej(q) {
  if (q.pemantau) pemantauSah_(q.pemantau); else auth_(q.token);
  const sah = {};
  rows_('Pengguna').forEach(u => { if (u.fotoId) sah[u.fotoId] = 1; });
  const tt = tetapan_();
  ['gbFoto', 'pkFoto', 'sukFoto'].forEach(k => { if (tt[k]) sah[tt[k]] = 1; });
  const out = {};
  const ids = (q.ids || []).slice(0, 20).filter(id => sah[id]);
  const c = cache_();
  const ada = ids.length ? c.getAll(ids.map(id => 'IMG_' + id)) : {};
  const simpan = {};
  ids.forEach(id => {
    if (ada['IMG_' + id]) { out[id] = ada['IMG_' + id]; return; }
    try {
      const b = DriveApp.getFileById(id).getBlob();
      out[id] = 'data:' + b.getContentType() + ';base64,' + Utilities.base64Encode(b.getBytes());
      if (out[id].length < 95000) simpan['IMG_' + id] = out[id];
    } catch (e) { /* */ }
  });
  if (Object.keys(simpan).length) try { c.putAll(simpan, CACHE_TTL); } catch (e) { /* */ }
  return out;
}

function muatTurun(q) {
  const f = failById_(q.id);
  const p = panitia_(f.panitiaId);
  if (q.pemantau) {
    const pm = pemantauSah_(q.pemantau);
    mesti_(pm.panitiaId === p.id && f.folderNo !== '9' && f.status === 'AKTIF');
  } else {
    const u = auth_(q.token);
    const a = akses_(u, p);
    mesti_(a.lihat && (f.folderNo !== '9' || a.kew));
  }
  mesti_(f.driveId, 'Item ini ialah pautan luar. Buka pautan untuk melihatnya.');
  let fd;
  try { fd = DriveApp.getFileById(f.driveId); } catch (e) { throw new Error('Fail tidak lagi wujud dalam Google Drive.'); }
  mesti_(fd.getSize() < 25 * 1048576, 'Fail terlalu besar untuk dimuat turun melalui aplikasi. Buka melalui Google Drive.');
  const b = fd.getBlob();
  return { nama: b.getName(), mime: b.getContentType(), data: Utilities.base64Encode(b.getBytes()) };
}

/* =====================================================================
 *  SYNC (satu panggilan) + BATCH (outbox)
 * ===================================================================== */
let _BATCH = false;
function balasDetail_(id, u) { if (_BATCH) return { ok: true }; flushLog_(); return panitiaDetail({ id }, u); }
function balasAdmin_(q, u) { return _BATCH ? { ok: true } : senaraiPengguna(q, u); }

/** ETag = versi semua jadual + pengguna. Jika sama dengan klien, tiada data dihantar. */
function etag_(u) {
  const g = gens_();
  return APP_VERSION + '|' + u.id + '|' + ['Tetapan', 'Pengguna', 'Panitia', 'Keahlian', 'Carta', 'Fail', 'Tugasan', 'Pemantau', 'Log'].map(n => g[n]).join('.');
}

/** Satu panggilan untuk semua data yang diperlukan skrin: init + butiran setiap panitia pengguna */
function sync(q, u) {
  const et = etag_(u);
  if (q.etag && q.etag === et) return { sama: true, etag: et };
  const d = init({}, u);
  const ids = d.panitia.filter(p => p.lihat && p.peranan && p.peranan !== 'MENUNGGU').map(p => p.id);
  if (q.aktif && ids.indexOf(q.aktif) < 0 && d.panitia.some(p => p.id === q.aktif && p.lihat)) ids.push(q.aktif);
  const detail = {};
  ids.slice(0, 12).forEach(id => { try { detail[id] = panitiaDetail({ id }, u); } catch (e) { /* abaikan */ } });
  return { etag: et, init: d, detail };
}

const BOLEH_BATCH = ['kemaskiniProfil', 'simpanCarta', 'lantikSU', 'sahkanAhli', 'tambahAhli', 'keluarkanAhli',
  'padamFail', 'pulihFail', 'padamKekal', 'tambahTugasan', 'padamTugasan', 'batalPemantau',
  'simpanPanitia', 'padamPanitia', 'lantikKP', 'kemaskiniPengguna', 'simpanTetapan'];

/** Jalankan beberapa perubahan dari outbox klien dalam satu permintaan. ID operasi menghalang hantar berganda. */
function batch(q, u) {
  const ops = (q.ops || []).slice(0, 50);
  const c = cache_();
  const sudah = ops.length ? c.getAll(ops.filter(o => o.id).map(o => 'OP_' + o.id)) : {};
  _BATCH = true;
  let hasil;
  try {
    hasil = ops.map(op => {
      try {
        mesti_(BOLEH_BATCH.indexOf(op.action) >= 0, 'Tindakan tidak dibenarkan: ' + op.action);
        if (op.id && sudah['OP_' + op.id]) return { id: op.id, ok: true, ulang: true };
        ACTIONS[op.action](Object.assign({}, op.data || {}, { token: q.token }), userById_(u.id) || u);
        if (op.id) c.put('OP_' + op.id, '1', CACHE_TTL);
        return { id: op.id, ok: true };
      } catch (e) {
        return { id: op.id, ok: false, error: mesejRalat_(e) };
      }
    });
  } finally { _BATCH = false; }
  return { hasil, perluSync: true };
}

/* =====================================================================
 *  JADUAL WAKTU GURU (dibaca daripada Google Sheet CSV yang diterbitkan)
 *  Lajur: Sesi, Hari, Period, Masa, Kelas, Subjek, Guru
 *  Hasil dicache 30 minit dalam CacheService.
 * ===================================================================== */
const JADUAL_TTL = 1800;

function tokenNama_(s) {
  return String(s || '').toUpperCase().replace(/[^A-Z0-9 ]+/g, ' ').split(/\s+/).filter(Boolean)
    .map(t => (t === 'BT' || t === 'BTE' || t === 'BINTE' || t === 'BNT') ? 'BINTI' : (t === 'B' || t === 'BN') ? 'BIN' : t);
}
function kunciNama_(s) { return tokenNama_(s).join(' '); }

function jadualIndeks_(paksa) {
  const url = String(tetapan_().jadualCsvUrl || '').trim();
  if (!url) return null;
  const kunci = 'JW_' + Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, url)).slice(0, 16);
  if (!paksa) {
    const c = cacheBaca_(kunci); if (c) return c;
    const gagal = cache_().get(kunci + '_GAGAL');
    if (gagal) throw new Error(gagal);
  }
  let res;
  try { res = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true }); }
  catch (e) { cache_().put(kunci + '_GAGAL', 'Jadual waktu tidak dapat dibaca: ' + e.message, 300); throw e; }
  if (res.getResponseCode() !== 200) {
    const m = 'Jadual waktu tidak dapat dibaca (HTTP ' + res.getResponseCode() + '). Pastikan Google Sheet diterbitkan sebagai CSV.';
    cache_().put(kunci + '_GAGAL', m, 300); // elak cuba semula setiap permintaan selama 5 minit
    throw new Error(m);
  }
  const csv = Utilities.parseCsv(res.getContentText('UTF-8'));
  mesti_(csv.length > 1, 'Fail CSV jadual waktu kosong.');
  const h = csv[0].map(x => String(x).trim().toLowerCase());
  const kol = n => h.indexOf(n);
  const iS = kol('sesi'), iH = kol('hari'), iP = kol('period'), iM = kol('masa'), iK = kol('kelas'), iSub = kol('subjek'), iG = kol('guru');
  mesti_(iH >= 0 && iM >= 0 && iG >= 0, 'Tajuk lajur CSV mesti mengandungi Hari, Masa dan Guru.');
  const nama = {}, slot = {};
  for (let r = 1; r < csv.length; r++) {
    const row = csv[r];
    const g = String(row[iG] || '').trim();
    if (!g) continue;
    const k = kunciNama_(g);
    if (!k) continue;
    if (!slot[k]) { slot[k] = []; nama[k] = g.replace(/\s+/g, ' '); }
    slot[k].push([
      iS >= 0 ? String(row[iS] || '').trim().toLowerCase() : '',
      String(row[iH] || '').trim().toUpperCase(),
      iP >= 0 ? String(row[iP] || '').trim() : '',
      String(row[iM] || '').trim(),
      iK >= 0 ? String(row[iK] || '').trim() : '',
      iSub >= 0 ? String(row[iSub] || '').replace(/\s+/g, ' ').trim() : ''
    ]);
  }
  const out = { nama, slot, dikemas: now_() };
  cacheTulis_(kunci, out, JADUAL_TTL);
  return out;
}

/** Padankan guru sistem dengan nama dalam jadual: manual → tepat → hampir (token) */
function padanJadual_(u, idx) {
  if (!idx || !u) return null;
  if (u.namaJadual) { const k = kunciNama_(u.namaJadual); if (idx.slot[k]) return { k, cara: 'manual' }; }
  const k = kunciNama_(u.nama);
  if (!k) return null;
  if (idx.slot[k]) return { k, cara: 'tepat' };
  const HENTI = { BIN: 1, BINTI: 1, A: 1, P: 1, L: 1, AL: 1, AP: 1 };
  const set = x => new Set(x.split(' ').filter(t => !HENTI[t]));
  const A = set(k);
  if (!A.size) return null;
  let terbaik = null, b1 = 0, b2 = 0;
  Object.keys(idx.slot).forEach(kk => {
    const B = set(kk);
    let sama = 0;
    A.forEach(t => { if (B.has(t)) sama++; });
    const j = sama / (A.size + B.size - sama);
    if (j > b1) { b2 = b1; b1 = j; terbaik = kk; } else if (j > b2) b2 = j;
  });
  if (terbaik && b1 >= 0.75 && b1 - b2 >= 0.15) return { k: terbaik, cara: 'hampir' };
  return null;
}

function jadualGuru(q, u) {
  const s = q.userId ? userById_(q.userId) : u;
  mesti_(s, 'Guru tidak dijumpai.');
  mesti_(s.id === u.id || bolehUrusGuru_(u, s.id) || senaraiAktif_().some(p => akses_(u, p).lihat && ahliPanitia_(p).some(a => a.id === s.id)));
  const idx = jadualIndeks_(!!q.segar && u.peranan === 'SUPER_ADMIN');
  if (!idx) return { tiadaSumber: true, userId: s.id, nama: s.nama };
  const m = padanJadual_(s, idx);
  const bolehPilih = s.id === u.id || u.peranan === 'SUPER_ADMIN';
  return {
    userId: s.id, nama: s.nama, padan: m ? m.cara : null, namaJadual: m ? idx.nama[m.k] : '',
    slot: m ? idx.slot[m.k] : [], dikemas: idx.dikemas, bilGuru: Object.keys(idx.slot).length, bolehPilih,
    calon: bolehPilih && (!m || q.calon) ? Object.keys(idx.nama).map(k => idx.nama[k]).sort() : []
  };
}

function setNamaJadual(q, u) {
  const s = q.userId ? userById_(q.userId) : u;
  mesti_(s, 'Guru tidak dijumpai.');
  mesti_(s.id === u.id || u.peranan === 'SUPER_ADMIN', 'Hanya guru sendiri atau Super Admin boleh menukar padanan jadual.');
  const nm = bersih_(q.nama, 150);
  if (nm) { const idx = jadualIndeks_(); mesti_(idx && idx.slot[kunciNama_(nm)], 'Nama tidak dijumpai dalam jadual waktu.'); }
  update_('Pengguna', s.id, { namaJadual: nm });
  return jadualGuru({ userId: s.id }, userById_(u.id));
}

/* =====================================================================
 *  PANITIA, CARTA & AHLI
 * ===================================================================== */
function panitiaDetail(q, u) {
  const p = panitia_(q.id);
  const a = akses_(u, p);
  mesti_(a.lihat || a.r === 'MENUNGGU', 'Anda bukan ahli panitia ini.');
  if (!a.lihat) return { panitia: { id: p.id, nama: p.nama, kod: p.kod }, akses: a, menungguSaya: true };

  const ahli = ahliPanitia_(p);
  const keah = rows_('Keahlian').filter(k => k.panitiaId === p.id);
  const failSemua = rows_('Fail').filter(f => f.panitiaId === p.id);
  const fail = failSemua.filter(f => f.status === 'AKTIF' && (a.kew || f.folderNo !== '9'))
    .map(failPub_).sort((x, y) => (y.dicipta > x.dicipta ? 1 : -1));
  const tugasan = rows_('Tugasan').filter(t => t.panitiaId === p.id)
    .map(t => ({ id: t.id, tajuk: t.tajuk, folderNo: Number(t.folderNo), tarikhAkhir: t.tarikhAkhir }));
  let idxJW = null;
  if (a.su) { try { idxJW = jadualIndeks_(); } catch (e) { idxJW = null; } }
  const semakan = a.su ? ahli.map(m => ({
    userId: m.id, nama: m.nama, jadual: !!m.jadualId || !!padanJadual_(m, idxJW),
    tugasan: tugasan.map(t => ({ id: t.id, siap: fail.some(f => f.tugasanId === t.id && f.pemilikId === m.id) }))
  })) : null;
  const kp = userById_(p.ketuaId);

  return {
    panitia: { id: p.id, nama: p.nama, kod: p.kod, ketuaId: p.ketuaId, setiausahaId: p.setiausahaId, ketua: kp ? kp.nama : '' },
    akses: a,
    sekolah: tetapan_().namaSekolah,
    ahli: ahli.map(m => Object.assign(pub_(m), { tahun: (keah.find(k => k.userId === m.id) || {}).tahun || '' })),
    menunggu: a.urus ? keah.filter(k => k.status === 'MENUNGGU')
      .map(k => ({ id: k.id, tahun: k.tahun, user: ringkas_(userById_(k.userId)) })).filter(x => x.user) : [],
    pentadbir: pentadbir_(),
    carta: cartaBina_(p, ahli),
    fail,
    tong: a.urus ? failSemua.filter(f => f.status === 'TRASH').map(failPub_) : [],
    tugasan,
    semakan,
    pemantau: a.urus ? rows_('Pemantau').filter(x => x.panitiaId === p.id && x.status === 'AKTIF' && new Date(x.tamat) > new Date())
      .map(x => ({ id: x.id, kod: x.kod, label: x.label, tamat: x.tamat })) : [],
    log: logTerkini_(400).filter(l => l.panitiaId === p.id && (a.kew || l.folderNo !== '9')).slice(-25).reverse()
      .map(l => ({ masa: l.masa, nama: l.nama, tindakan: l.tindakan, butiran: l.butiran, folderNo: l.folderNo })),
    dimuat: now_()
  };
}

/** Bina carta Tahap 3–5. Tahap 1–2 datang daripada Tetapan (pentadbir sekolah). */
function cartaBina_(p, ahli) {
  const ids = ahli.map(a => a.id);
  let nod = rows_('Carta').filter(c => c.panitiaId === p.id && ids.indexOf(c.userId) >= 0 && c.userId !== p.ketuaId)
    .map(c => ({ userId: c.userId, tahap: Number(c.tahap) === 4 ? 4 : 5, susunan: Number(c.susunan) || 0, jawatan: c.jawatan || 'Ahli Panitia' }));
  if (p.ketuaId && ids.indexOf(p.ketuaId) >= 0) nod.push({ userId: p.ketuaId, tahap: 3, susunan: 0, jawatan: 'Ketua Panitia' });
  if (p.setiausahaId && ids.indexOf(p.setiausahaId) >= 0 && p.setiausahaId !== p.ketuaId) {
    const s = nod.find(n => n.userId === p.setiausahaId);
    if (s) {
      if (s.tahap !== 4) { s.tahap = 4; s.susunan = -1; }
      if (!s.jawatan || s.jawatan === 'Ahli Panitia') s.jawatan = 'Setiausaha Panitia';
    } else nod.push({ userId: p.setiausahaId, tahap: 4, susunan: -1, jawatan: 'Setiausaha Panitia' });
  }
  ids.forEach(id => { if (!nod.find(n => n.userId === id)) nod.push({ userId: id, tahap: 5, susunan: 999, jawatan: 'Ahli Panitia' }); });
  return nod.sort((a, b) => a.tahap - b.tahap || a.susunan - b.susunan);
}

function simpanCarta(q, u) {
  const p = panitia_(q.panitiaId);
  mesti_(akses_(u, p).urus, 'Hanya Ketua Panitia boleh menyusun carta.');
  kunci_();
  const ids = ahliPanitia_(p).map(a => a.id);
  const nod = (q.nod || []).filter(n => ids.indexOf(n.userId) >= 0 && n.userId !== p.ketuaId);
  const kekal = rows_('Carta', true).filter(r => r.panitiaId !== p.id);
  const baru = nod.map((n, i) => ({
    id: uid_(), panitiaId: p.id, userId: n.userId, tahap: Number(n.tahap) === 4 ? 4 : 5, susunan: i,
    jawatan: bersih_(n.jawatan, 60) || 'Ahli Panitia'
  }));
  tulisSemula_('Carta', kekal.concat(baru));
  const su = baru.find(n => n.tahap === 4);
  update_('Panitia', p.id, { setiausahaId: su ? su.userId : '' });
  log_(u, p.id, '1', 'CARTA', 'Carta organisasi dikemas kini');
  return balasDetail_(p.id, u);
}

function lantikSU(q, u) {
  const p = panitia_(q.panitiaId);
  mesti_(akses_(u, p).urus, 'Hanya Ketua Panitia boleh melantik Setiausaha.');
  mesti_(ahliPanitia_(p).some(a => a.id === q.userId), 'Guru ini bukan ahli panitia.');
  mesti_(q.userId !== p.ketuaId, 'Ketua Panitia tidak boleh menjadi Setiausaha.');
  update_('Panitia', p.id, { setiausahaId: q.userId });
  const s = userById_(q.userId);
  log_(u, p.id, '1', 'LANTIK SU', s.nama + ' dilantik sebagai Setiausaha Panitia');
  return balasDetail_(p.id, u);
}

function sahkanAhli(q, u) {
  const k = rows_('Keahlian').find(x => x.id === q.id);
  mesti_(k, 'Permohonan tidak dijumpai.');
  const p = panitia_(k.panitiaId);
  mesti_(akses_(u, p).urus, 'Hanya Ketua Panitia boleh mengesahkan ahli.');
  const g = userById_(k.userId);
  if (q.lulus) {
    update_('Keahlian', k.id, { status: 'SAH' });
    log_(u, p.id, '1', 'SAHKAN AHLI', (g ? g.nama : '') + ' disahkan sebagai ahli');
  } else {
    deleteWhere_('Keahlian', x => x.id === k.id);
    log_(u, p.id, '1', 'TOLAK AHLI', (g ? g.nama : '') + ' ditolak');
  }
  return balasDetail_(p.id, u);
}

/** Senarai guru berdaftar yang belum menjadi ahli panitia (untuk KP menambah ahli) */
function calonAhli(q, u) {
  const p = panitia_(q.panitiaId);
  mesti_(akses_(u, p).urus, 'Hanya Ketua Panitia boleh menambah ahli.');
  const ahliIds = ahliPanitia_(p).map(a => a.id);
  const tunggu = {};
  rows_('Keahlian').forEach(k => { if (k.panitiaId === p.id && k.status === 'MENUNGGU') tunggu[k.userId] = 1; });
  return rows_('Pengguna').filter(x => x.status !== 'NYAHAKTIF' && ahliIds.indexOf(x.id) < 0)
    .map(x => ({ id: x.id, nama: x.nama, emel: x.emel, opsyen: x.opsyen, fotoId: x.fotoId, menunggu: !!tunggu[x.id] }))
    .sort((a, b) => (b.menunggu - a.menunggu) || a.nama.localeCompare(b.nama));
}

/** KP menambah guru terus sebagai ahli yang disahkan */
function tambahAhli(q, u) {
  const p = panitia_(q.panitiaId);
  mesti_(akses_(u, p).urus, 'Hanya Ketua Panitia boleh menambah ahli.');
  const ids = [...new Set((q.userIds || []).filter(Boolean))].slice(0, 200);
  mesti_(ids.length, 'Pilih sekurang-kurangnya seorang guru.');
  const tahun = bersih_(q.tahun, 60);
  const baru = [], nama = [];
  ids.forEach(id => {
    const g = userById_(id);
    if (!g || g.status === 'NYAHAKTIF') return;
    const k = rows_('Keahlian').find(x => x.panitiaId === p.id && x.userId === id);
    if (k) {
      const patch = { status: 'SAH' };
      if (tahun && !k.tahun) patch.tahun = tahun;
      update_('Keahlian', k.id, patch);
    } else {
      baru.push({ id: uid_(), panitiaId: p.id, userId: id, tahun, status: 'SAH', dicipta: now_() });
    }
    nama.push(g.nama);
  });
  insertMany_('Keahlian', baru);
  if (nama.length) log_(u, p.id, '1', 'TAMBAH AHLI', nama.length + ' guru: ' + nama.join(', '));
  return balasDetail_(p.id, u);
}

function keluarkanAhli(q, u) {
  const p = panitia_(q.panitiaId);
  mesti_(akses_(u, p).urus, 'Hanya Ketua Panitia boleh mengeluarkan ahli.');
  mesti_(q.userId !== p.ketuaId, 'Ketua Panitia tidak boleh dikeluarkan. Super Admin boleh menukar Ketua Panitia.');
  const g = userById_(q.userId);
  deleteWhere_('Keahlian', k => k.panitiaId === p.id && k.userId === q.userId);
  if (p.setiausahaId === q.userId) update_('Panitia', p.id, { setiausahaId: '' });
  log_(u, p.id, '1', 'KELUARKAN AHLI', g ? g.nama : q.userId);
  return balasDetail_(p.id, u);
}

/* =====================================================================
 *  FOLDER DIGITAL & FAIL
 * ===================================================================== */
function semakFolder_(u, p, no) {
  const n = Number(no);
  mesti_(n >= 1 && n <= 10, 'Folder tidak sah.');
  const a = akses_(u, p);
  mesti_(a.lihat, 'Anda bukan ahli panitia ini.');
  if (n === 9) mesti_(a.kew, 'Folder 09 terhad kepada Ketua Panitia, Setiausaha dan Pentadbir.');
  return String(n);
}

/** Pastikan tugasanId (jika ada) wujud dalam panitia yang sama */
function tugasanSah_(p, id) {
  id = bersih_(id, 20);
  if (!id) return '';
  return rows_('Tugasan').some(t => t.id === id && t.panitiaId === p.id) ? id : '';
}

function muatNaikFail(q, u) {
  const p = panitia_(q.panitiaId);
  const no = semakFolder_(u, p, q.folderNo);
  const f = q.fail || {};
  const saiz = saizB64_(f.data);
  mesti_(saiz > 0, 'Tiada fail dipilih.');
  mesti_(saiz <= MAKS_FAIL_MB * 1048576, 'Saiz fail melebihi had ' + MAKS_FAIL_MB + 'MB.');
  const file = simpanBlob_(folderPanitia_(p, no), f.data, f.mime, bersih_(f.nama, 150) || 'fail');
  const rec = {
    id: uid_(), panitiaId: p.id, folderNo: no, tajuk: bersih_(q.tajuk, 150) || file.getName(),
    penerangan: bersih_(q.penerangan, 500), jenis: 'DRIVE', url: 'https://drive.google.com/file/d/' + file.getId() + '/view',
    driveId: file.getId(), mime: f.mime || 'application/octet-stream', saiz: saiz, tugasanId: tugasanSah_(p, q.tugasanId), pemilikId: u.id,
    pemilikNama: u.nama, dicipta: now_(), status: 'AKTIF'
  };
  insert_('Fail', rec);
  log_(u, p.id, no, 'MUAT NAIK', rec.tajuk);
  return failPub_(rec);
}

function tambahPautan(q, u) {
  const p = panitia_(q.panitiaId);
  const no = semakFolder_(u, p, q.folderNo);
  const url = bersih_(q.url, 1000);
  mesti_(/^https?:\/\/\S+$/i.test(url), 'Pautan mesti bermula dengan http:// atau https://');
  const tajuk = bersih_(q.tajuk, 150);
  mesti_(tajuk, 'Isi tajuk dokumen.');
  const rec = {
    id: uid_(), panitiaId: p.id, folderNo: no, tajuk, penerangan: bersih_(q.penerangan, 500), jenis: 'LINK',
    url, driveId: '', mime: '', saiz: 0, tugasanId: tugasanSah_(p, q.tugasanId), pemilikId: u.id,
    pemilikNama: u.nama, dicipta: now_(), status: 'AKTIF'
  };
  insert_('Fail', rec);
  log_(u, p.id, no, 'TAMBAH PAUTAN', tajuk);
  return failPub_(rec);
}

function padamFail(q, u) {
  const f = failById_(q.id);
  mesti_(f.status === 'AKTIF', 'Fail sudah berada dalam tong sampah.');
  const p = panitia_(f.panitiaId);
  const a = akses_(u, p);
  mesti_(a.urus || (f.pemilikId === u.id && a.lihat), 'Hanya pemilik fail atau Ketua Panitia boleh memadam.');
  if (f.folderNo === '9') mesti_(a.kew);
  update_('Fail', f.id, { status: 'TRASH', dipadamPada: now_(), dipadamOleh: u.nama });
  log_(u, p.id, f.folderNo, 'PADAM', 'Ke tong sampah: ' + f.tajuk);
  return true;
}

function pulihFail(q, u) {
  const f = failById_(q.id);
  const p = panitia_(f.panitiaId);
  mesti_(akses_(u, p).urus, 'Hanya Ketua Panitia boleh memulihkan fail.');
  update_('Fail', f.id, { status: 'AKTIF', dipadamPada: '', dipadamOleh: '' });
  log_(u, p.id, f.folderNo, 'PULIH', f.tajuk);
  return true;
}

function padamKekal(q, u) {
  const f = failById_(q.id);
  const p = panitia_(f.panitiaId);
  mesti_(akses_(u, p).urus, 'Hanya Ketua Panitia boleh memadam kekal.');
  mesti_(f.status === 'TRASH', 'Pindahkan fail ke tong sampah dahulu.');
  buangDrive_(f.driveId);
  update_('Fail', f.id, { status: 'KEKAL_PADAM' });
  log_(u, p.id, f.folderNo, 'PADAM KEKAL', f.tajuk);
  return true;
}

/** Trigger harian: padam kekal fail > 30 hari dalam tong, buang sesi & pautan pemantau tamat */
function bersihkanTong() {
  ensureSchema_();
  const had = Date.now() - TONG_HARI * 864e5;
  rows_('Fail').filter(f => f.status === 'TRASH' && new Date(f.dipadamPada).getTime() < had).forEach(f => {
    buangDrive_(f.driveId);
    update_('Fail', f.id, { status: 'KEKAL_PADAM' });
    log_('Sistem', f.panitiaId, f.folderNo, 'PADAM KEKAL (AUTO)', f.tajuk);
  });
  deleteWhere_('Sesi', s => new Date(s.tamat).getTime() < Date.now());
  rows_('Pemantau').filter(x => x.status === 'AKTIF' && new Date(x.tamat) < new Date())
    .forEach(x => update_('Pemantau', x.id, { status: 'TAMAT' }));
  selesai_();
}

/* ---------- Tugasan wajib (senarai semak) ---------- */
function tambahTugasan(q, u) {
  const p = panitia_(q.panitiaId);
  mesti_(akses_(u, p).su, 'Hanya Ketua Panitia atau Setiausaha boleh menambah tugasan.');
  const tajuk = bersih_(q.tajuk, 120);
  mesti_(tajuk, 'Isi nama dokumen wajib.');
  let no = Math.floor(Number(q.folderNo)) || 4;
  if (no < 1 || no > 10) no = 4;
  if (no === 9) mesti_(akses_(u, p).kew, 'Folder 09 terhad.');
  insert_('Tugasan', { id: uid_(), panitiaId: p.id, tajuk, folderNo: no, tarikhAkhir: bersih_(q.tarikhAkhir, 20), diciptaOleh: u.nama, dicipta: now_() });
  log_(u, p.id, String(no), 'TUGASAN', 'Dokumen wajib baharu: ' + tajuk);
  return balasDetail_(p.id, u);
}

function padamTugasan(q, u) {
  const t = rows_('Tugasan').find(x => x.id === q.id);
  mesti_(t, 'Tugasan tidak dijumpai.');
  const p = panitia_(t.panitiaId);
  mesti_(akses_(u, p).su);
  deleteWhere_('Tugasan', x => x.id === t.id);
  log_(u, p.id, t.folderNo, 'PADAM TUGASAN', t.tajuk);
  return balasDetail_(p.id, u);
}

/* =====================================================================
 *  PENJANA LAPORAN SATU MUKA SURAT & EKSPORT CARTA
 * ===================================================================== */
function janaLaporan(q, u) {
  const p = panitia_(q.panitiaId);
  semakFolder_(u, p, 10);
  const d = {
    program: bersih_(q.program, 150), tarikh: bersih_(q.tarikh, 60), sasaran: bersih_(q.sasaran, 300),
    objektif: bersih_(q.objektif, 900), impak: bersih_(q.impak, 900)
  };
  mesti_(d.program, 'Isi nama program.');
  const gambar = (q.gambar || []).filter(Boolean).slice(0, 4);
  mesti_(gambar.length >= 1, 'Pilih sekurang-kurangnya satu gambar aktiviti.');
  gambar.forEach(g => mesti_(/^data:image\/(jpeg|png);base64,/.test(g), 'Format gambar tidak sah.'));

  const t = tetapan_();
  const sel = gambar.map(g => '<td style="width:50%;padding:4px;text-align:center"><img src="' + g + '" style="width:330px;height:236px"></td>');
  while (sel.length % 2) sel.push('<td></td>');
  let grid = '';
  for (let i = 0; i < sel.length; i += 2) grid += '<tr>' + sel[i] + sel[i + 1] + '</tr>';
  const baris = (l, v) => '<tr><td style="width:24%;padding:6px 8px;border:1px solid #9aa7a0;background:#eef3ef;font-weight:bold;vertical-align:top">' + l +
    '</td><td style="padding:6px 8px;border:1px solid #9aa7a0;vertical-align:top">' + esc_(v).replace(/\n/g, '<br>') + '</td></tr>';

  const html = '<html><head><meta charset="utf-8"><style>@page{size:A4;margin:14mm}body{font-family:Arial,sans-serif;font-size:10.5pt;color:#1c2b39}</style></head><body>' +
    '<div style="text-align:center;border-bottom:3px solid #145C43;padding-bottom:8px;margin-bottom:10px">' +
    '<div style="font-size:12pt;font-weight:bold">' + esc_(t.namaSekolah) + '</div>' +
    '<div style="font-size:15pt;font-weight:bold;margin-top:2px">LAPORAN AKTIVITI PANITIA ' + esc_(p.nama.toUpperCase()) + '</div></div>' +
    '<table style="width:100%;border-collapse:collapse;margin-bottom:10px">' +
    baris('Nama Program', d.program) + baris('Tarikh', d.tarikh) + baris('Sasaran', d.sasaran) +
    baris('Objektif', d.objektif) + baris('Impak / Rumusan', d.impak) + '</table>' +
    '<div style="font-weight:bold;margin:6px 0">Gambar Aktiviti</div>' +
    '<table style="width:100%;border-collapse:collapse">' + grid + '</table>' +
    '<table style="width:100%;margin-top:14px"><tr><td style="width:50%">Disediakan oleh:<br><br><b>' + esc_(u.nama) + '</b><br>' + esc_(u.jawatan || 'Guru') +
    '</td><td style="text-align:right;vertical-align:bottom;color:#6a7771;font-size:9pt">Dijana melalui e-Panitia<br>' +
    Utilities.formatDate(new Date(), 'Asia/Kuala_Lumpur', 'dd/MM/yyyy HH:mm') + '</td></tr></table></body></html>';

  const namaFail = ('Laporan ' + d.program + (d.tarikh ? ' (' + d.tarikh + ')' : '')).replace(/[\\/:*?"<>|]/g, '-') + '.pdf';
  const pdf = Utilities.newBlob(html, 'text/html', 'laporan.html').getAs('application/pdf').setName(namaFail);
  const file = folderPanitia_(p, '10').createFile(pdf);
  const rec = {
    id: uid_(), panitiaId: p.id, folderNo: '10', tajuk: 'Laporan: ' + d.program,
    penerangan: 'Laporan satu muka surat dijana automatik' + (d.tarikh ? ' — ' + d.tarikh : ''), jenis: 'DRIVE',
    url: file.getUrl(), driveId: file.getId(), mime: 'application/pdf', saiz: pdf.getBytes().length,
    tugasanId: tugasanSah_(p, q.tugasanId), pemilikId: u.id, pemilikNama: u.nama, dicipta: now_(), status: 'AKTIF'
  };
  insert_('Fail', rec);
  log_(u, p.id, '10', 'JANA LAPORAN', d.program);
  return { fail: failPub_(rec), nama: namaFail, data: Utilities.base64Encode(pdf.getBytes()) };
}

function eksportCartaPdf(q, u) {
  const p = panitia_(q.panitiaId);
  const a = akses_(u, p);
  mesti_(a.lihat, 'Anda bukan ahli panitia ini.');
  const ahli = ahliPanitia_(p);
  const carta = cartaBina_(p, ahli);
  const pt = pentadbir_();
  const t = tetapan_();
  const foto = id => {
    if (!id) return '';
    try { const b = DriveApp.getFileById(id).getBlob(); return 'data:' + b.getContentType() + ';base64,' + Utilities.base64Encode(b.getBytes()); }
    catch (e) { return ''; }
  };
  const kotak = (n, jawatan) => {
    const src = foto(n.fotoId);
    return '<td style="width:25%;padding:6px;text-align:center;vertical-align:top"><div style="border:1px solid #9aa7a0;padding:8px 4px;background:#f4f7f4">' +
      (src ? '<img src="' + src + '" style="width:78px;height:78px"><br>' : '<div style="width:78px;height:78px;margin:0 auto;background:#dcebe3"></div>') +
      '<div style="font-weight:bold;font-size:9.5pt;margin-top:4px">' + esc_(n.nama) + '</div>' +
      '<div style="font-size:8.5pt;color:#44545f">' + esc_(jawatan) + '</div></div></td>';
  };
  const tahap = (label, list) => {
    if (!list.length) return '';
    let h = '<div style="text-align:center;font-weight:bold;color:#145C43;margin:10px 0 4px">' + label + '</div><table style="margin:0 auto;border-collapse:collapse">';
    for (let i = 0; i < list.length; i += 4) h += '<tr>' + list.slice(i, i + 4).map(x => kotak(x.u, x.j)).join('') + '</tr>';
    return h + '</table><div style="width:2px;height:12px;background:#145C43;margin:0 auto"></div>';
  };
  const byId = id => ahli.find(x => x.id === id);
  const L3 = carta.filter(n => n.tahap === 3).map(n => ({ u: byId(n.userId), j: n.jawatan })).filter(x => x.u);
  const L4 = carta.filter(n => n.tahap === 4).map(n => ({ u: byId(n.userId), j: n.jawatan })).filter(x => x.u);
  const L5 = carta.filter(n => n.tahap === 5).map(n => ({ u: byId(n.userId), j: n.jawatan })).filter(x => x.u);
  const L1 = pt.gb ? [{ u: pt.gb, j: 'Guru Besar (Penasihat)' }] : [];
  const L2 = [pt.pk ? { u: pt.pk, j: 'PK Pentadbiran (Pengerusi)' } : null, pt.suk ? { u: pt.suk, j: 'SU Kurikulum (Penyelaras)' } : null].filter(Boolean);

  const html = '<html><head><meta charset="utf-8"><style>@page{size:A4;margin:12mm}body{font-family:Arial,sans-serif;color:#1c2b39}</style></head><body>' +
    '<div style="text-align:center;border-bottom:3px solid #145C43;padding-bottom:6px">' +
    '<div style="font-size:11pt;font-weight:bold">' + esc_(t.namaSekolah) + '</div>' +
    '<div style="font-size:15pt;font-weight:bold">CARTA ORGANISASI PANITIA ' + esc_(p.nama.toUpperCase()) + '</div>' +
    '<div style="font-size:10pt">Tahun ' + new Date().getFullYear() + '</div></div>' +
    tahap('Penasihat', L1) + tahap('Pengerusi & Penyelaras', L2) + tahap('Ketua Panitia', L3) + tahap('Setiausaha Panitia', L4) + tahap('Ahli Panitia', L5).replace(/<div style="width:2px[^>]*><\/div>$/, '') +
    '</body></html>';
  const nama = 'Carta Organisasi Panitia ' + p.nama + ' ' + Utilities.formatDate(new Date(), 'Asia/Kuala_Lumpur', 'yyyy-MM-dd') + '.pdf';
  const pdf = Utilities.newBlob(html, 'text/html', 'carta.html').getAs('application/pdf').setName(nama);
  let simpan = false;
  if (a.urus) {
    const file = folderPanitia_(p, '1').createFile(pdf);
    insert_('Fail', {
      id: uid_(), panitiaId: p.id, folderNo: '1', tajuk: 'Carta Organisasi (salinan rasmi)', penerangan: 'Dijana automatik ' +
        Utilities.formatDate(new Date(), 'Asia/Kuala_Lumpur', 'dd/MM/yyyy'), jenis: 'DRIVE', url: file.getUrl(), driveId: file.getId(),
      mime: 'application/pdf', saiz: pdf.getBytes().length, tugasanId: '', pemilikId: u.id, pemilikNama: u.nama, dicipta: now_(), status: 'AKTIF'
    });
    log_(u, p.id, '1', 'EKSPORT CARTA', 'Salinan rasmi PDF disimpan dalam Folder 01');
    simpan = true;
  }
  return { nama, data: Utilities.base64Encode(pdf.getBytes()), disimpan: simpan };
}

/* =====================================================================
 *  MOD PEMANTAU (PPD / JNJK / Nazir)
 * ===================================================================== */
function janaPemantau(q, u) {
  const p = panitia_(q.panitiaId);
  mesti_(akses_(u, p).urus, 'Hanya Ketua Panitia boleh menjana pautan pemantau.');
  const hari = Math.min(Math.max(Number(q.hari) || 3, 1), 30);
  const kod = uid_(20);
  const kata = String(100000 + parseInt(Utilities.getUuid().replace(/-/g, '').slice(0, 12), 16) % 900000);
  const salt = uid_(16);
  const tamat = new Date(Date.now() + hari * 864e5).toISOString();
  const label = bersih_(q.label, 80) || 'Pemantau';
  insert_('Pemantau', { id: uid_(), panitiaId: p.id, kod, hash: hash_(kata, salt), salt, label, tamat, diciptaOleh: u.nama, dicipta: now_(), status: 'AKTIF' });
  log_(u, p.id, '', 'PAUTAN PEMANTAU', label + ' (' + hari + ' hari)');
  return { kod, kataLaluan: kata, tamat, label };
}

function batalPemantau(q, u) {
  const x = rows_('Pemantau').find(r => r.id === q.id);
  mesti_(x, 'Pautan tidak dijumpai.');
  mesti_(akses_(u, panitia_(x.panitiaId)).urus);
  update_('Pemantau', x.id, { status: 'BATAL' });
  log_(u, x.panitiaId, '', 'BATAL PEMANTAU', x.label);
  return true;
}

function pemantauSah_(c) {
  c = c || {};
  const x = rows_('Pemantau').find(r => r.kod === String(c.kod || ''));
  mesti_(x && x.status === 'AKTIF' && new Date(x.tamat).getTime() > Date.now(), 'Pautan pemantau tidak sah atau telah tamat tempoh.');
  semakCuba_('P_' + x.id);
  if (x.hash !== hash_(String(c.kataLaluan || ''), x.salt)) {
    gagalCuba_('P_' + x.id);
    throw new Error('Kata laluan pemantau tidak betul.');
  }
  return x;
}

function pemantauData(q) {
  const x = pemantauSah_(q);
  const p = panitia_(x.panitiaId);
  const ahli = ahliPanitia_(p);
  const keah = rows_('Keahlian').filter(k => k.panitiaId === p.id);
  const kp = userById_(p.ketuaId);
  if (q.catat) log_('Pemantau: ' + x.label, p.id, '', 'AKSES PEMANTAU', 'Mod baca sahaja dibuka');
  return {
    panitia: { id: p.id, nama: p.nama, kod: p.kod, ketuaId: p.ketuaId, setiausahaId: p.setiausahaId, ketua: kp ? kp.nama : '' },
    akses: { r: 'PEMANTAU', lihat: true, urus: false, su: false, kew: false, admin: false },
    sekolah: tetapan_().namaSekolah,
    label: x.label, tamat: x.tamat,
    ahli: ahli.map(m => Object.assign(ringkas_(m), { gred: m.gred, tahun: (keah.find(k => k.userId === m.id) || {}).tahun || '' })),
    pentadbir: pentadbir_(),
    carta: cartaBina_(p, ahli),
    fail: rows_('Fail').filter(f => f.panitiaId === p.id && f.status === 'AKTIF' && f.folderNo !== '9').map(failPub_)
      .sort((a, b) => (b.dicipta > a.dicipta ? 1 : -1)),
    tugasan: [], tong: [], menunggu: [], pemantau: [], log: [], semakan: null
  };
}

/* =====================================================================
 *  SUPER ADMIN
 * ===================================================================== */
function senaraiPengguna(q, u) {
  admin_(u);
  const g = gens_();
  const etag = 'A|' + APP_VERSION + '|' + ['Pengguna', 'Keahlian', 'Panitia', 'Tetapan'].map(n => g[n]).join('.');
  if (q && q.etagAdm && q.etagAdm === etag) return { sama: true, etag };
  const kotor = ['Pengguna', 'Keahlian', 'Panitia', 'Tetapan'].some(n => _KOTOR[n]);
  return {
    etag: kotor ? '' : etag, // generasi akan bertukar di selesai_(); etag kosong = muat penuh kali seterusnya
    pengguna: rows_('Pengguna').map(x => pub_(x, true)).sort((a, b) => a.nama.localeCompare(b.nama)),
    keahlian: rows_('Keahlian').map(k => ({ panitiaId: k.panitiaId, userId: k.userId, status: k.status })),
    panitia: senaraiAktif_().map(p => ({ id: p.id, nama: p.nama, kod: p.kod, ketuaId: p.ketuaId, setiausahaId: p.setiausahaId })),
    tetapan: tetapan_()
  };
}

function simpanPanitia(q, u) {
  admin_(u);
  const nama = bersih_(q.nama, 80);
  const kod = bersih_(q.kod, 10).toUpperCase();
  mesti_(nama, 'Isi nama panitia.');
  mesti_(!senaraiAktif_().find(p => p.nama.toLowerCase() === nama.toLowerCase() && p.id !== q.id), 'Panitia dengan nama ini sudah wujud.');
  if (q.id) {
    const p = panitia_(q.id);
    update_('Panitia', p.id, { nama, kod });
    if (p.folderId) { try { DriveApp.getFolderById(p.folderId).setName('PANITIA ' + nama.toUpperCase()); } catch (e) { /* */ } }
    log_(u, p.id, '', 'KEMAS KINI PANITIA', nama);
  } else {
    const id = uid_();
    insert_('Panitia', { id, nama, kod, status: 'AKTIF', dicipta: now_() });
    log_(u, id, '', 'PANITIA BAHARU', nama);
  }
  return balasAdmin_(q, u);
}

function padamPanitia(q, u) {
  admin_(u);
  const p = panitia_(q.id);
  update_('Panitia', p.id, { status: 'PADAM' });
  log_(u, p.id, '', 'PADAM PANITIA', p.nama + ' (fail Drive dikekalkan)');
  return balasAdmin_(q, u);
}

function lantikKP(q, u) {
  admin_(u);
  const p = panitia_(q.panitiaId);
  if (!q.userId) {
    update_('Panitia', p.id, { ketuaId: '' });
    log_(u, p.id, '', 'LANTIK KP', 'Ketua Panitia dikosongkan');
    return balasAdmin_(q, u);
  }
  const g = userById_(q.userId);
  mesti_(g && g.status !== 'NYAHAKTIF', 'Guru tidak dijumpai.');
  const patch = { ketuaId: g.id };
  if (p.setiausahaId === g.id) patch.setiausahaId = '';
  update_('Panitia', p.id, patch);
  const k = rows_('Keahlian').find(x => x.panitiaId === p.id && x.userId === g.id);
  if (k) update_('Keahlian', k.id, { status: 'SAH' });
  else insert_('Keahlian', { id: uid_(), panitiaId: p.id, userId: g.id, tahun: '', status: 'SAH', dicipta: now_() });
  log_(u, p.id, '1', 'LANTIK KP', g.nama + ' dilantik sebagai Ketua Panitia');
  return balasAdmin_(q, u);
}

function kemaskiniPengguna(q, u) {
  admin_(u);
  const g = userById_(q.userId);
  mesti_(g, 'Pengguna tidak dijumpai.');
  const patch = {};
  if (q.peranan) patch.peranan = q.peranan === 'SUPER_ADMIN' ? 'SUPER_ADMIN' : 'GURU';
  if (q.status) patch.status = q.status === 'NYAHAKTIF' ? 'NYAHAKTIF' : 'AKTIF';
  const bakiAdmin = rows_('Pengguna').filter(x => x.peranan === 'SUPER_ADMIN' && x.status !== 'NYAHAKTIF' && x.id !== g.id).length;
  if ((patch.peranan === 'GURU' || patch.status === 'NYAHAKTIF') && g.peranan === 'SUPER_ADMIN') mesti_(bakiAdmin > 0, 'Sistem mesti mempunyai sekurang-kurangnya seorang Super Admin aktif.');
  update_('Pengguna', g.id, patch);
  if (patch.status === 'NYAHAKTIF') deleteWhere_('Sesi', s => s.userId === g.id);
  log_(u, '', '', 'KEMAS KINI PENGGUNA', g.nama + ' ' + JSON.stringify(patch));
  return balasAdmin_(q, u);
}

function resetKataLaluan(q, u) {
  admin_(u);
  const g = userById_(q.userId);
  mesti_(g, 'Pengguna tidak dijumpai.');
  const baru = uid_(8);
  const salt = uid_(16);
  update_('Pengguna', g.id, { hash: hash_(baru, salt), salt, wajibTukar: 'YA' });
  deleteWhere_('Sesi', s => s.userId === g.id);
  log_(u, '', '', 'RESET KATA LALUAN', g.nama);
  return { kataLaluan: baru };
}

/** Super Admin memuat naik foto GB / PK Pentadbiran / SU Kurikulum untuk carta organisasi */
function muatNaikFotoPentadbir(q, u) {
  admin_(u);
  const p = String(q.peranan || '');
  mesti_(['gb', 'pk', 'suk'].indexOf(p) >= 0, 'Jawatan tidak sah.');
  mesti_(saizB64_(q.data) < 3 * 1048576, 'Gambar terlalu besar.');
  const t = tetapan_();
  const pengguna = userById_(t[p + 'Id']);
  const f = simpanBlob_(folderSistem_('_FOTO_PROFIL'), q.data, 'image/jpeg', 'foto_' + p + '_' + Date.now() + '.jpg');
  const durl = 'data:image/jpeg;base64,' + String(q.data).replace(/^data:[^,]*,/, '');
  if (durl.length < 95000) try { cache_().put('IMG_' + f.getId(), durl, CACHE_TTL); } catch (e) { /* */ }
  if (pengguna) {
    // Dipautkan kepada pengguna berdaftar: foto profil guru itu dikemas kini
    const lama = pengguna.fotoId;
    update_('Pengguna', pengguna.id, { fotoId: f.getId(), dikemaskini: now_() });
    buangDrive_(lama);
  } else {
    const lama = t[p + 'Foto'];
    setTetapan_(p + 'Foto', f.getId());
    buangDrive_(lama);
  }
  log_(u, '', '', 'FOTO PENTADBIR', { gb: 'Guru Besar', pk: 'PK Pentadbiran', suk: 'SU Kurikulum' }[p]);
  const r = senaraiPengguna(q, u);
  r.fotoBaharu = f.getId();
  return r;
}

function simpanTetapan(q, u) {
  admin_(u);
  const d = q.tetapan || {};
  ['namaSekolah', 'kodSekolah', 'gbId', 'gbNama', 'pkId', 'pkNama', 'sukId', 'sukNama'].forEach(k => {
    if (d[k] !== undefined) setTetapan_(k, bersih_(d[k], 150));
  });
  if (d.jadualCsvUrl !== undefined) {
    const url = bersih_(d.jadualCsvUrl, 600);
    mesti_(!url || /^https:\/\/\S+$/i.test(url), 'Pautan CSV jadual mesti bermula dengan https://');
    setTetapan_('jadualCsvUrl', url);
  }
  log_(u, '', '', 'TETAPAN', 'Tetapan sekolah dikemas kini');
  return balasAdmin_(q, u);
}
