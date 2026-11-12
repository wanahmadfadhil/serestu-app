/**
 * =====================================================================
 *  e-KURIKULUM — Sistem Pengurusan Kurikulum & Panitia Sekolah Rendah
 *  Backend Google Apps Script (Google Sheets + Google Drive)
 *  Versi 2.0.0
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
 *     Pendaftaran seterusnya perlu DILULUSKAN oleh Super Admin
 *     (Profil > Pentadbiran sistem > Pengguna).
 *  7. Guru yang diimport dari tab "Data Guru": kata laluan awal ialah
 *     6 DIGIT AKHIR No. KP dan wajib ditukar semasa log masuk pertama.
 *
 *  e-KURIKULUM FASA 1 (versi 2.0.0)
 *   - Menu: Utama · Panitia · Takwim · Kurikulum · Profil
 *   - Pentadbir kurikulum = Super Admin + GB, PK Pentadbiran, SU Kurikulum
 *     (Tetapan) + pengguna dalam "Akses pentadbir kurikulum tambahan".
 *     Mereka boleh melihat semua panitia, mengurus JK Kurikulum, tindakan
 *     susulan, dokumen wajib sekolah dan program peringkat sekolah.
 *   - Jadual baharu (dijana automatik): Program, Mesyuarat, Tindakan, DokSekolah.
 *
 *  TRIGGER AUTOMATIK (dipasang oleh setupSistem)
 *   - bersihkanTong       : setiap hari 2 pagi
 *   - peringatanTugasan   : setiap hari 7 pagi (emel dokumen wajib hampir tarikh akhir)
 *   - sandaranMingguan    : setiap Ahad 4 pagi (salinan Sheet ke folder _SANDARAN)
 *   - arkibLog            : 1 haribulan setiap bulan (log > 1 tahun ke tab "Arkib Log")
 *
 *  Nota: setiap kali kod dikemas kini, buat "Manage deployments > Edit >
 *  Version: New version" supaya URL Web App kekal sama.
 * =====================================================================
 */

const APP_VERSION = '2.0.0';
const KATA_LALUAN_LALAI = '12345678';
const SESI_HARI = 7;
const TONG_HARI = 30;
const MAKS_FAIL_MB = 10;
const HASH_ULANG = 1000;        // bilangan ulangan hash kata laluan (lebih tinggi = lebih selamat tetapi log masuk lebih lambat)
const ARKIB_LOG_HARI = 365;     // log lebih lama daripada ini dipindahkan ke tab "Arkib Log"
const SIMPAN_SANDARAN = 8;      // bilangan sandaran mingguan yang disimpan
const HARI_PERINGATAN = [7, 3, 1, 0]; // emel peringatan dihantar X hari sebelum tarikh akhir dokumen wajib
const ZON = 'Asia/Kuala_Lumpur';

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
  Log:      ['id', 'masa', 'userId', 'nama', 'panitiaId', 'folderNo', 'tindakan', 'butiran'],
  Laporan:  ['id', 'panitiaId', 'tarikh', 'kategori', 'tajuk', 'tarikhMula', 'tarikhTamat', 'tarikhProgram', 'tempat',
             'objektif', 'tentatif', 'catatan', 'fotoIds', 'lakaranId', 'disediakanOleh', 'disediakanId', 'failId', 'driveId',
             'dicipta', 'dikemaskini', 'status'],
  Program:  ['id', 'panitiaId', 'tajuk', 'kategori', 'tarikhMula', 'tarikhTamat', 'masa', 'tempat', 'catatan', 'laporanId',
             'diciptaId', 'diciptaOleh', 'dicipta', 'dikemaskini', 'status'],
  Mesyuarat: ['id', 'tahun', 'bil', 'tajuk', 'tarikh', 'masa', 'tempat', 'status', 'minitUrl', 'minitDriveId', 'minitNama', 'minitMime',
              'catatan', 'diciptaOleh', 'dicipta', 'dikemaskini'],
  Tindakan: ['id', 'mesyuaratId', 'perkara', 'pemilikId', 'pemilikNama', 'tarikhAkhir', 'status', 'catatan', 'selesaiPada', 'selesaiOleh', 'dicipta'],
  DokSekolah: ['id', 'tajuk', 'folderNo', 'tarikhAkhir', 'catatan', 'diciptaOleh', 'dicipta', 'status']
};

const KATEGORI_LAPORAN = ['LAPORAN', 'PROGRAM', 'PERTANDINGAN', 'KURSUS / LDP', 'MESYUARAT', 'BENGKEL', 'LAWATAN', 'LAIN-LAIN'];
const BULAN_MS = ['JANUARI', 'FEBRUARI', 'MAC', 'APRIL', 'MEI', 'JUN', 'JULAI', 'OGOS', 'SEPTEMBER', 'OKTOBER', 'NOVEMBER', 'DISEMBER'];

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
  ['aksesKurikulum', '', 'ID pengguna tambahan (dipisah koma) yang mendapat akses pentadbir kurikulum, cth. Ketua Bidang'],
  ['domainEmel', '', 'Hadkan pendaftaran kepada domain emel ini, cth. moe-dl.edu.my (kosong = semua emel)'],
  ['emelPeringatan', 'YA', 'YA = hantar emel peringatan dokumen wajib kepada guru'],
  ['benamkan', 'TIDAK', 'YA hanya jika aplikasi dibenamkan dalam Google Sites / iframe laman lain'],
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
  const pasang = (fn, bina) => {
    ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === fn) ScriptApp.deleteTrigger(t); });
    bina(() => ScriptApp.newTrigger(fn));
  };
  pasang('bersihkanTong', b => b().timeBased().everyDays(1).atHour(2).inTimezone(ZON).create());
  pasang('peringatanTugasan', b => b().timeBased().everyDays(1).atHour(7).inTimezone(ZON).create());
  pasang('sandaranMingguan', b => b().timeBased().onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(4).inTimezone(ZON).create());
  pasang('arkibLog', b => b().timeBased().onMonthDay(1).atHour(3).inTimezone(ZON).create());
  pasang('cacheBerubah', b => {
    b().forSpreadsheet(getSS_()).onChange().create();
    b().forSpreadsheet(getSS_()).onEdit().create();
  });
  const nLalai = lalaiKeNoKP_();
  tutupKongsiJadual_();
  selesai_();
  if (nLalai) Logger.log('🔑 ' + nLalai + ' akaun yang belum log masuk kini guna kata laluan awal 6 digit akhir No. KP.');
  try { MailApp.getRemainingDailyQuota(); } catch (e) { Logger.log('⚠️ Kebenaran emel belum diberi: ' + e.message); }
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
  const guru = rows_('Pengguna').filter(aktif_);
  const tiada = guru.filter(u => !padanJadual_(u, idx)).map(u => u.nama);
  Logger.log('✅ ' + (guru.length - tiada.length) + '/' + guru.length + ' pengguna sistem dipadankan dengan jadual.');
  if (tiada.length) Logger.log('⚠️ Belum dipadankan (guru boleh pilih nama sendiri di Profil):\n' + tiada.join('\n'));
  Logger.log('Seterusnya: Manage deployments > Edit > Version: New version > Deploy.');
}

function onOpen() {
  try {
    SpreadsheetApp.getUi().createMenu('⚙️ e-Kurikulum')
      .addItem('Jana / baiki jadual', 'setupSistem')
      .addItem('Import guru dari tab Data Guru', 'importDataGuru')
      .addItem('Kosongkan cache', 'kosongkanCache')
      .addItem('Uji bacaan jadual waktu', 'ujiJadual')
      .addItem('Bersihkan tong sampah sekarang', 'bersihkanTong')
      .addSeparator()
      .addItem('Hantar emel peringatan dokumen wajib sekarang', 'peringatanTugasan')
      .addItem('Buat sandaran sekarang', 'sandaranMingguan')
      .addItem('Arkibkan log lama sekarang', 'arkibLog')
      .addToUi();
  } catch (e) { /* bukan konteks UI */ }
}

/**
 * IMPORT GURU PUKAL daripada tab "Data Guru" (dalam spreadsheet yang sama).
 * Lajur dikesan melalui tajuk baris pertama: Emel/Email, Nama, No KP/NoKP/IC.
 * Kata laluan awal setiap guru ialah 6 DIGIT AKHIR No. KP dan WAJIB
 * ditukar semasa log masuk kali pertama.
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
  let adaAdmin = ada.some(u => u.peranan === 'SUPER_ADMIN' && aktif_(u));

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
    let peranan = 'GURU';
    if (!adaAdmin && emel === pemilik) { peranan = 'SUPER_ADMIN'; adaAdmin = true; }
    baru.push({
      id: uid_(), nama, noKP: kp, emel, telefon: '', peranan, hash: 'LALAI', salt: '', // LALAI = 6 digit akhir No. KP
      status: 'AKTIF', dicipta: now_(), dikemaskini: now_(), wajibTukar: 'YA'
    });
    emelAda[emel] = 1; kpAda[kp] = 1;
  }
  insertMany_('Pengguna', baru);
  if (baru.length) log_('Sistem', '', '', 'IMPORT GURU', baru.length + ' guru didaftarkan dari tab ' + sh.getName());

  Logger.log('✅ ' + baru.length + ' guru berjaya didaftarkan (kata laluan awal: 6 digit akhir No. KP).');
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
  if (p.api === 'ping') return json_({ ok: true, app: 'e-Kurikulum', versi: APP_VERSION });
  let benam = false;
  try { ensureSchema_(); benam = String(tetapan_().benamkan || '').toUpperCase() === 'YA'; } catch (e) { /* guna lalai */ }
  try { selesai_(); } catch (e) { /* */ }
  try {
    const out = HtmlService.createHtmlOutputFromFile('index')
      .setTitle('e-Kurikulum')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
    // Benarkan dibenamkan dalam iframe (Google Sites) hanya jika ditetapkan — elak clickjacking
    if (benam) out.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    return out;
  } catch (err) {
    return json_({ ok: true, app: 'e-Kurikulum', versi: APP_VERSION, nota: 'Fail HTML "index" tiada. Guna mod hos luar.' });
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

const AWAM = ['awam', 'daftar', 'login', 'lupaKataLaluan', 'setSemulaKataLaluan', 'pemantauData', 'imej', 'muatTurun', 'laporanDetail'];

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
  awam, daftar, login, lupaKataLaluan, setSemulaKataLaluan, logout, init, kemaskiniProfil, tukarKataLaluan, muatNaikFoto, muatNaikJadual, lihatJadual,
  sync, batch, jadualGuru, setNamaJadual, muatNaikFotoPentadbir,
  imej, muatTurun, panitiaDetail, simpanCarta, lantikSU, sahkanAhli, calonAhli, tambahAhli, keluarkanAhli, muatNaikFail, tambahPautan, padamFail,
  pulihFail, padamKekal, tambahTugasan, padamTugasan, janaLaporan, eksportCartaPdf, janaPemantau, batalPemantau,
  pemantauData, senaraiPengguna, simpanPanitia, padamPanitia, lantikKP, kemaskiniPengguna, resetKataLaluan,
  simpanTetapan, luluskanPengguna, simpanLaporan, laporanDetail, padamLaporan,
  simpanProgram, padamProgram, simpanMesyuarat, padamMesyuarat, muatNaikMinit, lihatMinit,
  simpanTindakan, statusTindakan, padamTindakan, simpanDokSekolah, padamDokSekolah
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
  _PK = null;
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
/** Hash v2: SHA-256 berulang HASH_ULANG kali. Format: v2$<ulangan>$<hex> */
function hashV2_(pw, salt, n) {
  n = Number(n) || HASH_ULANG;
  const sb = Utilities.newBlob(String(salt)).getBytes();
  let h = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + '::' + pw, Utilities.Charset.UTF_8);
  for (let i = 1; i < n; i++) h = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, h.concat(sb));
  return 'v2$' + n + '$' + h.map(b => ('0' + (b & 255).toString(16)).slice(-2)).join('');
}
function kataBaru_(pw) { const salt = uid_(16); return { hash: hashV2_(String(pw), salt), salt }; }
function kataLalai_(u) { return String((u && u.noKP) || '').replace(/\D/g, '').slice(-6); }
/** Semak kata laluan: LALAI (6 digit akhir No. KP), v2 (berulang) atau format lama (SHA-256 sekali) */
function kataSah_(u, pw) {
  pw = String(pw || '');
  if (!u || !u.hash || !pw) return false;
  if (u.hash === 'LALAI') return kataLalai_(u).length === 6 && pw === kataLalai_(u);
  if (u.hash.indexOf('v2$') === 0) return hashV2_(pw, u.salt, u.hash.split('$')[1]) === u.hash;
  return hash_(pw, u.salt) === u.hash;
}
/** Pengguna boleh log masuk / dikira ahli: bukan NYAHAKTIF dan bukan MENUNGGU kelulusan */
function aktif_(u) { return !!u && u.status !== 'NYAHAKTIF' && u.status !== 'MENUNGGU'; }
function namaSekolah_() { return tetapan_().namaSekolah || 'Sekolah'; }

/** Hantar emel (MailApp). Gagal secara senyap jika kuota habis atau kebenaran belum diberi. */
function hantarEmel_(kepada, subjek, isiHtml) {
  try {
    if (!kepada || MailApp.getRemainingDailyQuota() < 1) return false;
    const url = urlApp_();
    MailApp.sendEmail({
      to: kepada, subject: subjek, name: 'e-Kurikulum ' + namaSekolah_(),
      htmlBody: '<div style="font-family:Arial,sans-serif;font-size:14px;color:#1c2b39;max-width:560px">' +
        '<div style="background:#145C43;color:#fff;padding:12px 16px;border-radius:10px 10px 0 0;font-weight:bold">e-Kurikulum · ' + esc_(namaSekolah_()) + '</div>' +
        '<div style="border:1px solid #d3dbd5;border-top:0;padding:16px;border-radius:0 0 10px 10px">' + isiHtml +
        (url ? '<p style="margin-top:18px"><a href="' + esc_(url) + '" style="background:#145C43;color:#fff;padding:10px 16px;border-radius:8px;text-decoration:none">Buka e-Kurikulum</a></p>' : '') +
        '<p style="color:#6a7771;font-size:12px;margin-top:18px">Emel automatik daripada sistem e-Kurikulum. Jangan balas emel ini.</p></div></div>'
    });
    return true;
  } catch (e) { Logger.log('Emel gagal ke ' + kepada + ': ' + e.message); return false; }
}

function mesti_(c, msg) { if (!c) throw new Error(msg || 'Anda tiada kebenaran untuk tindakan ini.'); }
function bersih_(s, max) { return String(s === null || s === undefined ? '' : s).trim().slice(0, max || 500); }
function esc_(s) {
  return String(s || '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}
function emelSah_(e) { return /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(e); }
function tetapan_() { const o = {}; rows_('Tetapan').forEach(r => { o[r.kunci] = r.nilai; }); return o; }
function setTetapan_(k, v) {
  _PK = null;
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
  mesti_(aktif_(u), 'SESI_TAMAT');
  return u;
}
function admin_(u) { mesti_(u && u.peranan === 'SUPER_ADMIN', 'Hanya Super Admin boleh melakukan tindakan ini.'); }

/** r: KP | SU | AHLI | MENUNGGU | null */
/** Pentadbir kurikulum: Super Admin, GB, PK Pentadbiran, SU Kurikulum dan pengguna dalam Tetapan.aksesKurikulum */
let _PK = null;
function pentadbirKur_(u) {
  if (!u) return false;
  if (u.peranan === 'SUPER_ADMIN') return true;
  if (!_PK) {
    const t = tetapan_();
    _PK = {};
    [t.gbId, t.pkId, t.sukId].concat(String(t.aksesKurikulum || '').split(/[\s,]+/)).forEach(id => { if (id) _PK[id] = 1; });
  }
  return !!_PK[u.id];
}
function kurAdmin_(u) { mesti_(pentadbirKur_(u), 'Hanya pentadbir kurikulum (GB, PK, SU Kurikulum) boleh melakukan tindakan ini.'); }

function akses_(u, p) {
  const admin = u.peranan === 'SUPER_ADMIN';
  const pk = !admin && pentadbirKur_(u);
  let r = null;
  if (p.ketuaId && p.ketuaId === u.id) r = 'KP';
  else if (p.setiausahaId && p.setiausahaId === u.id) r = 'SU';
  else {
    const k = rows_('Keahlian').find(x => x.panitiaId === p.id && x.userId === u.id);
    if (k) r = k.status === 'SAH' ? 'AHLI' : 'MENUNGGU';
  }
  const lihat = admin || pk || r === 'KP' || r === 'SU' || r === 'AHLI';
  return { r, admin, pentadbir: admin || pk, lihat, urus: admin || r === 'KP', su: admin || r === 'KP' || r === 'SU', kew: admin || pk || r === 'KP' || r === 'SU' };
}

function ahliPanitia_(p) {
  const ids = [];
  rows_('Keahlian').forEach(k => { if (k.panitiaId === p.id && k.status === 'SAH' && ids.indexOf(k.userId) < 0) ids.push(k.userId); });
  [p.ketuaId, p.setiausahaId].forEach(id => { if (id && ids.indexOf(id) < 0) ids.push(id); });
  return ids.map(userById_).filter(aktif_);
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
  return { namaSekolah: t.namaSekolah, versi: APP_VERSION, domainEmel: t.domainEmel || '', panitia: senaraiAktif_().map(p => ({ id: p.id, nama: p.nama, kod: p.kod })) };
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
  const pertama = !semua.length;
  const dom = String(tetapan_().domainEmel || '').toLowerCase().split(/[\s,]+/).filter(Boolean);
  if (dom.length && !pertama) mesti_(dom.some(d => emel.endsWith('@' + d)), 'Pendaftaran hanya untuk emel @' + dom.join(' atau @') + '.');
  const ada = semua.find(u => u.emel === emel);
  if (ada && ada.status === 'MENUNGGU') throw new Error('Pendaftaran dengan emel ini sedang menunggu kelulusan Super Admin.');
  mesti_(!ada, 'Emel ini sudah didaftarkan. Sila log masuk.');
  mesti_(!semua.find(u => u.noKP === noKP), 'No. Kad Pengenalan ini sudah didaftarkan.');
  const u = Object.assign({
    id: uid_(), nama, noKP, emel, telefon: tel, peranan: pertama ? 'SUPER_ADMIN' : 'GURU',
    status: pertama ? 'AKTIF' : 'MENUNGGU', dicipta: now_(), dikemaskini: now_()
  }, kataBaru_(q.kataLaluan));
  insert_('Pengguna', u);
  const disimpan = userById_(u.id);
  segerakSubjek_(disimpan, q.subjek);
  log_(disimpan, '', '', 'DAFTAR', pertama ? 'Super Admin pertama' : 'Pendaftaran baharu menunggu kelulusan');
  if (pertama) return bukaSesi_(disimpan);
  maklumAdminDaftar_(disimpan);
  return { menunggu: true, mesej: 'Pendaftaran diterima. Akaun anda akan diaktifkan selepas diluluskan oleh Super Admin sekolah. Anda akan menerima emel apabila diluluskan.' };
}

function maklumAdminDaftar_(g) {
  const admin = rows_('Pengguna').filter(x => x.peranan === 'SUPER_ADMIN' && aktif_(x) && x.emel).map(x => x.emel);
  if (!admin.length) return;
  hantarEmel_(admin.join(','), 'Pendaftaran baharu menunggu kelulusan: ' + g.nama,
    '<p>Seorang pengguna baharu telah mendaftar dan menunggu kelulusan anda:</p>' +
    '<table style="border-collapse:collapse;margin:8px 0"><tr><td style="padding:4px 12px 4px 0;color:#6a7771">Nama</td><td><b>' + esc_(g.nama) + '</b></td></tr>' +
    '<tr><td style="padding:4px 12px 4px 0;color:#6a7771">Emel</td><td>' + esc_(g.emel) + '</td></tr>' +
    '<tr><td style="padding:4px 12px 4px 0;color:#6a7771">Telefon</td><td>' + esc_(g.telefon) + '</td></tr></table>' +
    '<p>Buka <b>Profil &gt; Pentadbiran sistem &gt; Pengguna</b> untuk meluluskan atau menolak.</p>');
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
  let u = rows_('Pengguna').find(x => x.emel === emel);
  if (!kataSah_(u, q.kataLaluan)) {
    gagalCuba_('L_' + emel);
    throw new Error('Emel atau kata laluan tidak betul.');
  }
  cache_().remove('CUBA_L_' + emel);
  mesti_(u.status !== 'MENUNGGU', 'Akaun anda masih menunggu kelulusan Super Admin sekolah.');
  mesti_(u.status !== 'NYAHAKTIF', 'Akaun anda telah dinyahaktifkan. Hubungi pentadbir sistem.');
  if (u.hash !== 'LALAI' && String(u.hash).indexOf('v2$') !== 0) {
    u = update_('Pengguna', u.id, kataBaru_(q.kataLaluan)); // naik taraf hash lama secara senyap
  }
  return bukaSesi_(u);
}

/* ---------- Lupa kata laluan: kod 6 digit dihantar ke emel (sah 15 minit) ---------- */
const RESET_MIN = 15, RESET_MAKS_KOD = 3, RESET_MAKS_CUBA = 5;
const MESEJ_RESET = 'Jika emel ini berdaftar dan aktif, kod 6 digit telah dihantar. Semak peti masuk dan folder Spam.';

function lupaKataLaluan(q) {
  const emel = bersih_(q.emel, 120).toLowerCase();
  mesti_(emelSah_(emel), 'Isi emel yang sah.');
  const kKira = 'RESET_N_' + emel;
  const n = Number(cache_().get(kKira)) || 0;
  mesti_(n < RESET_MAKS_KOD, 'Terlalu banyak permintaan kod. Cuba lagi selepas ' + RESET_MIN + ' minit.');
  cache_().put(kKira, String(n + 1), RESET_MIN * 60);
  const u = rows_('Pengguna').find(x => x.emel === emel);
  // Jawapan sama sama ada emel wujud atau tidak — elak orang luar menyemak emel berdaftar
  if (!aktif_(u)) return { mesej: MESEJ_RESET };
  const kod = ('00000' + (parseInt(Utilities.getUuid().replace(/-/g, '').slice(0, 12), 16) % 1000000)).slice(-6);
  const salt = uid_(16);
  cache_().put('RESET_' + emel, JSON.stringify({ h: hash_(kod, salt), salt, cuba: 0 }), RESET_MIN * 60);
  const ok = hantarEmel_(u.emel, 'Kod set semula kata laluan e-Kurikulum: ' + kod,
    '<p>Salam sejahtera ' + esc_(u.nama) + ',</p><p>Kami menerima permintaan untuk menetapkan semula kata laluan akaun e-Kurikulum anda.</p>' +
    '<p style="font-size:30px;font-weight:bold;letter-spacing:8px;background:#eef3ef;padding:12px 16px;border-radius:10px;text-align:center;margin:16px 0">' + kod + '</p>' +
    '<p>Masukkan kod ini dalam aplikasi. Kod sah selama <b>' + RESET_MIN + ' minit</b>.</p>' +
    '<p style="color:#6a7771">Jika anda tidak meminta set semula kata laluan, abaikan emel ini. Kata laluan anda tidak berubah.</p>');
  mesti_(ok, 'Emel tidak dapat dihantar sekarang. Hubungi Super Admin sekolah untuk set semula kata laluan.');
  log_(u, '', '', 'LUPA KATA LALUAN', 'Kod set semula dihantar ke emel');
  return { mesej: MESEJ_RESET };
}

function setSemulaKataLaluan(q) {
  const emel = bersih_(q.emel, 120).toLowerCase();
  const kod = String(q.kod || '').replace(/\D/g, '');
  const k = 'RESET_' + emel;
  let r = null;
  try { r = JSON.parse(cache_().get(k) || 'null'); } catch (e) { r = null; }
  mesti_(r, 'Kod tidak sah atau telah tamat tempoh. Minta kod baharu.');
  if (r.cuba >= RESET_MAKS_CUBA || r.h !== hash_(kod, r.salt)) {
    r.cuba = (r.cuba || 0) + 1;
    if (r.cuba >= RESET_MAKS_CUBA) { cache_().remove(k); throw new Error('Terlalu banyak cubaan. Minta kod baharu.'); }
    cache_().put(k, JSON.stringify(r), RESET_MIN * 60);
    throw new Error('Kod tidak betul. Baki cubaan: ' + (RESET_MAKS_CUBA - r.cuba) + '.');
  }
  const u = rows_('Pengguna').find(x => x.emel === emel);
  mesti_(aktif_(u), 'Akaun tidak aktif. Hubungi Super Admin sekolah.');
  const baru = String(q.baru || '');
  mesti_(baru.length >= 6, 'Kata laluan baharu sekurang-kurangnya 6 aksara.');
  mesti_(baru !== KATA_LALUAN_LALAI && baru !== kataLalai_(u), 'Kata laluan baharu tidak boleh sama dengan kata laluan awal (6 digit akhir No. KP).');
  cache_().remove(k);
  cache_().remove('CUBA_L_' + emel);
  const v = update_('Pengguna', u.id, Object.assign(kataBaru_(baru), { wajibTukar: '', dikemaskini: now_() }));
  deleteWhere_('Sesi', x => x.userId === u.id); // log keluar semua peranti lain
  log_(u, '', '', 'SET SEMULA KATA LALUAN', 'Melalui kod emel');
  return bukaSesi_(v);
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
    jadualAda: (() => { try { return !!padanJadual_(u, jadualIndeks_()); } catch (e) { return false; } })(),
    pentadbirKur: pentadbirKur_(u)
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
  if (u.wajibTukar !== 'YA') mesti_(kataSah_(u, q.lama), 'Kata laluan semasa tidak betul.');
  const baru = String(q.baru || '');
  mesti_(baru.length >= 6, 'Kata laluan baharu sekurang-kurangnya 6 aksara.');
  mesti_(baru !== KATA_LALUAN_LALAI && baru !== kataLalai_(u), 'Kata laluan baharu tidak boleh sama dengan kata laluan awal (6 digit akhir No. KP).');
  mesti_(!kataSah_(u, baru), 'Kata laluan baharu mesti berbeza daripada yang lama.');
  update_('Pengguna', u.id, Object.assign(kataBaru_(baru), { wajibTukar: '', dikemaskini: now_() }));
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
  const f = simpanBlob_(folderSistem_('_JADUAL_WAKTU'), q.data, mime, 'Jadual_' + u.nama + '_' + Date.now());
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
  mesti_(fd.getSize() < 15 * 1048576, 'Salinan jadual terlalu besar untuk dipaparkan.');
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
  rows_('Laporan').forEach(r => {
    if (r.status !== 'AKTIF') return;
    if (r.lakaranId) sah[r.lakaranId] = 1;
    fotoLaporan_(r).forEach(id => { if (id) sah[id] = 1; });
  });
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
  return APP_VERSION + '|' + u.id + '|' + ['Tetapan', 'Pengguna', 'Panitia', 'Keahlian', 'Carta', 'Fail', 'Tugasan', 'Pemantau', 'Log', 'Laporan', 'Program', 'Mesyuarat', 'Tindakan', 'DokSekolah'].map(n => g[n]).join('.');
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
  return { etag: et, init: d, detail, kur: kurikulumData_(u) };
}

const BOLEH_BATCH = ['kemaskiniProfil', 'simpanCarta', 'lantikSU', 'sahkanAhli', 'tambahAhli', 'keluarkanAhli',
  'padamFail', 'pulihFail', 'padamKekal', 'tambahTugasan', 'padamTugasan', 'batalPemantau',
  'simpanPanitia', 'padamPanitia', 'lantikKP', 'kemaskiniPengguna', 'simpanTetapan', 'luluskanPengguna', 'padamLaporan',
  'simpanProgram', 'padamProgram', 'simpanMesyuarat', 'padamMesyuarat', 'simpanTindakan', 'statusTindakan', 'padamTindakan',
  'simpanDokSekolah', 'padamDokSekolah'];

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
      .filter(k => aktif_(userById_(k.userId)))
      .map(k => ({ id: k.id, tahun: k.tahun, user: ringkas_(userById_(k.userId)) })) : [],
    pentadbir: pentadbir_(),
    carta: cartaBina_(p, ahli),
    fail,
    tong: a.urus ? failSemua.filter(f => f.status === 'TRASH').map(failPub_) : [],
    tugasan,
    semakan,
    laporan: senaraiLaporan_(p),
    dokSekolah: rows_('DokSekolah').filter(d => d.status === 'AKTIF').map(d => {
      const f = failSemua.find(x => x.status === 'AKTIF' && x.tugasanId === d.id);
      return { id: d.id, tajuk: d.tajuk, folderNo: Number(d.folderNo) || 1, tarikhAkhir: d.tarikhAkhir, catatan: d.catatan, siap: !!f, failId: f ? f.id : '' };
    }),
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
  return rows_('Pengguna').filter(x => aktif_(x) && ahliIds.indexOf(x.id) < 0)
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
    if (!aktif_(g)) return;
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
  if (rows_('Tugasan').some(t => t.id === id && t.panitiaId === p.id)) return id;
  return rows_('DokSekolah').some(d => d.id === id && d.status === 'AKTIF') ? id : '';
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
  statusLaporanIkutFail_(f.id, 'AKTIF', 'PADAM');
  log_(u, p.id, f.folderNo, 'PADAM', 'Ke tong sampah: ' + f.tajuk);
  return true;
}

function pulihFail(q, u) {
  const f = failById_(q.id);
  const p = panitia_(f.panitiaId);
  mesti_(akses_(u, p).urus, 'Hanya Ketua Panitia boleh memulihkan fail.');
  update_('Fail', f.id, { status: 'AKTIF', dipadamPada: '', dipadamOleh: '' });
  statusLaporanIkutFail_(f.id, 'PADAM', 'AKTIF');
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
  statusLaporanIkutFail_(f.id, 'PADAM', 'KEKAL_PADAM');
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
    statusLaporanIkutFail_(f.id, 'PADAM', 'KEKAL_PADAM');
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
    '</td><td style="text-align:right;vertical-align:bottom;color:#6a7771;font-size:9pt">Dijana melalui e-Kurikulum<br>' +
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
    laporan: senaraiLaporan_(p),
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
  mesti_(aktif_(g), 'Guru tidak dijumpai atau belum diluluskan.');
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
  const bakiAdmin = rows_('Pengguna').filter(x => x.peranan === 'SUPER_ADMIN' && aktif_(x) && x.id !== g.id).length;
  if ((patch.peranan === 'GURU' || patch.status === 'NYAHAKTIF') && g.peranan === 'SUPER_ADMIN') mesti_(bakiAdmin > 0, 'Sistem mesti mempunyai sekurang-kurangnya seorang Super Admin aktif.');
  update_('Pengguna', g.id, patch);
  if (patch.status === 'NYAHAKTIF') deleteWhere_('Sesi', s => s.userId === g.id);
  log_(u, '', '', 'KEMAS KINI PENGGUNA', g.nama + ' ' + JSON.stringify(patch));
  return balasAdmin_(q, u);
}

/** Super Admin meluluskan atau menolak pendaftaran baharu */
function luluskanPengguna(q, u) {
  admin_(u);
  const g = userById_(q.userId);
  mesti_(g, 'Pengguna tidak dijumpai.');
  mesti_(g.status === 'MENUNGGU', 'Akaun ini tidak menunggu kelulusan.');
  if (q.lulus) {
    update_('Pengguna', g.id, { status: 'AKTIF', dikemaskini: now_() });
    log_(u, '', '', 'LULUS PENDAFTARAN', g.nama);
    hantarEmel_(g.emel, 'Akaun e-Kurikulum anda telah diluluskan',
      '<p>Assalamualaikum dan salam sejahtera ' + esc_(g.nama) + ',</p><p>Pendaftaran akaun e-Kurikulum anda telah <b>diluluskan</b>. ' +
      'Anda kini boleh log masuk menggunakan emel <b>' + esc_(g.emel) + '</b> dan kata laluan yang anda tetapkan semasa mendaftar.</p>' +
      '<p>Keahlian panitia yang anda pilih akan disahkan oleh Ketua Panitia masing-masing.</p>');
  } else {
    deleteWhere_('Keahlian', k => k.userId === g.id);
    deleteWhere_('Pengguna', x => x.id === g.id);
    log_(u, '', '', 'TOLAK PENDAFTARAN', g.nama + ' (' + g.emel + ')');
  }
  return balasAdmin_(q, u);
}

function resetKataLaluan(q, u) {
  admin_(u);
  const g = userById_(q.userId);
  mesti_(g, 'Pengguna tidak dijumpai.');
  const baru = uid_(8);
  update_('Pengguna', g.id, Object.assign(kataBaru_(baru), { wajibTukar: 'YA' }));
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
  if (d.domainEmel !== undefined) {
    const dm = bersih_(d.domainEmel, 200).toLowerCase().replace(/@/g, '').split(/[\s,]+/).filter(Boolean);
    dm.forEach(x => mesti_(/^[a-z0-9-]+(\.[a-z0-9-]+)+$/.test(x), 'Domain emel tidak sah: ' + x + '. Contoh: moe-dl.edu.my'));
    setTetapan_('domainEmel', dm.join(', '));
  }
  if (d.aksesKurikulum !== undefined) {
    const ada = {};
    rows_('Pengguna').forEach(x => { ada[x.id] = 1; });
    setTetapan_('aksesKurikulum', String(d.aksesKurikulum).split(/[\s,]+/).filter(id => ada[id]).join(','));
  }
  ['emelPeringatan', 'benamkan'].forEach(k => { if (d[k] !== undefined) setTetapan_(k, String(d[k]).toUpperCase() === 'YA' ? 'YA' : 'TIDAK'); });
  if (d.jadualCsvUrl !== undefined) {
    const url = bersih_(d.jadualCsvUrl, 600);
    mesti_(!url || /^https:\/\/\S+$/i.test(url), 'Pautan CSV jadual mesti bermula dengan https://');
    setTetapan_('jadualCsvUrl', url);
  }
  log_(u, '', '', 'TETAPAN', 'Tetapan sekolah dikemas kini');
  return balasAdmin_(q, u);
}

/* =====================================================================
 *  MIGRASI & AUTOMASI (trigger)
 * ===================================================================== */

/** Akaun import lama (kata laluan 12345678, belum ditukar) → kata laluan awal 6 digit akhir No. KP */
function lalaiKeNoKP_() {
  const sasar = rows_('Pengguna').filter(u => u.wajibTukar === 'YA' && u.hash !== 'LALAI' &&
    /^\d{12}$/.test(u.noKP) && u.salt && u.hash === hash_(KATA_LALUAN_LALAI, u.salt));
  sasar.forEach(u => update_('Pengguna', u.id, { hash: 'LALAI', salt: '' }));
  if (sasar.length) log_('Sistem', '', '', 'KATA LALUAN AWAL', sasar.length + ' akaun ditukar ke 6 digit akhir No. KP');
  return sasar.length;
}

/** Folder jadual waktu peribadi tidak lagi dikongsi melalui pautan awam (dipapar melalui aplikasi sahaja) */
function tutupKongsiJadual_() {
  try {
    const f = folderSistem_('_JADUAL_WAKTU');
    try { f.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE); } catch (e) { /* */ }
    const it = f.getFiles();
    while (it.hasNext()) { try { it.next().setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE); } catch (e) { /* */ } }
  } catch (e) { Logger.log('⚠️ Perkongsian folder jadual tidak dapat ditutup: ' + e.message); }
}

/** Trigger harian: emel peringatan kepada ahli yang belum menghantar dokumen wajib menjelang tarikh akhir */
function peringatanTugasan() {
  ensureSchema_();
  if (String(tetapan_().emelPeringatan || 'YA').toUpperCase() === 'TIDAK') { selesai_(); Logger.log('Emel peringatan dimatikan dalam Tetapan.'); return 0; }
  const hariIni = Utilities.formatDate(new Date(), ZON, 'yyyy-MM-dd');
  const beza = d => Math.round((Date.parse(d + 'T00:00:00Z') - Date.parse(hariIni + 'T00:00:00Z')) / 864e5);
  const sah = d => /^\d{4}-\d{2}-\d{2}$/.test(d || '') && HARI_PERINGATAN.indexOf(beza(d)) >= 0;
  const fail = rows_('Fail').filter(f => f.status === 'AKTIF' && f.tugasanId);
  const panitia = senaraiAktif_();
  const perGuru = {};
  const tambah = (u, item) => { if (u && u.emel && aktif_(u)) (perGuru[u.id] = perGuru[u.id] || { u, senarai: [] }).senarai.push(item); };
  const fol = n => ', Folder ' + ('0' + n).slice(-2);

  // 1. Dokumen wajib panitia (setiap ahli)
  rows_('Tugasan').forEach(tg => {
    if (!sah(tg.tarikhAkhir)) return;
    const p = panitia.find(x => x.id === tg.panitiaId);
    if (!p) return;
    ahliPanitia_(p).forEach(m => {
      if (fail.some(f => f.tugasanId === tg.id && f.pemilikId === m.id)) return;
      tambah(m, { tajuk: tg.tajuk, sumber: 'Dokumen wajib Panitia ' + p.nama + fol(tg.folderNo), tarikh: tg.tarikhAkhir, b: beza(tg.tarikhAkhir) });
    });
  });
  // 2. Dokumen wajib sekolah (Ketua Panitia & Setiausaha panitia yang belum menghantar)
  rows_('DokSekolah').filter(d => d.status === 'AKTIF' && sah(d.tarikhAkhir)).forEach(d => {
    panitia.forEach(p => {
      if (fail.some(f => f.tugasanId === d.id && f.panitiaId === p.id)) return;
      [p.ketuaId, p.setiausahaId].filter(Boolean).forEach(id => tambah(userById_(id),
        { tajuk: d.tajuk, sumber: 'Dokumen wajib sekolah, Panitia ' + p.nama + fol(d.folderNo), tarikh: d.tarikhAkhir, b: beza(d.tarikhAkhir) }));
    });
  });
  // 3. Tindakan susulan JK Kurikulum
  const mesy = {};
  rows_('Mesyuarat').forEach(m => { if (m.status !== 'PADAM') mesy[m.id] = m; });
  rows_('Tindakan').filter(t => t.status !== 'SELESAI' && mesy[t.mesyuaratId] && sah(t.tarikhAkhir)).forEach(t => {
    tambah(userById_(t.pemilikId), { tajuk: t.perkara, sumber: 'Tindakan ' + mesy[t.mesyuaratId].tajuk, tarikh: t.tarikhAkhir, b: beza(t.tarikhAkhir) });
  });

  let n = 0;
  Object.keys(perGuru).forEach(id => {
    const { u, senarai } = perGuru[id];
    senarai.sort((x, y) => x.b - y.b);
    const bila = b => b === 0 ? '<b style="color:#B83A2B">HARI INI</b>' : b === 1 ? '<b style="color:#B83A2B">esok</b>' : 'dalam ' + b + ' hari';
    const isi = '<p>Salam sejahtera ' + esc_(u.nama) + ',</p><p>Tugasan berikut <b>belum selesai</b> dalam e-Kurikulum:</p><ul>' +
      senarai.map(x => '<li style="margin-bottom:6px"><b>' + esc_(x.tajuk) + '</b><br><span style="color:#44545f">' + esc_(x.sumber) +
        ' · tarikh akhir ' + Utilities.formatDate(new Date(x.tarikh + 'T12:00:00Z'), ZON, 'dd/MM/yyyy') + ' (' + bila(x.b) + ')</span></li>').join('') +
      '</ul><p>Buka aplikasi dan lihat senarai <b>Tugasan saya</b> di Laman Utama.</p>';
    const segera = senarai[0].b <= 1 ? '[Segera] ' : '';
    if (hantarEmel_(u.emel, segera + 'Peringatan tugasan kurikulum (' + senarai.length + ')', isi)) n++;
  });
  if (n) log_('Sistem', '', '', 'PERINGATAN EMEL', n + ' guru diingatkan tentang tugasan kurikulum');
  selesai_();
  Logger.log('✅ Emel peringatan dihantar kepada ' + n + ' guru. Baki kuota emel hari ini: ' + MailApp.getRemainingDailyQuota());
  return n;
}

/** Trigger mingguan: salinan data (tanpa skrip) ke folder _SANDARAN. Simpan SIMPAN_SANDARAN salinan terkini. */
function sandaranMingguan() {
  ensureSchema_();
  const ss = getSS_();
  const folder = folderSistem_('_SANDARAN');
  const nama = 'Sandaran e-Panitia ' + Utilities.formatDate(new Date(), ZON, 'yyyy-MM-dd HHmm');
  const baru = SpreadsheetApp.create(nama);
  const asal = baru.getSheets()[0];
  asal.setName('__sementara_' + uid_(6));
  ss.getSheets().forEach(sh => { sh.copyTo(baru).setName(sh.getName()); });
  baru.deleteSheet(asal);
  DriveApp.getFileById(baru.getId()).moveTo(folder);
  const semua = [];
  const it = folder.getFiles();
  while (it.hasNext()) { const f = it.next(); if (/^Sandaran e-Panitia /.test(f.getName())) semua.push(f); }
  semua.sort((a, b) => b.getDateCreated() - a.getDateCreated()).slice(SIMPAN_SANDARAN).forEach(f => { try { f.setTrashed(true); } catch (e) { /* */ } });
  log_('Sistem', '', '', 'SANDARAN', nama);
  selesai_();
  Logger.log('✅ Sandaran dicipta: ' + baru.getUrl());
  return baru.getId();
}

/** Trigger bulanan: pindahkan log lebih lama daripada ARKIB_LOG_HARI ke tab "Arkib Log" */
function arkibLog() {
  ensureSchema_();
  kunci_();
  const had = Date.now() - ARKIB_LOG_HARI * 864e5;
  const t = table_('Log', true);
  const lama = t.rows.filter(r => { const m = new Date(r.masa).getTime(); return m && m < had; });
  if (!lama.length) { selesai_(); Logger.log('Tiada log lama untuk diarkib.'); return 0; }
  const ss = getSS_();
  let ar = ss.getSheetByName('Arkib Log');
  if (!ar) {
    ar = ss.insertSheet('Arkib Log');
    ar.getRange(1, 1, 1, t.head.length).setValues([t.head]).setFontWeight('bold').setBackground('#6a7771').setFontColor('#FFFFFF');
    ar.setFrozenRows(1);
  }
  const vals = lama.map(r => t.head.map(h => (r[h] === undefined ? '' : String(r[h]))));
  const mula = ar.getLastRow() + 1;
  if (mula + vals.length - 1 > ar.getMaxRows()) ar.insertRowsAfter(ar.getMaxRows(), mula + vals.length - 1 - ar.getMaxRows());
  ar.getRange(mula, 1, vals.length, t.head.length).setNumberFormat('@').setValues(vals);
  const buang = {};
  lama.forEach(r => { buang[r._row] = 1; });
  tulisSemula_('Log', t.rows.filter(r => !buang[r._row]));
  const sh = sh_('Log');
  const lebih = sh.getMaxRows() - Math.max(sh.getLastRow(), 1) - 100; // kecilkan helaian
  if (lebih > 0) sh.deleteRows(sh.getMaxRows() - lebih + 1, lebih);
  log_('Sistem', '', '', 'ARKIB LOG', lama.length + ' rekod lebih ' + ARKIB_LOG_HARI + ' hari dipindahkan ke tab Arkib Log');
  selesai_();
  Logger.log('✅ ' + lama.length + ' rekod log diarkib.');
  return lama.length;
}

/* =====================================================================
 *  LAPORAN AKTIVITI — ONE PAGE REPORT (OPR)
 *  Setiap laporan disimpan dalam jadual Laporan + PDF dalam Folder 10
 *  (direkod juga dalam jadual Fail supaya boleh dimuat turun / dipadam / dipulih).
 * ===================================================================== */
function fotoLaporan_(r) { try { const a = JSON.parse(r.fotoIds || '[]'); return Array.isArray(a) ? a : []; } catch (e) { return []; } }

function laporanPub_(r) {
  return {
    id: r.id, panitiaId: r.panitiaId, tarikh: r.tarikh, kategori: r.kategori, tajuk: r.tajuk,
    tarikhMula: r.tarikhMula, tarikhTamat: r.tarikhTamat, tarikhProgram: r.tarikhProgram, tempat: r.tempat,
    lakaranId: r.lakaranId, bilFoto: fotoLaporan_(r).filter(Boolean).length, disediakanOleh: r.disediakanOleh,
    disediakanId: r.disediakanId, failId: r.failId, driveId: r.driveId, dicipta: r.dicipta
  };
}
function senaraiLaporan_(p) {
  return rows_('Laporan').filter(r => r.panitiaId === p.id && r.status === 'AKTIF').map(laporanPub_)
    .sort((a, b) => (b.tarikhMula || b.dicipta) > (a.tarikhMula || a.dicipta) ? 1 : -1);
}
/** Kemas kini status laporan yang dipautkan kepada fail PDF (padam / pulih / padam kekal melalui Folder 10) */
function statusLaporanIkutFail_(failId, dari, ke) {
  rows_('Laporan').filter(r => r.failId === failId && r.status === dari).forEach(r => {
    update_('Laporan', r.id, { status: ke, dikemaskini: now_() });
    if (ke === 'KEKAL_PADAM') { fotoLaporan_(r).forEach(buangDrive_); buangDrive_(r.lakaranId); }
  });
}

function tarikhSah_(s) { return /^\d{4}-\d{2}-\d{2}$/.test(String(s || '')) && !isNaN(Date.parse(s + 'T00:00:00Z')) ? String(s) : ''; }
/** "6 - 31 JULAI 2026", "28 JUN - 3 JULAI 2026", "9 SEPTEMBER 2026" */
function julatTarikh_(mula, tamat) {
  const p = x => { const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(x || ''); return m ? { y: +m[1], m: +m[2], d: +m[3] } : null; };
  const a = p(mula), b = p(tamat);
  if (!a) return '';
  const penuh = x => x.d + ' ' + BULAN_MS[x.m - 1] + ' ' + x.y;
  if (!b || mula === tamat) return penuh(a);
  if (a.y === b.y && a.m === b.m) return a.d + ' - ' + b.d + ' ' + BULAN_MS[a.m - 1] + ' ' + a.y;
  if (a.y === b.y) return a.d + ' ' + BULAN_MS[a.m - 1] + ' - ' + penuh(b);
  return penuh(a) + ' - ' + penuh(b);
}
function tarikhDMY_(iso) { const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso || ''); return m ? (+m[3]) + '/' + (+m[2]) + '/' + m[1] : ''; }

/** Folder Drive "Foto Laporan" di dalam Folder 10 panitia (ID disimpan dalam peta subFolder, kunci "fl") */
function folderFotoLaporan_(p) {
  const induk = folderPanitia_(p, '10');
  let peta = petaFolder_(rows_('Panitia').find(x => x.id === p.id) || p);
  if (peta.fl) { try { return DriveApp.getFolderById(peta.fl); } catch (e) { /* cipta semula */ } }
  kunci_();
  const segar = rows_('Panitia').find(x => x.id === p.id) || p;
  peta = petaFolder_(segar);
  if (peta.fl) { try { return DriveApp.getFolderById(peta.fl); } catch (e) { /* */ } }
  const it = induk.getFoldersByName('Foto Laporan');
  const f = it.hasNext() ? it.next() : induk.createFolder('Foto Laporan');
  peta.fl = f.getId();
  update_('Panitia', p.id, { subFolder: JSON.stringify(peta) });
  return f;
}

function simpanLaporan(q, u) {
  const p = panitia_(q.panitiaId);
  const a = akses_(u, p);
  mesti_(a.lihat, 'Anda bukan ahli panitia ini.');
  let lama = null;
  if (q.id) {
    lama = rows_('Laporan').find(x => x.id === q.id && x.panitiaId === p.id && x.status === 'AKTIF');
    mesti_(lama, 'Laporan tidak dijumpai.');
    mesti_(lama.disediakanId === u.id || a.su, 'Hanya penyedia laporan, Ketua Panitia atau Setiausaha boleh mengubah laporan ini.');
  }
  const t = tetapan_();
  const hariIni = Utilities.formatDate(new Date(), ZON, 'yyyy-MM-dd');
  const d = {
    tarikh: tarikhSah_(q.tarikh) || hariIni,
    kategori: bersih_(q.kategori, 40).toUpperCase() || 'LAPORAN',
    tajuk: bersih_(q.tajuk, 200),
    tarikhMula: tarikhSah_(q.tarikhMula),
    tarikhTamat: tarikhSah_(q.tarikhTamat),
    tempat: bersih_(q.tempat, 200) || String(t.namaSekolah || ''),
    objektif: bersih_(q.objektif, 1200),
    tentatif: bersih_(q.tentatif, 1800),
    catatan: bersih_(q.catatan, 600),
    disediakanOleh: bersih_(q.disediakanOleh, 120) || u.nama
  };
  mesti_(d.tajuk, 'Isi tajuk program / bahan.');
  mesti_(d.tarikhMula, 'Isi tarikh program.');
  if (d.tarikhTamat && d.tarikhTamat < d.tarikhMula) d.tarikhTamat = '';
  d.tarikhProgram = julatTarikh_(d.tarikhMula, d.tarikhTamat);

  // Gambar: [ {data} baharu | {id} dikekalkan | null ] x 4
  const fotoLama = lama ? fotoLaporan_(lama) : [];
  const g = (Array.isArray(q.gambar) ? q.gambar : []).slice(0, 4);
  while (g.length < 4) g.push(null);
  g.forEach(x => {
    if (x && x.data) {
      mesti_(/^data:image\/(jpeg|png);base64,/.test(x.data), 'Format gambar tidak sah.');
      mesti_(saizB64_(x.data) < 3 * 1048576, 'Gambar terlalu besar.');
    } else if (x && x.id) mesti_(fotoLama.indexOf(x.id) >= 0, 'Gambar asal tidak dijumpai.');
  });
  mesti_(g.some(x => x && (x.data || x.id)), 'Pilih sekurang-kurangnya satu foto aktiviti.');

  const id = lama ? lama.id : uid_();
  const folderFoto = folderFotoLaporan_(p);
  const fotoIds = [], fotoData = [];
  g.forEach((x, i) => {
    if (x && x.data) {
      const f = simpanBlob_(folderFoto, x.data, 'image/jpeg', 'OPR_' + id + '_foto' + (i + 1) + '_' + Date.now() + '.jpg');
      fotoIds.push(f.getId()); fotoData.push(x.data);
      const durl = 'data:image/jpeg;base64,' + String(x.data).replace(/^data:[^,]*,/, '');
      if (durl.length < 95000) try { cache_().put('IMG_' + f.getId(), durl, CACHE_TTL); } catch (e) { /* */ }
    } else if (x && x.id) {
      fotoIds.push(x.id);
      let durl = '';
      try { const b = DriveApp.getFileById(x.id).getBlob(); durl = 'data:' + b.getContentType() + ';base64,' + Utilities.base64Encode(b.getBytes()); } catch (e) { /* hilang */ }
      fotoData.push(durl);
    } else { fotoIds.push(''); fotoData.push(''); }
  });
  let lakaranId = lama ? lama.lakaranId : '';
  if (q.lakaran && /^data:image\/jpeg;base64,/.test(q.lakaran) && saizB64_(q.lakaran) < 300000) {
    const lk = simpanBlob_(folderFoto, q.lakaran, 'image/jpeg', 'OPR_' + id + '_lakaran_' + Date.now() + '.jpg');
    if (lakaranId) buangDrive_(lakaranId);
    lakaranId = lk.getId();
    try { cache_().put('IMG_' + lakaranId, q.lakaran, CACHE_TTL); } catch (e) { /* */ }
  }

  // Jana PDF OPR
  const html = oprHtml_(p, d, fotoData, t);
  const namaFail = ('OPR ' + d.tajuk + ' (' + d.tarikhProgram + ')').replace(/[\\/:*?"<>|]/g, '-').slice(0, 150) + '.pdf';
  const pdf = Utilities.newBlob(html, 'text/html', 'opr.html').getAs('application/pdf').setName(namaFail);
  const file = folderPanitia_(p, '10').createFile(pdf);
  const failMeta = {
    tajuk: 'OPR: ' + d.tajuk, penerangan: d.kategori + ' — ' + d.tarikhProgram + (d.tempat ? ', ' + d.tempat : ''),
    url: file.getUrl(), driveId: file.getId(), mime: 'application/pdf', saiz: pdf.getBytes().length
  };
  let failRec = lama && lama.failId ? rows_('Fail').find(f => f.id === lama.failId && f.status !== 'KEKAL_PADAM') : null;
  if (failRec) {
    buangDrive_(failRec.driveId);
    failRec = update_('Fail', failRec.id, Object.assign(failMeta, { status: 'AKTIF', dipadamPada: '', dipadamOleh: '' }));
  } else {
    failRec = Object.assign({
      id: uid_(), panitiaId: p.id, folderNo: '10', jenis: 'DRIVE', tugasanId: tugasanSah_(p, q.tugasanId),
      pemilikId: u.id, pemilikNama: u.nama, dicipta: now_(), status: 'AKTIF'
    }, failMeta);
    insert_('Fail', failRec);
  }
  const rec = Object.assign({}, d, {
    id, panitiaId: p.id, fotoIds: JSON.stringify(fotoIds), lakaranId, failId: failRec.id, driveId: file.getId(),
    disediakanId: lama ? lama.disediakanId : u.id, dicipta: lama ? lama.dicipta : now_(), dikemaskini: now_(), status: 'AKTIF'
  });
  if (lama) {
    update_('Laporan', id, rec);
    fotoLama.filter(x => x && fotoIds.indexOf(x) < 0).forEach(buangDrive_); // foto yang diganti / dibuang
  } else {
    insert_('Laporan', rec);
  }
  if (q.programId) {
    const pg = rows_('Program').find(x => x.id === q.programId && x.panitiaId === p.id && x.status === 'AKTIF');
    if (pg) update_('Program', pg.id, { laporanId: id, dikemaskini: now_() });
  }
  log_(u, p.id, '10', lama ? 'KEMAS KINI OPR' : 'LAPORAN OPR', d.tajuk);
  return { laporan: laporanPub_(rec), fail: failPub_(failRec) };
}

/** Butiran penuh satu laporan (ahli panitia atau pemantau) */
function laporanDetail(q) {
  const r = rows_('Laporan').find(x => x.id === q.id && x.status === 'AKTIF');
  mesti_(r, 'Laporan tidak dijumpai atau telah dipadam.');
  const p = panitia_(r.panitiaId);
  let boleh = false;
  if (q.pemantau) { mesti_(pemantauSah_(q.pemantau).panitiaId === p.id, 'Tiada akses.'); }
  else {
    const u = auth_(q.token);
    const a = akses_(u, p);
    mesti_(a.lihat, 'Anda bukan ahli panitia ini.');
    boleh = r.disediakanId === u.id || a.su;
  }
  return Object.assign(laporanPub_(r), {
    objektif: r.objektif, tentatif: r.tentatif, catatan: r.catatan, fotoIds: fotoLaporan_(r),
    panitia: p.nama, sekolah: tetapan_().namaSekolah, bolehUbah: boleh
  });
}

function padamLaporan(q, u) {
  const r = rows_('Laporan').find(x => x.id === q.id && x.status === 'AKTIF');
  mesti_(r, 'Laporan tidak dijumpai.');
  const p = panitia_(r.panitiaId);
  const a = akses_(u, p);
  mesti_(a.lihat && (r.disediakanId === u.id || a.su), 'Hanya penyedia laporan, Ketua Panitia atau Setiausaha boleh memadam laporan ini.');
  update_('Laporan', r.id, { status: 'PADAM', dikemaskini: now_() });
  const f = r.failId ? rows_('Fail').find(x => x.id === r.failId && x.status === 'AKTIF') : null;
  if (f) update_('Fail', f.id, { status: 'TRASH', dipadamPada: now_(), dipadamOleh: u.nama });
  log_(u, p.id, '10', 'PADAM OPR', r.tajuk + ' (PDF ke tong sampah)');
  return balasDetail_(p.id, u);
}

/** HTML OPR satu muka surat A4 — saiz huruf & foto dikecilkan automatik mengikut panjang teks */
function oprHtml_(p, d, foto, t) {
  // Bajet ketinggian A4 (px @96dpi): ruang cetak ~1039px; simpan margin keselamatan untuk penukar PDF Google
  const baris = (s, lebar) => String(s || '').split('\n').reduce((n, l) => n + Math.max(1, Math.ceil(l.length / lebar)), 0);
  const bilFoto = foto.filter(Boolean).length;
  const barisFoto = Math.max(1, Math.ceil(bilFoto / 2));
  const TETAP = 335 + (d.tajuk.length > 55 ? 22 : 0) + (d.tempat.length > 60 ? 20 : 0);
  const tinggiTeks = saiz => {
    const lebar = Math.round(50 * 10 / saiz);                   // aksara sebaris (separuh lebar)
    return (Math.max(baris(d.objektif, lebar), baris(d.tentatif, lebar)) + (d.catatan ? baris(d.catatan, lebar * 2) : 1)) * saiz * 1.333 * 1.3;
  };
  let fs = 10, ft = 0, sebaris = 2;                              // sebaris = bilangan foto dalam satu baris
  for (const saiz of [10, 9.5, 9, 8.5, 8, 7.5]) {
    fs = saiz;
    ft = Math.floor((970 - TETAP - tinggiTeks(fs)) / barisFoto) - 34; // tinggi setiap foto (grid 2 x 2)
    if (ft >= 140) break;
  }
  if (ft < 115 && 970 - TETAP - tinggiTeks(7) - (bilFoto > 2 ? 154 : 119 * barisFoto) < 0) {
    // Mod padat: terlalu banyak baris pendek — gabungkan baris dengan " • " supaya kekal satu muka surat
    const padat = x => String(x || '').split('\n').map(l => l.trim()).filter(Boolean).join('  •  ');
    d = Object.assign({}, d, { objektif: padat(d.objektif), tentatif: padat(d.tentatif), catatan: padat(d.catatan) });
    for (const saiz of [9, 8.5, 8, 7.5]) { fs = saiz; ft = Math.floor((970 - TETAP - tinggiTeks(fs)) / barisFoto) - 34; if (ft >= 150) break; }
  }
  if (ft < 115 && bilFoto > 2) {
    // Teks sangat panjang: susun 4 foto dalam satu baris supaya kekal satu muka surat
    sebaris = 4;
    for (const saiz of [9, 8.5, 8, 7.5, 7]) { fs = saiz; ft = Math.floor(970 - TETAP - tinggiTeks(fs)) - 34; if (ft >= 115) break; }
    ft = Math.max(80, Math.min(ft, 120));
  } else ft = Math.max(85, Math.min(ft, barisFoto === 1 ? 260 : 230));
  const fw = Math.round(ft * 4 / 3);
  const B = '1px solid #8fa39a';
  const teks = s => esc_(s || '-').replace(/\n/g, '<br>');
  const lbl = (l, extra) => '<td style="width:17%;padding:5px 7px;border:' + B + ';background:#e3eee8;font-weight:bold;font-size:' + (fs - 1) + 'pt;color:#145C43;vertical-align:top' + (extra || '') + '">' + l + '</td>';
  const val = (v, span, gaya) => '<td' + (span ? ' colspan="' + span + '"' : '') + ' style="padding:5px 7px;border:' + B + ';vertical-align:top;' + (gaya || '') + '">' + v + '</td>';

  const sel = [];
  foto.forEach((src, i) => {
    if (!src) return;
    sel.push('<td style="width:' + (100 / sebaris) + '%;padding:5px;text-align:center;vertical-align:top;border:' + B + '">' +
      '<img src="' + src + '" style="width:' + fw + 'px;height:' + ft + 'px"><div style="font-size:' + (fs - 1.5) + 'pt;font-weight:bold;color:#145C43;margin-top:3px">FOTO ' + (i + 1) + '</div></td>');
  });
  while (sel.length % sebaris) sel.push('<td style="border:' + B + '"></td>');
  let grid = '';
  for (let i = 0; i < sel.length; i += sebaris) grid += '<tr>' + sel.slice(i, i + sebaris).join('') + '</tr>';

  return '<!DOCTYPE html><html><head><meta charset="utf-8"><style>@page{size:A4;margin:11mm 12mm}body{font-family:Arial,sans-serif;font-size:' + fs + 'pt;color:#1c2b39;margin:0}table{border-collapse:collapse;width:100%}td{font-size:' + fs + 'pt}</style></head><body>' +
    '<table style="margin-bottom:8px"><tr><td style="background:#145C43;color:#fff;padding:9px 12px;text-align:center">' +
    '<div style="font-size:10.5pt;font-weight:bold;letter-spacing:.5px">' + esc_(String(t.namaSekolah || '').toUpperCase()) + '</div>' +
    '<div style="font-size:15pt;font-weight:bold;margin-top:2px">LAPORAN AKTIVITI PANITIA ' + esc_(p.nama.toUpperCase()) + '</div>' +
    '<div style="font-size:9pt;color:#F2B705;font-weight:bold;letter-spacing:2px;margin-top:2px">ONE PAGE REPORT (OPR)</div></td></tr></table>' +
    '<table>' +
    '<tr>' + lbl('PANITIA') + val('PANITIA ' + esc_(p.nama.toUpperCase())) + lbl('TARIKH') + val(esc_(tarikhDMY_(d.tarikh)), 0, 'width:22%') + '</tr>' +
    '<tr>' + lbl('KATEGORI') + val(esc_(d.kategori)) + lbl('TARIKH PROGRAM') + val(esc_(d.tarikhProgram), 0, 'width:22%') + '</tr>' +
    '<tr>' + lbl('TAJUK BAHAN') + val('<b style="font-size:' + (fs + 1) + 'pt">' + esc_(d.tajuk.toUpperCase()) + '</b>', 3) + '</tr>' +
    '<tr>' + lbl('TEMPAT / LOKASI') + val(esc_(d.tempat.toUpperCase()), 3) + '</tr></table>' +
    '<table style="margin-top:6px"><tr>' +
    '<td style="width:50%;padding:5px 7px;border:' + B + ';background:#e3eee8;font-weight:bold;color:#145C43;font-size:' + (fs - 1) + 'pt">OBJEKTIF</td>' +
    '<td style="width:50%;padding:5px 7px;border:' + B + ';background:#e3eee8;font-weight:bold;color:#145C43;font-size:' + (fs - 1) + 'pt">TENTATIF / AKTIVITI</td></tr>' +
    '<tr>' + val(teks(d.objektif), 0, 'width:50%;line-height:1.35') + val(teks(d.tentatif), 0, 'width:50%;line-height:1.35') + '</tr>' +
    '<tr><td colspan="2" style="padding:5px 7px;border:' + B + '"><b style="color:#145C43;font-size:' + (fs - 1) + 'pt">CATATAN / IMPAK: </b>' + teks(d.catatan) + '</td></tr></table>' +
    '<table style="margin-top:6px">' + grid + '</table>' +
    '<table style="margin-top:8px"><tr><td style="width:60%;vertical-align:top;font-size:' + (fs - 0.5) + 'pt">' +
    '<b style="color:#145C43">DISEDIAKAN OLEH:</b><br><b>' + esc_(d.disediakanOleh.toUpperCase()) + '</b></td>' +
    '<td style="text-align:right;vertical-align:bottom;color:#6a7771;font-size:7.5pt">Dijana melalui e-Kurikulum<br>' +
    Utilities.formatDate(new Date(), ZON, 'dd/MM/yyyy HH:mm') + '</td></tr></table></body></html>';
}

/* =====================================================================
 *  E-KURIKULUM FASA 1: TAKWIM, JK KURIKULUM, TINDAKAN SUSULAN,
 *  DOKUMEN WAJIB SEKOLAH DAN PAPAN PEMUKA KURIKULUM
 *  ID rekod baharu dijana oleh klien (paparan serta-merta / outbox),
 *  disahkan formatnya di sini.
 * ===================================================================== */
function idBaru_(id, jadual) {
  id = String(id || '');
  mesti_(/^[A-Za-z0-9_]{8,40}$/.test(id), 'ID rekod tidak sah.');
  mesti_(!rows_(jadual).some(x => x.id === id), 'Rekod sudah wujud.');
  return id;
}
function tarikhWajib_(s, msg) { const t = tarikhSah_(s); mesti_(t, msg || 'Isi tarikh yang sah.'); return t; }

function programPub_(x, nmP) {
  return { id: x.id, panitiaId: x.panitiaId, panitia: x.panitiaId ? (nmP[x.panitiaId] || '') : '', tajuk: x.tajuk, kategori: x.kategori,
    tarikhMula: x.tarikhMula, tarikhTamat: x.tarikhTamat, masa: x.masa, tempat: x.tempat, catatan: x.catatan, laporanId: x.laporanId,
    diciptaId: x.diciptaId, diciptaOleh: x.diciptaOleh };
}
function mesyuaratPub_(m) {
  return { id: m.id, tahun: m.tahun, bil: Number(m.bil) || 0, tajuk: m.tajuk, tarikh: m.tarikh, masa: m.masa, tempat: m.tempat, status: m.status,
    minitUrl: m.minitUrl, minitDriveId: m.minitDriveId, minitNama: m.minitNama, catatan: m.catatan };
}
function tindakanPub_(t) {
  return { id: t.id, mesyuaratId: t.mesyuaratId, perkara: t.perkara, pemilikId: t.pemilikId, pemilikNama: t.pemilikNama,
    tarikhAkhir: t.tarikhAkhir, status: t.status || 'BELUM', catatan: t.catatan, selesaiPada: t.selesaiPada, selesaiOleh: t.selesaiOleh };
}

/** Data tab Takwim & Kurikulum (dihantar dalam sync) */
function kurikulumData_(u) {
  const pk = pentadbirKur_(u);
  const panitia = senaraiAktif_();
  const nmP = {};
  panitia.forEach(p => { nmP[p.id] = p.nama; });
  const dariTahun = String(new Date().getFullYear() - 1);
  const program = rows_('Program').filter(x => x.status === 'AKTIF' && String(x.tarikhMula).slice(0, 4) >= dariTahun && (!x.panitiaId || nmP[x.panitiaId]))
    .map(x => programPub_(x, nmP));
  const mesy = rows_('Mesyuarat').filter(m => m.status !== 'PADAM');
  const adaMesy = {};
  mesy.forEach(m => { adaMesy[m.id] = 1; });
  const tindakanSemua = rows_('Tindakan').filter(t => adaMesy[t.mesyuaratId]);
  const panitiaSaya = panitia.filter(p => p.ketuaId === u.id || p.setiausahaId === u.id).map(p => p.id);
  const dok = rows_('DokSekolah').filter(d => d.status === 'AKTIF');
  const failTanda = rows_('Fail').filter(f => f.status === 'AKTIF' && f.tugasanId);
  const sasaranDok = pk ? panitia : panitia.filter(p => panitiaSaya.indexOf(p.id) >= 0);
  const out = {
    akses: pk ? 'PENTADBIR' : 'GURU',
    program,
    mesyuarat: mesy.map(mesyuaratPub_).sort((a, b) => (b.tarikh > a.tarikh ? 1 : -1)),
    tindakan: (pk ? tindakanSemua : tindakanSemua.filter(t => t.pemilikId === u.id)).map(tindakanPub_),
    dokSekolah: dok.map(d => ({
      id: d.id, tajuk: d.tajuk, folderNo: Number(d.folderNo) || 1, tarikhAkhir: d.tarikhAkhir, catatan: d.catatan,
      status: sasaranDok.map(p => {
        const f = failTanda.find(x => x.tugasanId === d.id && x.panitiaId === p.id);
        return { panitiaId: p.id, nama: p.nama, siap: !!f, dicipta: f ? f.dicipta : '' };
      })
    })),
    panitiaSaya
  };
  if (pk) {
    out.papan = papanKurikulum_(panitia, dok, failTanda);
    out.guru = rows_('Pengguna').filter(aktif_).map(x => ({ id: x.id, nama: x.nama })).sort((a, b) => a.nama.localeCompare(b.nama));
  }
  return out;
}

function papanKurikulum_(panitia, dok, failTanda) {
  const th = String(new Date().getFullYear());
  const lap = rows_('Laporan').filter(r => r.status === 'AKTIF' && String(r.tarikhMula || r.dicipta).slice(0, 4) === th);
  const fail = rows_('Fail').filter(f => f.status === 'AKTIF');
  return panitia.map(p => {
    const fP = fail.filter(f => f.panitiaId === p.id);
    const kp = userById_(p.ketuaId);
    return {
      id: p.id, nama: p.nama, kod: p.kod, ketua: kp ? kp.nama : '', bilAhli: ahliPanitia_(p).length, bilFail: fP.length,
      opr: lap.filter(r => r.panitiaId === p.id).length,
      dokSiap: dok.filter(d => failTanda.some(f => f.tugasanId === d.id && f.panitiaId === p.id)).length, dokJumlah: dok.length,
      terakhir: fP.reduce((m, f) => (f.dicipta > m ? f.dicipta : m), '')
    };
  });
}

/* ---------- Takwim: program ---------- */
function bolehUrusProgram_(u, x) {
  if (pentadbirKur_(u)) return true;
  if (x.diciptaId === u.id) return true;
  if (!x.panitiaId) return false;
  const p = rows_('Panitia').find(y => y.id === x.panitiaId && y.status !== 'PADAM');
  return !!p && akses_(u, p).su;
}
/** Ketua Panitia atau Setiausaha bagi mana-mana panitia aktif */
function kpAtauSu_(u) { return senaraiAktif_().some(p => p.ketuaId === u.id || p.setiausahaId === u.id); }
function simpanProgram(q, u) {
  const panitiaId = bersih_(q.panitiaId, 20);
  if (panitiaId) mesti_(akses_(u, panitia_(panitiaId)).su || pentadbirKur_(u), 'Hanya Ketua Panitia, Setiausaha atau pentadbir boleh menambah program panitia.');
  else mesti_(pentadbirKur_(u) || kpAtauSu_(u), 'Hanya pentadbir, Ketua Panitia atau Setiausaha Panitia boleh menambah aktiviti peringkat sekolah.');
  const d = {
    panitiaId, tajuk: bersih_(q.tajuk, 200), kategori: bersih_(q.kategori, 40).toUpperCase() || 'PROGRAM',
    tarikhMula: tarikhWajib_(q.tarikhMula, 'Isi tarikh program.'), tarikhTamat: tarikhSah_(q.tarikhTamat),
    masa: bersih_(q.masa, 60), tempat: bersih_(q.tempat, 200), catatan: bersih_(q.catatan, 500), dikemaskini: now_()
  };
  mesti_(d.tajuk, 'Isi nama program.');
  if (d.tarikhTamat && d.tarikhTamat < d.tarikhMula) d.tarikhTamat = '';
  const lama = rows_('Program').find(x => x.id === q.id && x.status === 'AKTIF');
  if (lama) {
    mesti_(bolehUrusProgram_(u, lama), 'Anda tiada kebenaran mengubah program ini.');
    update_('Program', lama.id, d);
    log_(u, panitiaId, '', 'KEMAS KINI PROGRAM', d.tajuk);
  } else {
    insert_('Program', Object.assign(d, { id: idBaru_(q.id, 'Program'), laporanId: '', diciptaId: u.id, diciptaOleh: u.nama, dicipta: now_(), status: 'AKTIF' }));
    log_(u, panitiaId, '', 'PROGRAM BAHARU', d.tajuk + ' (' + d.tarikhMula + ')');
  }
  return { ok: true };
}
function padamProgram(q, u) {
  const x = rows_('Program').find(y => y.id === q.id && y.status === 'AKTIF');
  mesti_(x, 'Program tidak dijumpai.');
  mesti_(bolehUrusProgram_(u, x), 'Anda tiada kebenaran memadam program ini.');
  update_('Program', x.id, { status: 'BATAL', dikemaskini: now_() });
  log_(u, x.panitiaId, '', 'PADAM PROGRAM', x.tajuk);
  return { ok: true };
}

/* ---------- JK Kurikulum: mesyuarat, minit, tindakan susulan ---------- */
function simpanMesyuarat(q, u) {
  kurAdmin_(u);
  const tarikh = tarikhWajib_(q.tarikh, 'Isi tarikh mesyuarat.');
  const bil = Math.max(1, Math.min(20, Math.floor(Number(q.bil)) || 1));
  const tahun = tarikh.slice(0, 4);
  const d = {
    tahun, bil, tarikh, tajuk: bersih_(q.tajuk, 200) || ('Mesyuarat JK Kurikulum Bil. ' + bil + '/' + tahun),
    masa: bersih_(q.masa, 60), tempat: bersih_(q.tempat, 200), catatan: bersih_(q.catatan, 500),
    status: q.status === 'SELESAI' ? 'SELESAI' : 'DIRANCANG', dikemaskini: now_()
  };
  if (q.minitUrl !== undefined) {
    const url = bersih_(q.minitUrl, 1000);
    mesti_(!url || /^https?:\/\/\S+$/i.test(url), 'Pautan minit mesti bermula dengan http:// atau https://');
    d.minitUrl = url;
  }
  const lama = rows_('Mesyuarat').find(x => x.id === q.id && x.status !== 'PADAM');
  if (lama) update_('Mesyuarat', lama.id, d);
  else insert_('Mesyuarat', Object.assign(d, { id: idBaru_(q.id, 'Mesyuarat'), diciptaOleh: u.nama, dicipta: now_() }));
  log_(u, '', '', lama ? 'KEMAS KINI MESYUARAT' : 'MESYUARAT BAHARU', d.tajuk);
  return { ok: true };
}
function padamMesyuarat(q, u) {
  kurAdmin_(u);
  const m = rows_('Mesyuarat').find(x => x.id === q.id && x.status !== 'PADAM');
  mesti_(m, 'Mesyuarat tidak dijumpai.');
  update_('Mesyuarat', m.id, { status: 'PADAM', dikemaskini: now_() });
  log_(u, '', '', 'PADAM MESYUARAT', m.tajuk);
  return { ok: true };
}
/** Minit dimuat naik ke folder Drive _JK_KURIKULUM (tidak dikongsi awam) */
function muatNaikMinit(q, u) {
  kurAdmin_(u);
  const m = rows_('Mesyuarat').find(x => x.id === q.id && x.status !== 'PADAM');
  mesti_(m, 'Mesyuarat tidak dijumpai. Tunggu sebentar sehingga mesyuarat disimpan, kemudian cuba lagi.');
  const f = q.fail || {};
  const saiz = saizB64_(f.data);
  mesti_(saiz > 0, 'Tiada fail dipilih.');
  mesti_(saiz <= MAKS_FAIL_MB * 1048576, 'Saiz fail melebihi had ' + MAKS_FAIL_MB + 'MB.');
  const file = simpanBlob_(folderSistem_('_JK_KURIKULUM'), f.data, f.mime, bersih_(f.nama, 150) || ('Minit ' + m.tajuk));
  const lama = m.minitDriveId;
  update_('Mesyuarat', m.id, { minitDriveId: file.getId(), minitNama: file.getName(), minitMime: f.mime || '', status: 'SELESAI', dikemaskini: now_() });
  buangDrive_(lama);
  log_(u, '', '', 'MINIT JK', m.tajuk);
  return mesyuaratPub_(rows_('Mesyuarat').find(x => x.id === m.id));
}
function lihatMinit(q, u) {
  const m = rows_('Mesyuarat').find(x => x.id === q.id && x.status !== 'PADAM');
  mesti_(m && m.minitDriveId, 'Minit belum dimuat naik.');
  let fd;
  try { fd = DriveApp.getFileById(m.minitDriveId); } catch (e) { throw new Error('Fail minit tidak lagi wujud dalam Google Drive.'); }
  mesti_(fd.getSize() < 20 * 1048576, 'Fail minit terlalu besar untuk dipaparkan.');
  const b = fd.getBlob();
  return { nama: b.getName(), mime: b.getContentType(), data: Utilities.base64Encode(b.getBytes()) };
}
function simpanTindakan(q, u) {
  kurAdmin_(u);
  const m = rows_('Mesyuarat').find(x => x.id === q.mesyuaratId && x.status !== 'PADAM');
  mesti_(m, 'Mesyuarat tidak dijumpai.');
  const pemilik = q.pemilikId ? userById_(q.pemilikId) : null;
  const d = {
    mesyuaratId: m.id, perkara: bersih_(q.perkara, 400), pemilikId: pemilik ? pemilik.id : '',
    pemilikNama: pemilik ? pemilik.nama : bersih_(q.pemilikNama, 120), tarikhAkhir: tarikhSah_(q.tarikhAkhir), catatan: bersih_(q.catatan, 400)
  };
  mesti_(d.perkara, 'Isi perkara tindakan.');
  mesti_(d.pemilikNama, 'Pilih atau isi nama yang bertanggungjawab.');
  const lama = rows_('Tindakan').find(x => x.id === q.id);
  if (lama) update_('Tindakan', lama.id, d);
  else insert_('Tindakan', Object.assign(d, { id: idBaru_(q.id, 'Tindakan'), status: 'BELUM', dicipta: now_() }));
  log_(u, '', '', lama ? 'KEMAS KINI TINDAKAN' : 'TINDAKAN BAHARU', d.perkara + ' → ' + d.pemilikNama);
  return { ok: true };
}
function statusTindakan(q, u) {
  const t = rows_('Tindakan').find(x => x.id === q.id);
  mesti_(t, 'Tindakan tidak dijumpai.');
  mesti_(t.pemilikId === u.id || pentadbirKur_(u), 'Hanya penerima tindakan atau pentadbir boleh mengemas kini status.');
  const selesai = q.status === 'SELESAI';
  const patch = { status: selesai ? 'SELESAI' : 'BELUM', selesaiPada: selesai ? now_() : '', selesaiOleh: selesai ? u.nama : '' };
  if (q.catatan !== undefined) patch.catatan = bersih_(q.catatan, 400);
  update_('Tindakan', t.id, patch);
  log_(u, '', '', selesai ? 'TINDAKAN SELESAI' : 'TINDAKAN DIBUKA SEMULA', t.perkara);
  return { ok: true };
}
function padamTindakan(q, u) {
  kurAdmin_(u);
  const t = rows_('Tindakan').find(x => x.id === q.id);
  mesti_(t, 'Tindakan tidak dijumpai.');
  deleteWhere_('Tindakan', x => x.id === t.id);
  log_(u, '', '', 'PADAM TINDAKAN', t.perkara);
  return { ok: true };
}

/* ---------- Dokumen wajib sekolah (satu penghantaran bagi setiap panitia) ---------- */
function simpanDokSekolah(q, u) {
  kurAdmin_(u);
  let no = Math.floor(Number(q.folderNo)) || 4;
  if (no < 1 || no > 10) no = 4;
  const d = { tajuk: bersih_(q.tajuk, 150), folderNo: no, tarikhAkhir: tarikhSah_(q.tarikhAkhir), catatan: bersih_(q.catatan, 400) };
  mesti_(d.tajuk, 'Isi nama dokumen.');
  const lama = rows_('DokSekolah').find(x => x.id === q.id && x.status === 'AKTIF');
  if (lama) update_('DokSekolah', lama.id, d);
  else insert_('DokSekolah', Object.assign(d, { id: idBaru_(q.id, 'DokSekolah'), diciptaOleh: u.nama, dicipta: now_(), status: 'AKTIF' }));
  log_(u, '', String(no), lama ? 'KEMAS KINI DOK. WAJIB SEKOLAH' : 'DOK. WAJIB SEKOLAH', d.tajuk);
  return { ok: true };
}
function padamDokSekolah(q, u) {
  kurAdmin_(u);
  const d = rows_('DokSekolah').find(x => x.id === q.id && x.status === 'AKTIF');
  mesti_(d, 'Dokumen tidak dijumpai.');
  update_('DokSekolah', d.id, { status: 'PADAM' });
  log_(u, '', '', 'PADAM DOK. WAJIB SEKOLAH', d.tajuk);
  return { ok: true };
}
