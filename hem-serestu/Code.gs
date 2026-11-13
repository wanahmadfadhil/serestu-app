/**
 * ==================================================================
 *  SISTEM PENGURUSAN HEM PINTAR — Code.gs (Backend Lengkap, 1 fail)
 *  Versi 1.0 | Google Apps Script + Google Sheets
 * ------------------------------------------------------------------
 *  CARA PASANG
 *  1. Buka Google Sheet baharu → Extensions → Apps Script.
 *  2. Padam kod asal, tampal KESELURUHAN fail ini → Simpan.
 *  3. Pilih fungsi setup → Run (beri kebenaran bila diminta).
 *  4. Muat semula Google Sheet → menu "🎓 HEM Pintar" muncul.
 *  5. Isi tbl_tetapan (nama sekolah, penandatangan surat, dll).
 *
 *  CARA IMPORT DATA IDME
 *  1. Muat turun senarai murid dari IDME. JANGAN buka/simpan dalam Excel.
 *  2. Buka helaian IDME_IMPORT → File → Import → Upload fail tersebut.
 *  3. Pilih "Replace current sheet", NYAHTANDA
 *     "Convert text to numbers, dates, and formulas" → Import data.
 *  4. Menu 🎓 HEM Pintar → 2. Import Data IDME.
 *  5. Tetapkan lajur Sesi (Pagi/Petang) dalam tbl_kelas.
 *  Import boleh diulang bila-bila masa: rekod dikemas kini ikut
 *  ID MURID, mata merit & data sistem TIDAK terpadam.
 *
 *  CARA DEPLOY
 *  Deploy → New deployment → Web app
 *    Execute as: Me   |   Who has access: Anyone
 *  Salin URL /exec ke frontend.
 *
 *  NAIK TARAF KE v1.1
 *  • Jalankan setup sekali (beri kebenaran MailApp untuk emel pengesahan pendaftaran).
 *  • Deploy → Manage deployments → Edit → Version: New version (URL /exec kekal).
 *  • Kata laluan sedia ada kekal sah; hash dinaik taraf secara automatik semasa log masuk.
 *  • Menu "Arkib Kehadiran Tahun Lepas" (jalankan setiap awal tahun) memastikan sistem kekal pantas.
 *
 *  CONTOH PANGGILAN DARI FRONTEND
 *  fetch(URL, {
 *    method: 'POST',
 *    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
 *    body: JSON.stringify({ action: 'login', payload: { emel, pin } })
 *  }).then(r => r.json());
 *  Tindakan lain: { action, token, payload }
 * ==================================================================
 */

const APP = {
  NAMA: 'e-HEM Pintar',
  VERSI: '1.1',
  TZ: 'Asia/Kuala_Lumpur',
  TOKEN_SAAT: 6 * 3600,          // sesi log masuk 6 jam (had CacheService)
  MIN_KATA_LALUAN: 8,
  HASH_ULANG: 200,                // bilangan pusingan SHA-256 bagi setiap kata laluan (seimbang dgn had masa GAS)
  SHEET_IMPORT: 'IDME_IMPORT',
  SHEET_PIN_SEMENTARA: 'KATA_LALUAN_SEMENTARA',
  CACHE_DASHBOARD_SAAT: 300,
  MAKS_FOTO_MB: 2,
  MAKS_FOTO_PUKAL_MB: 20,         // jumlah saiz satu respons lihatFotoPukal
  MAKS_JSON_SEL: 40000            // had aksara JSON dalam satu sel (had Sheets 50,000)
};

/* ------------------------------------------------------------------
 *  SKEMA JADUAL — setup() membina/menambah lajur mengikut senarai ini.
 *  "text" = lajur diformat teks biasa (elak ID/telefon/tarikh rosak).
 * ------------------------------------------------------------------ */
const SCHEMA = {
  tbl_tetapan: {
    color: '#334155',
    headers: ['Kunci', 'Nilai', 'Keterangan'],
    text: ['Kunci', 'Nilai']
  },
  tbl_guru: {
    color: '#0284c7',
    headers: ['ID_Guru', 'Nama_Penuh', 'Emel', 'No_Telefon', 'Jawatan', 'Gred', 'Subjek',
      'Peranan', 'Sesi', 'Foto_ID', 'PIN_Hash', 'Tukar_PIN', 'Status_Aktif', 'Tarikh_Daftar', 'No_KP',
      'Tema', 'Latar', 'Header_ID', 'Header_URL', 'Kelas_Mohon', 'Cara_Daftar'],
    text: ['ID_Guru', 'No_Telefon', 'PIN_Hash', 'No_KP', 'Header_ID']
  },
  tbl_murid: {
    color: '#0ea5e9',
    headers: ['ID_Murid', 'Nama', 'Jantina', 'Tarikh_Lahir', 'Tahun', 'ID_Kelas', 'Nama_Kelas',
      'Status', 'Status_IDME', 'Warganegara', 'Negara_Asal', 'OKU', 'Kategori_OKU', 'Yatim',
      'Nama_Penjaga', 'Hubungan_Penjaga', 'Tel_Penjaga', 'Pendapatan_Isi_Rumah',
      'Kategori_Pendapatan', 'Tanggungan', 'Mata_Merit', 'Tag', 'Foto_ID', 'Tarikh_Import',
      'No_KP', 'Jenis_KP', 'Nama_Penjaga_1', 'Hubungan_Penjaga_1', 'Tel_Penjaga_1',
      'Nama_Penjaga_2', 'Hubungan_Penjaga_2', 'Tel_Penjaga_2', 'Bil_Isi_Rumah', 'Pendapatan_Per_Kapita',
      'Pendapatan_Penjaga_1', 'Pendapatan_Penjaga_2'],
    text: ['ID_Murid', 'Tarikh_Lahir', 'ID_Kelas', 'Tel_Penjaga', 'No_KP', 'Tel_Penjaga_1', 'Tel_Penjaga_2']
  },
  tbl_alamat_murid: {
    color: '#0891b2',
    headers: ['ID_Murid', 'Alamat_1', 'Alamat_2', 'Alamat_3', 'Poskod', 'Bandar', 'Daerah',
      'Negeri', 'Tarikh_Kemaskini', 'Lat', 'Lng', 'Sumber_Lokasi'],
    text: ['ID_Murid', 'Poskod', 'Lat', 'Lng']
  },
  tbl_kelas: {
    color: '#6366f1',
    headers: ['ID_Kelas', 'Nama_Kelas', 'Tahun', 'Sesi', 'Nama_Guru_Kelas_IDME', 'ID_Guru_Kelas',
      'Carta_Org_JSON', 'Jadual_Tugas_JSON', 'Pelan_Meja_JSON', 'Status'],
    text: ['ID_Kelas', 'ID_Guru_Kelas']
  },
  tbl_kehadiran: {
    color: '#10b981',
    headers: ['ID_Rekod', 'Tarikh', 'ID_Murid', 'ID_Kelas', 'Status', 'Sebab',
      'ID_Guru_Perekod', 'Timestamp'],
    text: ['ID_Rekod', 'Tarikh', 'ID_Murid', 'ID_Kelas']
  },
  tbl_kategori_bantuan: {
    color: '#f59e0b',
    headers: ['ID_Kategori', 'Kod', 'Nama_Bantuan', 'Pemberi', 'Kriteria', 'Status_Aktif', 'Had_Isi_Rumah', 'Had_Per_Kapita', 'Syarat_Had', 'ID_Guru_PIC'],
    text: ['ID_Kategori']
  },
  tbl_agihan_bantuan: {
    color: '#d97706',
    headers: ['ID_Agihan', 'ID_Kategori', 'ID_Murid', 'Tahun', 'Status', 'Tarikh_Serahan',
      'Amaun', 'Catatan', 'ID_Guru_Rekod', 'Timestamp'],
    text: ['ID_Agihan', 'ID_Kategori', 'ID_Murid', 'Tahun', 'Tarikh_Serahan']
  },
  tbl_kebersihan: {
    color: '#059669',
    headers: ['ID_Nilai', 'Tarikh', 'Bulan', 'Minggu_Ke', 'ID_Kelas', 'Skor_Lantai', 'Skor_Meja',
      'Skor_Sudut', 'Skor_Bacaan', 'Jumlah_Skor', 'Foto_ID', 'ID_Guru_Penilai', 'Timestamp'],
    text: ['ID_Nilai', 'Tarikh', 'Bulan', 'ID_Kelas']
  },
  tbl_kategori_disiplin: {
    color: '#be123c',
    headers: ['ID_Kategori', 'Jenis', 'Nama', 'Mata', 'Status_Aktif'],
    text: ['ID_Kategori']
  },
  tbl_disiplin: {
    color: '#e11d48',
    headers: ['ID_Kes', 'Tarikh_Masa', 'ID_Murid', 'Jenis', 'ID_Kategori', 'Kategori',
      'Mata_Kesan', 'Baki_Selepas', 'Deskripsi', 'Foto_ID', 'ID_Guru_Lapor', 'Status_Tindakan'],
    text: ['ID_Kes', 'ID_Murid', 'ID_Kategori']
  },
  tbl_log_audit: {
    color: '#475569',
    headers: ['Timestamp', 'ID_Guru', 'Tindakan', 'Butiran'],
    text: ['ID_Guru']
  },
  // Jadual waktu kelas (dimuat naik daripada XML aSc Timetables; dihurai di pelayar)
  tbl_jadual_waktu: {
    color: '#7c3aed',
    headers: ['ID_Kelas', 'Kelas_aSc', 'Jadual_JSON', 'Nama_Fail', 'ID_Guru_Muat_Naik', 'Tarikh_Kemaskini'],
    text: ['ID_Kelas', 'Kelas_aSc']
  },
  // SPBT: satu baris bagi setiap murid setahun. Rekod_JSON = { "Nama buku": 1 pinjam | 2 pulang | 3 kedua-dua }
  tbl_spbt: {
    color: '#c2410c',
    headers: ['ID_Murid', 'Tahun_Sesi', 'ID_Kelas', 'Rekod_JSON', 'ID_Guru_Rekod', 'Timestamp'],
    text: ['ID_Murid', 'Tahun_Sesi', 'ID_Kelas']
  },
  // RMT: satu baris sehari. Rekod_JSON = { idMurid: 1 makan | 0 tidak makan }; Skor_JSON = 4 kriteria borang C8 (1-4)
  tbl_rmt_harian: {
    color: '#ea580c',
    headers: ['Tarikh', 'Menu', 'ID_Guru_Bertugas', 'Bil_Penerima', 'Bil_Hadir', 'Rekod_JSON', 'Skor_JSON', 'Catatan', 'Tiada_RMT', 'ID_Guru_Rekod', 'Timestamp'],
    text: ['Tarikh', 'ID_Guru_Bertugas']
  },
  // Guru bertugas: satu baris bagi setiap minggu & sesi. Pelapor_JSON = { Isnin: idGuru, ... } (guru yang mengisi laporan hari itu)
  tbl_bertugas_minggu: {
    color: '#0d9488',
    headers: ['Minggu_Mula', 'Sesi', 'Ahli_JSON', 'ID_Ketua', 'Pelapor_JSON', 'ID_Guru_Rekod', 'Timestamp'],
    text: ['Minggu_Mula', 'Sesi', 'ID_Ketua']
  },
  // Laporan guru bertugas harian: Draf → Dihantar → Disahkan (atau Dipulangkan oleh PK HEM)
  tbl_laporan_bertugas: {
    color: '#0f766e',
    headers: ['ID_Laporan', 'Tarikh', 'Sesi', 'ID_Pelapor', 'Status', 'Data_JSON', 'Foto_ID', 'Ulasan', 'ID_Pengesah', 'Tarikh_Hantar', 'Tarikh_Sah', 'Timestamp'],
    text: ['ID_Laporan', 'Tarikh', 'Sesi', 'ID_Pelapor', 'Foto_ID', 'ID_Pengesah']
  },
  // Unit HEM: carta organisasi & skop tugas setiap unit disimpan sebagai JSON
  tbl_unit_hem: {
    color: '#9333ea',
    headers: ['ID_Unit', 'Nama', 'Singkatan', 'Ikon', 'Warna', 'ID_Ketua', 'Susunan', 'Status', 'Carta_JSON', 'Skop_JSON', 'Timestamp'],
    text: ['ID_Unit', 'ID_Ketua']
  },
  // Laporan aktiviti unit (One Page Report): Draf → Dihantar → Disahkan (atau Dipulangkan oleh ketua unit)
  tbl_opr: {
    color: '#7e22ce',
    headers: ['ID_OPR', 'ID_Unit', 'Tarikh', 'Program', 'Data_JSON', 'Foto_JSON', 'Status', 'ID_Penyedia', 'ID_Pengesah', 'Ulasan', 'Tarikh_Hantar', 'Tarikh_Sah', 'Timestamp'],
    text: ['ID_OPR', 'ID_Unit', 'Tarikh', 'ID_Penyedia', 'ID_Pengesah']
  },
  // Rekod kawalan kelas: satu baris bagi setiap check in guru (kelas, tarikh, waktu jadual)
  tbl_kawalan_kelas: {
    color: '#0e7490',
    headers: ['ID_Rekod', 'Tarikh', 'ID_Kelas', 'Waktu', 'Label_Waktu', 'Masa_Mula', 'Masa_Tamat', 'Subjek', 'Guru_Jadual', 'ID_Guru', 'Ganti', 'Masa_Masuk', 'Status_Masa', 'Tajuk', 'Catatan', 'Timestamp'],
    text: ['ID_Rekod', 'Tarikh', 'ID_Kelas', 'Label_Waktu', 'Masa_Mula', 'Masa_Tamat', 'ID_Guru', 'Masa_Masuk']
  }
};

/* ------------------------------------------------------------------
 *  PERANAN & KEBENARAN (disemak di pelayan, bukan di pelayar)
 *  Guru kelas ditentukan melalui tbl_kelas.ID_Guru_Kelas.
 * ------------------------------------------------------------------ */
const PERANAN = {
  superadmin: 'Super Admin',
  pkhem: 'PK HEM',
  gurudisiplin: 'Guru Disiplin / UBK',
  gurubertugas: 'Guru Bertugas',
  guru: 'Guru'
};

const KEBENARAN = {
  urus_tetapan: ['superadmin'],
  urus_guru: ['superadmin', 'pkhem'],
  urus_kelas_semua: ['superadmin', 'pkhem'],
  urus_bantuan: ['superadmin', 'pkhem'],
  urus_kategori: ['superadmin', 'pkhem'],
  lihat_sensitif_semua: ['superadmin', 'pkhem', 'gurudisiplin'],
  rekod_kebersihan: ['superadmin', 'pkhem', 'gurudisiplin', 'gurubertugas', 'guru'],
  surat_amaran: ['superadmin', 'pkhem', 'gurudisiplin'],
  ubah_mata_disiplin: ['superadmin', 'pkhem', 'gurudisiplin'],
  tindakan_kes: ['superadmin', 'pkhem', 'gurudisiplin'],
  urus_jadual_waktu: ['superadmin', 'pkhem'],
  urus_bertugas: ['superadmin', 'pkhem'],
  urus_carta_hem: ['superadmin', 'pkhem'],
  urus_unit_hem: ['superadmin', 'pkhem'],
  pantau_kawalan: ['superadmin', 'pkhem'],
  lihat_audit: ['superadmin', 'pkhem'],
  laporan_hem: ['superadmin', 'pkhem']
};

const STATUS_KEHADIRAN = ['Hadir', 'Lewat', 'Tidak Hadir'];
// Rekod lama (TH_Bersebab / Ponteng) dibaca sebagai "Tidak Hadir"
function statusH_(s) {
  const v = String(s || '');
  if (STATUS_KEHADIRAN.indexOf(v) >= 0) return v;
  return v === '' ? '' : 'Tidak Hadir';
}
// Lewat dikira HADIR (murid berada di sekolah) dalam peratus kehadiran
function hadirH_(s) { const v = statusH_(s); return v === 'Hadir' || v === 'Lewat'; }

/* ---- Pendapatan per kapita ----
 * Per kapita = Pendapatan_Isi_Rumah ÷ bilangan ahli isi rumah.
 * Ahli isi rumah = Tanggungan (IDME) + bilangan penjaga yang direkodkan (Penjaga 1/2) jika tetapan
 * PERKAPITA_KIRA_PENJAGA = 'Ya' (lalai), atau Tanggungan sahaja jika 'Tidak'.
 * Pulangkan null jika pendapatan atau tanggungan tidak direkodkan. */
function perKapita_(m, t) {
  const p = m.Pendapatan_Isi_Rumah;
  const tg = Math.round(num_(m.Tanggungan));
  const penjaga = [m.Nama_Penjaga_1 || m.Nama_Penjaga, m.Nama_Penjaga_2].filter(v => String(v || '').trim()).length;
  if (!(tg > 0)) return { bil: null, perKapita: null };     // tanggungan tidak direkod → tidak boleh dikira
  const bil = tg + (String(t.PERKAPITA_KIRA_PENJAGA || 'Ya') === 'Ya' ? penjaga : 0);
  if (p === '' || p === null || p === undefined || isNaN(Number(p))) return { bil: bil, perKapita: null };
  return { bil: bil, perKapita: Math.round(Number(p) / bil * 100) / 100 };
}

/* ---- Syarat bantuan ----
 * Semua jenis bantuan: warganegara sahaja. BAP & SPBT: semua murid warganegara layak tanpa syarat pendapatan. */
const SKIM_SEMUA_LAYAK = ['BAP', 'SPBT'];
function semuaLayakSkim_(k) { return !!k && SKIM_SEMUA_LAYAK.indexOf(String(k.Kod).toUpperCase()) >= 0; }
function warganegara_(m) { const v = String(m.Warganegara || '').toUpperCase(); return /WARGANEGARA|MALAYSIA/.test(v) && !/BUKAN/.test(v); }
// Boleh dikira sebagai penerima: warganegara, dan ada maklumat pendapatan kecuali skim semua layak
function layakBantuan_(m, k) { return warganegara_(m) && (semuaLayakSkim_(k) || !kosongPendapatan_(m)); }

// Kelayakan skim bantuan ikut had pendapatan isi rumah & per kapita (null = skim tiada had ditetapkan)
function layakSkim_(pendapatan, perKapita, k) {
  if (semuaLayakSkim_(k)) return true;
  const hI = Number(k.Had_Isi_Rumah) || 0, hP = Number(k.Had_Per_Kapita) || 0;
  const syarat = [];
  if (hI > 0) syarat.push(pendapatan != null && pendapatan <= hI);
  if (hP > 0) syarat.push(perKapita != null && perKapita <= hP);
  if (!syarat.length) return null;
  return String(k.Syarat_Had) === 'SALAH_SATU' ? syarat.some(Boolean) : syarat.every(Boolean);
}

/* ---- Tema paparan (sama seperti STMS) ---- */
const TEMA_SAH = ['biru', 'zamrud', 'lautan', 'lavender', 'sakura', 'senja', 'marun', 'hutan', 'grafit'];
const LATAR_SAH = ['lembut', 'aurora', 'titik', 'garis', 'polos'];

const TETAPAN_LALAI = [
  ['NAMA_SEKOLAH', 'SEKOLAH KEBANGSAAN SUNGAI TUA BAHARU', 'Nama penuh pada kepala surat'],
  ['KOD_SEKOLAH', '', 'Kod sekolah KPM'],
  ['ALAMAT_SEKOLAH', '', 'Alamat pada kepala surat'],
  ['TEL_SEKOLAH', '', 'No. telefon pada kepala surat'],
  ['NAMA_PENANDATANGAN', '', 'Nama PK HEM pada surat amaran'],
  ['JAWATAN_PENANDATANGAN', 'Penolong Kanan Hal Ehwal Murid', 'Jawatan penandatangan surat'],
  ['RUJUKAN_SURAT', 'HEM/DIS', 'Awalan no. rujukan surat amaran'],
  ['MERIT_AWAL', '100', 'Mata merit permulaan murid baharu'],
  ['MERIT_MAKS', '150', 'Had maksimum mata merit'],
  ['AMBANG_AMARAN', '70', 'Baki ≤ nilai ini = status Amaran'],
  ['AMBANG_KRITIKAL', '50', 'Baki di BAWAH nilai ini = Kritikal (layak surat amaran)'],
  ['HAD_B40', '5249', 'Pendapatan isi rumah ≤ nilai ini = B40 (semak angka DOSM terkini)'],
  ['HAD_M40', '11819', 'Pendapatan ≤ nilai ini = M40, melebihi = T20'],
  ['PERKAPITA_KIRA_PENJAGA', 'Ya', 'Ahli isi rumah untuk per kapita: Ya = tanggungan + bilangan penjaga; Tidak = tanggungan sahaja'],
  ['TEMA_LALAI', 'biru', 'Tema warna untuk skrin log masuk dan guru yang belum memilih tema'],
  ['LATAR_LALAI', 'lembut', 'Latar belakang lalai'],
  ['SESI_TAHUN', '{}', 'Sesi persekolahan ikut tahun (diurus di Tetapan Sistem)'],
  ['PENDAFTARAN_GURU', 'Kelulusan', 'Pendaftaran akaun sendiri: Kelulusan (PK HEM luluskan dahulu), Automatik, atau Tutup'],
  ['DOMAIN_EMEL', 'moe-dl.edu.my', 'Domain emel MOE untuk pendaftaran guru (cth moe-dl.edu.my)'],
  ['NAMA_APP', 'e-HEM Pintar', 'Nama aplikasi (log masuk, menu sisi, tajuk tab)'],
  ['LOGO_ID', '', 'Logo sekolah (diurus di Tetapan Asas)'],
  ['SPBT_BUKU', '', 'Senarai buku teks ikut tahun (diurus dalam tab SPBT)'],
  ['RMT_TETAPAN', '', 'Jadual guru bertugas & menu RMT (diurus dalam tab RMT)'],
  ['CARTA_HEM', '', 'Carta organisasi HEM (diurus di Dashboard)'],
  ['FOLDER_SANDARAN_ID', '', 'Folder sandaran harian (dijana automatik)'],
  ['SANDARAN_TERAKHIR', '', 'Masa sandaran terakhir (dijana automatik)'],
  ['NAIK_KELAS_TAHUN', '', 'Tahun terakhir murid dinaikkan kelas (dijana automatik)'],
  ['FOLDER_FOTO_ID', '', 'Dijana automatik — jangan ubah']
];

// Tambah kunci tetapan yang belum wujud (tanpa mengubah nilai sedia ada)
function ensureTetapan_() {
  const ada = new Set(read_('tbl_tetapan').rows.map(r => String(r.Kunci)));
  const baru = TETAPAN_LALAI.filter(r => !ada.has(r[0])).map(r => ({ Kunci: r[0], Nilai: r[1], Keterangan: r[2] }));
  if (baru.length) insert_('tbl_tetapan', baru);
}

/* ==================================================================
 *  1. MENU & SETUP
 * ================================================================== */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🎓 HEM Pintar')
    .addItem('1. Setup / Bina Jadual', 'setup')
    .addItem('2. Import Data IDME', 'importIDME')
    .addItem('3. Import Data Guru', 'importGuru')
    .addSeparator()
    .addItem('Tetapkan Semula Kata Laluan Guru', 'menuResetPin')
    .addItem('Kosongkan Cache Dashboard', 'kosongCache')
    .addItem('Arkib Kehadiran Tahun Lepas', 'arkibKehadiran')
    .addItem('Sandaran Sekarang', 'menuSandaran')
    .addToUi();
}

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone(APP.TZ);

  Object.keys(SCHEMA).forEach(name => ensureSheet_(ss, name, SCHEMA[name]));
  ensureImportSheet_(ss);

  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('SALT')) props.setProperty('SALT', Utilities.getUuid());

  const info = seedDefaults_();

  ['Sheet1', 'Helaian1'].forEach(n => {
    const sh = ss.getSheetByName(n);
    if (sh && sh.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(sh);
  });

  SpreadsheetApp.flush();
  let m = '✅ Setup HEM Pintar selesai.\n\n' + Object.keys(SCHEMA).length +
    ' jadual + helaian IDME_IMPORT sedia digunakan.';
  if (info.admin) {
    m += '\n\nAkaun Super Admin dicipta:\nEmel: ' + info.admin +
      '\nKata laluan sementara: ' + info.pin + '\n(Salin sekarang. Anda akan diminta tukar kata laluan semasa log masuk pertama.)';
  }
  m += '\n\nLangkah seterusnya: isi tbl_tetapan, kemudian import data IDME.';
  msg_(m);
}

function ensureSheet_(ss, name, def) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);

  const lastCol = sh.getLastColumn();
  let existing = lastCol ? sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String) : [];
  if (existing.every(h => h === '')) existing = [];

  const missing = def.headers.filter(h => existing.indexOf(h) === -1);
  if (missing.length) sh.getRange(1, existing.length + 1, 1, missing.length).setValues([missing]);

  const all = existing.concat(missing);
  sh.getRange(1, 1, 1, all.length)
    .setBackground(def.color).setFontColor('#ffffff')
    .setFontWeight('bold').setHorizontalAlignment('center');
  sh.setFrozenRows(1);

  const rows = Math.max(sh.getMaxRows() - 1, 1);
  (def.text || []).forEach(h => {
    const c = all.indexOf(h) + 1;
    if (c > 0) sh.getRange(2, c, rows, 1).setNumberFormat('@');
  });
  return sh;
}

function ensureImportSheet_(ss) {
  let sh = ss.getSheetByName(APP.SHEET_IMPORT);
  if (!sh) sh = ss.insertSheet(APP.SHEET_IMPORT);
  if (sh.getLastRow() === 0) {
    sh.getRange('A1').setValue(
      'Helaian ini untuk import IDME. File → Import → Upload → "Replace current sheet", ' +
      'nyahtanda "Convert text to numbers, dates, and formulas". Kemudian menu 🎓 HEM Pintar → 2. Import Data IDME.'
    ).setFontWeight('bold').setFontColor('#b45309');
  }
  sh.setTabColor('#f59e0b');
}

function seedDefaults_() {
  const info = {};

  ensureTetapan_();

  if (!read_('tbl_guru').rows.length) {
    const emel = (Session.getEffectiveUser().getEmail() || 'admin@sekolah.edu.my').toLowerCase();
    const pin = pinSementara_();
    insert_('tbl_guru', [{
      ID_Guru: 'G001', Nama_Penuh: 'Pentadbir Sistem', Emel: emel, No_Telefon: '',
      Jawatan: 'Pentadbir Sistem', Gred: '', Subjek: '', Peranan: 'superadmin', Sesi: '',
      Foto_ID: '', PIN_Hash: hash_(pin), Tukar_PIN: 'TRUE',
      Status_Aktif: 'Aktif', Tarikh_Daftar: new Date()
    }]);
    info.admin = emel; info.pin = pin;
  }

  if (!read_('tbl_kategori_bantuan').rows.length) {
    insert_('tbl_kategori_bantuan', [
      ['KB01', 'RMT', 'Rancangan Makanan Tambahan', 'KPM'],
      ['KB02', 'KWAMP', 'Kumpulan Wang Amanah Pelajar Miskin', 'KPM'],
      ['KB03', 'SPBT', 'Skim Pinjaman Buku Teks', 'KPM'],
      ['KB04', 'BAP', 'Bantuan Awal Persekolahan', 'KPM']
    ].map(r => ({ ID_Kategori: r[0], Kod: r[1], Nama_Bantuan: r[2], Pemberi: r[3], Kriteria: '', Status_Aktif: 'Aktif' })));
  }

  if (!read_('tbl_kategori_disiplin').rows.length) {
    insert_('tbl_kategori_disiplin', [
      ['DM01', 'Demerit', 'Kekemasan Diri / Pakaian Seragam', 5],
      ['DM02', 'Demerit', 'Datang Lewat Berulang Kali', 10],
      ['DM03', 'Demerit', 'Ponteng Kelas / Sekolah', 15],
      ['DM04', 'Demerit', 'Tingkah Laku Biadab terhadap Guru', 20],
      ['DM05', 'Demerit', 'Vandalisme Harta Sekolah', 25],
      ['DM06', 'Demerit', 'Buli / Pergaduhan Fizikal', 30],
      ['MR01', 'Merit', 'Membantu Guru / Rakan', 5],
      ['MR02', 'Merit', 'Memulangkan Barang Tercicir', 10],
      ['MR03', 'Merit', 'Kepimpinan / Tugas Kelas Cemerlang', 5],
      ['MR04', 'Merit', 'Mewakili Sekolah', 10]
    ].map(r => ({ ID_Kategori: r[0], Jenis: r[1], Nama: r[2], Mata: r[3], Status_Aktif: 'Aktif' })));
  }

  return info;
}

function menuResetPin() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('Tetapkan Semula Kata Laluan Guru', 'Masukkan emel guru:', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const emel = r.getResponseText().trim().toLowerCase();
  const g = read_('tbl_guru').rows.find(x => String(x.Emel).trim().toLowerCase() === emel && String(x.Status_Aktif) !== 'Dipadam');
  if (!g) return ui.alert('Emel tidak dijumpai dalam tbl_guru.');
  const pin = pinSementara_();
  g.PIN_Hash = hash_(pin);
  g.Tukar_PIN = 'TRUE';
  updateRow_('tbl_guru', g);
  batalSesi_(g.ID_Guru);
  audit_('MENU', 'RESET_PIN', emel);
  ui.alert('Kata laluan ' + g.Nama_Penuh + ' ditetapkan semula.\n\nKata laluan sementara: ' + pin +
    '\n\nBerikan kepada guru secara peribadi. Guru wajib menukarnya semasa log masuk.');
}

function kosongCache() {
  CacheService.getScriptCache().remove('dashboard');
  msg_('Cache dashboard dikosongkan.');
}

/* ---------- Arkib kehadiran: pindah rekod tahun lepas ke helaian arkib_kehadiran_YYYY ---------- */
// tbl_kehadiran kekal kecil (hanya tahun semasa) supaya setiap panggilan API kekal pantas
function arkibKehadiran() {
  denganLock_('Arkib kehadiran', arkibKehadiranCore_);
}

function arkibKehadiranCore_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tahunIni = today_().slice(0, 4);
  const tb = read_('tbl_kehadiran');
  const ikutTahun = {}, kekal = [];
  tb.rows.forEach(r => {
    const th = fmtTarikh_(r.Tarikh).slice(0, 4);
    if (/^\d{4}$/.test(th) && th < tahunIni) (ikutTahun[th] = ikutTahun[th] || []).push(r);
    else kekal.push(r);
  });
  const tahun = Object.keys(ikutTahun).sort();
  if (!tahun.length) return 'Tiada rekod kehadiran tahun lepas untuk diarkib.';

  const H = tb.headers;
  tahun.forEach(th => {
    const nama = namaArkibKehadiran_(th);
    let sh = ss.getSheetByName(nama);
    if (!sh) {
      sh = ss.insertSheet(nama);
      sh.getRange(1, 1, 1, H.length).setValues([H]).setFontWeight('bold').setBackground('#475569').setFontColor('#ffffff');
      sh.setFrozenRows(1);
    }
    const rows = ikutTahun[th], start = sh.getLastRow() + 1;
    ensureRows_(sh, start + rows.length - 1);
    applyTextFormat_(sh, 'tbl_kehadiran', H, start, rows.length);
    sh.getRange(start, 1, rows.length, H.length).setValues(rows.map(o => H.map(h => cell_(o[h]))));
  });
  writeAll_('tbl_kehadiran', kekal);
  CacheService.getScriptCache().remove('dashboard');
  audit_('MENU', 'ARKIB_KEHADIRAN', tahun.map(th => `${th}=${ikutTahun[th].length}`).join(', '));
  return '✅ Arkib selesai.\n\n' + tahun.map(th => `• ${th}: ${ikutTahun[th].length} rekod → ${namaArkibKehadiran_(th)}`).join('\n') +
    `\n\n${kekal.length} rekod tahun ${tahunIni} kekal dalam tbl_kehadiran.`;
}

function namaArkibKehadiran_(tahun) { return 'arkib_kehadiran_' + tahun; }

// Rekod kehadiran bagi satu tahun: tbl_kehadiran + helaian arkib tahun itu (jika ada)
function kehadiranTahun_(tahun) {
  const utama = read_('tbl_kehadiran').rows;
  const nama = namaArkibKehadiran_(tahun);
  if (!SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nama)) return utama;
  return utama.concat(read_(nama).rows);
}

function tahunDiarkib_(tahun) {
  return !!SpreadsheetApp.getActiveSpreadsheet().getSheetByName(namaArkibKehadiran_(tahun));
}

/* ==================================================================
 *  2. IMPORT DATA IDME
 * ================================================================== */
/* ==================================================================
 *  2b. IMPORT DATA GURU (helaian "data guru": emel, nama, no KP)
 * ================================================================== */
function importGuru() {
  let nama = 'data guru';
  try {
    const ui = SpreadsheetApp.getUi();
    const r = ui.prompt('Import Data Guru', 'Nama helaian sumber.\nBiarkan kosong untuk guna "data guru".', ui.ButtonSet.OK_CANCEL);
    if (r.getSelectedButton() !== ui.Button.OK) return;
    nama = r.getResponseText().trim() || 'data guru';
  } catch (e) { /* dijalankan dari editor */ }
  denganLock_('Import guru', () => importGuruCore_(nama));
}

// Jalankan tugas menu dengan script lock; ralat (termasuk lock tamat masa) dipaparkan dengan kemas
function denganLock_(label, fn) {
  const lock = LockService.getScriptLock();
  let dapat = false;
  try {
    lock.waitLock(30000); dapat = true;
    msg_(fn());
  } catch (e) {
    msg_('❌ ' + label + ' gagal:\n\n' + (dapat ? e.message : 'Sistem sedang sibuk. Cuba lagi sebentar.'));
  } finally {
    if (dapat) lock.releaseLock();
  }
}

function importGuruCore_(namaHelaian) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(SCHEMA).forEach(n => ensureSheet_(ss, n, SCHEMA[n]));
  const src = ss.getSheets().find(s => s.getName().trim().toLowerCase() === String(namaHelaian).trim().toLowerCase());
  if (!src) throw new Error('Helaian "' + namaHelaian + '" tidak dijumpai dalam fail ini.');
  if (src.getLastRow() < 2) throw new Error('Helaian "' + namaHelaian + '" kosong.');

  const vals = bacaNilai_(src);
  const kunci = s => String(s).toUpperCase().replace(/[^A-Z0-9]/g, '');
  const CALON = {
    emel: ['EMEL', 'EMAIL', 'EMELGURU', 'EMAILGURU', 'ALAMATEMEL', 'EMELDELIMA', 'IDDELIMA', 'EMAILADDRESS'],
    nama: ['NAMA', 'NAMAGURU', 'NAMAPENUH', 'NAMAPENUHGURU'],
    kp: ['NOKP', 'NOKADPENGENALAN', 'NOPENGENALAN', 'KADPENGENALAN', 'KP', 'NOMYKAD', 'MYKAD', 'NOIC', 'IC']
  };
  let hIdx = -1, col = {};
  for (let i = 0; i < Math.min(vals.length, 10); i++) {
    const H = vals[i].map(kunci);
    const c = {};
    Object.keys(CALON).forEach(k => { c[k] = H.findIndex(h => CALON[k].indexOf(h) >= 0); });
    if (c.emel >= 0 && c.nama >= 0) { hIdx = i; col = c; break; }
  }
  if (hIdx < 0) throw new Error('Lajur EMEL dan NAMA tidak dijumpai dalam 10 baris pertama helaian.');

  const tg = read_('tbl_guru');
  const guru = tg.rows.slice();
  const ikutEmel = new Map(guru.filter(g => String(g.Status_Aktif) !== 'Dipadam').map(g => [String(g.Emel).trim().toLowerCase(), g]));
  const ids = guru.map(g => g.ID_Guru);
  let baru = 0, kemas = 0;
  const rosak = [], ulang = [], seen = new Set(), pinBaru = [];

  vals.slice(hIdx + 1).forEach((r, i) => {
    const emel = String(r[col.emel] || '').trim().toLowerCase();
    const nama = String(r[col.nama] || '').replace(/\s+/g, ' ').trim();
    if (!emel && !nama) return;
    const baris = hIdx + 2 + i;
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(emel)) { rosak.push('baris ' + baris); return; }
    if (seen.has(emel)) { ulang.push(emel); return; }
    seen.add(emel);
    const kp = col.kp >= 0 ? noKP_(r[col.kp]) : '';
    const kpSah = /^\d{12}$/.test(kp);

    const lama = ikutEmel.get(emel);
    if (lama) {
      if (nama) lama.Nama_Penuh = nama;
      if (kpSah) lama.No_KP = kp;
      kemas++;
    } else {
      const id = nextId_(ids, 'G', 3); ids.push(id);
      const pin = pinSementara_();      // kata laluan sementara rawak (bukan No. KP yang mudah diteka)
      const g = {
        ID_Guru: id, Nama_Penuh: nama, Emel: emel, No_Telefon: '', Jawatan: 'Guru', Gred: '', Subjek: '',
        Peranan: 'guru', Sesi: '', Foto_ID: '', PIN_Hash: hash_(pin),
        Tukar_PIN: 'TRUE', Status_Aktif: 'Aktif', Tarikh_Daftar: new Date(), No_KP: kpSah ? kp : ''
      };
      guru.push(g); ikutEmel.set(emel, g); baru++;
      pinBaru.push([nama, emel, pin]);
    }
  });
  writeAll_('tbl_guru', guru);
  if (pinBaru.length) tulisPinSementara_(pinBaru);

  // Padankan guru kelas daripada nama dalam IDME (kelas yang belum ada guru kelas)
  const ikutNama = new Map(guru.map(g => [namaKunci_(g.Nama_Penuh), String(g.ID_Guru)]));
  const tk = read_('tbl_kelas');
  const kelasUbah = tk.rows.filter(k => !k.ID_Guru_Kelas && ikutNama.get(namaKunci_(k.Nama_Guru_Kelas_IDME)));
  kelasUbah.forEach(k => { k.ID_Guru_Kelas = ikutNama.get(namaKunci_(k.Nama_Guru_Kelas_IDME)); });
  updateRows_('tbl_kelas', kelasUbah);

  CacheService.getScriptCache().remove('dashboard');
  audit_('MENU', 'IMPORT_GURU', `baru=${baru}, kemas=${kemas}, kelas=${kelasUbah.length}`);

  let m = '✅ Import guru selesai.\n\n' +
    `• Guru baharu: ${baru}\n• Guru dikemas kini: ${kemas}\n• Guru kelas dipadankan automatik: ${kelasUbah.length} kelas`;
  if (rosak.length) m += `\n• ⚠️ Emel tidak sah, dilangkau: ${rosak.slice(0, 8).join(', ')}${rosak.length > 8 ? ', ...' : ''}`;
  if (ulang.length) m += `\n• ⚠️ Emel berulang, dilangkau: ${Array.from(new Set(ulang)).slice(0, 5).join(', ')}`;
  if (pinBaru.length) {
    m += `\n\nKata laluan sementara (rawak) bagi ${pinBaru.length} guru baharu ditulis dalam helaian "${APP.SHEET_PIN_SEMENTARA}". ` +
      'Edarkan kepada setiap guru secara peribadi, kemudian PADAM helaian tersebut. ' +
      'Sistem akan memaksa mereka menukar kata laluan semasa log masuk pertama.';
  }
  m += '\nKata laluan guru sedia ada tidak diubah.';
  return m;
}

// Senarai kata laluan sementara untuk diedarkan oleh pentadbir (helaian dilindungi, padam selepas guna)
function tulisPinSementara_(baris) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(APP.SHEET_PIN_SEMENTARA);
  if (!sh) {
    sh = ss.insertSheet(APP.SHEET_PIN_SEMENTARA);
    sh.getRange(1, 1, 1, 4).setValues([['Nama', 'Emel', 'Kata_Laluan_Sementara', 'Dijana']])
      .setFontWeight('bold').setBackground('#b91c1c').setFontColor('#ffffff');
    sh.setTabColor('#b91c1c');
    try {
      const pr = sh.protect().setDescription('Sulit: kata laluan sementara');
      pr.removeEditors(pr.getEditors()); pr.addEditor(Session.getEffectiveUser());
    } catch (e) { /* perlindungan tidak tersedia */ }
  }
  const stamp = Utilities.formatDate(new Date(), APP.TZ, 'yyyy-MM-dd HH:mm');
  const start = sh.getLastRow() + 1;
  sh.getRange(start, 1, baris.length, 4).setNumberFormat('@')
    .setValues(baris.map(r => [r[0], r[1], r[2], stamp]));
}

// Baca helaian: nombor panjang & tarikh dibaca ikut nilai sebenar, bukan paparan sel
function bacaNilai_(sh) {
  const rng = sh.getDataRange();
  const raw = rng.getValues(), disp = rng.getDisplayValues();
  return raw.map((r, i) => r.map((v, j) => {
    if (v instanceof Date) return Utilities.formatDate(v, APP.TZ, 'yyyy-MM-dd');
    if (typeof v === 'number' && Number.isInteger(v) && Math.abs(v) >= 1e6) return String(v);
    return disp[i][j];
  }));
}

// Kunci padanan nama: abaikan huruf besar/kecil, ruang dan tanda baca
function namaKunci_(n) { return String(n || '').toUpperCase().replace(/[^A-Z0-9]/g, ''); }

function importIDME() {
  let nama = APP.SHEET_IMPORT;
  try {
    const ui = SpreadsheetApp.getUi();
    const r = ui.prompt('Import Data IDME',
      'Nama helaian sumber (contoh: data murid).\nBiarkan kosong untuk guna "' + APP.SHEET_IMPORT + '".',
      ui.ButtonSet.OK_CANCEL);
    if (r.getSelectedButton() !== ui.Button.OK) return;
    nama = r.getResponseText().trim() || APP.SHEET_IMPORT;
  } catch (e) { /* dijalankan dari editor tanpa UI */ }
  denganLock_('Import', () => importIDMECore_(nama));
}

function importIDMECore_(namaHelaian) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const nama = String(namaHelaian || APP.SHEET_IMPORT).trim();
  const src = ss.getSheets().find(s => s.getName().trim().toLowerCase() === nama.toLowerCase());
  if (!src) throw new Error('Helaian "' + nama + '" tidak dijumpai dalam fail ini.');
  Object.keys(SCHEMA).forEach(n => ensureSheet_(ss, n, SCHEMA[n]));   // tambah lajur baharu jika belum ada
  if (src.getLastRow() < 2) throw new Error('Helaian "' + nama + '" kosong.');

  // Baca nilai sebenar: nombor panjang (ID/telefon) & tarikh tidak bergantung pada paparan sel
  return prosesImportIDME_(bacaNilai_(src));
}

// Teras import murid: menerima jadual (baris × lajur) daripada helaian Google Sheet atau fail Excel
function prosesImportIDME_(vals) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(SCHEMA).forEach(n => ensureSheet_(ss, n, SCHEMA[n]));
  const norm = s => String(s).toUpperCase().replace(/\s+/g, ' ').trim();

  // Cari baris tajuk (IDME kadang-kadang ada baris judul di atas jadual)
  let hIdx = -1;
  for (let i = 0; i < Math.min(vals.length, 20); i++) {
    if (vals[i].some(c => norm(c) === 'ID MURID')) { hIdx = i; break; }
  }
  if (hIdx < 0) throw new Error('Lajur "ID MURID" tidak dijumpai dalam 20 baris pertama.');

  const H = vals[hIdx].map(norm);
  const col = name => H.indexOf(name);
  const wajib = ['ID MURID', 'NAMA', 'NAMA KELAS', 'JANTINA'];
  const tiada = wajib.filter(n => col(n) < 0);
  if (tiada.length) throw new Error('Lajur wajib tiada: ' + tiada.join(', '));

  const g = (row, name) => { const i = col(name); return i < 0 ? '' : String(row[i] || '').trim(); };

  const body = vals.slice(hIdx + 1)
    .map((r, i) => ({ r, baris: hIdx + 2 + i }))
    .filter(x => g(x.r, 'ID MURID') !== '');
  if (!body.length) throw new Error('Tiada baris murid dijumpai di bawah baris tajuk.');

  // ---- Semakan data rosak (kesan Excel) & ID berulang ----
  const rosak = [], ulang = [], seen = {};
  body.forEach(({ r, baris }) => {
    const id = g(r, 'ID MURID');
    if (/e\+/i.test(id) || !/^[0-9A-Za-z\-]+$/.test(id)) rosak.push('baris ' + baris);
    if (seen[id]) ulang.push(id);
    seen[id] = true;
  });
  if (rosak.length) {
    throw new Error('ID MURID rosak (notasi saintifik seperti 2.51E+11) pada ' + rosak.length +
      ' baris, contohnya ' + rosak.slice(0, 5).join(', ') + '.\n\n' +
      'Muat turun semula dari IDME, JANGAN buka dalam Excel, dan import dengan pilihan ' +
      '"Convert text to numbers, dates, and formulas" dinyahtanda.');
  }
  if (ulang.length) {
    throw new Error('ID MURID berulang: ' + Array.from(new Set(ulang)).slice(0, 5).join(', '));
  }

  const t = tetapan_();
  const merit0 = Number(t.MERIT_AWAL) || 100;
  const hadB40 = Number(t.HAD_B40) || 5249;
  const hadM40 = Number(t.HAD_M40) || 11819;
  const stamp = Utilities.formatDate(new Date(), APP.TZ, 'yyyy-MM-dd HH:mm');

  // ---- Murid ----
  const tm = read_('tbl_murid');
  const murid = new Map();
  tm.rows.forEach(o => murid.set(String(o.ID_Murid), o));

  const ta = read_('tbl_alamat_murid');
  const alamat = new Map();
  ta.rows.forEach(o => alamat.set(String(o.ID_Murid), o));

  const kelasIDME = new Map();
  const dalamImport = new Set();
  let baru = 0, kemas = 0, tanpaTel = 0;

  body.forEach(({ r }) => {
    const id = g(r, 'ID MURID');
    dalamImport.add(id);

    // Kelas dikenal pasti ikut TAHUN + NAMA KELAS (IDME guna nama sama, cth "ARIF", untuk setiap tahun)
    const tahun = g(r, 'TAHUN / TINGKATAN') || g(r, 'TAHUN/TINGKATAN') || g(r, 'TAHUN');
    const namaKelas = namaKelasPaparan_(tahun, g(r, 'NAMA KELAS'));
    const idKelas = kelasId_(tahun, namaKelas);
    if (!kelasIDME.has(idKelas)) {
      kelasIDME.set(idKelas, { nama: namaKelas, tahun: tahun, guru: g(r, 'NAMA GURU KELAS') });
    }

    const s1 = g(r, 'PENDAPATAN PENJAGA 1'), s2 = g(r, 'PENDAPATAN PENJAGA 2');
    const adaPendapatan = s1 !== '' || s2 !== '';
    const jumlah = num_(s1) + num_(s2);
    const kategori = !adaPendapatan ? 'Tiada Maklumat'
      : jumlah <= hadB40 ? 'B40' : jumlah <= hadM40 ? 'M40' : 'T20';

    const tel1 = normTel_(g(r, 'NO. TEL. BIMBIT PENJAGA 1')) || normTel_(g(r, 'NO. TEL. PEJABAT PENJAGA 1'));
    const tel2 = normTel_(g(r, 'NO. TEL. BIMBIT PENJAGA 2')) || normTel_(g(r, 'NO. TEL. PEJABAT PENJAGA 2'));
    const tel = tel1 || tel2;
    if (!tel) tanpaTel++;
    const guna1 = g(r, 'PENJAGA 1') !== '';
    const statusIDME = g(r, 'STATUS PENGAJIAN');

    const data = {
      ID_Murid: id,
      Nama: g(r, 'NAMA'),
      Jantina: jantina_(g(r, 'JANTINA')),
      Tarikh_Lahir: tarikh_(g(r, 'TARIKH LAHIR')),
      Tahun: tahun,
      ID_Kelas: idKelas,
      Nama_Kelas: namaKelas,
      Status: (!statusIDME || norm(statusIDME) === 'BERSEKOLAH') ? 'Aktif' : 'Tidak Aktif',
      Status_IDME: statusIDME,
      Warganegara: g(r, 'WARGANEGARA'),
      Negara_Asal: g(r, 'NEGARA ASAL'),
      OKU: g(r, 'STATUS OKU'),
      Kategori_OKU: g(r, 'KATEGORI KETIDAKUPAYAAN'),
      Yatim: g(r, 'STATUS YATIM'),
      Nama_Penjaga: guna1 ? g(r, 'PENJAGA 1') : g(r, 'PENJAGA 2'),
      Hubungan_Penjaga: guna1 ? g(r, 'HUBUNGAN PENJAGA 1') : g(r, 'HUBUNGAN PENJAGA 2'),
      Tel_Penjaga: tel,
      No_KP: noKP_(g(r, 'NO. PENGENALAN')),
      Jenis_KP: g(r, 'JENIS PENGENALAN'),
      Nama_Penjaga_1: g(r, 'PENJAGA 1'), Hubungan_Penjaga_1: g(r, 'HUBUNGAN PENJAGA 1'), Tel_Penjaga_1: tel1,
      Nama_Penjaga_2: g(r, 'PENJAGA 2'), Hubungan_Penjaga_2: g(r, 'HUBUNGAN PENJAGA 2'), Tel_Penjaga_2: tel2,
      Pendapatan_Isi_Rumah: adaPendapatan ? jumlah : '',
      Pendapatan_Penjaga_1: s1 !== '' ? num_(s1) : '', Pendapatan_Penjaga_2: s2 !== '' ? num_(s2) : '',
      Kategori_Pendapatan: kategori,
      Tanggungan: g(r, 'TANGGUNGAN'),
      Tarikh_Import: stamp
    };
    const pk = perKapita_(data, t);       // simpan juga dalam helaian untuk rujukan
    data.Bil_Isi_Rumah = pk.bil == null ? '' : pk.bil;
    data.Pendapatan_Per_Kapita = pk.perKapita == null ? '' : pk.perKapita;

    const lama = murid.get(id);
    if (lama) {
      Object.assign(lama, data);          // Mata_Merit, Tag, Foto_ID kekal
      kemas++;
    } else {
      murid.set(id, Object.assign({ Mata_Merit: merit0, Tag: '', Foto_ID: '' }, data));
      baru++;
    }

    // ---- Alamat ----
    const a = {
      ID_Murid: id,
      Alamat_1: bersih_(g(r, 'ALAMAT 1')),
      Alamat_2: bersih_(g(r, 'ALAMAT 2')),
      Alamat_3: bersih_(g(r, 'ALAMAT 3')),
      Poskod: poskod_(g(r, 'POSKOD')),
      Bandar: bersih_(g(r, 'BANDAR')),
      Daerah: bersih_(g(r, 'DAERAH')),
      Negeri: bersih_(g(r, 'NEGERI')),
      Tarikh_Kemaskini: stamp
    };
    if (col('ALAMAT 1') >= 0) {
      const lamaA = alamat.get(id);
      // Alamat berubah → lokasi auto dikira semula; lokasi yang ditetapkan guru (MANUAL) dikekalkan
      if (lamaA && teksAlamat_(lamaA) !== teksAlamat_(a) && String(lamaA.Sumber_Lokasi) !== 'MANUAL') {
        a.Lat = ''; a.Lng = ''; a.Sumber_Lokasi = '';
      }
      alamat.set(id, Object.assign(lamaA || {}, a));
    }
  });

  writeAll_('tbl_murid', Array.from(murid.values()));
  writeAll_('tbl_alamat_murid', Array.from(alamat.values()));

  // ---- Kelas (auto-bina, Sesi & data guru kelas dikekalkan) ----
  const tk = read_('tbl_kelas');
  const kelas = new Map();
  tk.rows.forEach(o => kelas.set(String(o.ID_Kelas), o));
  const guruNama = new Map();
  read_('tbl_guru').rows.forEach(x => guruNama.set(namaKunci_(x.Nama_Penuh), String(x.ID_Guru)));

  let kelasBaru = 0;
  const sesiTahun = petaSesiTahun_();
  kelasIDME.forEach((k, idKelas) => {
    const lama = kelas.get(idKelas);
    const idGuru = guruNama.get(namaKunci_(k.guru)) || '';
    if (lama) {
      lama.Nama_Kelas = k.nama;
      lama.Tahun = k.tahun;
      lama.Nama_Guru_Kelas_IDME = k.guru;
      lama.Status = 'Aktif';
      if (!lama.ID_Guru_Kelas) lama.ID_Guru_Kelas = idGuru;
      if (!lama.Sesi) lama.Sesi = sesiTahun[namaKunci_(k.tahun)] || '';
    } else {
      kelas.set(idKelas, {
        ID_Kelas: idKelas, Nama_Kelas: k.nama, Tahun: k.tahun, Sesi: sesiTahun[namaKunci_(k.tahun)] || '',
        Nama_Guru_Kelas_IDME: k.guru, ID_Guru_Kelas: idGuru,
        Carta_Org_JSON: '', Jadual_Tugas_JSON: '', Pelan_Meja_JSON: '', Status: 'Aktif'
      });
      kelasBaru++;
    }
  });
  // Nyahaktif kelas lama yang tiada lagi murid aktif (cth kelas bercantum daripada import versi lama)
  const bilMurid = {};
  murid.forEach(o => { if (String(o.Status) === 'Aktif') bilMurid[String(o.ID_Kelas)] = (bilMurid[String(o.ID_Kelas)] || 0) + 1; });
  let kelasTutup = 0;
  kelas.forEach((k, idKelas) => {
    if (!kelasIDME.has(idKelas) && !bilMurid[idKelas] && String(k.Status) !== 'Tidak Aktif') { k.Status = 'Tidak Aktif'; kelasTutup++; }
  });
  const senaraiKelas = Array.from(kelas.values())
    .sort((a, b) => tahunSort_(a.Tahun) - tahunSort_(b.Tahun) || String(a.Nama_Kelas).localeCompare(String(b.Nama_Kelas)));
  writeAll_('tbl_kelas', senaraiKelas);

  const kelasAktif = senaraiKelas.filter(k => String(k.Status) !== 'Tidak Aktif');
  const tanpaSesi = kelasAktif.filter(k => !k.Sesi).map(k => k.Nama_Kelas);
  const bilTahun = new Set(kelasAktif.map(k => String(k.Tahun)).filter(Boolean)).size;
  const tiadaDalamImport = tm.rows.filter(o => !dalamImport.has(String(o.ID_Murid))).length;

  CacheService.getScriptCache().remove('dashboard');
  audit_('MENU', 'IMPORT_IDME', `baru=${baru}, kemas=${kemas}, kelasBaru=${kelasBaru}`);

  let m = '✅ Import IDME selesai.\n\n' +
    `• Murid baharu: ${baru}\n• Murid dikemas kini: ${kemas}\n` +
    `• Kelas baharu: ${kelasBaru} (jumlah ${kelasAktif.length} kelas aktif dalam ${bilTahun} tahun)`;
  if (kelasTutup) m += `\n• ${kelasTutup} kelas lama tanpa murid dinyahaktifkan`;
  if (tanpaTel) m += `\n• ⚠️ ${tanpaTel} murid tiada no. telefon penjaga`;
  if (tiadaDalamImport) {
    m += `\n• ℹ️ ${tiadaDalamImport} murid dalam sistem tiada dalam fail ini (tidak diubah). ` +
      'Jika mereka sudah berpindah, status IDME akan dikemas kini pada import penuh seterusnya.';
  }
  if (tanpaSesi.length) {
    m += `\n\n⚠️ Sila tetapkan lajur Sesi (Pagi/Petang) dalam tbl_kelas untuk ${tanpaSesi.length} kelas:\n` +
      tanpaSesi.slice(0, 15).join(', ') + (tanpaSesi.length > 15 ? ', ...' : '');
  }
  return m;
}

/* ==================================================================
 *  3. WEB APP API
 * ================================================================== */
// [fungsi, adakah menulis data (perlu lock)]
const API = {
  bootstrap: [apiBootstrap_, false],
  tukarPin: [apiTukarPin_, true],
  logout: [apiLogout_, false],
  senaraiMurid: [apiSenaraiMurid_, false],
  dashboard: [apiDashboard_, false],

  getKelas: [apiGetKelas_, false],
  simpanKelas: [apiSimpanKelas_, true],

  getKehadiran: [apiGetKehadiran_, false],
  simpanKehadiran: [apiSimpanKehadiran_, true],

  simpanDisiplin: [apiSimpanDisiplin_, true],
  senaraiDisiplin: [apiSenaraiDisiplin_, false],
  tindakanKes: [apiTindakanKes_, true],
  suratAmaran: [apiSuratAmaran_, false],
  getAlamat: [apiGetAlamat_, false],

  simpanKebersihan: [apiSimpanKebersihan_, true],
  kedudukanKebersihan: [apiKedudukanKebersihan_, false],
  senaraiKebersihan: [apiSenaraiKebersihan_, false],

  senaraiAgihan: [apiSenaraiAgihan_, false],
  simpanAgihan: [apiSimpanAgihan_, true],
  tambahKategoriBantuan: [apiTambahKategoriBantuan_, true],
  simpanLogo: [apiSimpanLogo_, true],
  daftarAgihanPukal: [apiDaftarAgihanPukal_, true],
  calonSkim: [apiCalonSkim_, false],
  getBap: [apiGetBap_, false],
  simpanBap: [apiSimpanBap_, true],
  getRmt: [apiGetRmt_, false],
  simpanRmt: [apiSimpanRmt_, true],
  julatRmt: [apiJulatRmt_, false],
  simpanTetapanRmt: [apiSimpanTetapanRmt_, true],
  getBertugas: [apiGetBertugas_, false],
  simpanBertugas: [apiSimpanBertugas_, true],
  getLaporanBertugas: [apiGetLaporanBertugas_, false],
  simpanLaporanBertugas: [apiSimpanLaporanBertugas_, true],
  sahLaporanBertugas: [apiSahLaporanBertugas_, true],
  senaraiLaporanBertugas: [apiSenaraiLaporanBertugas_, false],
  getCartaHem: [apiGetCartaHem_, false],
  simpanCartaHem: [apiSimpanCartaHem_, true],
  senaraiUnitHem: [apiSenaraiUnitHem_, false],
  getUnitHem: [apiGetUnitHem_, false],
  simpanUnitHem: [apiSimpanUnitHem_, true],
  susunUnitHem: [apiSusunUnitHem_, true],
  simpanCartaUnit: [apiSimpanCartaUnit_, true],
  simpanSkopUnit: [apiSimpanSkopUnit_, true],
  senaraiOpr: [apiSenaraiOpr_, false],
  simpanOpr: [apiSimpanOpr_, true],
  sahOpr: [apiSahOpr_, true],
  padamOpr: [apiPadamOpr_, true],
  getKawalan: [apiGetKawalan_, false],
  kawalanMinggu: [apiKawalanMinggu_, false],
  jadualSaya: [apiJadualSaya_, false],
  checkIn: [apiCheckIn_, true],
  kemaskiniKawalan: [apiKemaskiniKawalan_, true],
  padamKawalan: [apiPadamKawalan_, true],
  pantauKawalan: [apiPantauKawalan_, false],
  infoImbas: [apiInfoImbas_, false],
  kenalanMurid: [apiKenalanMurid_, false],
  rmtImbas: [apiRmtImbas_, true],
  statusSandaran: [apiStatusSandaran_, false],
  sandaranSekarang: [apiSandaranSekarang_, true],
  pasangSandaran: [apiPasangSandaran_, true],
  arkibTahunLepas: [apiArkibTahunLepas_, true],
  senaraiAudit: [apiSenaraiAudit_, false],
  laporanBulanan: [apiLaporanBulanan_, false],
  muridBerisiko: [apiMuridBerisiko_, false],
  naikKelas: [apiNaikKelas_, true],

  simpanGuru: [apiSimpanGuru_, true],
  resetPinGuru: [apiResetPinGuru_, true],
  padamGuru: [apiPadamGuru_, true],
  senaraiPengguna: [apiSenaraiPengguna_, false],

  simpanTetapan: [apiSimpanTetapan_, true],

  getTetapan: [apiGetTetapan_, false],
  simpanSesiTahun: [apiSimpanSesiTahun_, true],
  importMuridExcel: [apiImportMuridExcel_, true],
  senaraiPendaftaran: [apiSenaraiPendaftaran_, false],
  putusPendaftaran: [apiPutusPendaftaran_, true],
  kemaskiniMurid: [apiKemaskiniMurid_, true],
  fotoMurid: [apiFotoMurid_, true],

  muatNaikFoto: [apiMuatNaikFoto_, true],
  lihatFoto: [apiLihatFoto_, false],
  lihatFotoPukal: [apiLihatFotoPukal_, false],

  profilMurid: [apiProfilMurid_, false],     // lock diambil sendiri hanya semasa menyimpan koordinat
  simpanLokasi: [apiSimpanLokasi_, true],
  simpanKelasJagaan: [apiSimpanKelasJagaan_, true],
  senaraiNamaKP: [apiSenaraiNamaKP_, false],
  analisisKehadiran: [apiAnalisisKehadiran_, false],
  simpanTema: [apiSimpanTema_, true],
  simpanHeader: [apiSimpanHeader_, true],
  kehadiranBulanan: [apiKehadiranBulanan_, false],
  kehadiranMurid: [apiKehadiranMurid_, false],
  dataMuridKelas: [apiDataMuridKelas_, false],
  getSpbt: [apiGetSpbt_, false],
  simpanSpbt: [apiSimpanSpbt_, true],
  simpanBukuSpbt: [apiSimpanBukuSpbt_, true],

  getJadualWaktu: [apiGetJadualWaktu_, false],
  simpanJadualWaktu: [apiSimpanJadualWaktu_, true],
  padamJadualWaktu: [apiPadamJadualWaktu_, true],
  senaraiJadualWaktu: [apiSenaraiJadualWaktu_, false]
};

function doGet(e) {
  const kod = e && e.parameter && e.parameter.sahkan;
  if (kod) return sahkanEmel_(String(kod));
  return out_({ ok: true, app: APP.NAMA, versi: APP.VERSI, masa: new Date().toISOString() });
}

// Pautan pengesahan emel (pendaftaran mod Automatik): aktifkan akaun selepas guru membuka pautan dalam emel MOE
function sahkanEmel_(kod) {
  const halaman = (tajuk, teks) => HtmlService.createHtmlOutput(
    '<meta name="viewport" content="width=device-width,initial-scale=1">' +
    '<div style="font-family:system-ui,sans-serif;max-width:480px;margin:48px auto;padding:0 16px;text-align:center">' +
    '<h2>' + escHtml_(tajuk) + '</h2><p>' + escHtml_(teks) + '</p></div>').setTitle(APP.NAMA);
  if (!/^[a-f0-9]{32}$/.test(kod)) return halaman('Pautan tidak sah', 'Pautan pengesahan tidak sah.');
  const cache = CacheService.getScriptCache();
  const idGuru = cache.get('sahkan_' + kod);
  if (!idGuru) return halaman('Pautan tamat tempoh', 'Pautan ini telah tamat tempoh. Akaun anda masih boleh diluluskan oleh PK HEM.');
  const lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch (err) { return halaman('Sistem sibuk', 'Sila buka semula pautan ini sebentar lagi.'); }
  try {
    const g = read_('tbl_guru').rows.find(r => String(r.ID_Guru) === idGuru && String(r.Status_Aktif) === 'Menunggu');
    if (!g) return halaman('Sudah diproses', 'Akaun ini sudah aktif atau telah diproses.');
    g.Status_Aktif = 'Aktif';
    updateRow_('tbl_guru', g);
    const diberi = beriKelas_(String(g.ID_Guru), parseJSON_(g.Kelas_Mohon, []).map(String));
    cache.remove('sahkan_' + kod);
    audit_(String(g.ID_Guru), 'SAHKAN_EMEL', `kelas=${diberi.join(',')}`);
    return halaman('Emel disahkan', 'Akaun anda kini aktif. Anda boleh log masuk ke ' + APP.NAMA + '.');
  } finally {
    lock.releaseLock();
  }
}

function escHtml_(s) {
  return String(s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

function doPost(e) {
  let req;
  try {
    req = JSON.parse(e.postData.contents);
  } catch (err) {
    return out_({ ok: false, ralat: 'Permintaan tidak sah.' });
  }

  const action = String(req.action || '');
  const p = req.payload || {};
  let lock = null;

  try {
    if (action === 'login') return out_({ ok: true, data: login_(p) });
    if (action === 'infoDaftar') return out_({ ok: true, data: infoDaftar_() });
    if (action === 'infoAwam') return out_({ ok: true, data: infoAwam_(p) });
    if (action === 'lupaPin') return out_({ ok: true, data: lupaPin_(p) });
    if (action === 'resetPinKod') {
      lock = LockService.getScriptLock(); lock.waitLock(20000);
      return out_({ ok: true, data: resetPinKod_(p) });
    }
    if (action === 'daftarGuru') {
      lock = LockService.getScriptLock(); lock.waitLock(20000);
      return out_({ ok: true, data: daftarGuru_(p) });
    }

    const user = auth_(req.token);
    const entry = API[action];
    if (!entry) throw err_('Tindakan tidak dikenali: ' + action);
    // Kata laluan sementara: hanya tukar kata laluan / log keluar dibenarkan (disemak di pelayan)
    if (user.tukarPin && action !== 'tukarPin' && action !== 'logout') {
      throw err_('Sila tukar kata laluan sementara anda dahulu.', 'TUKAR_PIN');
    }

    if (entry[1]) {
      lock = LockService.getScriptLock();
      // Waktu sibuk (cth kehadiran pagi): tunggu hingga 20 saat; jika masih sibuk, frontend cuba semula sendiri
      if (!lock.tryLock(20000)) { lock = null; throw err_('Sistem sedang sibuk. Sila cuba sebentar lagi.', 'SIBUK'); }
    }
    const data = entry[0](p, user, req.token);
    return out_({ ok: true, data: data });
  } catch (err) {
    return out_({ ok: false, ralat: err.message, kod: err.kod || '' });
  } finally {
    if (lock) lock.releaseLock();
  }
}

/* ---------- Log masuk & sesi ---------- */
function login_(p) {
  const emel = String(p.emel || '').trim().toLowerCase();
  const pin = String(p.pin || '');
  if (!emel || !pin) throw err_('Sila isi emel dan kata laluan.');

  const cache = CacheService.getScriptCache();
  const kCuba = 'cuba_' + emel;
  const cuba = Number(cache.get(kCuba) || 0);
  if (cuba >= 5) throw err_('Terlalu banyak cubaan. Cuba lagi selepas 10 minit.');

  const g0 = read_('tbl_guru').rows.find(r => String(r.Emel).trim().toLowerCase() === emel && String(r.Status_Aktif) !== 'Dipadam');
  if (g0 && String(g0.Status_Aktif) === 'Menunggu' && semakPin_(g0.PIN_Hash, pin)) {
    throw err_('Akaun anda sedang menunggu pengesahan emel atau kelulusan PK HEM. Sila semak emel MOE anda.');
  }
  const g = g0 && String(g0.Status_Aktif) === 'Aktif' ? g0 : null;
  if (!g || !semakPin_(g.PIN_Hash, pin)) {
    cache.put(kCuba, String(cuba + 1), 600);
    throw err_('Emel atau kata laluan salah.');
  }
  cache.remove(kCuba);
  // Naik taraf hash format lama (satu pusingan, salt dikongsi) kepada format v2 secara senyap
  if (String(g.PIN_Hash).indexOf('v2$') !== 0) {
    const lock = LockService.getScriptLock();
    if (lock.tryLock(5000)) {
      try { g.PIN_Hash = hash_(pin); updateRow_('tbl_guru', g); } finally { lock.releaseLock(); }
    }
  }

  const user = {
    id: String(g.ID_Guru), nama: String(g.Nama_Penuh), emel: emel,
    peranan: PERANAN[g.Peranan] ? String(g.Peranan) : 'guru',
    // Paksa tukar jika masih kata laluan sementara ATAU kata laluan lama kurang 8 aksara
    tukarPin: isTrue_(g.Tukar_PIN) || pin.length < APP.MIN_KATA_LALUAN
  };
  user.masuk = Date.now();
  const token = Utilities.getUuid();
  cache.put('tok_' + token, JSON.stringify(user), APP.TOKEN_SAAT);
  audit_(user.id, 'LOGIN', '');
  return { token: token, user: user };
}

function auth_(token) {
  if (!token) throw err_('Sila log masuk.', 'AUTH');
  const cache = CacheService.getScriptCache();
  const c = cache.get('tok_' + token);
  if (!c) throw err_('Sesi tamat. Sila log masuk semula.', 'AUTH');
  const user = JSON.parse(c);
  const rev = Number(cache.get('rev_' + user.id) || 0);
  if (rev && rev > Number(user.masuk || 0)) {
    cache.remove('tok_' + token);
    throw err_('Akaun anda telah dikemas kini oleh pentadbir. Sila log masuk semula.', 'AUTH');
  }
  cache.put('tok_' + token, c, APP.TOKEN_SAAT);   // lanjutkan sesi
  return user;
}

function apiLogout_(p, u, token) {
  CacheService.getScriptCache().remove('tok_' + token);
  return { keluar: true };
}

function apiTukarPin_(p, u, token) {
  const lama = String(p.pinLama || ''), baru = String(p.pinBaru || '');
  if (baru.length < APP.MIN_KATA_LALUAN) throw err_('Kata laluan baharu mesti sekurang-kurangnya ' + APP.MIN_KATA_LALUAN + ' aksara.');
  if (baru.length > 64) throw err_('Kata laluan terlalu panjang (maksimum 64 aksara).');
  if (baru === lama) throw err_('Kata laluan baharu mesti berbeza daripada kata laluan semasa.');
  const g = read_('tbl_guru').rows.find(r => String(r.ID_Guru) === u.id);
  if (!g || !semakPin_(g.PIN_Hash, lama)) throw err_('Kata laluan semasa salah.');
  if (String(g.No_KP || '') && baru === String(g.No_KP)) throw err_('Kata laluan tidak boleh sama dengan No. KP.');
  g.PIN_Hash = hash_(baru);
  g.Tukar_PIN = 'FALSE';
  updateRow_('tbl_guru', g);
  u.tukarPin = false;
  CacheService.getScriptCache().put('tok_' + token, JSON.stringify(u), APP.TOKEN_SAAT);
  audit_(u.id, 'TUKAR_PIN', '');
  return { berjaya: true };
}

/* ---------- Kebenaran ---------- */
function boleh_(u, k) { return (KEBENARAN[k] || []).indexOf(u.peranan) >= 0; }
function wajib_(u, k) { if (!boleh_(u, k)) throw err_('Anda tiada kebenaran untuk tindakan ini.', 'FORBIDDEN'); }
function kelasSaya_(u) {
  return read_('tbl_kelas').rows.filter(k => String(k.ID_Guru_Kelas) === u.id).map(k => String(k.ID_Kelas));
}
function bolehSensitif_(u, idKelas, kelasSaya) {
  if (boleh_(u, 'lihat_sensitif_semua')) return true;
  return (kelasSaya || kelasSaya_(u)).indexOf(String(idKelas)) >= 0;
}

/* ---------- Bootstrap (data awal frontend) ---------- */
function apiBootstrap_(p, u) {
  const t = tetapan_();
  const kelasSaya = kelasSaya_(u);
  const guru = read_('tbl_guru').rows;
  const namaGuru = new Map(guru.map(g => [String(g.ID_Guru), String(g.Nama_Penuh)]));

  return {
    user: Object.assign({}, u, (() => {
      const me = guru.find(g => String(g.ID_Guru) === u.id) || {};
      return {
        labelPeranan: PERANAN[u.peranan], kelasSaya: kelasSaya,
        tema: TEMA_SAH.indexOf(String(me.Tema)) >= 0 ? String(me.Tema) : '',
        latar: LATAR_SAH.indexOf(String(me.Latar)) >= 0 ? String(me.Latar) : '',
        headerId: String(me.Header_ID || ''), headerUrl: String(me.Header_URL || '')
      };
    })()),
    tetapan: {
      namaApp: t.NAMA_APP || 'e-HEM Pintar', logoVersi: t.LOGO_ID || '',
      namaSekolah: t.NAMA_SEKOLAH || '', kodSekolah: t.KOD_SEKOLAH || '',
      alamatSekolah: t.ALAMAT_SEKOLAH || '', telSekolah: t.TEL_SEKOLAH || '',
      meritAwal: Number(t.MERIT_AWAL) || 100, meritMaks: Number(t.MERIT_MAKS) || 150,
      ambangAmaran: Number(t.AMBANG_AMARAN) || 70, ambangKritikal: Number(t.AMBANG_KRITIKAL) || 50,
      hariIni: today_(), perkapitaPenjaga: t.PERKAPITA_KIRA_PENJAGA !== 'Tidak',
      temaLalai: TEMA_SAH.indexOf(t.TEMA_LALAI) >= 0 ? t.TEMA_LALAI : 'biru',
      latarLalai: LATAR_SAH.indexOf(t.LATAR_LALAI) >= 0 ? t.LATAR_LALAI : 'lembut'
    },
    kelas: read_('tbl_kelas').rows.filter(k => String(k.Status) !== 'Tidak Aktif').map(k => ({
      id: String(k.ID_Kelas), nama: String(k.Nama_Kelas), tahun: String(k.Tahun), sesi: String(k.Sesi),
      idGuruKelas: String(k.ID_Guru_Kelas),
      namaGuruKelas: namaGuru.get(String(k.ID_Guru_Kelas)) || String(k.Nama_Guru_Kelas_IDME || '')
    })),
    murid: senaraiMurid_(u, kelasSaya, t),
    // Guru biasa tidak menerima emel/telefon guru lain (hanya pentadbir)
    guru: guru.filter(g => String(g.Status_Aktif) === 'Aktif').map(guruAwam_).map(g =>
      boleh_(u, 'urus_guru') || g.id === u.id ? g : Object.assign(g, { emel: '', tel: '' })),
    kategoriBantuan: read_('tbl_kategori_bantuan').rows.filter(k => String(k.Status_Aktif) === 'Aktif').map(k => ({
      id: String(k.ID_Kategori), kod: String(k.Kod), nama: String(k.Nama_Bantuan),
      pemberi: String(k.Pemberi), kriteria: String(k.Kriteria),
      hadIsiRumah: Number(k.Had_Isi_Rumah) || 0, hadPerKapita: Number(k.Had_Per_Kapita) || 0, syarat: String(k.Syarat_Had || 'KEDUA'),
      idPIC: String(k.ID_Guru_PIC || ''), namaPIC: namaGuru.get(String(k.ID_Guru_PIC || '')) || ''
    })),
    kategoriDisiplin: read_('tbl_kategori_disiplin').rows.filter(k => String(k.Status_Aktif) === 'Aktif').map(k => ({
      id: String(k.ID_Kategori), jenis: String(k.Jenis), nama: String(k.Nama), mata: Number(k.Mata) || 0
    })),
    kebenaran: Object.keys(KEBENARAN).filter(k => boleh_(u, k)),
    rmt: rmtInfo_(u),
    bertugas: bertugasInfo_(u),
    unitHem: unitInfo_(u)
  };
}

function apiSenaraiMurid_(p, u) {
  return senaraiMurid_(u, kelasSaya_(u), tetapan_());
}

function senaraiMurid_(u, kelasSaya, t) {
  return read_('tbl_murid').rows
    .filter(m => String(m.Status) === 'Aktif')
    .map(m => {
      const merit = Number(m.Mata_Merit);
      const o = {
        id: String(m.ID_Murid), nama: String(m.Nama), jantina: String(m.Jantina),
        tahun: String(m.Tahun), idKelas: String(m.ID_Kelas), namaKelas: String(m.Nama_Kelas),
        warganegara: String(m.Warganegara), oku: String(m.OKU),
        namaPenjaga: String(m.Nama_Penjaga), telPenjaga: String(m.Tel_Penjaga),
        merit: merit, statusMerit: statusMerit_(merit, t),
        tag: String(m.Tag || ''), fotoId: String(m.Foto_ID || '')
      };
      if (bolehSensitif_(u, m.ID_Kelas, kelasSaya)) {
        const pk = perKapita_(m, t);
        o.sensitif = {
          pendapatan: m.Pendapatan_Isi_Rumah === '' ? null : Number(m.Pendapatan_Isi_Rumah),
          kategoriPendapatan: String(m.Kategori_Pendapatan), yatim: String(m.Yatim),
          tanggungan: String(m.Tanggungan), kategoriOKU: String(m.Kategori_OKU),
          perKapita: pk.perKapita, bilIsiRumah: pk.bil,
          hubunganPenjaga: String(m.Hubungan_Penjaga)
        };
      }
      return o;
    });
}

function guruAwam_(g) {
  return {
    id: String(g.ID_Guru), nama: String(g.Nama_Penuh), emel: String(g.Emel), tel: String(g.No_Telefon),
    jawatan: String(g.Jawatan), gred: String(g.Gred), subjek: String(g.Subjek),
    peranan: String(g.Peranan), labelPeranan: PERANAN[g.Peranan] || 'Guru',
    sesi: String(g.Sesi), fotoId: String(g.Foto_ID || '')
  };
}

/* ---------- Dashboard ---------- */
function apiDashboard_() {
  const cache = CacheService.getScriptCache();
  const c = cache.get('dashboard');
  if (c) return JSON.parse(c);

  const t = tetapan_();
  const hariIni = today_();
  const bulanIni = hariIni.slice(0, 7);
  const tahunIni = hariIni.slice(0, 4);
  const minggu = new Set();
  for (let i = 0; i < 7; i++) minggu.add(fmtTarikh_(new Date(Date.now() - i * 86400000)));

  const kelas = new Map(read_('tbl_kelas').rows.map(k => [String(k.ID_Kelas), {
    nama: String(k.Nama_Kelas), sesi: String(k.Sesi || 'Belum Ditetapkan'), aktif: String(k.Status) !== 'Tidak Aktif'
  }]));
  const murid = read_('tbl_murid').rows.filter(m => String(m.Status) === 'Aktif');
  const namaMurid = new Map(murid.map(m => [String(m.ID_Murid), String(m.Nama)]));

  // Enrolmen
  const tahunMap = {};
  let L = 0, P = 0, warga = 0, kritikal = 0;
  murid.forEach(m => {
    const th = String(m.Tahun || '-');
    tahunMap[th] = tahunMap[th] || { L: 0, P: 0 };
    if (m.Jantina === 'L') { L++; tahunMap[th].L++; } else if (m.Jantina === 'P') { P++; tahunMap[th].P++; }
    if (warganegara_(m)) warga++;
    if (statusMerit_(Number(m.Mata_Merit), t) === 'Kritikal') kritikal++;
  });
  const tahunList = Object.keys(tahunMap).sort((a, b) => tahunSort_(a) - tahunSort_(b));
  // Kelas aktif yang mempunyai murid (sama seperti Analisis Kehadiran)
  const adaMurid = new Set(murid.map(m => String(m.ID_Kelas)));
  let kelasAktif = 0;
  kelas.forEach((k, id) => { if (k.aktif && adaMurid.has(id)) kelasAktif++; });

  // Kehadiran
  let hHadir = 0, hJumlah = 0;
  const kelasDirekod = new Set();
  const sesiMinggu = {}, kelasBulan = {};
  read_('tbl_kehadiran').rows.forEach(r => {
    const tr = fmtTarikh_(r.Tarikh);
    const hadir = hadirH_(r.Status) ? 1 : 0;
    const idK = String(r.ID_Kelas);
    if (tr === hariIni) { hJumlah++; hHadir += hadir; kelasDirekod.add(idK); }
    if (minggu.has(tr)) {
      const s = (kelas.get(idK) || {}).sesi || 'Belum Ditetapkan';
      sesiMinggu[s] = sesiMinggu[s] || { hadir: 0, jumlah: 0 };
      sesiMinggu[s].hadir += hadir; sesiMinggu[s].jumlah++;
    }
    if (tr.slice(0, 7) === bulanIni) {
      kelasBulan[idK] = kelasBulan[idK] || { hadir: 0, jumlah: 0 };
      kelasBulan[idK].hadir += hadir; kelasBulan[idK].jumlah++;
    }
  });
  const pct = (h, j) => j ? Math.round(h / j * 1000) / 10 : null;

  const terbaik = {};
  Object.keys(kelasBulan).forEach(idK => {
    const k = kelas.get(idK); if (!k) return;
    const p = pct(kelasBulan[idK].hadir, kelasBulan[idK].jumlah);
    if (!terbaik[k.sesi] || p > terbaik[k.sesi].peratus) terbaik[k.sesi] = { kelas: k.nama, peratus: p };
  });

  // Kebersihan bulan ini
  const bersih = {};
  read_('tbl_kebersihan').rows.forEach(r => {
    if (String(r.Bulan) !== bulanIni) return;
    const idK = String(r.ID_Kelas);
    bersih[idK] = bersih[idK] || { jumlah: 0, bil: 0 };
    bersih[idK].jumlah += Number(r.Jumlah_Skor) || 0; bersih[idK].bil++;
  });
  let terbersih = null;
  Object.keys(bersih).forEach(idK => {
    const purata = Math.round(bersih[idK].jumlah / bersih[idK].bil * 10) / 10;
    if (!terbersih || purata > terbersih.purata) {
      terbersih = { kelas: (kelas.get(idK) || {}).nama || idK, purata: purata };
    }
  });

  // Disiplin terkini
  const kes = read_('tbl_disiplin').rows
    .sort((a, b) => new Date(b.Tarikh_Masa) - new Date(a.Tarikh_Masa))
    .slice(0, 5)
    .map(k => ({
      idKes: String(k.ID_Kes), namaMurid: namaMurid.get(String(k.ID_Murid)) || String(k.ID_Murid),
      kategori: String(k.Kategori), mata: Number(k.Mata_Kesan), tarikh: fmtTarikh_(k.Tarikh_Masa)
    }));

  // Bantuan tahun ini
  const katB = new Map(read_('tbl_kategori_bantuan').rows.map(k => [String(k.ID_Kategori), k]));
  const muridId = new Map(murid.map(m => [String(m.ID_Murid), m]));
  const bantuan = {};
  // Hanya penerima yang layak dikira (sama seperti senarai penerima)
  read_('tbl_agihan_bantuan').rows.forEach(a => {
    const k = katB.get(String(a.ID_Kategori)), m = muridId.get(String(a.ID_Murid));
    if (String(a.Tahun) !== tahunIni || !m || !layakBantuan_(m, k)) return;
    const kod = k ? String(k.Kod || k.Nama_Bantuan) : String(a.ID_Kategori);
    bantuan[kod] = (bantuan[kod] || 0) + 1;
  });

  const d = {
    dikira: new Date().toISOString(),
    enrolmen: { jumlah: murid.length, lelaki: L, perempuan: P, warganegara: warga, bukanWarganegara: murid.length - warga },
    enrolmenTahun: tahunList.map(th => ({ tahun: th, lelaki: tahunMap[th].L, perempuan: tahunMap[th].P })),
    kehadiranHariIni: {
      hadir: hHadir, tidakHadir: hJumlah - hHadir, peratus: pct(hHadir, hJumlah),
      kelasDirekod: kelasDirekod.size, jumlahKelas: kelasAktif
    },
    kehadiranSesiMinggu: Object.keys(sesiMinggu).map(s => ({ sesi: s, peratus: pct(sesiMinggu[s].hadir, sesiMinggu[s].jumlah) })),
    kehadiranTerbaikBulan: terbaik,
    kelasTerbersihBulan: terbersih,
    kesTerkini: kes,
    bantuanTahunIni: bantuan,
    bilanganKritikal: kritikal
  };
  try { cache.put('dashboard', JSON.stringify(d), APP.CACHE_DASHBOARD_SAAT); } catch (e) { /* terlalu besar untuk cache */ }
  return d;
}

/* ---------- Kelas ---------- */
function apiGetKelas_(p) {
  const k = read_('tbl_kelas').rows.find(r => String(r.ID_Kelas) === String(p.idKelas));
  if (!k) throw err_('Kelas tidak dijumpai.');
  return {
    id: String(k.ID_Kelas), nama: String(k.Nama_Kelas), tahun: String(k.Tahun), sesi: String(k.Sesi),
    idGuruKelas: String(k.ID_Guru_Kelas),
    cartaOrg: parseJSON_(k.Carta_Org_JSON, {}),
    jadualTugas: parseJSON_(k.Jadual_Tugas_JSON, {}),
    pelanMeja: parseJSON_(k.Pelan_Meja_JSON, [])
  };
}

/* ---------- Jadual waktu kelas (import XML aSc Timetables) ----------
 * Fail XML dihurai di pelayar; pelayan hanya menerima jadual yang sudah dipetakan ke ID_Kelas, membersihkannya
 * dan menyimpan satu baris JSON bagi setiap kelas dalam tbl_jadual_waktu.
 * Format JSON: { hari: [nama], waktu: [{label, mula, tamat, rehat}], sel: [{h, w, n, s: [{k, nm, g, b}]}] }
 *   h = indeks hari, w = indeks lajur waktu, n = bilangan waktu bersambung, s = subjek (k singkatan, nm nama,
 *   g guru, b bilik) — lebih daripada satu subjek dalam satu sel bagi kumpulan (cth PI / PM) */
const SHEET_JW = 'tbl_jadual_waktu';

function adaSheetJW_() { return !!SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_JW); }

function apiGetJadualWaktu_(p) {
  const idKelas = String(p.idKelas || '');
  if (!adaSheetJW_()) return { idKelas: idKelas, jadual: null, kelasAda: [] };
  const rows = read_(SHEET_JW).rows;
  const r = rows.find(x => String(x.ID_Kelas) === idKelas);
  return {
    idKelas: idKelas,
    jadual: r ? parseJSON_(r.Jadual_JSON, null) : null,
    kelasAsc: r ? String(r.Kelas_aSc || '') : '',
    namaFail: r ? String(r.Nama_Fail || '') : '',
    tarikh: r ? fmtTarikh_(r.Tarikh_Kemaskini) : '',
    kelasAda: rows.map(x => String(x.ID_Kelas))
  };
}

function apiSimpanJadualWaktu_(p, u) {
  wajib_(u, 'urus_jadual_waktu');
  const senarai = Array.isArray(p.kelas) ? p.kelas : [];
  if (!senarai.length) throw err_('Tiada jadual kelas untuk disimpan.');
  if (senarai.length > 200) throw err_('Terlalu banyak kelas dalam satu muat naik.');
  const kelasWujud = new Set(read_('tbl_kelas').rows.map(k => String(k.ID_Kelas)));
  const masuk = new Map();
  senarai.forEach(x => {
    const id = String(x.idKelas || '');
    if (!kelasWujud.has(id)) throw err_('Kelas tidak dijumpai: ' + id);
    if (masuk.has(id)) throw err_('Kelas yang sama dipadankan lebih daripada sekali: ' + id);
    masuk.set(id, { asc: String(x.kelasAsc || '').slice(0, 60), json: jsonSel_(bersihJadual_(x.jadual), 'Jadual waktu ' + id) });
  });

  if (!adaSheetJW_()) ensureSheet_(SpreadsheetApp.getActiveSpreadsheet(), SHEET_JW, SCHEMA[SHEET_JW]);
  const stamp = new Date(), namaFail = String(p.namaFail || '').slice(0, 120);
  const rows = read_(SHEET_JW).rows.filter(r => !p.ganti || masuk.has(String(r.ID_Kelas)));   // ganti = buang kelas lain
  const ikutId = new Map(rows.map(r => [String(r.ID_Kelas), r]));
  masuk.forEach((v, id) => {
    const r = ikutId.get(id) || {};
    Object.assign(r, { ID_Kelas: id, Kelas_aSc: v.asc, Jadual_JSON: v.json, Nama_Fail: namaFail, ID_Guru_Muat_Naik: u.id, Tarikh_Kemaskini: stamp });
    ikutId.set(id, r);
  });
  writeAll_(SHEET_JW, Array.from(ikutId.values()));
  audit_(u.id, 'MUAT_NAIK_JADUAL_WAKTU', `${masuk.size} kelas ${namaFail}${p.ganti ? ' (ganti semua)' : ''}`);
  return { disimpan: masuk.size, jumlah: ikutId.size };
}

// Status jadual waktu bagi semua kelas (Tetapan Admin → Jadual waktu)
function apiSenaraiJadualWaktu_(p, u) {
  wajib_(u, 'urus_jadual_waktu');
  if (!adaSheetJW_()) return [];
  return read_(SHEET_JW).rows.map(r => {
    const j = parseJSON_(r.Jadual_JSON, null);
    return { idKelas: String(r.ID_Kelas), kelasAsc: String(r.Kelas_aSc || ''), namaFail: String(r.Nama_Fail || ''),
      tarikh: fmtTarikh_(r.Tarikh_Kemaskini), bilWaktu: j && Array.isArray(j.sel) ? j.sel.reduce((a, c) => a + (Number(c.n) || 1), 0) : 0 };
  });
}

function apiPadamJadualWaktu_(p, u) {
  wajib_(u, 'urus_jadual_waktu');
  if (!adaSheetJW_()) return { dipadam: 0 };
  const rows = read_(SHEET_JW).rows;
  const kekal = p.semua ? [] : rows.filter(r => String(r.ID_Kelas) !== String(p.idKelas || ''));
  writeAll_(SHEET_JW, kekal);
  audit_(u.id, 'PADAM_JADUAL_WAKTU', p.semua ? 'semua' : String(p.idKelas || ''));
  return { dipadam: rows.length - kekal.length };
}

// Bersihkan jadual daripada pelayar: had saiz, jenis data & julat indeks
function bersihJadual_(j) {
  j = j && typeof j === 'object' ? j : {};
  const teks = (v, n) => String(v == null ? '' : v).replace(/\s+/g, ' ').trim().slice(0, n);
  const masa = v => /^\d{1,2}:\d{2}$/.test(String(v || '')) ? String(v) : '';
  const hari = (Array.isArray(j.hari) ? j.hari : []).slice(0, 7).map(h => teks(h, 20));
  const waktu = (Array.isArray(j.waktu) ? j.waktu : []).slice(0, 24).map(w => ({
    label: teks(w && w.label, 20), mula: masa(w && w.mula), tamat: masa(w && w.tamat), rehat: !!(w && w.rehat)
  }));
  if (!hari.length || !waktu.length) throw err_('Jadual tidak lengkap (tiada hari atau waktu).');
  const bulat = v => Math.round(Number(v));
  const sel = (Array.isArray(j.sel) ? j.sel : []).slice(0, 400).map(c => ({
    h: bulat(c && c.h), w: bulat(c && c.w), n: bulat(c && c.n) || 1,
    s: (Array.isArray(c && c.s) ? c.s : []).slice(0, 6).map(x => ({ k: teks(x && x.k, 20), nm: teks(x && x.nm, 80), g: teks(x && x.g, 80), b: teks(x && x.b, 40) }))
  })).filter(c => c.h >= 0 && c.h < hari.length && c.w >= 0 && c.w < waktu.length && c.n >= 1 && c.w + c.n <= waktu.length && c.s.length);
  return { hari: hari, waktu: waktu, sel: sel };
}

/* ---------- Tema & header peribadi guru ---------- */
function guruSaya_(u) {
  const g = read_('tbl_guru').rows.find(r => String(r.ID_Guru) === u.id);
  if (!g) throw err_('Profil guru tidak dijumpai.');
  return g;
}

function apiSimpanTema_(p, u) {
  const g = guruSaya_(u);
  if (p.tema !== undefined) g.Tema = TEMA_SAH.indexOf(p.tema) >= 0 ? p.tema : '';
  if (p.latar !== undefined) g.Latar = LATAR_SAH.indexOf(p.latar) >= 0 ? p.latar : '';
  updateRow_('tbl_guru', g);
  return { tema: String(g.Tema || ''), latar: String(g.Latar || '') };
}

// Header: {fotoId} dari muat naik, {url} pautan gambar, atau {buang: true}
function apiSimpanHeader_(p, u) {
  const g = guruSaya_(u);
  if (p.buang) { g.Header_ID = ''; g.Header_URL = ''; }
  else if (p.fotoId) { g.Header_ID = idFail_(p.fotoId); g.Header_URL = ''; }
  else if (p.url) { g.Header_URL = pautanGambar_(p.url); g.Header_ID = ''; }
  else throw err_('Tiada gambar diberi.');
  updateRow_('tbl_guru', g);
  audit_(u.id, 'SIMPAN_HEADER', p.buang ? 'buang' : p.fotoId ? 'muat naik' : 'pautan');
  return { headerId: String(g.Header_ID || ''), headerUrl: String(g.Header_URL || '') };
}

// Pautan kongsi Google Drive → pautan paparan terus; pautan lain mesti https
function pautanGambar_(url) {
  const v = String(url || '').trim();
  const m = v.match(/drive\.google\.com\/(?:file\/d\/|open\?id=|uc\?(?:export=\w+&)?id=)([\w-]{20,})/);
  if (m) return 'https://drive.google.com/thumbnail?id=' + m[1] + '&sz=w2000';
  if (!/^https:\/\/[^\s"'<>]+$/i.test(v)) throw err_('Pautan gambar mesti bermula dengan https://');
  return v.slice(0, 1000);
}

/* ---------- Analisis kehadiran harian (ikut tahun, kelas terbaik, kelas belum isi) ---------- */
function apiAnalisisKehadiran_(p) {
  const tarikh = p.tarikh ? semakTarikh_(p.tarikh) : today_();
  const cache = CacheService.getScriptCache();
  const ck = 'ah2:' + tarikh;
  const c = cache.get(ck);
  if (c) return JSON.parse(c);

  const bulan = tarikh.slice(0, 7);
  const namaGuru = new Map(read_('tbl_guru').rows.map(g => [String(g.ID_Guru), String(g.Nama_Penuh)]));
  const enrol = {};
  read_('tbl_murid').rows.forEach(m => { if (String(m.Status) === 'Aktif') enrol[String(m.ID_Kelas)] = (enrol[String(m.ID_Kelas)] || 0) + 1; });

  const K = {};
  read_('tbl_kelas').rows.forEach(k => {
    const id = String(k.ID_Kelas);
    if (String(k.Status) === 'Tidak Aktif' || !enrol[id]) return;
    K[id] = {
      idKelas: id, nama: String(k.Nama_Kelas), tahun: String(k.Tahun), sesi: String(k.Sesi || ''),
      idGuruKelas: String(k.ID_Guru_Kelas || ''),
      guruKelas: namaGuru.get(String(k.ID_Guru_Kelas)) || String(k.Nama_Guru_Kelas_IDME || ''),
      enrolmen: enrol[id], direkod: false, hadir: 0, lewat: 0, tidakHadir: 0, bH: 0, bJ: 0, bHari: {}
    };
  });

  read_('tbl_kehadiran').rows.forEach(r => {
    const k = K[String(r.ID_Kelas)]; if (!k) return;
    const tr = fmtTarikh_(r.Tarikh), st = statusH_(r.Status);
    if (tr === tarikh) {
      k.direkod = true;
      if (st === 'Hadir') k.hadir++; else if (st === 'Lewat') k.lewat++; else k.tidakHadir++;
    }
    if (tr.slice(0, 7) === bulan && tr <= tarikh) { k.bJ++; if (hadirH_(st)) k.bH++; k.bHari[tr] = 1; }
  });

  const pct = (h, j) => j ? Math.round(h / j * 1000) / 10 : null;
  const perKelas = Object.keys(K).map(id => K[id])
    .sort((a, b) => tahunSort_(a.tahun) - tahunSort_(b.tahun) || a.nama.localeCompare(b.nama))
    .map(k => ({
      idKelas: k.idKelas, nama: k.nama, tahun: k.tahun, sesi: k.sesi, idGuruKelas: k.idGuruKelas, guruKelas: k.guruKelas,
      enrolmen: k.enrolmen, direkod: k.direkod, hadir: k.hadir, lewat: k.lewat, tidakHadir: k.tidakHadir,
      peratus: pct(k.hadir + k.lewat, k.hadir + k.lewat + k.tidakHadir),
      bulanPeratus: pct(k.bH, k.bJ), bulanHari: Object.keys(k.bHari).length
    }));

  const T = {};
  perKelas.forEach(k => {
    const t = T[k.tahun] = T[k.tahun] || { tahun: k.tahun, enrolmen: 0, kelas: 0, kelasDirekod: 0, hadir: 0, lewat: 0, tidakHadir: 0 };
    t.enrolmen += k.enrolmen; t.kelas++;
    if (k.direkod) { t.kelasDirekod++; t.hadir += k.hadir; t.lewat += k.lewat; t.tidakHadir += k.tidakHadir; }
  });
  const perTahun = Object.keys(T).map(t => T[t]).sort((a, b) => tahunSort_(a.tahun) - tahunSort_(b.tahun))
    .map(t => Object.assign(t, { peratus: pct(t.hadir + t.lewat, t.hadir + t.lewat + t.tidakHadir) }));

  const jum = perTahun.reduce((a, t) => ({ h: a.h + t.hadir, l: a.l + t.lewat, x: a.x + t.tidakHadir }), { h: 0, l: 0, x: 0 });
  const out = {
    tarikh: tarikh, bulan: bulan,
    sekolah: { peratus: pct(jum.h + jum.l, jum.h + jum.l + jum.x), hadir: jum.h, lewat: jum.l, tidakHadir: jum.x,
      kelasDirekod: perKelas.filter(k => k.direkod).length, jumlahKelas: perKelas.length },
    perTahun: perTahun,
    perKelas: perKelas,
    terbaikHari: perKelas.filter(k => k.direkod && k.peratus !== null)
      .sort((a, b) => b.peratus - a.peratus || b.enrolmen - a.enrolmen).slice(0, 10),
    terbaikBulan: perKelas.filter(k => k.bulanHari > 0)
      .sort((a, b) => b.bulanPeratus - a.bulanPeratus || b.bulanHari - a.bulanHari).slice(0, 10),
    belumIsi: perKelas.filter(k => !k.direkod)
  };
  try { cache.put(ck, JSON.stringify(out), 120); } catch (e) { /* data terlalu besar untuk cache */ }
  return out;
}

/* ---------- SPBT (buku teks): rekod pinjam & pulang sekelas ----------
 * Hanya guru kelas bagi kelas itu dan pentadbir (urus_kelas_semua). Senarai buku ikut tahun disimpan dalam
 * tetapan SPBT_BUKU; jika kosong, senarai lalai digunakan. Rekod ikut tahun sesi, jadi tahun baharu bermula kosong. */
const SHEET_SPBT = 'tbl_spbt';
const SPBT_BUKU_LALAI = {
  1: ['Bahasa Melayu Jilid 1', 'Bahasa Melayu Jilid 2', "Super Minds Student's Book 1", 'Matematik Jilid 1', 'Matematik Jilid 2', 'Sains', 'Pendidikan Islam', 'Pendidikan Moral', 'Pendidikan Kesenian', 'PJPK', 'Bahasa Arab', 'Buku Aktiviti Bahasa Melayu Jilid 1', 'Buku Aktiviti Bahasa Melayu Jilid 2'],
  2: ['Bahasa Melayu Jilid 1', 'Bahasa Melayu Jilid 2', "Super Minds Student's Book 2", 'Matematik Jilid 1', 'Matematik Jilid 2', 'Sains', 'Pendidikan Islam', 'Pendidikan Moral', 'Pendidikan Kesenian', 'PJPK', 'Bahasa Arab', 'Buku Aktiviti Bahasa Melayu Jilid 1', 'Buku Aktiviti Bahasa Melayu Jilid 2'],
  3: ['Bahasa Melayu Jilid 1', 'Bahasa Melayu Jilid 2', 'Get Smart Plus 3', 'Matematik Jilid 1', 'Matematik Jilid 2', 'Sains', 'Pendidikan Islam', 'Pendidikan Moral', 'Pendidikan Kesenian', 'PJPK', 'Bahasa Arab', 'Buku Aktiviti Bahasa Melayu Jilid 1', 'Buku Aktiviti Bahasa Melayu Jilid 2'],
  4: ['Bahasa Melayu', 'Get Smart Plus 4', 'Matematik', 'Sains', 'Pendidikan Islam', 'Pendidikan Moral', 'Pendidikan Muzik', 'Pendidikan Seni Visual', 'PJPK', 'Sejarah', 'Reka Bentuk dan Teknologi', 'Bahasa Arab'],
  5: ['Bahasa Melayu', 'English Plus 1', 'Matematik', 'Sains', 'Pendidikan Islam', 'Pendidikan Moral', 'Pendidikan Muzik', 'Pendidikan Seni Visual', 'PJPK', 'Sejarah', 'Reka Bentuk dan Teknologi', 'Bahasa Arab'],
  6: ['Bahasa Melayu', 'Academy Stars Year 6', 'Matematik', 'Sains', 'Pendidikan Islam', 'Pendidikan Moral', 'Pendidikan Muzik', 'Pendidikan Seni Visual', 'PJPK', 'Sejarah', 'Reka Bentuk dan Teknologi', 'Bahasa Arab']
};

function bukuSpbt_() {
  const simpan = parseJSON_(tetapan_().SPBT_BUKU, null) || {};
  const out = {};
  Object.keys(SPBT_BUKU_LALAI).forEach(n => { out[n] = Array.isArray(simpan[n]) ? simpan[n] : SPBT_BUKU_LALAI[n]; });
  return out;
}

function kelasSpbt_(idKelas, u) {
  const k = read_('tbl_kelas').rows.find(r => String(r.ID_Kelas) === idKelas);
  if (!k) throw err_('Kelas tidak dijumpai.');
  if (!boleh_(u, 'urus_kelas_semua') && kelasSaya_(u).indexOf(idKelas) < 0) {
    throw err_('Rekod SPBT hanya boleh dibuka oleh guru kelas tersebut atau PK HEM.', 'FORBIDDEN');
  }
  return k;
}

function apiGetSpbt_(p, u) {
  const idKelas = String(p.idKelas || '');
  const k = kelasSpbt_(idKelas, u);
  const tahunSesi = today_().slice(0, 4);
  const semua = bukuSpbt_();
  const dalamKelas = new Set(read_('tbl_murid').rows.filter(m => String(m.ID_Kelas) === idKelas && String(m.Status) === 'Aktif' && warganegara_(m)).map(m => String(m.ID_Murid)));
  const rekod = {};
  if (SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SPBT)) {
    read_(SHEET_SPBT).rows.forEach(r => {
      if (String(r.Tahun_Sesi) === tahunSesi && dalamKelas.has(String(r.ID_Murid))) rekod[String(r.ID_Murid)] = parseJSON_(r.Rekod_JSON, {});
    });
  }
  return { idKelas: idKelas, tahunSesi: tahunSesi, buku: semua[tahunSort_(k.Tahun)] || [], semuaBuku: boleh_(u, 'urus_kelas_semua') ? semua : null, rekod: rekod };
}

// ubah: [{ idMurid, buku, nilai }] — hanya sel yang berubah, jadi dua guru tidak menimpa sel masing-masing
function apiSimpanSpbt_(p, u) {
  const idKelas = String(p.idKelas || '');
  const k = kelasSpbt_(idKelas, u);
  const ubah = Array.isArray(p.ubah) ? p.ubah : [];
  if (!ubah.length) return { disimpan: 0 };
  if (ubah.length > 5000) throw err_('Terlalu banyak perubahan dalam satu simpanan.');
  const tahunSesi = today_().slice(0, 4);
  const bukuSah = new Set(bukuSpbt_()[tahunSort_(k.Tahun)] || []);
  const dalamKelas = new Set(read_('tbl_murid').rows.filter(m => String(m.ID_Kelas) === idKelas && String(m.Status) === 'Aktif' && warganegara_(m)).map(m => String(m.ID_Murid)));
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName(SHEET_SPBT)) ensureSheet_(ss, SHEET_SPBT, SCHEMA[SHEET_SPBT]);
  const sedia = new Map();
  read_(SHEET_SPBT).rows.forEach(r => { if (String(r.Tahun_Sesi) === tahunSesi) sedia.set(String(r.ID_Murid), r); });

  const perMurid = new Map();
  ubah.forEach(x => {
    const id = String(x.idMurid), buku = String(x.buku), n = Number(x.nilai);
    if (!dalamKelas.has(id)) throw err_('Murid ' + id + ' bukan dalam kelas ini.');
    if (!bukuSah.has(buku)) throw err_('Buku tidak dijumpai: ' + buku);
    if ([0, 1, 2, 3].indexOf(n) < 0) throw err_('Nilai rekod tidak sah.');
    if (!perMurid.has(id)) perMurid.set(id, []);
    perMurid.get(id).push([buku, n]);
  });

  const now = new Date(), baru = [], kemas = [];
  perMurid.forEach((senarai, id) => {
    const r = sedia.get(id);
    const rec = r ? parseJSON_(r.Rekod_JSON, {}) : {};
    senarai.forEach(([buku, n]) => { if (n) rec[buku] = n; else delete rec[buku]; });
    const data = { ID_Kelas: idKelas, Rekod_JSON: JSON.stringify(rec), ID_Guru_Rekod: u.id, Timestamp: now };
    if (r) { Object.assign(r, data); kemas.push(r); }
    else baru.push(Object.assign({ ID_Murid: id, Tahun_Sesi: tahunSesi }, data));
  });
  updateRows_(SHEET_SPBT, kemas);
  insert_(SHEET_SPBT, baru);
  audit_(u.id, 'SIMPAN_SPBT', `${idKelas} ${ubah.length} rekod`);
  return { disimpan: ubah.length };
}

// Senarai buku ikut tahun (PK HEM / Super Admin). Rekod dipadankan ikut nama buku.
function apiSimpanBukuSpbt_(p, u) {
  wajib_(u, 'urus_kelas_semua');
  ensureTetapan_();
  const masuk = p.buku && typeof p.buku === 'object' ? p.buku : {};
  const semua = bukuSpbt_();
  Object.keys(SPBT_BUKU_LALAI).forEach(n => {
    if (!Array.isArray(masuk[n])) return;
    const bersih = Array.from(new Set(masuk[n].map(b => String(b).replace(/\s+/g, ' ').trim()).filter(Boolean)));
    if (bersih.length > 30) throw err_('Maksimum 30 buku bagi setiap tahun.');
    if (bersih.some(b => b.length > 80)) throw err_('Nama buku terlalu panjang (maksimum 80 aksara).');
    semua[n] = bersih;
  });
  const r = read_('tbl_tetapan').rows.find(x => String(x.Kunci) === 'SPBT_BUKU');
  r.Nilai = jsonSel_(semua, 'Senarai buku');
  updateRow_('tbl_tetapan', r);
  audit_(u.id, 'SIMPAN_BUKU_SPBT', '');
  return { semuaBuku: semua };
}

/* ---------- Senarai nama & No. KP murid untuk dimuat turun (Excel) ---------- */
// No. KP hanya diisi bagi kelas yang pengguna berhak lihat maklumat sulit (guru kelas / guru disiplin / PK HEM)
/* ---------- Analisis data murid sekelas (tab Data murid dalam Pengurusan Kelas) ----------
 * Semua guru: bilangan murid, L/P & status merit. Guru kelas / PK HEM / guru disiplin (bolehSensitif_):
 * kategori pendapatan, penerima bantuan ikut skim, dan murid yang datanya tidak lengkap. */
function apiDataMuridKelas_(p, u) {
  const idKelas = String(p.idKelas || '');
  const k = read_('tbl_kelas').rows.find(r => String(r.ID_Kelas) === idKelas);
  if (!k) throw err_('Kelas tidak dijumpai.');
  const t = tetapan_();
  const tm = read_('tbl_murid');
  const murid = tm.rows.filter(m => String(m.ID_Kelas) === idKelas && String(m.Status) === 'Aktif')
    .sort((a, b) => String(a.Nama).localeCompare(String(b.Nama)));
  const kira = f => murid.filter(f).length;
  const ringkas = m => ({ id: String(m.ID_Murid), nama: String(m.Nama) });

  const out = {
    idKelas: idKelas, jumlah: murid.length,
    lelaki: kira(m => m.Jantina === 'L'), perempuan: kira(m => m.Jantina === 'P'),
    merit: { Baik: 0, Amaran: 0, Kritikal: 0 },
    penuh: bolehSensitif_(u, idKelas)
  };
  murid.forEach(m => { out.merit[statusMerit_(Number(m.Mata_Merit), t)]++; });
  if (!out.penuh) return out;

  // Profil isi rumah
  const kat = { B40: 0, M40: 0, T20: 0, 'Tiada Maklumat': 0 };
  murid.forEach(m => { const c = String(m.Kategori_Pendapatan || 'Tiada Maklumat'); kat[kat[c] === undefined ? 'Tiada Maklumat' : c]++; });
  out.pendapatan = kat;
  out.khas = { oku: kira(m => adaStatus_(m.OKU)), yatim: kira(m => adaStatus_(m.Yatim)), bukanWarganegara: kira(m => !warganegara_(m)) };

  // Penerima bantuan tahun ini ikut skim (status "Tidak Layak" tidak dikira)
  const tahun = /^\d{4}$/.test(String(p.tahun || '')) ? String(p.tahun) : today_().slice(0, 4);
  // Hanya penerima yang layak dikira: warganegara, dan ada maklumat pendapatan kecuali BAP/SPBT
  const dalamKelas = new Map(murid.map(m => [String(m.ID_Murid), m]));
  const katSkim = new Map(read_('tbl_kategori_bantuan').rows.map(c => [String(c.ID_Kategori), c]));
  const agihan = read_('tbl_agihan_bantuan').rows.filter(a => String(a.Tahun) === tahun && dalamKelas.has(String(a.ID_Murid)) && String(a.Status) !== 'Tidak Layak' &&
    layakBantuan_(dalamKelas.get(String(a.ID_Murid)), katSkim.get(String(a.ID_Kategori))));
  const unik = new Set(agihan.map(a => String(a.ID_Murid)));
  out.bantuan = {
    tahun: tahun, jumlahPenerima: unik.size,
    skim: read_('tbl_kategori_bantuan').rows.filter(c => String(c.Status_Aktif) === 'Aktif').map(c => {
      const a = agihan.filter(x => String(x.ID_Kategori) === String(c.ID_Kategori));
      return { id: String(c.ID_Kategori), kod: String(c.Kod), nama: String(c.Nama_Bantuan),
        penerima: new Set(a.map(x => String(x.ID_Murid))).size, selesai: a.filter(x => String(x.Status) === 'Selesai Agihan').length };
    })
  };

  // Data tidak lengkap
  const alamat = new Map(read_('tbl_alamat_murid').rows.map(a => [String(a.ID_Murid), a]));
  const kosong = v => String(v == null ? '' : v).trim() === '';
  // Pecahan pendapatan penjaga 1/2 hanya tersedia selepas import IDME versi ini
  const adaPecahan = tm.headers.indexOf('Pendapatan_Penjaga_1') >= 0 &&
    tm.rows.some(m => !kosong(m.Pendapatan_Penjaga_1) || !kosong(m.Pendapatan_Penjaga_2));
  const senarai = {
    tiadaAlamat: murid.filter(m => { const a = alamat.get(String(m.ID_Murid)); return !a || (kosong(a.Alamat_1) && kosong(a.Poskod)); }),
    tiadaTel1: murid.filter(m => kosong(m.Tel_Penjaga_1)),
    tiadaTel2: murid.filter(m => !kosong(m.Nama_Penjaga_2) && kosong(m.Tel_Penjaga_2)),
    tiadaTelLangsung: murid.filter(m => kosong(m.Tel_Penjaga_1) && kosong(m.Tel_Penjaga_2)),
    tiadaPendapatan: murid.filter(m => kosong(m.Pendapatan_Isi_Rumah)),
    tiadaPendapatan1: adaPecahan ? murid.filter(m => kosong(m.Pendapatan_Penjaga_1)) : null,
    tiadaPendapatan2: adaPecahan ? murid.filter(m => !kosong(m.Nama_Penjaga_2) && kosong(m.Pendapatan_Penjaga_2)) : null,
    tiadaTanggungan: murid.filter(m => !(num_(m.Tanggungan) > 0))
  };
  const tidakLengkap = new Set([].concat(senarai.tiadaAlamat, senarai.tiadaTelLangsung, senarai.tiadaPendapatan).map(m => String(m.ID_Murid)));
  out.tidakLengkap = {};
  Object.keys(senarai).forEach(key => { out.tidakLengkap[key] = senarai[key] ? senarai[key].map(ringkas) : null; });
  out.dataLengkap = murid.length - tidakLengkap.size;
  return out;
}

function kosongPendapatan_(m) { const v = m.Pendapatan_Isi_Rumah; return v === '' || v === null || v === undefined || isNaN(Number(v)); }

// Nilai IDME seperti "TIDAK" / "BUKAN OKU" bermaksud tiada status
function adaStatus_(v) { const s = String(v || '').trim(); return !!s && !/^(TIDAK|BUKAN|BUKAN OKU|TIADA|-|N|NO)$/i.test(s); }

function apiSenaraiNamaKP_(p, u) {
  const tahun = String(p.tahun || ''), idKelas = String(p.idKelas || '');
  const kelasSaya = kelasSaya_(u);
  const namaGuru = new Map(read_('tbl_guru').rows.map(g => [String(g.ID_Guru), String(g.Nama_Penuh)]));
  const ikutKelas = {};
  read_('tbl_murid').rows.filter(m => String(m.Status) === 'Aktif').forEach(m => {
    (ikutKelas[String(m.ID_Kelas)] = ikutKelas[String(m.ID_Kelas)] || []).push(m);
  });
  const out = read_('tbl_kelas').rows
    .filter(k => String(k.Status) !== 'Tidak Aktif' && (!tahun || String(k.Tahun) === tahun) && (!idKelas || String(k.ID_Kelas) === idKelas))
    .sort((a, b) => tahunSort_(a.Tahun) - tahunSort_(b.Tahun) || String(a.Nama_Kelas).localeCompare(String(b.Nama_Kelas)))
    .map(k => {
      const kp = bolehSensitif_(u, k.ID_Kelas, kelasSaya);
      const ms = (ikutKelas[String(k.ID_Kelas)] || []).sort((a, b) => String(a.Nama).localeCompare(String(b.Nama)));
      return {
        idKelas: String(k.ID_Kelas), namaKelas: String(k.Nama_Kelas), tahun: String(k.Tahun), sesi: String(k.Sesi || ''),
        guruKelas: namaGuru.get(String(k.ID_Guru_Kelas)) || String(k.Nama_Guru_Kelas_IDME || ''),
        kpDibenarkan: kp,
        murid: ms.map(m => ({ nama: String(m.Nama), noKP: kp ? String(m.No_KP || '') : '', jantina: String(m.Jantina) }))
      };
    });
  audit_(u.id, 'MUAT_TURUN_SENARAI', `${tahun || 'semua tahun'} ${idKelas || 'semua kelas'}`);
  return out;
}

/* ---------- Kelas jagaan (guru kelas) — dari Profil Guru ---------- */
function apiSimpanKelasJagaan_(p, u) {
  const idGuru = String(p.idGuru || u.id);
  const admin = boleh_(u, 'urus_kelas_semua');
  if (idGuru !== u.id && !admin) throw err_('Anda hanya boleh menetapkan kelas jagaan anda sendiri.', 'FORBIDDEN');
  const guru = read_('tbl_guru').rows;
  const g = guru.find(r => String(r.ID_Guru) === idGuru);
  if (!g) throw err_('Guru tidak dijumpai.');
  const namaGuru = new Map(guru.map(r => [String(r.ID_Guru), String(r.Nama_Penuh)]));

  const mahu = new Set((Array.isArray(p.kelas) ? p.kelas : []).map(String));
  if (!admin && mahu.size > 2) throw err_('Maksimum 2 kelas jagaan (cth satu sesi pagi, satu sesi petang).');

  const tk = read_('tbl_kelas');
  const wujud = new Set(tk.rows.map(k => String(k.ID_Kelas)));
  mahu.forEach(id => { if (!wujud.has(id)) throw err_('Kelas tidak dijumpai: ' + id); });

  const ubah = [];
  tk.rows.forEach(k => {
    const id = String(k.ID_Kelas), milik = String(k.ID_Guru_Kelas || '');
    if (mahu.has(id)) {
      if (String(k.Status) === 'Tidak Aktif') throw err_('Kelas ' + k.Nama_Kelas + ' tidak aktif.');
      if (milik && milik !== idGuru && !admin) {
        throw err_(`Kelas ${k.Nama_Kelas} sudah mempunyai guru kelas (${namaGuru.get(milik) || milik}). Minta PK HEM membuat perubahan.`);
      }
      if (milik !== idGuru) { k.ID_Guru_Kelas = idGuru; ubah.push(k); }
    } else if (milik === idGuru) {
      k.ID_Guru_Kelas = ''; ubah.push(k);
    }
  });
  updateRows_('tbl_kelas', ubah);
  CacheService.getScriptCache().remove('dashboard');
  audit_(u.id, 'KELAS_JAGAAN', `${idGuru}: ${Array.from(mahu).join(', ') || '(tiada)'}`);
  return { idGuru: idGuru, kelas: Array.from(mahu) };
}

function apiSimpanKelas_(p, u) {
  const k = read_('tbl_kelas').rows.find(r => String(r.ID_Kelas) === String(p.idKelas));
  if (!k) throw err_('Kelas tidak dijumpai.');
  const admin = boleh_(u, 'urus_kelas_semua');
  const guruKelas = String(k.ID_Guru_Kelas) === u.id;
  if (!admin && !guruKelas) throw err_('Hanya guru kelas atau PK HEM boleh mengubah kelas ini.', 'FORBIDDEN');

  if (p.sesi !== undefined || p.idGuruKelas !== undefined) {
    if (!admin) throw err_('Hanya PK HEM boleh menukar sesi atau guru kelas.', 'FORBIDDEN');
    if (p.sesi !== undefined) {
      if (['Pagi', 'Petang', ''].indexOf(p.sesi) < 0) throw err_('Sesi mesti Pagi atau Petang.');
      k.Sesi = p.sesi;
    }
    if (p.idGuruKelas !== undefined) {
      const idG = String(p.idGuruKelas || '');
      if (idG && !read_('tbl_guru').rows.some(g => String(g.ID_Guru) === idG && String(g.Status_Aktif) === 'Aktif')) {
        throw err_('Guru kelas tidak dijumpai atau tidak aktif.');
      }
      k.ID_Guru_Kelas = idG;
    }
  }
  if (p.cartaOrg !== undefined) k.Carta_Org_JSON = jsonSel_(p.cartaOrg, 'Carta organisasi');
  if (p.jadualTugas !== undefined) k.Jadual_Tugas_JSON = jsonSel_(p.jadualTugas, 'Jadual bertugas');
  if (p.pelanMeja !== undefined) {
    // Format baharu: { lajur, baris, meja: [idMurid | ''] }; format lama (senarai ID) masih diterima
    const pm = p.pelanMeja;
    if (Array.isArray(pm)) k.Pelan_Meja_JSON = JSON.stringify(pm.map(String).slice(0, 200));
    else if (pm && typeof pm === 'object' && Array.isArray(pm.meja)) {
      const lajur = Math.max(1, Math.min(12, Math.round(Number(pm.lajur) || 6)));
      const baris = Math.max(1, Math.min(15, Math.round(Number(pm.baris) || 5)));
      k.Pelan_Meja_JSON = JSON.stringify({ lajur: lajur, baris: baris, meja: pm.meja.slice(0, lajur * baris).map(x => String(x || '')) });
    } else throw err_('Format pelan meja tidak sah.');
  }
  updateRow_('tbl_kelas', k);
  CacheService.getScriptCache().remove('dashboard');
  audit_(u.id, 'SIMPAN_KELAS', String(k.ID_Kelas));
  return apiGetKelas_(p);
}

/* ---------- Kehadiran ---------- */
function apiGetKehadiran_(p) {
  const tarikh = semakTarikh_(p.tarikh);
  const idKelas = String(p.idKelas || '');
  const rekod = {};
  let perekod = '', masa = null;
  kehadiranTahun_(tarikh.slice(0, 4)).forEach(r => {
    if (fmtTarikh_(r.Tarikh) === tarikh && String(r.ID_Kelas) === idKelas) {
      rekod[String(r.ID_Murid)] = { status: statusH_(r.Status), sebab: String(r.Sebab) };
      if (!masa || new Date(r.Timestamp) > new Date(masa)) { masa = r.Timestamp; perekod = String(r.ID_Guru_Perekod); }
    }
  });
  return { tarikh: tarikh, idKelas: idKelas, rekod: rekod, perekod: perekod, masa: masa, sudahDirekod: !!masa };
}

/* ---------- Rekod bulanan sekelas (semak kehadiran lepas) ---------- */
function apiKehadiranBulanan_(p) {
  const bulan = /^\d{4}-\d{2}$/.test(String(p.bulan || '')) ? String(p.bulan) : today_().slice(0, 7);
  const idKelas = String(p.idKelas || '');
  const hari = {}, rekod = {};
  kehadiranTahun_(bulan.slice(0, 4)).forEach(r => {
    const tr = fmtTarikh_(r.Tarikh);
    if (String(r.ID_Kelas) !== idKelas || tr.slice(0, 7) !== bulan) return;
    hari[tr] = 1;
    const id = String(r.ID_Murid);
    (rekod[id] = rekod[id] || {})[tr] = statusH_(r.Status);
  });
  return { bulan: bulan, idKelas: idKelas, hari: Object.keys(hari).sort(), rekod: rekod };
}

/* ---------- Analisis kehadiran seorang murid (setahun) ---------- */
function apiKehadiranMurid_(p) {
  const id = String(p.idMurid || '');
  const tahun = /^\d{4}$/.test(String(p.tahun || '')) ? String(p.tahun) : today_().slice(0, 4);
  const rekod = kehadiranTahun_(tahun)
    .filter(r => String(r.ID_Murid) === id && fmtTarikh_(r.Tarikh).slice(0, 4) === tahun)
    .map(r => ({ tarikh: fmtTarikh_(r.Tarikh), status: statusH_(r.Status), sebab: String(r.Sebab || ''), idKelas: String(r.ID_Kelas) }))
    .sort((a, b) => a.tarikh < b.tarikh ? -1 : 1);
  return { idMurid: id, tahun: tahun, rekod: rekod };
}

// Dasar sekolah: SEMUA guru boleh merekod kehadiran mana-mana kelas (cth guru ganti) — tiada semakan peranan disengajakan
function apiSimpanKehadiran_(p, u) {
  const tarikh = semakTarikh_(p.tarikh);
  if (tarikh > today_()) throw err_('Tidak boleh merekod kehadiran untuk tarikh akan datang.');
  if (tahunDiarkib_(tarikh.slice(0, 4))) throw err_('Rekod kehadiran tahun ' + tarikh.slice(0, 4) + ' telah diarkib dan tidak boleh diubah.');
  const idKelas = String(p.idKelas || '');
  const senarai = Array.isArray(p.rekod) ? p.rekod : [];
  if (!senarai.length) throw err_('Tiada rekod kehadiran dihantar.');

  const muridKelas = new Set(read_('tbl_murid').rows
    .filter(m => String(m.ID_Kelas) === idKelas && String(m.Status) === 'Aktif')
    .map(m => String(m.ID_Murid)));

  const tb = read_('tbl_kehadiran');
  const sedia = new Map();
  tb.rows.forEach(r => {
    if (fmtTarikh_(r.Tarikh) === tarikh) sedia.set(String(r.ID_Murid), r);
  });

  const now = new Date();
  const baru = [], ubah = [];
  senarai.forEach(x => {
    const id = String(x.idMurid);
    if (!muridKelas.has(id)) throw err_('Murid ' + id + ' bukan dalam kelas ini.');
    if (STATUS_KEHADIRAN.indexOf(x.status) < 0) throw err_('Status kehadiran tidak sah: ' + x.status);
    const sebab = String(x.sebab || '').slice(0, 300);
    const lama = sedia.get(id);
    if (lama) {
      if (lama.Status !== x.status || String(lama.Sebab) !== sebab || String(lama.ID_Kelas) !== idKelas) {
        lama.Status = x.status; lama.Sebab = sebab; lama.ID_Kelas = idKelas;
        lama.ID_Guru_Perekod = u.id; lama.Timestamp = now;
        ubah.push(lama);
      }
    } else {
      baru.push({
        ID_Rekod: genId_('KH'), Tarikh: tarikh, ID_Murid: id, ID_Kelas: idKelas,
        Status: x.status, Sebab: sebab, ID_Guru_Perekod: u.id, Timestamp: now
      });
    }
  });
  updateRows_('tbl_kehadiran', ubah);
  insert_('tbl_kehadiran', baru);
  CacheService.getScriptCache().removeAll(['dashboard', 'ah2:' + tarikh]);
  audit_(u.id, 'SIMPAN_KEHADIRAN', `${idKelas} ${tarikh} baru=${baru.length} kemas=${ubah.length}`);
  return { baru: baru.length, dikemaskini: ubah.length };
}

/* ---------- Disiplin & Merit ---------- */
// Dasar sekolah: SEMUA guru boleh merekod merit/demerit untuk mana-mana murid; hanya mengubah nilai mata
// lalai kategori memerlukan kebenaran ubah_mata_disiplin
function apiSimpanDisiplin_(p, u) {
  const t = tetapan_();
  const kat = read_('tbl_kategori_disiplin').rows.find(k => String(k.ID_Kategori) === String(p.idKategori));
  if (!kat || String(kat.Status_Aktif) !== 'Aktif') throw err_('Kategori tidak sah.');

  const tm = read_('tbl_murid');
  const m = tm.rows.find(r => String(r.ID_Murid) === String(p.idMurid) && String(r.Status) === 'Aktif');
  if (!m) throw err_('Murid tidak dijumpai.');

  let mata = Number(kat.Mata) || 0;
  if (p.mata !== undefined && p.mata !== null && p.mata !== '') {
    wajib_(u, 'ubah_mata_disiplin');
    mata = Math.abs(Number(p.mata)) || 0;
  }
  const merit = String(kat.Jenis) === 'Merit';
  const kesan = merit ? mata : -mata;
  const maks = Number(t.MERIT_MAKS) || 150;
  const baki = Math.max(0, Math.min(maks, (Number(m.Mata_Merit) || 0) + kesan));

  m.Mata_Merit = baki;
  updateRow_('tbl_murid', m);

  const idKes = genId_(merit ? 'MR' : 'DS');
  insert_('tbl_disiplin', [{
    ID_Kes: idKes, Tarikh_Masa: new Date(), ID_Murid: String(m.ID_Murid), Jenis: String(kat.Jenis),
    ID_Kategori: String(kat.ID_Kategori), Kategori: String(kat.Nama), Mata_Kesan: kesan,
    Baki_Selepas: baki, Deskripsi: String(p.deskripsi || '').slice(0, 1000),
    Foto_ID: idFail_(p.fotoId), ID_Guru_Lapor: u.id,
    Status_Tindakan: merit ? 'Selesai' : 'Baharu'
  }]);
  CacheService.getScriptCache().remove('dashboard');
  audit_(u.id, 'REKOD_DISIPLIN', `${idKes} ${m.ID_Murid} ${kesan}`);
  return { idKes: idKes, baki: baki, statusMerit: statusMerit_(baki, t) };
}

function apiSenaraiDisiplin_(p, u) {
  const kelasSaya = kelasSaya_(u);
  const semua = boleh_(u, 'lihat_sensitif_semua');
  // Imbas QR: guru yang mengimbas murid melihat semua rekod merit & salah laku murid itu (direkod dalam log audit)
  const imbas = !!p.imbas && !!p.idMurid;
  if (imbas) audit_(u.id, 'LIHAT_SEJARAH_IMBAS', String(p.idMurid));
  const murid = new Map(read_('tbl_murid').rows.map(m => [String(m.ID_Murid), m]));
  const guru = new Map(read_('tbl_guru').rows.map(g => [String(g.ID_Guru), String(g.Nama_Penuh)]));
  const had = Math.min(Number(p.had) || 50, p.idKelas || p.analisis ? 5000 : 500);
  // Tapisan pilihan: idKelas (murid yang kini dalam kelas itu, untuk tab Sahsiah & Disiplin) dan dari (YYYY-MM-DD)
  const idKelas = String(p.idKelas || ''), dari = /^\d{4}-\d{2}-\d{2}$/.test(String(p.dari || '')) ? String(p.dari) : '';

  return read_('tbl_disiplin').rows
    .filter(k => !p.idMurid || String(k.ID_Murid) === String(p.idMurid))
    .filter(k => { if (!idKelas) return true; const m = murid.get(String(k.ID_Murid)); return m && String(m.ID_Kelas) === idKelas; })
    .filter(k => !dari || fmtTarikh_(k.Tarikh_Masa) >= dari)
    .filter(k => {
      if (semua || imbas || String(k.ID_Guru_Lapor) === u.id) return true;
      const m = murid.get(String(k.ID_Murid));
      return m && kelasSaya.indexOf(String(m.ID_Kelas)) >= 0;
    })
    .sort((a, b) => new Date(b.Tarikh_Masa) - new Date(a.Tarikh_Masa))
    .slice(0, had)
    .map(k => {
      const m = murid.get(String(k.ID_Murid)) || {};
      return {
        idKes: String(k.ID_Kes), tarikhMasa: k.Tarikh_Masa, idMurid: String(k.ID_Murid),
        namaMurid: String(m.Nama || k.ID_Murid), namaKelas: String(m.Nama_Kelas || ''),
        jenis: String(k.Jenis), kategori: String(k.Kategori), mata: Number(k.Mata_Kesan),
        baki: Number(k.Baki_Selepas), deskripsi: String(k.Deskripsi), fotoId: String(k.Foto_ID || ''),
        pelapor: guru.get(String(k.ID_Guru_Lapor)) || String(k.ID_Guru_Lapor),
        statusTindakan: String(k.Status_Tindakan)
      };
    });
}

function apiTindakanKes_(p, u) {
  wajib_(u, 'tindakan_kes');
  const sah = ['Baharu', 'Dalam Siasatan', 'Kaunseling', 'Ibu Bapa Dipanggil', 'Selesai'];
  if (sah.indexOf(p.status) < 0) throw err_('Status tindakan tidak sah.');
  const k = read_('tbl_disiplin').rows.find(r => String(r.ID_Kes) === String(p.idKes));
  if (!k) throw err_('Kes tidak dijumpai.');
  k.Status_Tindakan = p.status;
  updateRow_('tbl_disiplin', k);
  audit_(u.id, 'TINDAKAN_KES', `${k.ID_Kes} → ${p.status}`);
  return { berjaya: true };
}

/* ---------- Profil murid (klik nama dalam Pengurusan Kelas) ---------- */
function apiProfilMurid_(p, u) {
  const t = tetapan_();
  const m = read_('tbl_murid').rows.find(r => String(r.ID_Murid) === String(p.idMurid));
  if (!m) throw err_('Murid tidak dijumpai.');
  const penuh = bolehSensitif_(u, m.ID_Kelas);
  const merit = Number(m.Mata_Merit);

  const penjaga = [];
  const p1 = String(m.Nama_Penjaga_1 || m.Nama_Penjaga || ''), p2 = String(m.Nama_Penjaga_2 || '');
  if (p1) penjaga.push({ nama: p1, hubungan: String(m.Hubungan_Penjaga_1 || m.Hubungan_Penjaga || ''), tel: String(m.Tel_Penjaga_1 || m.Tel_Penjaga || '') });
  if (penuh && p2) penjaga.push({ nama: p2, hubungan: String(m.Hubungan_Penjaga_2 || ''), tel: String(m.Tel_Penjaga_2 || '') });

  const out = {
    id: String(m.ID_Murid), nama: String(m.Nama), jantina: String(m.Jantina), tahun: String(m.Tahun),
    idKelas: String(m.ID_Kelas), namaKelas: String(m.Nama_Kelas), warganegara: String(m.Warganegara),
    oku: String(m.OKU), tag: String(m.Tag || ''), fotoId: String(m.Foto_ID || ''),
    merit: merit, statusMerit: statusMerit_(merit, t), penuh: penuh, penjaga: penjaga
  };
  if (!penuh) return out;

  Object.assign(out, {
    noKP: String(m.No_KP || ''), jenisKP: String(m.Jenis_KP || ''), tarikhLahir: String(m.Tarikh_Lahir || ''),
    negaraAsal: String(m.Negara_Asal || ''), kategoriOKU: String(m.Kategori_OKU || ''), yatim: String(m.Yatim || ''),
    kategoriPendapatan: String(m.Kategori_Pendapatan || ''), tanggungan: String(m.Tanggungan || ''),
    pendapatan: m.Pendapatan_Isi_Rumah === '' ? null : Number(m.Pendapatan_Isi_Rumah),
    perKapita: perKapita_(m, t).perKapita, bilIsiRumah: perKapita_(m, t).bil
  });

  const a = read_('tbl_alamat_murid').rows.find(r => String(r.ID_Murid) === String(m.ID_Murid));
  if (a) {
    // Cari koordinat sekali sahaja, kemudian simpan (Geocoder terbina dalam Apps Script)
    const tiada = v => v === '' || v === null || v === undefined;
    if ((tiada(a.Lat) || tiada(a.Lng)) && teksAlamat_(a) && a.Lat !== undefined) {
      const g = geocode_(a);
      if (g) {
        a.Lat = String(g.lat); a.Lng = String(g.lng); a.Sumber_Lokasi = g.sumber;
        // Tulis di bawah lock dengan baris terkini (elak menimpa perubahan serentak, cth lokasi MANUAL)
        const lock = LockService.getScriptLock();
        if (lock.tryLock(10000)) {
          try {
            delete _MEMO['r:tbl_alamat_murid'];
            const kini = read_('tbl_alamat_murid').rows.find(r => String(r.ID_Murid) === String(m.ID_Murid));
            if (kini && (tiada(kini.Lat) || tiada(kini.Lng))) {
              Object.assign(kini, { Lat: a.Lat, Lng: a.Lng, Sumber_Lokasi: a.Sumber_Lokasi });
              updateRow_('tbl_alamat_murid', kini);
            }
          } finally { lock.releaseLock(); }
        }
      }
    }
    out.alamat = Object.assign(alamatObj_(a), {
      teks: teksAlamat_(a),
      lat: tiada(a.Lat) ? null : Number(a.Lat), lng: tiada(a.Lng) ? null : Number(a.Lng),
      sumber: String(a.Sumber_Lokasi || '')
    });
  }
  audit_(u.id, 'LIHAT_PROFIL', String(m.ID_Murid));
  return out;
}

function apiSimpanLokasi_(p, u) {
  const m = read_('tbl_murid').rows.find(r => String(r.ID_Murid) === String(p.idMurid));
  if (!m) throw err_('Murid tidak dijumpai.');
  if (!bolehSensitif_(u, m.ID_Kelas)) throw err_('Hanya guru kelas atau PK HEM boleh menetapkan lokasi rumah.', 'FORBIDDEN');
  const lat = Number(p.lat), lng = Number(p.lng);
  if (!(lat > 0.5 && lat < 7.6 && lng > 99.5 && lng < 119.5)) throw err_('Koordinat di luar Malaysia. Semak semula.');
  const ta = read_('tbl_alamat_murid');
  let a = ta.rows.find(r => String(r.ID_Murid) === String(m.ID_Murid));
  const data = { Lat: lat.toFixed(6), Lng: lng.toFixed(6), Sumber_Lokasi: 'MANUAL', Tarikh_Kemaskini: Utilities.formatDate(new Date(), APP.TZ, 'yyyy-MM-dd HH:mm') };
  if (a) { Object.assign(a, data); updateRow_('tbl_alamat_murid', a); }
  else insert_('tbl_alamat_murid', [Object.assign({ ID_Murid: String(m.ID_Murid) }, data)]);
  audit_(u.id, 'SIMPAN_LOKASI', String(m.ID_Murid));
  return { lat: Number(data.Lat), lng: Number(data.Lng), sumber: 'MANUAL' };
}

function teksAlamat_(a) {
  return [a.Alamat_1, a.Alamat_2, a.Alamat_3, [a.Poskod, a.Bandar].filter(Boolean).join(' '), a.Negeri]
    .map(x => String(x || '').trim()).filter(Boolean).join(', ');
}

function geocode_(a) {
  const cuba = [
    teksAlamat_(a),
    [a.Alamat_2, a.Alamat_3, a.Poskod, a.Bandar, a.Negeri].map(x => String(x || '').trim()).filter(Boolean).join(', '),
    [a.Poskod, a.Bandar, a.Negeri].map(x => String(x || '').trim()).filter(Boolean).join(', ')
  ].filter((v, i, arr) => v && arr.indexOf(v) === i);
  for (let i = 0; i < cuba.length; i++) {
    try {
      const r = Maps.newGeocoder().setRegion('my').geocode(cuba[i] + ', Malaysia');
      if (r && r.status === 'OK' && r.results.length) {
        const loc = r.results[0].geometry.location;
        const jenis = r.results[0].geometry.location_type;   // ROOFTOP = tepat
        return { lat: loc.lat, lng: loc.lng, sumber: i === 0 && jenis === 'ROOFTOP' ? 'AUTO_TEPAT' : i === 0 ? 'AUTO_ANGGARAN' : 'AUTO_KAWASAN' };
      }
    } catch (e) { return null; }   // kuota harian habis / ralat perkhidmatan
  }
  return null;
}

function apiGetAlamat_(p, u) {
  const m = read_('tbl_murid').rows.find(r => String(r.ID_Murid) === String(p.idMurid));
  if (!m) throw err_('Murid tidak dijumpai.');
  if (!bolehSensitif_(u, m.ID_Kelas)) throw err_('Anda tiada kebenaran melihat alamat murid ini.', 'FORBIDDEN');
  const a = read_('tbl_alamat_murid').rows.find(r => String(r.ID_Murid) === String(p.idMurid));
  audit_(u.id, 'LIHAT_ALAMAT', String(p.idMurid));
  return a ? alamatObj_(a) : null;
}

function apiSuratAmaran_(p, u) {
  wajib_(u, 'surat_amaran');
  const t = tetapan_();
  const m = read_('tbl_murid').rows.find(r => String(r.ID_Murid) === String(p.idMurid));
  if (!m) throw err_('Murid tidak dijumpai.');
  const a = read_('tbl_alamat_murid').rows.find(r => String(r.ID_Murid) === String(p.idMurid));
  const kes = read_('tbl_disiplin').rows
    .filter(k => String(k.ID_Murid) === String(p.idMurid) && Number(k.Mata_Kesan) < 0)
    .sort((x, y) => new Date(x.Tarikh_Masa) - new Date(y.Tarikh_Masa))
    .map(k => ({ tarikh: fmtTarikh_(k.Tarikh_Masa), kategori: String(k.Kategori), mata: Number(k.Mata_Kesan), deskripsi: String(k.Deskripsi) }));
  const tahun = today_().slice(0, 4);
  audit_(u.id, 'JANA_SURAT', String(m.ID_Murid));
  return {
    sekolah: {
      nama: t.NAMA_SEKOLAH || '', alamat: t.ALAMAT_SEKOLAH || '', tel: t.TEL_SEKOLAH || '',
      penandatangan: t.NAMA_PENANDATANGAN || '', jawatan: t.JAWATAN_PENANDATANGAN || ''
    },
    noRujukan: `${t.RUJUKAN_SURAT || 'HEM/DIS'}/${tahun}/${String(m.ID_Murid).slice(-6)}`,
    tarikh: today_(),
    murid: {
      id: String(m.ID_Murid), nama: String(m.Nama), kelas: String(m.Nama_Kelas),
      baki: Number(m.Mata_Merit), meritAwal: Number(t.MERIT_AWAL) || 100
    },
    penjaga: { nama: String(m.Nama_Penjaga), hubungan: String(m.Hubungan_Penjaga) },
    alamat: a ? alamatObj_(a) : null,
    kes: kes
  };
}

function alamatObj_(a) {
  return {
    alamat1: String(a.Alamat_1), alamat2: String(a.Alamat_2), alamat3: String(a.Alamat_3),
    poskod: String(a.Poskod), bandar: String(a.Bandar), daerah: String(a.Daerah), negeri: String(a.Negeri)
  };
}

/* ---------- Kebersihan 3K ---------- */
// Semua guru boleh menilai. Hanya guru yang mengisi sesuatu penilaian boleh mengubahnya.
function apiSimpanKebersihan_(p, u) {
  const s = p.skor || {};
  const skor = ['lantai', 'meja', 'sudut', 'bacaan'].map(k => {
    const v = Math.round(Number(s[k]));
    if (isNaN(v) || v < 0 || v > 25) throw err_('Skor ' + k + ' mesti 0 hingga 25.');
    return v;
  });
  const jumlah = skor.reduce((a, b) => a + b, 0);
  const tb = read_('tbl_kebersihan');

  // Tentukan rekod: edit (idNilai), atau penilaian sedia ada guru ini bagi kelas & tarikh yang sama
  let rec = null, tarikh, kelas;
  if (p.idNilai) {
    rec = tb.rows.find(r => String(r.ID_Nilai) === String(p.idNilai));
    if (!rec) throw err_('Penilaian tidak dijumpai.');
    if (String(rec.ID_Guru_Penilai) !== u.id) throw err_('Hanya guru yang mengisi penilaian ini boleh mengubahnya.', 'FORBIDDEN');
    tarikh = fmtTarikh_(rec.Tarikh);
  } else {
    kelas = read_('tbl_kelas').rows.find(k => String(k.ID_Kelas) === String(p.idKelas));
    if (!kelas) throw err_('Kelas tidak dijumpai.');
    tarikh = p.tarikh ? semakTarikh_(p.tarikh) : today_();
    if (tarikh > today_()) throw err_('Tidak boleh menilai untuk tarikh akan datang.');
    rec = tb.rows.find(r => String(r.ID_Guru_Penilai) === u.id && String(r.ID_Kelas) === String(kelas.ID_Kelas) && fmtTarikh_(r.Tarikh) === tarikh) || null;
  }

  const data = {
    Skor_Lantai: skor[0], Skor_Meja: skor[1], Skor_Sudut: skor[2], Skor_Bacaan: skor[3],
    Jumlah_Skor: jumlah, Timestamp: new Date()
  };
  if (p.fotoId !== undefined) data.Foto_ID = idFail_(p.fotoId);

  let id, dikemaskini = false;
  if (rec) {
    Object.assign(rec, data);
    updateRow_('tbl_kebersihan', rec);
    id = String(rec.ID_Nilai); dikemaskini = true;
  } else {
    id = genId_('KB');
    insert_('tbl_kebersihan', [Object.assign({
      ID_Nilai: id, Tarikh: tarikh, Bulan: tarikh.slice(0, 7),
      Minggu_Ke: Number(Utilities.formatDate(parseTarikh_(tarikh), APP.TZ, 'w')),
      ID_Kelas: String(kelas.ID_Kelas), Foto_ID: '', ID_Guru_Penilai: u.id
    }, data)]);
  }
  CacheService.getScriptCache().remove('dashboard');
  audit_(u.id, dikemaskini ? 'UBAH_KEBERSIHAN' : 'NILAI_KEBERSIHAN', id);
  return { idNilai: id, jumlah: jumlah, dikemaskini: dikemaskini };
}

// Senarai penilaian sebulan: semua guru boleh lihat; "milikSaya" menentukan butang ubah
function apiSenaraiKebersihan_(p, u) {
  const bulan = /^\d{4}-\d{2}$/.test(String(p.bulan || '')) ? p.bulan : today_().slice(0, 7);
  const kelas = new Map(read_('tbl_kelas').rows.map(k => [String(k.ID_Kelas), k]));
  const guru = new Map(read_('tbl_guru').rows.map(g => [String(g.ID_Guru), String(g.Nama_Penuh)]));
  return read_('tbl_kebersihan').rows
    .filter(r => String(r.Bulan) === bulan)
    .sort((a, b) => fmtTarikh_(b.Tarikh).localeCompare(fmtTarikh_(a.Tarikh)) || new Date(b.Timestamp) - new Date(a.Timestamp))
    .map(r => {
      const k = kelas.get(String(r.ID_Kelas)) || {};
      return {
        idNilai: String(r.ID_Nilai), tarikh: fmtTarikh_(r.Tarikh), idKelas: String(r.ID_Kelas),
        namaKelas: String(k.Nama_Kelas || r.ID_Kelas), tahun: String(k.Tahun || ''), sesi: String(k.Sesi || ''),
        skor: { lantai: Number(r.Skor_Lantai), meja: Number(r.Skor_Meja), sudut: Number(r.Skor_Sudut), bacaan: Number(r.Skor_Bacaan) },
        jumlah: Number(r.Jumlah_Skor), fotoId: String(r.Foto_ID || ''),
        penilai: guru.get(String(r.ID_Guru_Penilai)) || String(r.ID_Guru_Penilai),
        milikSaya: String(r.ID_Guru_Penilai) === u.id, masa: r.Timestamp
      };
    });
}

function apiKedudukanKebersihan_(p) {
  const bulan = /^\d{4}-\d{2}$/.test(String(p.bulan || '')) ? p.bulan : today_().slice(0, 7);
  const kelas = new Map(read_('tbl_kelas').rows.map(k => [String(k.ID_Kelas), k]));
  const guru = new Map(read_('tbl_guru').rows.map(g => [String(g.ID_Guru), String(g.Nama_Penuh)]));
  const agg = {};
  read_('tbl_kebersihan').rows.forEach(r => {
    if (String(r.Bulan) !== bulan) return;
    const id = String(r.ID_Kelas);
    agg[id] = agg[id] || { jumlah: 0, bil: 0, akhir: null, penilai: '' };
    agg[id].jumlah += Number(r.Jumlah_Skor) || 0;
    agg[id].bil++;
    if (!agg[id].akhir || new Date(r.Timestamp) > new Date(agg[id].akhir)) {
      agg[id].akhir = r.Timestamp; agg[id].penilai = guru.get(String(r.ID_Guru_Penilai)) || '';
    }
  });
  const gred = v => v >= 90 ? 'A (Cemerlang)' : v >= 80 ? 'B (Baik)' : v >= 65 ? 'C (Sederhana)' : 'D (Perlu Tindakan)';
  return {
    bulan: bulan,
    kedudukan: Object.keys(agg).map(id => {
      const k = kelas.get(id) || {};
      const purata = Math.round(agg[id].jumlah / agg[id].bil * 10) / 10;
      return {
        idKelas: id, nama: String(k.Nama_Kelas || id), sesi: String(k.Sesi || ''),
        purata: purata, bilPenilaian: agg[id].bil, gred: gred(purata), penilaiTerakhir: agg[id].penilai
      };
    }).sort((a, b) => b.purata - a.purata)
  };
}

/* ---------- Bantuan & Kebajikan ---------- */
// Person In Charge: guru yang dilantik mengurus satu skim bantuan (penerima, status, had kelayakan skim itu)
function picSkim_(u, k) { return !!k && String(k.ID_Guru_PIC || '') === u.id; }
function wajibSkim_(u, k) {
  if (!boleh_(u, 'urus_bantuan') && !picSkim_(u, k)) throw err_('Hanya PIC skim ini atau PK HEM boleh membuat tindakan ini.', 'FORBIDDEN');
}

function apiSenaraiAgihan_(p, u) {
  const semua = boleh_(u, 'lihat_sensitif_semua') || boleh_(u, 'urus_bantuan');
  const kelasSaya = kelasSaya_(u);
  const kat = new Map(read_('tbl_kategori_bantuan').rows.map(k => [String(k.ID_Kategori), k]));
  const skimSaya = new Set(Array.from(kat.values()).filter(k => picSkim_(u, k)).map(k => String(k.ID_Kategori)));
  if (!semua && !kelasSaya.length && !skimSaya.size && !rmtAkses_(u).bertugas) throw err_('Anda tiada kebenaran melihat senarai bantuan.', 'FORBIDDEN');

  const tahun = String(p.tahun || today_().slice(0, 4));
  const t = tetapan_();
  const murid = new Map(read_('tbl_murid').rows.map(m => [String(m.ID_Murid), m]));

  return read_('tbl_agihan_bantuan').rows
    .filter(a => String(a.Tahun) === tahun)
    .filter(a => !p.idKategori || String(a.ID_Kategori) === String(p.idKategori))
    .filter(a => {
      if (semua || skimSaya.has(String(a.ID_Kategori))) return true;
      const m = murid.get(String(a.ID_Murid));
      return m && kelasSaya.indexOf(String(m.ID_Kelas)) >= 0;
    })
    .map(a => {
      const m = murid.get(String(a.ID_Murid)) || {};
      const k = kat.get(String(a.ID_Kategori)) || {};
      return {
        idAgihan: String(a.ID_Agihan), idKategori: String(a.ID_Kategori), kod: String(k.Kod || ''),
        namaBantuan: String(k.Nama_Bantuan || ''), idMurid: String(a.ID_Murid),
        namaMurid: String(m.Nama || a.ID_Murid), namaKelas: String(m.Nama_Kelas || ''),
        kategoriPendapatan: String(m.Kategori_Pendapatan || ''),
        pendapatan: m.Pendapatan_Isi_Rumah === '' || m.Pendapatan_Isi_Rumah === undefined ? null : Number(m.Pendapatan_Isi_Rumah),
        perKapita: m.ID_Murid ? perKapita_(m, t).perKapita : null, warganegara: warganegara_(m),
        status: String(a.Status), tarikhSerahan: String(a.Tarikh_Serahan),
        amaun: a.Amaun === '' ? null : Number(a.Amaun), catatan: String(a.Catatan)
      };
    });
}

function apiSimpanAgihan_(p, u) {
  const sah = ['Dalam Proses', 'Layak', 'Tidak Layak', 'Selesai Agihan'];
  if (sah.indexOf(p.status) < 0) throw err_('Status agihan tidak sah.');
  if (p.tarikhSerahan) semakTarikh_(p.tarikhSerahan);

  const tb = read_('tbl_agihan_bantuan');
  if (p.idAgihan) {
    const a = tb.rows.find(r => String(r.ID_Agihan) === String(p.idAgihan));
    if (!a) throw err_('Rekod agihan tidak dijumpai.');
    wajibSkim_(u, read_('tbl_kategori_bantuan').rows.find(k => String(k.ID_Kategori) === String(a.ID_Kategori)));
    a.Status = p.status;
    a.Tarikh_Serahan = String(p.tarikhSerahan || '');
    a.Amaun = p.amaun === undefined || p.amaun === null || p.amaun === '' ? '' : Number(p.amaun);
    a.Catatan = String(p.catatan || '').slice(0, 500);
    a.ID_Guru_Rekod = u.id; a.Timestamp = new Date();
    updateRow_('tbl_agihan_bantuan', a);
    CacheService.getScriptCache().remove('dashboard');
    audit_(u.id, 'KEMASKINI_AGIHAN', String(a.ID_Agihan));
    return { idAgihan: String(a.ID_Agihan) };
  }

  const kat = read_('tbl_kategori_bantuan').rows.find(k => String(k.ID_Kategori) === String(p.idKategori));
  if (!kat) throw err_('Kategori bantuan tidak sah.');
  wajibSkim_(u, kat);
  const m = read_('tbl_murid').rows.find(r => String(r.ID_Murid) === String(p.idMurid));
  if (!m) throw err_('Murid tidak dijumpai.');
  if (!warganegara_(m)) throw err_('Bantuan hanya untuk murid warganegara.');
  const tahun = today_().slice(0, 4);
  if (tb.rows.some(r => String(r.ID_Kategori) === String(p.idKategori) &&
      String(r.ID_Murid) === String(p.idMurid) && String(r.Tahun) === tahun)) {
    throw err_('Murid ini sudah didaftarkan untuk bantuan tersebut tahun ini.');
  }
  const id = genId_('AG');
  insert_('tbl_agihan_bantuan', [{
    ID_Agihan: id, ID_Kategori: String(kat.ID_Kategori), ID_Murid: String(m.ID_Murid), Tahun: tahun,
    Status: p.status, Tarikh_Serahan: String(p.tarikhSerahan || ''),
    Amaun: p.amaun === undefined || p.amaun === null || p.amaun === '' ? '' : Number(p.amaun),
    Catatan: String(p.catatan || '').slice(0, 500), ID_Guru_Rekod: u.id, Timestamp: new Date()
  }]);
  CacheService.getScriptCache().remove('dashboard');
  audit_(u.id, 'DAFTAR_AGIHAN', id);
  return { idAgihan: id };
}

// Tambah skim baharu, atau kemas kini skim sedia ada jika p.id diberi (termasuk had kelayakan)
// PIC boleh mengubah skimnya sendiri; hanya PK HEM / Super Admin boleh menambah skim dan melantik PIC
function apiTambahKategoriBantuan_(p, u) {
  const admin = boleh_(u, 'urus_bantuan');
  if (p.id) wajibSkim_(u, read_('tbl_kategori_bantuan').rows.find(r => String(r.ID_Kategori) === String(p.id)));
  else wajib_(u, 'urus_bantuan');
  pastiLajur_('tbl_kategori_bantuan');           // lajur Had_* / PIC mungkin belum wujud pada pemasangan lama
  const nama = String(p.nama || '').trim();
  const kod = String(p.kod || '').trim().toUpperCase();
  if (!nama || !kod) throw err_('Sila isi kod dan nama bantuan.');
  const had = v => {
    if (v === undefined || v === null || v === '') return '';
    const n = Number(v);
    if (isNaN(n) || n < 0 || n > 1e7) throw err_('Had pendapatan mesti nombor positif.');
    return Math.round(n * 100) / 100;
  };
  const data = {
    Kod: kod, Nama_Bantuan: nama, Pemberi: String(p.pemberi || '').slice(0, 100), Kriteria: String(p.kriteria || '').slice(0, 500),
    Had_Isi_Rumah: had(p.hadIsiRumah), Had_Per_Kapita: had(p.hadPerKapita), Syarat_Had: p.syarat === 'SALAH_SATU' ? 'SALAH_SATU' : 'KEDUA'
  };
  if (admin && p.idPIC !== undefined) {
    const idPIC = String(p.idPIC || '');
    if (idPIC && !read_('tbl_guru').rows.some(g => String(g.ID_Guru) === idPIC && String(g.Status_Aktif) === 'Aktif')) throw err_('PIC tidak dijumpai atau tidak aktif.');
    data.ID_Guru_PIC = idPIC;
  }
  const rows = read_('tbl_kategori_bantuan').rows;
  if (rows.some(r => String(r.Kod).toUpperCase() === kod && String(r.ID_Kategori) !== String(p.id || ''))) throw err_('Kod bantuan sudah wujud.');
  if (p.id) {
    const k = rows.find(r => String(r.ID_Kategori) === String(p.id));
    if (!k) throw err_('Skim bantuan tidak dijumpai.');
    Object.assign(k, data);
    updateRow_('tbl_kategori_bantuan', k);
    audit_(u.id, 'KEMASKINI_KATEGORI_BANTUAN', `${k.ID_Kategori} ${kod}`);
    return { id: String(k.ID_Kategori) };
  }
  const id = nextId_(rows.map(r => r.ID_Kategori), 'KB', 2);
  insert_('tbl_kategori_bantuan', [Object.assign({ ID_Kategori: id, Status_Aktif: 'Aktif' }, data)]);
  audit_(u.id, 'TAMBAH_KATEGORI_BANTUAN', `${id} ${kod}`);
  return { id: id };
}

// Daftar ramai murid yang layak sekali gus. Murid yang tidak memenuhi had skim ditolak di pelayan.
function apiDaftarAgihanPukal_(p, u) {
  const t = tetapan_();
  const sah = ['Dalam Proses', 'Layak'];
  const status = sah.indexOf(p.status) >= 0 ? p.status : 'Layak';
  const kat = read_('tbl_kategori_bantuan').rows.find(k => String(k.ID_Kategori) === String(p.idKategori));
  if (!kat) throw err_('Kategori bantuan tidak sah.');
  wajibSkim_(u, kat);
  const ids = Array.from(new Set((Array.isArray(p.idMurid) ? p.idMurid : []).map(String))).slice(0, 2000);
  if (!ids.length) throw err_('Tiada murid dipilih.');
  const murid = new Map(read_('tbl_murid').rows.filter(m => String(m.Status) === 'Aktif').map(m => [String(m.ID_Murid), m]));
  const tahun = today_().slice(0, 4);
  const sudah = new Set(read_('tbl_agihan_bantuan').rows
    .filter(r => String(r.ID_Kategori) === String(kat.ID_Kategori) && String(r.Tahun) === tahun).map(r => String(r.ID_Murid)));
  const now = new Date(), baru = [];
  let ditolak = 0;
  ids.forEach(id => {
    const m = murid.get(id);
    if (!m || sudah.has(id) || !warganegara_(m)) { ditolak++; return; }
    const pend = m.Pendapatan_Isi_Rumah === '' ? null : Number(m.Pendapatan_Isi_Rumah);
    if (layakSkim_(pend, perKapita_(m, t).perKapita, kat) === false) { ditolak++; return; }
    baru.push({
      ID_Agihan: genId_('AG'), ID_Kategori: String(kat.ID_Kategori), ID_Murid: id, Tahun: tahun, Status: status,
      Tarikh_Serahan: '', Amaun: '', Catatan: 'Didaftar ikut had kelayakan', ID_Guru_Rekod: u.id, Timestamp: now
    });
    sudah.add(id);
  });
  insert_('tbl_agihan_bantuan', baru);
  CacheService.getScriptCache().remove('dashboard');
  audit_(u.id, 'DAFTAR_AGIHAN_PUKAL', `${kat.Kod} ${baru.length} murid`);
  return { didaftar: baru.length, dilangkau: ditolak };
}

/* ---------- RMT: kehadiran harian & laporan harian (Buku Panduan RMT: borang C8, C9) ----------
 * Boleh merekod: PIC RMT, PK HEM / Super Admin, dan guru bertugas bagi hari tersebut.
 * Jadual guru bertugas & menu ikut hari (C3, C5) disimpan dalam tetapan RMT_TETAPAN. */
const SHEET_RMT = 'tbl_rmt_harian';
const HARI_RMT = ['Isnin', 'Selasa', 'Rabu', 'Khamis', 'Jumaat'];

function skimRmt_() { return read_('tbl_kategori_bantuan').rows.find(k => String(k.Kod).toUpperCase() === 'RMT' && String(k.Status_Aktif) === 'Aktif'); }
function tetapanRmt_() { const o = parseJSON_(tetapan_().RMT_TETAPAN, null) || {}; return { bertugas: o.bertugas || {}, menu: o.menu || {} }; }
function hariTarikh_(t) {
  const [y, m, d] = String(t).split('-').map(Number);
  return ['Ahad', 'Isnin', 'Selasa', 'Rabu', 'Khamis', 'Jumaat', 'Sabtu'][new Date(Date.UTC(y, m - 1, d)).getUTCDay()];
}
function rmtAkses_(u) {
  const k = skimRmt_();
  if (!k) return { urus: false, bertugas: false };
  const b = tetapanRmt_().bertugas, isnin = isninMinggu_(today_());
  return { urus: boleh_(u, 'urus_bantuan') || picSkim_(u, k),
    bertugas: Object.keys(b).some(h => b[h] === u.id) || HARI_RMT.some((h, i) => guruBertugasRmt_(tambahHari_(isnin, i)) === u.id) };
}
function wajibRmt_(u) {
  if (!skimRmt_()) throw err_('Skim bantuan berkod RMT tidak dijumpai.');
  const a = rmtAkses_(u);
  if (!a.urus && !a.bertugas) throw err_('Hanya PIC RMT, guru bertugas RMT atau PK HEM boleh membuka rekod ini.', 'FORBIDDEN');
  return a;
}
function adaSheetRmt_() { return !!SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_RMT); }
function barisRmt_(tarikh) { return adaSheetRmt_() ? read_(SHEET_RMT).rows.find(r => fmtTarikh_(r.Tarikh) === tarikh) || null : null; }

// Penerima RMT tahun itu: status Layak / Selesai Agihan, murid aktif yang ada maklumat pendapatan
function penerimaRmt_(tahun) {
  const k = skimRmt_();
  const ids = new Set(read_('tbl_agihan_bantuan').rows.filter(a => String(a.ID_Kategori) === String(k.ID_Kategori) && String(a.Tahun) === tahun &&
    ['Layak', 'Selesai Agihan'].indexOf(String(a.Status)) >= 0).map(a => String(a.ID_Murid)));
  return read_('tbl_murid').rows.filter(m => ids.has(String(m.ID_Murid)) && String(m.Status) === 'Aktif' && layakBantuan_(m, k))
    .sort((a, b) => tahunSort_(a.Tahun) - tahunSort_(b.Tahun) || String(a.Nama_Kelas).localeCompare(String(b.Nama_Kelas)) || String(a.Nama).localeCompare(String(b.Nama)));
}

function apiGetRmt_(p, u) {
  const a = wajibRmt_(u);
  const tarikh = p.tarikh ? semakTarikh_(p.tarikh) : today_();
  const ts = tetapanRmt_(), hari = hariTarikh_(tarikh), r = barisRmt_(tarikh);
  const sekolah = {};
  kehadiranTahun_(tarikh.slice(0, 4)).forEach(x => { if (fmtTarikh_(x.Tarikh) === tarikh) sekolah[String(x.ID_Murid)] = statusH_(x.Status); });
  const guru = new Map(read_('tbl_guru').rows.map(g => [String(g.ID_Guru), String(g.Nama_Penuh)]));
  const idGuru = r ? String(r.ID_Guru_Bertugas || '') : guruBertugasRmt_(tarikh);
  return {
    tarikh: tarikh, hari: hari, menu: r ? String(r.Menu || '') : String(ts.menu[hari] || ''),
    idGuruBertugas: idGuru, namaGuruBertugas: guru.get(idGuru) || '',
    penerima: penerimaRmt_(tarikh.slice(0, 4)).map(m => ({ id: String(m.ID_Murid), nama: String(m.Nama), kelas: String(m.Nama_Kelas), idKelas: String(m.ID_Kelas), sekolah: sekolah[String(m.ID_Murid)] || '' })),
    rekod: r ? parseJSON_(r.Rekod_JSON, null) : null, skor: r ? parseJSON_(r.Skor_JSON, null) : null,
    catatan: r ? String(r.Catatan || '') : '', tiadaRmt: r ? isTrue_(r.Tiada_RMT) : false,
    bilHadir: r ? Number(r.Bil_Hadir) || 0 : 0, urus: a.urus, bolehUbah: a.urus || idGuru === u.id, tetapan: ts
  };
}

// Simpan kehadiran dan/atau laporan bagi satu tarikh (medan yang tidak dihantar kekal)
function apiSimpanRmt_(p, u) {
  const a = wajibRmt_(u);
  const tarikh = semakTarikh_(p.tarikh);
  if (tarikh > today_()) throw err_('Tidak boleh merekod untuk tarikh akan datang.');
  const ts = tetapanRmt_(), r = barisRmt_(tarikh);
  const idGuruAsal = r ? String(r.ID_Guru_Bertugas || '') : guruBertugasRmt_(tarikh);
  if (!a.urus && idGuruAsal !== u.id) throw err_('Hanya guru bertugas hari ini, PIC RMT atau PK HEM boleh merekod.', 'FORBIDDEN');

  const data = { ID_Guru_Rekod: u.id, Timestamp: new Date() };
  if (p.menu !== undefined) data.Menu = String(p.menu).slice(0, 200);
  if (p.catatan !== undefined) data.Catatan = String(p.catatan).slice(0, 500);
  if (a.urus && p.idGuruBertugas !== undefined) {
    const id = String(p.idGuruBertugas || '');
    if (id && !read_('tbl_guru').rows.some(g => String(g.ID_Guru) === id && String(g.Status_Aktif) === 'Aktif')) throw err_('Guru bertugas tidak dijumpai.');
    data.ID_Guru_Bertugas = id;
  }
  const tiada = p.tiadaRmt !== undefined ? !!p.tiadaRmt : (r ? isTrue_(r.Tiada_RMT) : false);
  data.Tiada_RMT = tiada ? 'TRUE' : 'FALSE';
  if (tiada) { Object.assign(data, { Rekod_JSON: '{}', Skor_JSON: '', Bil_Hadir: 0, Bil_Penerima: 0 }); }
  else if (p.rekod && typeof p.rekod === 'object') {
    const sah = new Set(penerimaRmt_(tarikh.slice(0, 4)).map(m => String(m.ID_Murid)));
    Object.keys(r ? parseJSON_(r.Rekod_JSON, {}) : {}).forEach(id => sah.add(id));    // murid yang sudah direkod sebelum ini
    const rekod = {};
    Object.keys(p.rekod).forEach(id => {
      if (!sah.has(id)) throw err_('Murid ' + id + ' bukan penerima RMT.');
      rekod[id] = p.rekod[id] ? 1 : 0;
    });
    Object.assign(data, { Rekod_JSON: jsonSel_(rekod, 'Rekod kehadiran RMT'), Bil_Penerima: Object.keys(rekod).length, Bil_Hadir: Object.keys(rekod).filter(id => rekod[id]).length });
  }
  if (p.skor !== undefined && !tiada) {
    if (!r && data.Rekod_JSON === undefined) throw err_('Rekod kehadiran dahulu sebelum mengisi laporan.');
    const skor = (Array.isArray(p.skor) ? p.skor : []).map(Number);
    if (skor.length !== 4 || skor.some(n => [1, 2, 3, 4].indexOf(n) < 0)) throw err_('Skor keempat-empat kriteria mesti 1 hingga 4.');
    data.Skor_JSON = JSON.stringify(skor);
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!adaSheetRmt_()) ensureSheet_(ss, SHEET_RMT, SCHEMA[SHEET_RMT]);
  if (r) { Object.assign(r, data); updateRow_(SHEET_RMT, r); }
  else {
    insert_(SHEET_RMT, [Object.assign({ Tarikh: tarikh, Menu: String(ts.menu[hariTarikh_(tarikh)] || ''), ID_Guru_Bertugas: idGuruAsal }, data)]);
  }
  audit_(u.id, 'SIMPAN_RMT', tarikh);
  return apiGetRmt_({ tarikh: tarikh }, u);
}

// Rekod bagi satu julat tarikh: cetakan C8 (seminggu) & C9 (sebulan)
function apiJulatRmt_(p, u) {
  wajibRmt_(u);
  const dari = semakTarikh_(p.dari), hingga = semakTarikh_(p.hingga);
  if (hingga < dari) throw err_('Julat tarikh tidak sah.');
  const guru = new Map(read_('tbl_guru').rows.map(g => [String(g.ID_Guru), String(g.Nama_Penuh)]));
  const hari = adaSheetRmt_() ? read_(SHEET_RMT).rows.filter(r => { const t = fmtTarikh_(r.Tarikh); return t >= dari && t <= hingga; }).map(r => ({
    tarikh: fmtTarikh_(r.Tarikh), hari: hariTarikh_(fmtTarikh_(r.Tarikh)), menu: String(r.Menu || ''), guru: guru.get(String(r.ID_Guru_Bertugas)) || '',
    bilHadir: Number(r.Bil_Hadir) || 0, bilPenerima: Number(r.Bil_Penerima) || 0, rekod: parseJSON_(r.Rekod_JSON, {}), skor: parseJSON_(r.Skor_JSON, null),
    catatan: String(r.Catatan || ''), tiadaRmt: isTrue_(r.Tiada_RMT)
  })).sort((a, b) => a.tarikh < b.tarikh ? -1 : 1) : [];
  // Termasuk murid yang pernah direkod dalam julat ini walaupun bukan penerima lagi
  const senarai = penerimaRmt_(dari.slice(0, 4)), ada = new Set(senarai.map(m => String(m.ID_Murid)));
  const lain = new Set();
  hari.forEach(h => Object.keys(h.rekod).forEach(id => { if (!ada.has(id)) lain.add(id); }));
  const tambahan = read_('tbl_murid').rows.filter(m => lain.has(String(m.ID_Murid)));
  return { dari: dari, hingga: hingga, hari: hari, penerima: senarai.concat(tambahan).map(m => ({ id: String(m.ID_Murid), nama: String(m.Nama), kelas: String(m.Nama_Kelas) })) };
}

// Jadual guru bertugas & menu ikut hari (PIC / PK HEM)
function apiSimpanTetapanRmt_(p, u) {
  const k = skimRmt_();
  if (!k) throw err_('Skim bantuan berkod RMT tidak dijumpai.');
  wajibSkim_(u, k);
  ensureTetapan_();
  const aktif = new Set(read_('tbl_guru').rows.filter(g => String(g.Status_Aktif) === 'Aktif').map(g => String(g.ID_Guru)));
  const out = { bertugas: {}, menu: {} };
  HARI_RMT.forEach(h => {
    const id = String((p.bertugas || {})[h] || '');
    if (id && !aktif.has(id)) throw err_('Guru bertugas ' + h + ' tidak dijumpai.');
    if (id) out.bertugas[h] = id;
    const menu = String((p.menu || {})[h] || '').trim().slice(0, 200);
    if (menu) out.menu[h] = menu;
  });
  const r = read_('tbl_tetapan').rows.find(x => String(x.Kunci) === 'RMT_TETAPAN');
  r.Nilai = JSON.stringify(out);
  updateRow_('tbl_tetapan', r);
  audit_(u.id, 'TETAPAN_RMT', '');
  return out;
}

// Ringkasan untuk bootstrap: akses pengguna & sama ada rekod hari ini sudah diisi (peringatan Dashboard)
function rmtInfo_(u) {
  const a = rmtAkses_(u);
  if (!a.urus && !a.bertugas) return null;
  const t = today_(), hari = hariTarikh_(t), r = barisRmt_(t);
  return {
    urus: a.urus, pic: picSkim_(u, skimRmt_()), bertugas: a.bertugas, bertugasHariIni: guruBertugasRmt_(t) === u.id,
    hariSekolah: HARI_RMT.indexOf(hari) >= 0, hadirSudah: !!r && (isTrue_(r.Tiada_RMT) || !!parseJSON_(r.Rekod_JSON, null)),
    laporanSudah: !!r && (isTrue_(r.Tiada_RMT) || !!parseJSON_(r.Skor_JSON, null))
  };
}

/* ---------- Guru bertugas: jadual mingguan & laporan harian ----------
 * PK HEM menetapkan kumpulan guru bertugas bagi setiap minggu & sesi, dan seorang pelapor bagi setiap hari.
 * Pelapor mengisi laporan (Draf → Dihantar); PK HEM mengesahkan atau memulangkan untuk dibetulkan.
 * Pelapor hari itu juga menjadi guru bertugas RMT; jadual dalam tetapan RMT hanya digunakan jika hari itu tiada pelapor. */
const SHEET_BTG = 'tbl_bertugas_minggu', SHEET_LAP = 'tbl_laporan_bertugas';
const ITEM_PANTAU = ['ketibaan', 'lalulintas', 'kawasan', 'tandas', 'kantin', 'bersurai'];
const JENIS_KEJADIAN = ['Murid sakit', 'Kemalangan / kecederaan', 'Kerosakan harta benda', 'Pelawat', 'Lain-lain'];

function adaSheet_(n) { return !!SpreadsheetApp.getActiveSpreadsheet().getSheetByName(n); }
function pastiSheet_(n) { const ss = SpreadsheetApp.getActiveSpreadsheet(); if (!ss.getSheetByName(n)) ensureSheet_(ss, n, SCHEMA[n]); }
function tambahHari_(t, n) { const [y, m, d] = String(t).split('-').map(Number); return new Date(Date.UTC(y, m - 1, d + n)).toISOString().slice(0, 10); }
function isninMinggu_(t) {
  const [y, m, d] = String(t).split('-').map(Number);
  return tambahHari_(t, -((new Date(Date.UTC(y, m - 1, d)).getUTCDay() + 6) % 7));
}
function masaTeks_(v) { return v instanceof Date ? Utilities.formatDate(v, APP.TZ, 'yyyy-MM-dd HH:mm') : String(v || ''); }
function sesiSekolah_() {
  const s = [];
  read_('tbl_kelas').rows.forEach(k => { const v = String(k.Sesi || '').trim(); if (v && String(k.Status) !== 'Tidak Aktif' && s.indexOf(v) < 0) s.push(v); });
  return s.length ? s.sort() : ['Pagi'];
}
function semakSesi_(s) { const ada = sesiSekolah_(); return ada.indexOf(String(s || '')) >= 0 ? String(s) : ada[0]; }

function jadualMinggu_(minggu, sesi) {
  const r = adaSheet_(SHEET_BTG) ? read_(SHEET_BTG).rows.find(x => fmtTarikh_(x.Minggu_Mula) === minggu && String(x.Sesi) === sesi) : null;
  return { r: r, ahli: r ? parseJSON_(r.Ahli_JSON, []).map(String) : [], ketua: r ? String(r.ID_Ketua || '') : '', pelapor: r ? parseJSON_(r.Pelapor_JSON, {}) : {} };
}
function pelaporHari_(tarikh, sesi) { return String(jadualMinggu_(isninMinggu_(tarikh), sesi).pelapor[hariTarikh_(tarikh)] || ''); }
function guruBertugasRmt_(tarikh) {
  const sesi = sesiSekolah_();
  for (let i = 0; i < sesi.length; i++) { const id = pelaporHari_(tarikh, sesi[i]); if (id) return id; }
  return String(tetapanRmt_().bertugas[hariTarikh_(tarikh)] || '');
}
function barisLaporan_(tarikh, sesi) {
  return adaSheet_(SHEET_LAP) ? read_(SHEET_LAP).rows.find(r => fmtTarikh_(r.Tarikh) === tarikh && String(r.Sesi) === sesi) || null : null;
}
// Boleh membaca laporan: PK HEM, pelapor laporan itu, dan semua guru bertugas minggu tersebut
function aksesLaporan_(u, tarikh, sesi, r) {
  if (boleh_(u, 'urus_bertugas') || (r && String(r.ID_Pelapor) === u.id)) return true;
  return jadualMinggu_(isninMinggu_(tarikh), sesi).ahli.indexOf(u.id) >= 0;
}

function apiGetBertugas_(p, u) {
  const minggu = isninMinggu_(p.minggu ? semakTarikh_(p.minggu) : today_()), sesi = semakSesi_(p.sesi);
  const j = jadualMinggu_(minggu, sesi), akhir = tambahHari_(minggu, 6), laporan = {};
  if (adaSheet_(SHEET_LAP)) read_(SHEET_LAP).rows.forEach(r => {
    const t = fmtTarikh_(r.Tarikh);
    if (t >= minggu && t <= akhir && String(r.Sesi) === sesi) laporan[t] = String(r.Status);
  });
  return { minggu: minggu, sesi: sesi, senaraiSesi: sesiSekolah_(), ahli: j.ahli, ketua: j.ketua, pelapor: j.pelapor, laporan: laporan,
    ada: !!j.r, urus: boleh_(u, 'urus_bertugas'), hariIni: today_() };
}

function apiSimpanBertugas_(p, u) {
  wajib_(u, 'urus_bertugas');
  const minggu = isninMinggu_(semakTarikh_(p.minggu)), sesi = semakSesi_(p.sesi);
  const aktif = new Set(read_('tbl_guru').rows.filter(g => String(g.Status_Aktif) === 'Aktif').map(g => String(g.ID_Guru)));
  const ahli = [];
  (Array.isArray(p.ahli) ? p.ahli : []).map(String).forEach(id => {
    if (!aktif.has(id)) throw err_('Guru bertugas tidak dijumpai.');
    if (ahli.indexOf(id) < 0) ahli.push(id);
  });
  if (ahli.length > 15) throw err_('Terlalu ramai guru bertugas dalam seminggu (maksimum 15).');
  const ketua = String(p.ketua || '');
  if (ketua && ahli.indexOf(ketua) < 0) throw err_('Ketua mesti salah seorang guru bertugas minggu ini.');
  const pelapor = {};
  HARI_RMT.forEach(h => {
    const id = String((p.pelapor || {})[h] || '');
    if (!id) return;
    if (ahli.indexOf(id) < 0) throw err_('Pelapor hari ' + h + ' mesti salah seorang guru bertugas minggu ini.');
    pelapor[h] = id;
  });
  pastiSheet_(SHEET_BTG);
  const data = { Minggu_Mula: minggu, Sesi: sesi, Ahli_JSON: JSON.stringify(ahli), ID_Ketua: ketua, Pelapor_JSON: JSON.stringify(pelapor), ID_Guru_Rekod: u.id, Timestamp: new Date() };
  const r = jadualMinggu_(minggu, sesi).r;
  if (r) updateRow_(SHEET_BTG, Object.assign(r, data)); else insert_(SHEET_BTG, [data]);
  CacheService.getScriptCache().remove('dashboard');
  audit_(u.id, 'JADUAL_BERTUGAS', minggu + ' ' + sesi + ' ahli=' + ahli.length);
  return apiGetBertugas_({ minggu: minggu, sesi: sesi }, u);
}

// Data yang diambil terus daripada modul lain bagi tarikh & sesi itu
function autoLaporan_(tarikh, sesi) {
  const satuSesi = sesiSekolah_().length <= 1;
  const kelas = read_('tbl_kelas').rows.filter(k => String(k.Status) !== 'Tidak Aktif' && (satuSesi || String(k.Sesi) === sesi));
  const idKelas = new Set(kelas.map(k => String(k.ID_Kelas)));
  const murid = new Map(read_('tbl_murid').rows.filter(m => String(m.Status) === 'Aktif' && idKelas.has(String(m.ID_Kelas))).map(m => [String(m.ID_Murid), m]));
  let direkod = 0, hadir = 0;
  const lewat = [], kelasAda = new Set(), dikira = new Set();
  kehadiranTahun_(tarikh.slice(0, 4)).forEach(x => {
    const m = murid.get(String(x.ID_Murid));
    if (!m || fmtTarikh_(x.Tarikh) !== tarikh || dikira.has(String(x.ID_Murid))) return;
    dikira.add(String(x.ID_Murid)); direkod++; kelasAda.add(String(m.ID_Kelas));
    if (hadirH_(x.Status)) hadir++;
    if (statusH_(x.Status) === 'Lewat') lewat.push({ nama: String(m.Nama), kelas: String(m.Nama_Kelas) });
  });
  const adaMurid = new Set(Array.from(murid.values()).map(m => String(m.ID_Kelas)));
  const kelasBelum = kelas.filter(k => adaMurid.has(String(k.ID_Kelas)) && !kelasAda.has(String(k.ID_Kelas)))
    .sort((a, b) => tahunSort_(a.Tahun) - tahunSort_(b.Tahun) || String(a.Nama_Kelas).localeCompare(String(b.Nama_Kelas))).map(k => String(k.Nama_Kelas));
  const disiplin = {};
  let kes = 0, merit = 0;
  read_('tbl_disiplin').rows.forEach(k => {
    if (!murid.has(String(k.ID_Murid)) || fmtTarikh_(k.Tarikh_Masa) !== tarikh) return;
    if (String(k.Jenis) === 'Merit') { merit++; return; }
    kes++; const n = String(k.Kategori || 'Lain-lain'); disiplin[n] = (disiplin[n] || 0) + 1;
  });
  const r = barisRmt_(tarikh);
  return {
    murid: murid.size, direkod: direkod, hadir: hadir, lewat: lewat.slice(0, 100), kelasBelum: kelasBelum,
    kes: kes, disiplin: disiplin, merit: merit,
    rmt: r ? { tiada: isTrue_(r.Tiada_RMT), penerima: Number(r.Bil_Penerima) || 0, hadir: Number(r.Bil_Hadir) || 0, laporan: !!parseJSON_(r.Skor_JSON, null) } : null
  };
}

function bersihLaporan_(d) {
  d = d || {};
  const teks = (v, n) => String(v == null ? '' : v).trim().slice(0, n);
  const skala = {};
  ITEM_PANTAU.forEach(k => { const n = Number((d.skala || {})[k]); if ([1, 2, 3, 4].indexOf(n) >= 0) skala[k] = n; });
  const kejadian = (Array.isArray(d.kejadian) ? d.kejadian : []).slice(0, 20).map(x => ({
    jenis: JENIS_KEJADIAN.indexOf(String(x && x.jenis)) >= 0 ? String(x.jenis) : 'Lain-lain',
    masa: /^\d{1,2}:\d{2}$/.test(String(x && x.masa || '')) ? String(x.masa) : '',
    butiran: teks(x && x.butiran, 300), tindakan: teks(x && x.tindakan, 300)
  })).filter(x => x.butiran);
  return {
    guruTidakHadir: teks(d.guruTidakHadir, 300),
    perhimpunan: ['Rasmi', 'Ringkas', 'Tiada'].indexOf(String(d.perhimpunan)) >= 0 ? String(d.perhimpunan) : '',
    pengacara: teks(d.pengacara, 100), pengumuman: teks(d.pengumuman, 500),
    skala: skala, catatanPantau: teks(d.catatanPantau, 500), kejadian: kejadian,
    isu: teks(d.isu, 1000), cadangan: teks(d.cadangan, 500)
  };
}

function apiGetLaporanBertugas_(p, u) {
  const tarikh = p.tarikh ? semakTarikh_(p.tarikh) : today_(), sesi = semakSesi_(p.sesi);
  const r = barisLaporan_(tarikh, sesi);
  if (!aksesLaporan_(u, tarikh, sesi, r)) throw err_('Laporan ini hanya boleh dibuka oleh guru bertugas minggu tersebut dan PK HEM.', 'FORBIDDEN');
  const urus = boleh_(u, 'urus_bertugas'), pelapor = pelaporHari_(tarikh, sesi), j = jadualMinggu_(isninMinggu_(tarikh), sesi);
  const status = r ? String(r.Status) : '', terbuka = !r || status === 'Draf' || status === 'Dipulangkan';
  const nama = new Map(read_('tbl_guru').rows.map(g => [String(g.ID_Guru), String(g.Nama_Penuh)]));
  const data = r ? parseJSON_(r.Data_JSON, {}) : {};
  const idPelapor = r && r.ID_Pelapor ? String(r.ID_Pelapor) : pelapor;
  return {
    tarikh: tarikh, sesi: sesi, hari: hariTarikh_(tarikh), senaraiSesi: sesiSekolah_(),
    idPelapor: idPelapor, namaPelapor: nama.get(idPelapor) || '', ahli: j.ahli.map(id => nama.get(id) || id), ketua: nama.get(j.ketua) || '',
    status: status, data: bersihLaporan_(data), fotoId: r ? String(r.Foto_ID || '') : '',
    // Selepas dihantar, data automatik dibekukan seperti semasa laporan dihantar
    auto: !terbuka && data.auto ? data.auto : autoLaporan_(tarikh, sesi),
    ulasan: r ? String(r.Ulasan || '') : '', namaPengesah: r ? nama.get(String(r.ID_Pengesah || '')) || '' : '',
    tarikhHantar: r ? masaTeks_(r.Tarikh_Hantar) : '', tarikhSah: r ? masaTeks_(r.Tarikh_Sah) : '',
    bolehIsi: terbuka && tarikh <= today_() && (urus || (!!pelapor && pelapor === u.id)),
    bolehSah: urus && status === 'Dihantar', urus: urus, hariIni: today_()
  };
}

// hantar = true: hantar kepada PK HEM (laporan dikunci). Selainnya simpan sebagai draf.
function apiSimpanLaporanBertugas_(p, u) {
  const tarikh = semakTarikh_(p.tarikh), sesi = semakSesi_(p.sesi);
  if (tarikh > today_()) throw err_('Laporan tidak boleh diisi untuk tarikh akan datang.');
  const urus = boleh_(u, 'urus_bertugas'), pelapor = pelaporHari_(tarikh, sesi);
  if (!urus && (!pelapor || pelapor !== u.id)) throw err_('Hanya guru pelapor hari ini atau PK HEM boleh mengisi laporan ini.', 'FORBIDDEN');
  pastiSheet_(SHEET_LAP);
  const r = barisLaporan_(tarikh, sesi);
  if (r && ['Dihantar', 'Disahkan'].indexOf(String(r.Status)) >= 0) throw err_('Laporan ini sudah dihantar dan tidak boleh diubah lagi.');
  const data = bersihLaporan_(p.data), hantar = !!p.hantar, now = new Date();
  if (hantar) {
    if (!data.perhimpunan) throw err_('Pilih status perhimpunan sebelum menghantar.');
    if (ITEM_PANTAU.some(k => !data.skala[k])) throw err_('Lengkapkan semua item pemantauan (skala 1 hingga 4) sebelum menghantar.');
    data.auto = autoLaporan_(tarikh, sesi);
  }
  const o = { Tarikh: tarikh, Sesi: sesi, ID_Pelapor: pelapor || (r && String(r.ID_Pelapor)) || u.id,
    Status: hantar ? 'Dihantar' : (r ? String(r.Status) : 'Draf'), Data_JSON: jsonSel_(data, 'Laporan'), Timestamp: now };
  if (p.fotoId !== undefined) o.Foto_ID = idFail_(p.fotoId);
  if (hantar) o.Tarikh_Hantar = now;
  if (r) updateRow_(SHEET_LAP, Object.assign(r, o));
  else insert_(SHEET_LAP, [Object.assign({ ID_Laporan: genId_('LB'), Foto_ID: '', Ulasan: '', ID_Pengesah: '', Tarikh_Hantar: '', Tarikh_Sah: '' }, o)]);
  if (hantar) audit_(u.id, 'HANTAR_LAPORAN_BERTUGAS', tarikh + ' ' + sesi);
  return { status: o.Status };
}

function apiSahLaporanBertugas_(p, u) {
  wajib_(u, 'urus_bertugas');
  const tarikh = semakTarikh_(p.tarikh), sesi = semakSesi_(p.sesi), r = barisLaporan_(tarikh, sesi);
  if (!r || String(r.Status) !== 'Dihantar') throw err_('Hanya laporan yang sudah dihantar boleh disahkan atau dipulangkan.');
  const pulang = p.tindakan === 'pulang', ulasan = String(p.ulasan || '').trim().slice(0, 1000);
  if (pulang && !ulasan) throw err_('Nyatakan perkara yang perlu dibetulkan dalam ulasan.');
  Object.assign(r, { Status: pulang ? 'Dipulangkan' : 'Disahkan', Ulasan: ulasan, ID_Pengesah: u.id, Tarikh_Sah: pulang ? '' : new Date(), Timestamp: new Date() });
  updateRow_(SHEET_LAP, r);
  audit_(u.id, pulang ? 'PULANG_LAPORAN_BERTUGAS' : 'SAH_LAPORAN_BERTUGAS', tarikh + ' ' + sesi);
  return { status: r.Status };
}

// Laporan dalam julat tarikh (maksimum 62 hari) + hari yang laporannya belum dihantar.
// PK HEM melihat semua; guru lain hanya laporan minggu dia bertugas.
function apiSenaraiLaporanBertugas_(p, u) {
  const dari = semakTarikh_(p.dari), hingga = semakTarikh_(p.hingga);
  if (hingga < dari) throw err_('Tarikh akhir mesti selepas tarikh mula.');
  if (tambahHari_(dari, 62) < hingga) throw err_('Julat maksimum ialah 62 hari.');
  const nama = new Map(read_('tbl_guru').rows.map(g => [String(g.ID_Guru), String(g.Nama_Penuh)]));
  const lap = adaSheet_(SHEET_LAP) ? read_(SHEET_LAP).rows.filter(r => { const t = fmtTarikh_(r.Tarikh); return t >= dari && t <= hingga; }) : [];
  const siap = new Set(lap.filter(r => ['Dihantar', 'Disahkan'].indexOf(String(r.Status)) >= 0).map(r => fmtTarikh_(r.Tarikh) + '|' + r.Sesi));
  const laporan = lap.filter(r => aksesLaporan_(u, fmtTarikh_(r.Tarikh), String(r.Sesi), r)).map(r => {
    const d = parseJSON_(r.Data_JSON, {}), t = fmtTarikh_(r.Tarikh);
    return { tarikh: t, hari: hariTarikh_(t), sesi: String(r.Sesi), status: String(r.Status), namaPelapor: nama.get(String(r.ID_Pelapor)) || '',
      data: bersihLaporan_(d), auto: d.auto || null, fotoId: String(r.Foto_ID || ''), ulasan: String(r.Ulasan || ''),
      namaPengesah: nama.get(String(r.ID_Pengesah || '')) || '', tarikhHantar: masaTeks_(r.Tarikh_Hantar) };
  }).sort((a, b) => b.tarikh.localeCompare(a.tarikh) || a.sesi.localeCompare(b.sesi));
  const belum = [], hariIni = today_();
  if (boleh_(u, 'urus_bertugas') && adaSheet_(SHEET_BTG)) {
    read_(SHEET_BTG).rows.forEach(w => {
      const minggu = fmtTarikh_(w.Minggu_Mula), pel = parseJSON_(w.Pelapor_JSON, {});
      HARI_RMT.forEach((h, i) => {
        const t = tambahHari_(minggu, i);
        if (!pel[h] || t < dari || t > hingga || t >= hariIni || siap.has(t + '|' + w.Sesi)) return;
        belum.push({ tarikh: t, hari: h, sesi: String(w.Sesi), namaPelapor: nama.get(String(pel[h])) || '' });
      });
    });
  }
  return { laporan: laporan, belum: belum.sort((a, b) => b.tarikh.localeCompare(a.tarikh)) };
}

// Ringkasan untuk bootstrap (peringatan Dashboard)
function bertugasInfo_(u) {
  const urus = boleh_(u, 'urus_bertugas'), t = today_();
  const out = { urus: urus, tugas: [], dipulangkan: 0, menunggu: 0 };
  if (adaSheet_(SHEET_BTG) && HARI_RMT.indexOf(hariTarikh_(t)) >= 0) {
    out.tugas = sesiSekolah_().filter(s => pelaporHari_(t, s) === u.id).map(s => { const r = barisLaporan_(t, s); return { sesi: s, status: r ? String(r.Status) : '' }; });
  }
  if (adaSheet_(SHEET_LAP)) read_(SHEET_LAP).rows.forEach(r => {
    if (String(r.Status) === 'Dipulangkan' && String(r.ID_Pelapor) === u.id) out.dipulangkan++;
    if (urus && String(r.Status) === 'Dihantar') out.menunggu++;
  });
  return out;
}

/* ---------- Carta organisasi HEM (Dashboard) ----------
 * Disimpan sebagai JSON dalam tetapan CARTA_HEM. nodes = senarai rata; induk ditentukan oleh "parent",
 * susunan adik-beradik ikut susunan dalam senarai. Semua guru boleh melihat; PK HEM / Super Admin menyunting. */
const REKA_CARTA_HEM = ['ceria', 'konfeti', 'kartun', 'formal', 'eksekutif'];
function apiGetCartaHem_(p, u) {
  const c = parseJSON_(tetapan_().CARTA_HEM, null);
  return { carta: c && Array.isArray(c.nodes) ? c : null, urus: boleh_(u, 'urus_carta_hem') };
}

function apiSimpanCartaHem_(p, u) {
  wajib_(u, 'urus_carta_hem');
  const out = bersihCarta_(p.carta, 'Carta Organisasi Hal Ehwal Murid');
  ensureTetapan_();
  const r = read_('tbl_tetapan').rows.find(x => String(x.Kunci) === 'CARTA_HEM');
  r.Nilai = jsonSel_(out, 'Carta organisasi');
  updateRow_('tbl_tetapan', r);
  audit_(u.id, 'CARTA_HEM', 'kotak=' + out.nodes.length);
  return { carta: out, urus: true };
}

// Sahkan & bersihkan data carta organisasi (dikongsi oleh carta HEM dan carta setiap unit)
function bersihCarta_(c, tajukLalai) {
  c = c || {};
  const teks = (v, n) => String(v == null ? '' : v).trim().slice(0, n);
  const warna = v => /^#[0-9a-f]{6}$/i.test(String(v || '')) ? String(v).toLowerCase() : '';
  const guru = new Set(read_('tbl_guru').rows.map(g => String(g.ID_Guru)));
  const senarai = (Array.isArray(c.nodes) ? c.nodes : []);
  if (senarai.length > 100) throw err_('Carta terlalu besar (maksimum 100 kotak).');
  const ids = new Set();
  const nodes = senarai.map(n => {
    const id = teks(n && n.id, 20);
    if (!/^[\w-]{1,20}$/.test(id) || ids.has(id)) throw err_('ID kotak carta tidak sah.');
    ids.add(id);
    return {
      id: id, parent: teks(n.parent, 20), jawatan: teks(n.jawatan, 80), susun: n.susun === 'menegak' ? 'menegak' : 'mendatar',
      ahli: (Array.isArray(n.ahli) ? n.ahli : []).slice(0, 12).map(a => ({
        idGuru: guru.has(String(a && a.idGuru)) ? String(a.idGuru) : '', nama: teks(a && a.nama, 100), fotoId: idFail_(a && a.fotoId)
      })).filter(a => a.idGuru || a.nama)
    };
  });
  // Induk mesti wujud dan tiada kitaran (kotak tidak boleh menjadi anak kepada keturunannya sendiri)
  const induk = new Map(nodes.map(n => [n.id, n.parent]));
  nodes.forEach(n => {
    if (n.parent && !ids.has(n.parent)) n.parent = '';
    const dilawat = new Set([n.id]);
    let x = n.parent;
    while (x) { if (dilawat.has(x)) throw err_('Susunan carta tidak sah (kitaran).'); dilawat.add(x); x = induk.get(x) || ''; }
  });
  const out = {
    tajuk: teks(c.tajuk, 100) || tajukLalai, sub: teks(c.sub, 100),
    reka: REKA_CARTA_HEM.indexOf(String(c.reka)) >= 0 ? String(c.reka) : 'eksekutif', latar: warna(c.latar), kotak: warna(c.kotak),
    // Warna utama (tajuk, garisan & label), gambar latar sendiri dan lapisan cerah di atas gambar itu
    utama: warna(c.utama), latarFoto: idFail_(c.latarFoto), lapisan: [0, 0.2, 0.4, 0.6].indexOf(Number(c.lapisan)) >= 0 ? Number(c.lapisan) : 0.4,
    nodes: nodes, dikemaskini: Utilities.formatDate(new Date(), APP.TZ, 'yyyy-MM-dd HH:mm')
  };
  return out;
}

/* ---------- Unit HEM: carta organisasi, skop tugas & laporan aktiviti (OPR) setiap unit ----------
 * PK HEM / Super Admin mengurus unit dan melantik ketua unit. Ketua unit menyunting carta & skop tugas unitnya
 * dan mengesahkan OPR. Ahli unit (guru dalam carta unit) mengisi OPR. Semua guru melihat carta, skop & OPR yang disahkan.
 * Jadual unit dicipta beserta unit permulaan pada kali pertama data unit disimpan. */
const SHEET_UNIT = 'tbl_unit_hem', SHEET_OPR = 'tbl_opr';
const IKON_UNIT = ['fa-scale-balanced', 'fa-comments', 'fa-shield-halved', 'fa-star', 'fa-broom', 'fa-kit-medical', 'fa-mosquito', 'fa-ban',
  'fa-medal', 'fa-futbol', 'fa-book-open', 'fa-leaf', 'fa-people-group', 'fa-bullhorn', 'fa-heart-pulse', 'fa-hands-holding-child', 'fa-fire-extinguisher', 'fa-bus'];
const WARNA_UNIT = ['rose', 'violet', 'amber', 'sky', 'emerald', 'teal', 'lime', 'slate', 'orange', 'indigo', 'pink', 'cyan'];
const UNIT_LALAI = [
  ['U01', 'Disiplin & Pengawas', 'DISIPLIN', 'fa-scale-balanced', 'rose'], ['U02', 'Kaunseling', 'KAUNSELING', 'fa-comments', 'violet'],
  ['U03', 'Keselamatan', 'KESELAMATAN', 'fa-shield-halved', 'amber'], ['U04', 'MOIES', 'MOIES', 'fa-star', 'sky'],
  ['U05', 'Kebersihan & Keceriaan', 'KEBERSIHAN', 'fa-broom', 'emerald'], ['U06', 'Kesihatan & Pengurusan Wabak', 'KESIHATAN', 'fa-kit-medical', 'teal'],
  ['U07', 'PPDe', 'PPDE', 'fa-mosquito', 'lime'], ['U08', 'PPDa', 'PPDA', 'fa-ban', 'slate']];
const STATUS_OPR_TERBUKA = ['Draf', 'Dipulangkan'];

function unitLalai_() {
  return UNIT_LALAI.map((x, i) => ({ ID_Unit: x[0], Nama: x[1], Singkatan: x[2], Ikon: x[3], Warna: x[4], ID_Ketua: '', Susunan: i + 1, Status: 'Aktif', Carta_JSON: '', Skop_JSON: '', Timestamp: new Date() }));
}
function unitSemua_() {
  return (adaSheet_(SHEET_UNIT) ? read_(SHEET_UNIT).rows.slice() : unitLalai_()).sort((a, b) => Number(a.Susunan) - Number(b.Susunan));
}
function pastiUnit_() { if (!adaSheet_(SHEET_UNIT)) { pastiSheet_(SHEET_UNIT); insert_(SHEET_UNIT, unitLalai_()); } }
function cariUnit_(id) {
  const r = unitSemua_().find(x => String(x.ID_Unit) === String(id || ''));
  if (!r) throw err_('Unit tidak dijumpai.');
  return r;
}
function oprSemua_() { return adaSheet_(SHEET_OPR) ? read_(SHEET_OPR).rows : []; }
function namaGuruMap_() { return new Map(read_('tbl_guru').rows.map(g => [String(g.ID_Guru), String(g.Nama_Penuh)])); }
function aksesUnit_(u, r) {
  const urus = boleh_(u, 'urus_unit_hem'), ketua = !!r.ID_Ketua && String(r.ID_Ketua) === u.id;
  const c = parseJSON_(r.Carta_JSON, null);
  const ahli = ketua || (!!c && Array.isArray(c.nodes) && c.nodes.some(n => (n.ahli || []).some(a => a.idGuru === u.id)));
  return { urus: urus, ketua: ketua, ahli: ahli, sunting: urus || ketua, opr: urus || ketua || ahli, sah: urus || ketua };
}
function ringkasUnit_(r, nama) {
  const c = parseJSON_(r.Carta_JSON, null), ahli = new Set();
  if (c && Array.isArray(c.nodes)) c.nodes.forEach(n => (n.ahli || []).forEach(a => ahli.add(a.idGuru || a.nama)));
  return { id: String(r.ID_Unit), nama: String(r.Nama), singkatan: String(r.Singkatan || ''), ikon: String(r.Ikon || 'fa-star'), warna: String(r.Warna || 'sky'),
    idKetua: String(r.ID_Ketua || ''), namaKetua: nama.get(String(r.ID_Ketua || '')) || '', status: String(r.Status || 'Aktif'), bilAhli: ahli.size };
}

function apiSenaraiUnitHem_(p, u) {
  const urus = boleh_(u, 'urus_unit_hem'), tahun = today_().slice(0, 4), nama = namaGuruMap_(), bil = {}, tunggu = {};
  oprSemua_().forEach(o => {
    const id = String(o.ID_Unit);
    if (String(o.Status) === 'Disahkan' && fmtTarikh_(o.Tarikh).slice(0, 4) === tahun) bil[id] = (bil[id] || 0) + 1;
    if (String(o.Status) === 'Dihantar') tunggu[id] = (tunggu[id] || 0) + 1;
  });
  return { urus: urus, tahun: tahun, unit: unitSemua_().filter(r => urus || String(r.Status) !== 'Arkib').map(r =>
    Object.assign(ringkasUnit_(r, nama), { bilOpr: bil[String(r.ID_Unit)] || 0, menunggu: aksesUnit_(u, r).sah ? tunggu[String(r.ID_Unit)] || 0 : 0 })) };
}

function apiGetUnitHem_(p, u) {
  const r = cariUnit_(p.id), a = aksesUnit_(u, r);
  if (String(r.Status) === 'Arkib' && !a.urus) throw err_('Unit ini telah diarkibkan.');
  return Object.assign(ringkasUnit_(r, namaGuruMap_()), { carta: parseJSON_(r.Carta_JSON, null), skop: parseJSON_(r.Skop_JSON, null) || { umum: '', jawatan: {} }, akses: a });
}

function apiSimpanUnitHem_(p, u) {
  wajib_(u, 'urus_unit_hem');
  pastiUnit_();
  const semua = unitSemua_(), nama = String(p.nama || '').trim().replace(/^unit\s+/i, '').slice(0, 60);
  if (!nama) throw err_('Isi nama unit.');
  if (semua.some(x => String(x.ID_Unit) !== String(p.id || '') && String(x.Nama).toLowerCase() === nama.toLowerCase())) throw err_('Nama unit sudah wujud.');
  const ketua = String(p.idKetua || '');
  if (ketua && !read_('tbl_guru').rows.some(g => String(g.ID_Guru) === ketua && String(g.Status_Aktif) === 'Aktif')) throw err_('Ketua unit tidak dijumpai.');
  const data = { Nama: nama, Singkatan: String(p.singkatan || '').trim().toUpperCase().slice(0, 12), Ikon: IKON_UNIT.indexOf(p.ikon) >= 0 ? p.ikon : 'fa-star',
    Warna: WARNA_UNIT.indexOf(p.warna) >= 0 ? p.warna : 'sky', ID_Ketua: ketua, Status: p.status === 'Arkib' ? 'Arkib' : 'Aktif', Timestamp: new Date() };
  let r;
  if (p.id) { r = cariUnit_(p.id); updateRow_(SHEET_UNIT, Object.assign(r, data)); }
  else {
    r = Object.assign({ ID_Unit: genId_('U'), Susunan: semua.reduce((m, x) => Math.max(m, Number(x.Susunan) || 0), 0) + 1, Carta_JSON: '', Skop_JSON: '' }, data);
    insert_(SHEET_UNIT, [r]);
  }
  audit_(u.id, 'UNIT_HEM', r.ID_Unit + ' ' + nama);
  return ringkasUnit_(r, namaGuruMap_());
}

// ids = susunan baharu semua unit
function apiSusunUnitHem_(p, u) {
  wajib_(u, 'urus_unit_hem');
  pastiUnit_();
  const ids = (Array.isArray(p.ids) ? p.ids : []).map(String), rows = unitSemua_();
  rows.forEach(r => { const i = ids.indexOf(String(r.ID_Unit)); r.Susunan = i >= 0 ? i + 1 : ids.length + Number(r.Susunan); });
  updateRows_(SHEET_UNIT, rows);
  return { ok: true };
}

function apiSimpanCartaUnit_(p, u) {
  pastiUnit_();
  const r = cariUnit_(p.id);
  if (!aksesUnit_(u, r).sunting) throw err_('Hanya ketua unit atau PK HEM boleh menyunting carta unit ini.', 'FORBIDDEN');
  const out = bersihCarta_(p.carta, 'Carta Organisasi Unit ' + r.Nama);
  r.Carta_JSON = jsonSel_(out, 'Carta organisasi unit'); r.Timestamp = new Date();
  updateRow_(SHEET_UNIT, r);
  audit_(u.id, 'CARTA_UNIT', r.ID_Unit);
  return { carta: out };
}

// skop = { umum: teks, jawatan: { idKotakCarta: [tugas, ...] } }
function apiSimpanSkopUnit_(p, u) {
  pastiUnit_();
  const r = cariUnit_(p.id);
  if (!aksesUnit_(u, r).sunting) throw err_('Hanya ketua unit atau PK HEM boleh menyunting skop tugas unit ini.', 'FORBIDDEN');
  const s = p.skop || {}, jawatan = {};
  Object.keys(s.jawatan || {}).slice(0, 100).forEach(k => {
    if (!/^[\w-]{1,20}$/.test(k)) return;
    const t = (Array.isArray(s.jawatan[k]) ? s.jawatan[k] : []).map(x => String(x || '').trim().slice(0, 300)).filter(Boolean).slice(0, 40);
    if (t.length) jawatan[k] = t;
  });
  const out = { umum: String(s.umum || '').trim().slice(0, 3000), jawatan: jawatan };
  r.Skop_JSON = jsonSel_(out, 'Skop tugas'); r.Timestamp = new Date();
  updateRow_(SHEET_UNIT, r);
  audit_(u.id, 'SKOP_UNIT', r.ID_Unit);
  return { skop: out };
}

function bersihOpr_(d) {
  d = d || {};
  const t = (k, n) => String(d[k] == null ? '' : d[k]).trim().slice(0, n);
  return { masaMula: t('masaMula', 20), masaTamat: t('masaTamat', 20), tempat: t('tempat', 150), sasaran: t('sasaran', 200), peserta: t('peserta', 150),
    anjuran: t('anjuran', 150), objektif: t('objektif', 1000), ringkasan: t('ringkasan', 2000), kekuatan: t('kekuatan', 1000), kelemahan: t('kelemahan', 1000), cadangan: t('cadangan', 1000) };
}
function oprKeluar_(o, u, a, nama) {
  const st = String(o.Status), terbuka = STATUS_OPR_TERBUKA.indexOf(st) >= 0;
  return { id: String(o.ID_OPR), idUnit: String(o.ID_Unit), tarikh: fmtTarikh_(o.Tarikh), program: String(o.Program), data: bersihOpr_(parseJSON_(o.Data_JSON, {})),
    foto: parseJSON_(o.Foto_JSON, []), status: st, idPenyedia: String(o.ID_Penyedia), namaPenyedia: nama.get(String(o.ID_Penyedia)) || '',
    namaPengesah: nama.get(String(o.ID_Pengesah || '')) || '', ulasan: String(o.Ulasan || ''), tarikhHantar: masaTeks_(o.Tarikh_Hantar), tarikhSah: masaTeks_(o.Tarikh_Sah),
    bolehUbah: terbuka && (String(o.ID_Penyedia) === u.id || a.sah), bolehSah: a.sah && st === 'Dihantar', bolehPadam: st === 'Draf' && (String(o.ID_Penyedia) === u.id || a.sah) };
}

function apiSenaraiOpr_(p, u) {
  const r = cariUnit_(p.idUnit), a = aksesUnit_(u, r), nama = namaGuruMap_();
  const tahun = /^\d{4}$/.test(String(p.tahun || '')) ? String(p.tahun) : today_().slice(0, 4);
  const laporan = oprSemua_().filter(o => String(o.ID_Unit) === String(r.ID_Unit) && fmtTarikh_(o.Tarikh).slice(0, 4) === tahun)
    // Guru di luar unit hanya melihat laporan yang telah disahkan; draf hanya dilihat penyedia, ketua & PK HEM
    .filter(o => String(o.Status) === 'Disahkan' || (a.opr && (String(o.Status) !== 'Draf' || a.sah || String(o.ID_Penyedia) === u.id)))
    .sort((x, y) => fmtTarikh_(y.Tarikh).localeCompare(fmtTarikh_(x.Tarikh))).map(o => oprKeluar_(o, u, a, nama));
  return { tahun: tahun, laporan: laporan, akses: a };
}

function apiSimpanOpr_(p, u) {
  pastiUnit_();
  const r = cariUnit_(p.idUnit), a = aksesUnit_(u, r);
  if (!a.opr) throw err_('Hanya ahli unit, ketua unit atau PK HEM boleh mengisi laporan aktiviti unit ini.', 'FORBIDDEN');
  const tarikh = semakTarikh_(p.tarikh), program = String(p.program || '').trim().slice(0, 150);
  if (!program) throw err_('Isi nama program / aktiviti.');
  const data = bersihOpr_(p.data), hantar = !!p.hantar;
  const foto = (Array.isArray(p.foto) ? p.foto : []).slice(0, 4).map(idFail_).filter(Boolean);
  if (hantar && (!data.objektif || !data.ringkasan)) throw err_('Isi objektif dan ringkasan aktiviti sebelum menghantar.');
  pastiSheet_(SHEET_OPR);
  const now = new Date();
  let o = p.id ? oprSemua_().find(x => String(x.ID_OPR) === String(p.id)) : null;
  if (p.id && (!o || String(o.ID_Unit) !== String(r.ID_Unit))) throw err_('Laporan tidak dijumpai.');
  const d = { Tarikh: tarikh, Program: program, Data_JSON: jsonSel_(data, 'Laporan aktiviti'), Foto_JSON: JSON.stringify(foto), Timestamp: now };
  if (hantar) Object.assign(d, { Status: 'Dihantar', Tarikh_Hantar: now });
  if (o) {
    if (STATUS_OPR_TERBUKA.indexOf(String(o.Status)) < 0) throw err_('Laporan ini sudah dihantar dan tidak boleh diubah.');
    if (String(o.ID_Penyedia) !== u.id && !a.sah) throw err_('Hanya penyedia laporan atau ketua unit boleh mengubahnya.', 'FORBIDDEN');
    updateRow_(SHEET_OPR, Object.assign(o, d));
  } else {
    o = Object.assign({ ID_OPR: genId_('OPR'), ID_Unit: String(r.ID_Unit), Status: 'Draf', ID_Penyedia: u.id, ID_Pengesah: '', Ulasan: '', Tarikh_Hantar: '', Tarikh_Sah: '' }, d);
    insert_(SHEET_OPR, [o]);
  }
  if (hantar) audit_(u.id, 'HANTAR_OPR', o.ID_OPR);
  return { id: String(o.ID_OPR), status: String(o.Status) };
}

function apiSahOpr_(p, u) {
  const o = oprSemua_().find(x => String(x.ID_OPR) === String(p.id || ''));
  if (!o) throw err_('Laporan tidak dijumpai.');
  if (!aksesUnit_(u, cariUnit_(o.ID_Unit)).sah) throw err_('Hanya ketua unit atau PK HEM boleh mengesahkan laporan ini.', 'FORBIDDEN');
  if (String(o.Status) !== 'Dihantar') throw err_('Hanya laporan yang sudah dihantar boleh disahkan atau dipulangkan.');
  const pulang = p.tindakan === 'pulang', ulasan = String(p.ulasan || '').trim().slice(0, 1000);
  if (pulang && !ulasan) throw err_('Nyatakan perkara yang perlu dibetulkan dalam ulasan.');
  Object.assign(o, { Status: pulang ? 'Dipulangkan' : 'Disahkan', Ulasan: ulasan, ID_Pengesah: u.id, Tarikh_Sah: pulang ? '' : new Date(), Timestamp: new Date() });
  updateRow_(SHEET_OPR, o);
  audit_(u.id, pulang ? 'PULANG_OPR' : 'SAH_OPR', o.ID_OPR);
  return { status: o.Status };
}

function apiPadamOpr_(p, u) {
  const o = oprSemua_().find(x => String(x.ID_OPR) === String(p.id || ''));
  if (!o) throw err_('Laporan tidak dijumpai.');
  const a = aksesUnit_(u, cariUnit_(o.ID_Unit));
  if (String(o.Status) !== 'Draf') throw err_('Hanya draf boleh dipadam.');
  if (String(o.ID_Penyedia) !== u.id && !a.sah) throw err_('Hanya penyedia laporan atau ketua unit boleh memadamnya.', 'FORBIDDEN');
  sheet_(SHEET_OPR).deleteRow(o._row);
  delete _MEMO['r:' + SHEET_OPR];
  audit_(u.id, 'PADAM_OPR', o.ID_OPR);
  return { ok: true };
}

// Ringkasan untuk bootstrap (peringatan Dashboard): OPR menunggu pengesahan & OPR dipulangkan kepada penyedia
function unitInfo_(u) {
  const out = { menunggu: 0, dipulangkan: 0 };
  if (!adaSheet_(SHEET_OPR)) return out;
  const unit = new Map(unitSemua_().map(r => [String(r.ID_Unit), r])), sah = {};
  oprSemua_().forEach(o => {
    const id = String(o.ID_Unit), r = unit.get(id); if (!r) return;
    if (sah[id] === undefined) sah[id] = aksesUnit_(u, r).sah;
    if (String(o.Status) === 'Dihantar' && sah[id]) out.menunggu++;
    if (String(o.Status) === 'Dipulangkan' && String(o.ID_Penyedia) === u.id) out.dipulangkan++;
  });
  return out;
}

/* ---------- Rekod kawalan kelas (check in guru masuk kelas) ----------
 * Waktu & subjek diambil daripada jadual waktu aSc kelas itu; kehadiran murid daripada rekod kehadiran harian.
 * Semua guru boleh check in mana-mana kelas pada hari ini (guru ganti dikesan automatik jika nama tidak sepadan jadual).
 * Masa masuk ditentukan oleh pelayan; lewat jika lebih 10 minit selepas waktu bermula. */
const SHEET_KK = 'tbl_kawalan_kelas', LEWAT_MINIT = 10;
const GELARAN_GURU = ['PN', 'PUAN', 'EN', 'ENCIK', 'CIK', 'CIKGU', 'CG', 'USTAZ', 'USTAZAH', 'UST', 'DR', 'TN', 'HJ', 'HAJI', 'HAJAH', 'HJH', 'MR', 'MRS', 'MS', 'MISS', 'SIR', 'BIN', 'BINTI', 'BT', 'AL', 'AP'];
function tokenNama_(s) { return String(s || '').toUpperCase().replace(/[^A-Z0-9\s]/g, ' ').split(/\s+/).filter(t => t.length >= 2 && GELARAN_GURU.indexOf(t) < 0); }
// Nama guru dalam aSc (cth "Pn. Nor Aini") sepadan jika semua perkataannya ada dalam nama penuh guru
function padanGuru_(namaAsc, namaPenuh) {
  const n = tokenNama_(namaPenuh);
  return String(namaAsc || '').split('/').some(x => { const a = tokenNama_(x); return a.length > 0 && a.every(t => n.indexOf(t) >= 0); });
}
function minit_(hhmm) { const m = /^(\d{1,2}):(\d{2})$/.exec(String(hhmm || '')); return m ? Number(m[1]) * 60 + Number(m[2]) : null; }
function masaKini_() { return Utilities.formatDate(new Date(), APP.TZ, 'HH:mm'); }
function jadualKelasMap_() {
  const out = new Map();
  if (adaSheetJW_()) read_(SHEET_JW).rows.forEach(r => { const j = parseJSON_(r.Jadual_JSON, null); if (j) out.set(String(r.ID_Kelas), j); });
  return out;
}
// Indeks hari dalam jadual bagi tarikh itu (ikut nama hari BM/BI, atau susunan Isnin = 0)
function hariJadual_(j, tarikh) {
  const [y, m, d] = String(tarikh).split('-').map(Number), dow = new Date(Date.UTC(y, m - 1, d)).getUTCDay();
  const nama = [['AHAD', 'SUN', 'MIN'], ['ISN', 'MON'], ['SEL', 'TUE'], ['RAB', 'WED'], ['KHA', 'THU'], ['JUM', 'FRI'], ['SAB', 'SAT']][dow];
  const i = j.hari.findIndex(h => nama.some(n => String(h).toUpperCase().indexOf(n) === 0));
  if (i >= 0) return i;
  const k = (dow + 6) % 7;
  return k < j.hari.length && dow !== 0 && dow !== 6 ? k : -1;
}
// Senarai waktu bagi satu hari: sel bersebelahan dengan subjek & guru yang sama digabungkan; waktu rehat disertakan
function slotHari_(j, tarikh) {
  if (!j || !Array.isArray(j.hari) || !Array.isArray(j.waktu)) return [];
  const h = hariJadual_(j, tarikh); if (h < 0) return [];
  const kunci = c => c.s.map(x => x.k + '|' + x.g).join(';');
  // Waktu rehat: lajur rehat jadual, waktu berlabel "Rehat", atau sel yang subjeknya Rehat / Waktu makan (tiada check in)
  const REHAT = /^\s*(rehat|recess|break|waktu makan|makan)\b/i;
  const rehatW = i => !!(j.waktu[i] && (j.waktu[i].rehat || REHAT.test(String(j.waktu[i].label || ''))));
  const rehatSel = c => c.s.length > 0 && c.s.every(x => REHAT.test(String(x.nm || x.k || '')));
  // Pecahkan setiap sel kepada satu unit bagi setiap waktu (unit pada waktu rehat dibuang), kemudian gabung unit bersebelahan yang sama
  const unit = new Map(), rehat = new Set();
  j.sel.filter(c => c.h === h).forEach(c => {
    for (let i = c.w; i < c.w + c.n; i++) {
      if (rehatW(i) || rehatSel(c)) { rehat.add(i); continue; }
      if (!unit.has(i)) unit.set(i, c.s);
    }
  });
  j.waktu.forEach((x, i) => { if (rehatW(i)) rehat.add(i); });
  const gab = [];
  Array.from(unit.keys()).sort((a, b) => a - b).forEach(i => {
    const s = unit.get(i), akhir = gab[gab.length - 1];
    if (akhir && akhir.w + akhir.n === i && kunci(akhir) === kunci({ s: s })) akhir.n++;
    else gab.push({ w: i, n: 1, s: s });
  });
  const out = gab.map(c => {
    const a = j.waktu[c.w] || {}, b = j.waktu[c.w + c.n - 1] || {};
    return { w: c.w, label: c.n > 1 ? `${a.label}–${b.label}` : String(a.label || ''), mula: a.mula || '', tamat: b.tamat || '',
      subjek: c.s.map(x => x.nm || x.k).filter(Boolean).join(' / '), guru: c.s.map(x => x.g).filter(Boolean).join(' / ') };
  });
  // Waktu rehat bersebelahan dipaparkan sebagai satu baris
  Array.from(rehat).filter(i => !unit.has(i) && j.waktu[i]).sort((a, b) => a - b).forEach(i => {
    const akhir = out.find(x => x.rehat && x.hujung === i - 1);
    if (akhir) { akhir.tamat = j.waktu[i].tamat; akhir.hujung = i; return; }
    out.push({ w: i, hujung: i, label: 'REHAT', mula: j.waktu[i].mula, tamat: j.waktu[i].tamat, rehat: true });
  });
  return out.sort((a, b) => a.w - b.w);
}
function kehadiranKelas_(idKelas, tarikh) {
  const murid = new Map(read_('tbl_murid').rows.filter(m => String(m.ID_Kelas) === idKelas && String(m.Status) === 'Aktif').map(m => [String(m.ID_Murid), m]));
  let direkod = 0, hadir = 0; const tidakHadir = [];
  kehadiranTahun_(tarikh.slice(0, 4)).forEach(r => {
    if (fmtTarikh_(r.Tarikh) !== tarikh || String(r.ID_Kelas) !== idKelas || !murid.has(String(r.ID_Murid))) return;
    direkod++;
    if (hadirH_(r.Status)) hadir++; else tidakHadir.push(String(murid.get(String(r.ID_Murid)).Nama));
  });
  return { jumlah: murid.size, direkod: direkod, hadir: hadir, tidakHadir: tidakHadir.sort() };
}
function rekodKK_() { return adaSheet_(SHEET_KK) ? read_(SHEET_KK).rows : []; }
function rekodKeluar_(r, u, nama) {
  const t = fmtTarikh_(r.Tarikh), milik = String(r.ID_Guru) === u.id;
  return { id: String(r.ID_Rekod), tarikh: t, idKelas: String(r.ID_Kelas), waktu: r.Waktu === '' ? -1 : Number(r.Waktu), label: String(r.Label_Waktu || ''),
    mula: String(r.Masa_Mula || ''), tamat: String(r.Masa_Tamat || ''), subjek: String(r.Subjek || ''), guruJadual: String(r.Guru_Jadual || ''),
    idGuru: String(r.ID_Guru), namaGuru: nama.get(String(r.ID_Guru)) || '', ganti: isTrue_(r.Ganti), masuk: String(r.Masa_Masuk || ''),
    status: String(r.Status_Masa || ''), tajuk: String(r.Tajuk || ''), catatan: String(r.Catatan || ''),
    bolehUbah: boleh_(u, 'pantau_kawalan') || (milik && t === today_()) };
}
function hariKawalan_(idKelas, tarikh, u, jm, nama) {
  const rekod = rekodKKTahun_(tarikh.slice(0, 4)).filter(r => String(r.ID_Kelas) === idKelas && fmtTarikh_(r.Tarikh) === tarikh).map(r => rekodKeluar_(r, u, nama));
  return { tarikh: tarikh, hari: hariTarikh_(tarikh), slot: slotHari_(jm.get(idKelas), tarikh), rekod: rekod, kehadiran: kehadiranKelas_(idKelas, tarikh) };
}

function apiGetKawalan_(p, u) {
  const idKelas = String(p.idKelas || ''), tarikh = p.tarikh ? semakTarikh_(p.tarikh) : today_();
  if (!read_('tbl_kelas').rows.some(k => String(k.ID_Kelas) === idKelas)) throw err_('Kelas tidak dijumpai.');
  const jm = jadualKelasMap_();
  return Object.assign(hariKawalan_(idKelas, tarikh, u, jm, namaGuruMap_()), { idKelas: idKelas, adaJadual: jm.has(idKelas),
    hariIni: today_(), masaKini: masaKini_(), pantau: boleh_(u, 'pantau_kawalan') });
}

// Isnin–Jumaat bagi minggu tarikh itu (untuk cetakan buku rekod)
function apiKawalanMinggu_(p, u) {
  const idKelas = String(p.idKelas || ''), isnin = isninMinggu_(p.tarikh ? semakTarikh_(p.tarikh) : today_());
  const jm = jadualKelasMap_(), nama = namaGuruMap_();
  return { idKelas: idKelas, minggu: isnin, hari: [0, 1, 2, 3, 4].map(i => hariKawalan_(idKelas, tambahHari_(isnin, i), u, jm, nama)) };
}

// Jadual guru ini hari ini (kad "Kelas anda sekarang" di Dashboard)
function apiJadualSaya_(p, u) {
  const t = today_(), jm = jadualKelasMap_(), me = namaGuruMap_().get(u.id) || u.nama || '';
  const kelas = new Map(read_('tbl_kelas').rows.filter(k => String(k.Status) !== 'Tidak Aktif').map(k => [String(k.ID_Kelas), k]));
  const sudah = new Map(rekodKK_().filter(r => fmtTarikh_(r.Tarikh) === t).map(r => [String(r.ID_Kelas) + '|' + r.Waktu, r]));
  const slot = [];
  jm.forEach((j, idKelas) => {
    if (!kelas.has(idKelas)) return;
    slotHari_(j, t).forEach(s => {
      if (s.rehat || !padanGuru_(s.guru, me)) return;
      const r = sudah.get(idKelas + '|' + s.w);
      slot.push(Object.assign(s, { idKelas: idKelas, namaKelas: String(kelas.get(idKelas).Nama_Kelas), masuk: r ? String(r.Masa_Masuk) : '', olehSaya: !!r && String(r.ID_Guru) === u.id }));
    });
  });
  return { tarikh: t, masaKini: masaKini_(), slot: slot.sort((a, b) => (minit_(a.mula) === null ? 1e4 : minit_(a.mula)) - (minit_(b.mula) === null ? 1e4 : minit_(b.mula)) || a.w - b.w) };
}

// waktu = indeks waktu jadual (>= 0), atau -1 untuk kelas tanpa jadual (guru isi masa & subjek)
function apiCheckIn_(p, u) {
  const idKelas = String(p.idKelas || ''), t = today_();
  if (!read_('tbl_kelas').rows.some(k => String(k.ID_Kelas) === idKelas)) throw err_('Kelas tidak dijumpai.');
  const waktu = Number(p.waktu), nama = namaGuruMap_(), me = nama.get(u.id) || '';
  let slot;
  if (waktu >= 0) {
    slot = slotHari_(jadualKelasMap_().get(idKelas), t).find(s => s.w === waktu && !s.rehat);
    if (!slot) throw err_('Waktu tidak dijumpai dalam jadual hari ini.');
  } else {
    const mula = String(p.mula || '');
    if (minit_(mula) === null) throw err_('Isi masa mula.');
    slot = { w: '', label: '', mula: mula, tamat: minit_(p.tamat) === null ? '' : String(p.tamat), subjek: '', guru: '' };
  }
  const subjek = String(p.subjek || '').trim().slice(0, 80) || slot.subjek;
  if (!subjek) throw err_('Isi subjek.');
  const kunci = r => String(r.ID_Kelas) === idKelas && fmtTarikh_(r.Tarikh) === t && (waktu >= 0 ? String(r.Waktu) === String(waktu) : String(r.Waktu) === '' && String(r.Masa_Mula) === slot.mula);
  const ada = rekodKK_().find(kunci);
  if (ada) throw err_(`Waktu ini sudah di-check in oleh ${nama.get(String(ada.ID_Guru)) || 'guru lain'} pada ${ada.Masa_Masuk}.`);
  const masuk = masaKini_(), mm = minit_(slot.mula);
  pastiSheet_(SHEET_KK);
  const r = { ID_Rekod: genId_('KK'), Tarikh: t, ID_Kelas: idKelas, Waktu: slot.w, Label_Waktu: slot.label, Masa_Mula: slot.mula, Masa_Tamat: slot.tamat,
    Subjek: subjek, Guru_Jadual: slot.guru, ID_Guru: u.id, Ganti: slot.guru && !padanGuru_(slot.guru, me) ? 'TRUE' : (p.ganti ? 'TRUE' : 'FALSE'),
    Masa_Masuk: masuk, Status_Masa: mm !== null && minit_(masuk) > mm + LEWAT_MINIT ? 'Lewat' : 'Tepat',
    Tajuk: String(p.tajuk || '').trim().slice(0, 200), Catatan: String(p.catatan || '').trim().slice(0, 500), Timestamp: new Date() };
  insert_(SHEET_KK, [r]);
  audit_(u.id, 'CHECK_IN', idKelas + ' ' + t + ' ' + (slot.label || slot.mula));
  return rekodKeluar_(r, u, nama);
}

function rekodMilik_(p, u) {
  const r = rekodKK_().find(x => String(x.ID_Rekod) === String(p.id || ''));
  if (!r) throw err_('Rekod tidak dijumpai.');
  if (!boleh_(u, 'pantau_kawalan') && !(String(r.ID_Guru) === u.id && fmtTarikh_(r.Tarikh) === today_())) throw err_('Rekod hanya boleh diubah oleh guru yang check in pada hari yang sama, atau PK HEM.', 'FORBIDDEN');
  return r;
}
function apiKemaskiniKawalan_(p, u) {
  const r = rekodMilik_(p, u);
  if (p.tajuk !== undefined) r.Tajuk = String(p.tajuk || '').trim().slice(0, 200);
  if (p.catatan !== undefined) r.Catatan = String(p.catatan || '').trim().slice(0, 500);
  r.Timestamp = new Date();
  updateRow_(SHEET_KK, r);
  return rekodKeluar_(r, u, namaGuruMap_());
}
function apiPadamKawalan_(p, u) {
  const r = rekodMilik_(p, u);
  sheet_(SHEET_KK).deleteRow(r._row);
  delete _MEMO['r:' + SHEET_KK];
  audit_(u.id, 'PADAM_CHECK_IN', String(r.ID_Rekod));
  return { ok: true };
}

// Pemantauan PK HEM: kelas tanpa guru sekarang, ringkasan setiap kelas hari ini & peratus waktu diisi minggu ini
function apiPantauKawalan_(p, u) {
  wajib_(u, 'pantau_kawalan');
  const t = today_(), kini = minit_(masaKini_()), jm = jadualKelasMap_(), isnin = isninMinggu_(t);
  const rekod = rekodKK_().filter(r => { const d = fmtTarikh_(r.Tarikh); return d >= isnin && d <= t; });
  const ada = new Set(rekod.map(r => String(r.ID_Kelas) + '|' + fmtTarikh_(r.Tarikh) + '|' + r.Waktu));
  const kelas = read_('tbl_kelas').rows.filter(k => String(k.Status) !== 'Tidak Aktif')
    .sort((a, b) => tahunSort_(a.Tahun) - tahunSort_(b.Tahun) || String(a.Nama_Kelas).localeCompare(String(b.Nama_Kelas)));
  const tanpaGuru = [], ringkas = [];
  let berjalan = 0, diisiKini = 0, mingguSlot = 0, mingguIsi = 0;
  kelas.forEach(k => {
    const id = String(k.ID_Kelas), j = jm.get(id); if (!j) return;
    for (let d = isnin; d <= t; d = tambahHari_(d, 1)) {
      slotHari_(j, d).forEach(s => {
        if (s.rehat || (d === t && minit_(s.mula) > kini)) return;
        mingguSlot++; if (ada.has(id + '|' + d + '|' + s.w)) mingguIsi++;
      });
    }
    const slot = slotHari_(j, t).filter(s => !s.rehat), bermula = slot.filter(s => minit_(s.mula) !== null && minit_(s.mula) <= kini);
    const isi = bermula.filter(s => ada.has(id + '|' + t + '|' + s.w));
    const ganti = rekod.filter(r => String(r.ID_Kelas) === id && fmtTarikh_(r.Tarikh) === t && isTrue_(r.Ganti)).length;
    slot.forEach(s => {
      const m = minit_(s.mula), e = minit_(s.tamat);
      if (m === null || e === null || kini < m || kini >= e) return;
      berjalan++;
      if (ada.has(id + '|' + t + '|' + s.w)) diisiKini++;
      else if (kini >= m + LEWAT_MINIT) tanpaGuru.push({ idKelas: id, namaKelas: String(k.Nama_Kelas), label: s.label, mula: s.mula, tamat: s.tamat, subjek: s.subjek, guru: s.guru });
    });
    const h = kehadiranKelas_(id, t);
    ringkas.push({ idKelas: id, namaKelas: String(k.Nama_Kelas), slot: slot.length, bermula: bermula.length, diisi: isi.length, ganti: ganti, hadir: h.hadir, direkod: h.direkod, jumlah: h.jumlah });
  });
  return { tarikh: t, masaKini: masaKini_(), berjalan: berjalan, diisiKini: diisiKini, tanpaGuru: tanpaGuru, ringkas: ringkas,
    peratusMinggu: mingguSlot ? Math.round(mingguIsi / mingguSlot * 100) : null, adaJadual: jm.size > 0 };
}

/* ---------- Imbas QR murid (menu pantas) ----------
 * Status kehadiran hari ini & bilangan lewat bulan ini dipaparkan pada menu selepas imbasan. */
function muridAktif_(id) {
  const m = read_('tbl_murid').rows.find(r => String(r.ID_Murid) === String(id || '') && String(r.Status) === 'Aktif');
  if (!m) throw err_('Murid tidak dijumpai.');
  return m;
}
function apiInfoImbas_(p, u) {
  const m = muridAktif_(p.idMurid), t = today_(), bulan = t.slice(0, 7);
  let hariIni = '', lewat = 0;
  kehadiranTahun_(t.slice(0, 4)).forEach(r => {
    if (String(r.ID_Murid) !== String(m.ID_Murid)) return;
    const d = fmtTarikh_(r.Tarikh);
    if (d === t) hariIni = statusH_(r.Status);
    if (d.slice(0, 7) === bulan && statusH_(r.Status) === 'Lewat') lewat++;
  });
  return { idMurid: String(m.ID_Murid), hariIni: hariIni, lewatBulan: lewat, masaKini: Utilities.formatDate(new Date(), APP.TZ, 'HH:mm'), rmt: rmtImbasInfo_(m, u) };
}

/* ---- RMT melalui imbasan QR: tanda murid makan / tidak makan hari ini ----
 * Hanya untuk PIC RMT / PK HEM / guru bertugas RMT hari ini, dan murid penerima RMT. */
function rmtImbasInfo_(m, u) {
  if (!skimRmt_()) return null;
  const a = rmtAkses_(u), t = today_();
  if (!a.urus && guruBertugasRmt_(t) !== u.id) return null;
  const penerima = penerimaRmt_(t.slice(0, 4)).some(x => String(x.ID_Murid) === String(m.ID_Murid));
  const r = barisRmt_(t), rekod = r ? parseJSON_(r.Rekod_JSON, {}) || {} : {};
  return { penerima: penerima, tiadaRmt: r ? isTrue_(r.Tiada_RMT) : false, makan: rekod[String(m.ID_Murid)] === 1, bilHadir: r ? Number(r.Bil_Hadir) || 0 : 0, bilPenerima: r ? Number(r.Bil_Penerima) || 0 : 0 };
}
function apiRmtImbas_(p, u) {
  if (!skimRmt_()) throw err_('Skim bantuan berkod RMT tidak dijumpai.');
  const t = today_(), a = rmtAkses_(u);
  if (!a.urus && guruBertugasRmt_(t) !== u.id) throw err_('Hanya guru bertugas RMT hari ini, PIC RMT atau PK HEM boleh merekod.', 'FORBIDDEN');
  const id = String(p.idMurid || ''), penerima = penerimaRmt_(t.slice(0, 4)).map(x => String(x.ID_Murid));
  if (penerima.indexOf(id) < 0) throw err_('Murid ini bukan penerima RMT.');
  const r = barisRmt_(t);
  if (r && isTrue_(r.Tiada_RMT)) throw err_('Hari ini ditanda tiada RMT.');
  // Murid penerima yang belum diimbas dikira tidak makan; rekod sedia ada dikekalkan
  const rekod = Object.assign({}, r ? parseJSON_(r.Rekod_JSON, {}) || {} : {});
  penerima.forEach(x => { if (rekod[x] === undefined) rekod[x] = 0; });
  const sudah = rekod[id] === 1;
  rekod[id] = p.makan === false ? 0 : 1;
  const data = { Rekod_JSON: jsonSel_(rekod, 'Rekod kehadiran RMT'), Bil_Penerima: Object.keys(rekod).length, Bil_Hadir: Object.keys(rekod).filter(x => rekod[x]).length,
    Tiada_RMT: 'FALSE', ID_Guru_Rekod: u.id, Timestamp: new Date() };
  if (!adaSheetRmt_()) ensureSheet_(SpreadsheetApp.getActiveSpreadsheet(), SHEET_RMT, SCHEMA[SHEET_RMT]);
  if (r) { Object.assign(r, data); updateRow_(SHEET_RMT, r); }
  else insert_(SHEET_RMT, [Object.assign({ Tarikh: t, Menu: String(tetapanRmt_().menu[hariTarikh_(t)] || ''), ID_Guru_Bertugas: guruBertugasRmt_(t) }, data)]);
  return { makan: rekod[id] === 1, sudah: sudah, bilHadir: data.Bil_Hadir, bilPenerima: data.Bil_Penerima };
}

// No. telefon penjaga untuk semua guru (kecemasan); setiap paparan direkod dalam log audit
function apiKenalanMurid_(p, u) {
  const m = muridAktif_(p.idMurid), penjaga = [];
  [1, 2].forEach(i => {
    const nama = String(m['Nama_Penjaga_' + i] || '').trim(), tel = String(m['Tel_Penjaga_' + i] || '').trim();
    if (nama || tel) penjaga.push({ nama: nama, hubungan: String(m['Hubungan_Penjaga_' + i] || ''), tel: tel });
  });
  if (!penjaga.length && (m.Nama_Penjaga || m.Tel_Penjaga)) penjaga.push({ nama: String(m.Nama_Penjaga || ''), hubungan: String(m.Hubungan_Penjaga || ''), tel: String(m.Tel_Penjaga || '') });
  audit_(u.id, 'LIHAT_KENALAN_MURID', String(m.ID_Murid));
  return { idMurid: String(m.ID_Murid), nama: String(m.Nama), namaKelas: String(m.Nama_Kelas), penjaga: penjaga };
}

/* ---------- Sandaran harian ----------
 * Salinan penuh spreadsheet ke folder Drive "HEM Pintar - Sandaran (SULIT)" setiap malam (pencetus masa, 1–2 pagi).
 * Salinan lebih 30 hari dibuang ke tong sampah Drive. */
const SANDARAN_HARI = 30, AWALAN_SANDARAN = 'Sandaran HEM ';
function simpanTetapanKunci_(kunci, nilai) {
  ensureTetapan_();
  const r = read_('tbl_tetapan').rows.find(x => String(x.Kunci) === kunci);
  if (r) { r.Nilai = String(nilai); updateRow_('tbl_tetapan', r); }
}
function folderSandaran_() {
  const id = tetapan_().FOLDER_SANDARAN_ID;
  if (id) { try { const f = DriveApp.getFolderById(id); if (!f.isTrashed()) return f; } catch (e) { /* cipta semula */ } }
  const f = DriveApp.createFolder('HEM Pintar - Sandaran (SULIT)');
  simpanTetapanKunci_('FOLDER_SANDARAN_ID', f.getId());
  return f;
}
function buatSandaran_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), f = folderSandaran_(), masa = Utilities.formatDate(new Date(), APP.TZ, 'yyyy-MM-dd HH:mm');
  const salinan = DriveApp.getFileById(ss.getId()).makeCopy(AWALAN_SANDARAN + masa, f);
  const had = Date.now() - SANDARAN_HARI * 864e5;
  let dibuang = 0;
  const it = f.getFiles();
  while (it.hasNext()) {
    const x = it.next();
    if (x.getName().indexOf(AWALAN_SANDARAN) === 0 && x.getDateCreated().getTime() < had) { x.setTrashed(true); dibuang++; }
  }
  simpanTetapanKunci_('SANDARAN_TERAKHIR', masa);
  return { masa: masa, nama: salinan.getName(), dibuang: dibuang };
}
// Dipanggil oleh pencetus masa harian (tiada pengguna log masuk)
function sandaranHarian() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(120000)) return;
  try { buatSandaran_(); audit_('SISTEM', 'SANDARAN_HARIAN', ''); } finally { lock.releaseLock(); }
}
function menuSandaran() { denganLock_('Sandaran', () => '✅ Sandaran dibuat: ' + buatSandaran_().nama); }
function pencetusSandaran_() { return ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'sandaranHarian'); }
function bilTahunLepas_(nama) {
  if (!adaSheet_(nama)) return 0;
  const t = today_().slice(0, 4);
  return read_(nama).rows.filter(r => { const th = fmtTarikh_(r.Tarikh).slice(0, 4); return /^\d{4}$/.test(th) && th < t; }).length;
}
function apiStatusSandaran_(p, u) {
  wajib_(u, 'urus_tetapan');
  const t = tetapan_();
  let bil = 0, url = '';
  if (t.FOLDER_SANDARAN_ID) {
    try {
      const f = DriveApp.getFolderById(t.FOLDER_SANDARAN_ID), it = f.getFiles(); url = f.getUrl();
      while (it.hasNext()) if (it.next().getName().indexOf(AWALAN_SANDARAN) === 0) bil++;
    } catch (e) { /* folder dipadam */ }
  }
  return { aktif: pencetusSandaran_().length > 0, terakhir: t.SANDARAN_TERAKHIR || '', bilangan: bil, folderUrl: url, simpanHari: SANDARAN_HARI,
    arkib: { kehadiran: bilTahunLepas_('tbl_kehadiran'), kawalan: bilTahunLepas_(SHEET_KK) } };
}
function apiSandaranSekarang_(p, u) {
  wajib_(u, 'urus_tetapan');
  const r = buatSandaran_();
  audit_(u.id, 'SANDARAN_MANUAL', r.nama);
  return r;
}
function apiPasangSandaran_(p, u) {
  wajib_(u, 'urus_tetapan');
  pencetusSandaran_().forEach(t => ScriptApp.deleteTrigger(t));
  if (p.aktif) ScriptApp.newTrigger('sandaranHarian').timeBased().everyDays(1).atHour(1).inTimezone(APP.TZ).create();
  audit_(u.id, 'SANDARAN_HARIAN', p.aktif ? 'aktif' : 'henti');
  return apiStatusSandaran_(p, u);
}

/* ---------- Arkib tahunan rekod kawalan kelas (seperti arkib kehadiran) ---------- */
function namaArkibKK_(tahun) { return 'arkib_kawalan_' + tahun; }
// Rekod kawalan kelas bagi satu tahun: jadual utama + helaian arkib tahun itu (jika ada)
function rekodKKTahun_(tahun) {
  const utama = rekodKK_();
  return adaSheet_(namaArkibKK_(tahun)) ? utama.concat(read_(namaArkibKK_(tahun)).rows) : utama;
}
function arkibKawalanCore_() {
  if (!adaSheet_(SHEET_KK)) return {};
  const ss = SpreadsheetApp.getActiveSpreadsheet(), tahunIni = today_().slice(0, 4), tb = read_(SHEET_KK);
  const ikutTahun = {}, kekal = [];
  tb.rows.forEach(r => {
    const th = fmtTarikh_(r.Tarikh).slice(0, 4);
    if (/^\d{4}$/.test(th) && th < tahunIni) (ikutTahun[th] = ikutTahun[th] || []).push(r); else kekal.push(r);
  });
  const H = tb.headers, bil = {};
  Object.keys(ikutTahun).forEach(th => {
    let sh = ss.getSheetByName(namaArkibKK_(th));
    if (!sh) { sh = ss.insertSheet(namaArkibKK_(th)); sh.getRange(1, 1, 1, H.length).setValues([H]).setFontWeight('bold'); sh.setFrozenRows(1); }
    const rows = ikutTahun[th], start = sh.getLastRow() + 1;
    ensureRows_(sh, start + rows.length - 1);
    applyTextFormat_(sh, SHEET_KK, H, start, rows.length);
    sh.getRange(start, 1, rows.length, H.length).setValues(rows.map(o => H.map(h => cell_(o[h]))));
    bil[th] = rows.length;
  });
  if (Object.keys(bil).length) writeAll_(SHEET_KK, kekal);
  return bil;
}
function apiArkibTahunLepas_(p, u) {
  wajib_(u, 'urus_tetapan');
  const kehadiran = bilTahunLepas_('tbl_kehadiran') ? arkibKehadiranCore_() : 'Tiada rekod kehadiran tahun lepas.';
  const kawalan = arkibKawalanCore_();
  audit_(u.id, 'ARKIB_TAHUN_LEPAS', JSON.stringify(kawalan));
  return { kehadiran: kehadiran, kawalan: kawalan };
}

/* ---------- Log audit (PK HEM / Super Admin) ---------- */
// Dibaca dari bawah (5000 rekod terkini) supaya pantas walaupun log sangat panjang
function apiSenaraiAudit_(p, u) {
  wajib_(u, 'lihat_audit');
  const sh = sheet_('tbl_log_audit'), last = sh.getLastRow();
  if (last < 2) return { rekod: [], jenis: [], diimbas: 0 };
  const H = headers_(sh), n = Math.min(last - 1, 5000), vals = sh.getRange(last - n + 1, 1, n, H.length).getValues();
  const iT = H.indexOf('Timestamp'), iG = H.indexOf('ID_Guru'), iA = H.indexOf('Tindakan'), iB = H.indexOf('Butiran');
  const nama = namaGuruMap_(), dari = String(p.dari || ''), hingga = String(p.hingga || ''), tindakan = String(p.tindakan || ''), idGuru = String(p.idGuru || ''), q = String(p.q || '').toLowerCase();
  const rekod = [];
  for (let i = vals.length - 1; i >= 0 && rekod.length < 500; i--) {
    const r = vals[i], masa = r[iT] instanceof Date ? Utilities.formatDate(r[iT], APP.TZ, 'yyyy-MM-dd HH:mm') : String(r[iT]);
    if ((dari && masa.slice(0, 10) < dari) || (hingga && masa.slice(0, 10) > hingga) || (tindakan && String(r[iA]) !== tindakan) || (idGuru && String(r[iG]) !== idGuru)) continue;
    const nm = nama.get(String(r[iG])) || String(r[iG]);
    if (q && (String(r[iB]) + ' ' + nm + ' ' + r[iA]).toLowerCase().indexOf(q) < 0) continue;
    rekod.push({ masa: masa, idGuru: String(r[iG]), namaGuru: nm, tindakan: String(r[iA]), butiran: String(r[iB]) });
  }
  return { rekod: rekod, jenis: Array.from(new Set(vals.map(r => String(r[iA])))).filter(Boolean).sort(), diimbas: n };
}

/* ---------- Murid berisiko (amaran awal) ----------
 * Kriteria: lewat ≥ 3 kali sebulan, tidak hadir ≥ 3 hari sebulan, tidak hadir 3 hari rekod berturut-turut (terkini),
 * salah laku ≥ 3 kali sebulan, atau merit pada tahap Amaran / Kritikal. */
const HAD_RISIKO = { lewat: 3, tidakHadir: 3, berturut: 3, salahLaku: 3 };
function senaraiBerisiko_(murid, bulan, t) {
  const ids = new Set(murid.map(m => String(m.ID_Murid))), st = {};
  const s = id => st[id] = st[id] || { lewat: 0, tidakHadir: 0, salahLaku: 0, rekod: [] };
  kehadiranTahun_(bulan.slice(0, 4)).forEach(r => {
    const id = String(r.ID_Murid); if (!ids.has(id)) return;
    const d = fmtTarikh_(r.Tarikh), status = statusH_(r.Status), x = s(id);
    x.rekod.push([d, status]);
    if (d.slice(0, 7) === bulan) { if (status === 'Lewat') x.lewat++; if (status === 'Tidak Hadir') x.tidakHadir++; }
  });
  read_('tbl_disiplin').rows.forEach(k => {
    const id = String(k.ID_Murid);
    if (ids.has(id) && Number(k.Mata_Kesan) < 0 && fmtTarikh_(k.Tarikh_Masa).slice(0, 7) === bulan) s(id).salahLaku++;
  });
  return murid.map(m => {
    const id = String(m.ID_Murid), x = st[id] || { lewat: 0, tidakHadir: 0, salahLaku: 0, rekod: [] }, sebab = [];
    let berturut = 0;
    x.rekod.sort((a, b) => b[0].localeCompare(a[0]));
    for (let i = 0; i < x.rekod.length && x.rekod[i][1] === 'Tidak Hadir'; i++) berturut++;
    const sm = statusMerit_(Number(m.Mata_Merit), t);
    if (berturut >= HAD_RISIKO.berturut) sebab.push({ jenis: 'berturut', teks: `Tidak hadir ${berturut} hari berturut-turut`, berat: 3 });
    if (x.tidakHadir >= HAD_RISIKO.tidakHadir) sebab.push({ jenis: 'tidakHadir', teks: `Tidak hadir ${x.tidakHadir} hari bulan ini`, berat: 2 });
    if (x.lewat >= HAD_RISIKO.lewat) sebab.push({ jenis: 'lewat', teks: `Lewat ${x.lewat} kali bulan ini`, berat: 1 });
    if (x.salahLaku >= HAD_RISIKO.salahLaku) sebab.push({ jenis: 'salahLaku', teks: `${x.salahLaku} salah laku bulan ini`, berat: 2 });
    if (sm === 'Kritikal' || sm === 'Amaran') sebab.push({ jenis: 'merit', teks: `Merit ${Number(m.Mata_Merit)} (${sm})`, berat: sm === 'Kritikal' ? 3 : 1 });
    return { id: id, nama: String(m.Nama), idKelas: String(m.ID_Kelas), namaKelas: String(m.Nama_Kelas), sebab: sebab, skor: sebab.reduce((a, b) => a + b.berat, 0) };
  }).filter(x => x.sebab.length).sort((a, b) => b.skor - a.skor || a.nama.localeCompare(b.nama));
}
// PK HEM / guru disiplin: semua murid; guru kelas: murid kelasnya sahaja
function apiMuridBerisiko_(p, u) {
  const semua = boleh_(u, 'lihat_sensitif_semua'), kelasSaya = kelasSaya_(u), bulan = /^\d{4}-\d{2}$/.test(String(p.bulan || '')) ? String(p.bulan) : today_().slice(0, 7);
  if (!semua && !kelasSaya.length) return { skop: 'tiada', bulan: bulan, senarai: [], had: HAD_RISIKO };
  const murid = read_('tbl_murid').rows.filter(m => String(m.Status) === 'Aktif' && (semua || kelasSaya.indexOf(String(m.ID_Kelas)) >= 0));
  return { skop: semua ? 'semua' : 'kelas', bulan: bulan, senarai: senaraiBerisiko_(murid, bulan, tetapan_()), had: HAD_RISIKO };
}

/* ---------- Laporan bulanan HEM (PK HEM / Super Admin) ---------- */
function apiLaporanBulanan_(p, u) {
  wajib_(u, 'laporan_hem');
  const bulan = /^\d{4}-\d{2}$/.test(String(p.bulan || '')) ? String(p.bulan) : today_().slice(0, 7), tahun = bulan.slice(0, 4), hariIni = today_();
  const murid = read_('tbl_murid').rows.filter(m => String(m.Status) === 'Aktif'), mm = new Map(murid.map(m => [String(m.ID_Murid), m]));
  const kelas = read_('tbl_kelas').rows.filter(k => String(k.Status) !== 'Tidak Aktif');
  // Kehadiran
  const ikutKelas = {}, lewatMurid = {}, hari = new Set();
  let direkod = 0, hadir = 0, lewat = 0, tidakHadir = 0;
  kehadiranTahun_(tahun).forEach(r => {
    const d = fmtTarikh_(r.Tarikh), m = mm.get(String(r.ID_Murid));
    if (d.slice(0, 7) !== bulan || !m) return;
    const s = statusH_(r.Status), c = ikutKelas[String(m.ID_Kelas)] = ikutKelas[String(m.ID_Kelas)] || { direkod: 0, hadir: 0 };
    hari.add(d); direkod++; c.direkod++;
    if (hadirH_(r.Status)) { hadir++; c.hadir++; }
    if (s === 'Lewat') { lewat++; lewatMurid[String(m.ID_Murid)] = (lewatMurid[String(m.ID_Murid)] || 0) + 1; }
    if (s === 'Tidak Hadir') tidakHadir++;
  });
  const pct = (a, b) => b ? Math.round(a / b * 1000) / 10 : null;
  const kehadiranKelas = kelas.map(k => { const c = ikutKelas[String(k.ID_Kelas)] || { direkod: 0, hadir: 0 }; return { namaKelas: String(k.Nama_Kelas), tahun: String(k.Tahun), peratus: pct(c.hadir, c.direkod) }; })
    .filter(x => x.peratus !== null).sort((a, b) => b.peratus - a.peratus);
  const topLewat = Object.keys(lewatMurid).map(id => ({ nama: String(mm.get(id).Nama), namaKelas: String(mm.get(id).Nama_Kelas), bil: lewatMurid[id] }))
    .sort((a, b) => b.bil - a.bil || a.nama.localeCompare(b.nama)).slice(0, 10);
  // Disiplin & merit
  const katM = {}, katD = {}, demeritMurid = {};
  let meritBil = 0, meritMata = 0, demeritBil = 0, demeritMata = 0;
  read_('tbl_disiplin').rows.forEach(k => {
    if (fmtTarikh_(k.Tarikh_Masa).slice(0, 7) !== bulan) return;
    const mata = Number(k.Mata_Kesan) || 0, n = String(k.Kategori || 'Lain-lain');
    if (mata >= 0) { meritBil++; meritMata += mata; katM[n] = (katM[n] || 0) + 1; }
    else { demeritBil++; demeritMata += mata; katD[n] = (katD[n] || 0) + 1; const id = String(k.ID_Murid); demeritMurid[id] = (demeritMurid[id] || 0) + mata; }
  });
  const susunKat = o => Object.keys(o).map(k => ({ kategori: k, bil: o[k] })).sort((a, b) => b.bil - a.bil).slice(0, 8);
  const topDemerit = Object.keys(demeritMurid).filter(id => mm.has(id)).map(id => ({ nama: String(mm.get(id).Nama), namaKelas: String(mm.get(id).Nama_Kelas), mata: demeritMurid[id] }))
    .sort((a, b) => a.mata - b.mata).slice(0, 10);
  // Kawalan kelas: waktu jadual yang sudah berlalu dalam bulan ini berbanding check in
  const jm = jadualKelasMap_(), rekodKK = rekodKKTahun_(tahun).filter(r => fmtTarikh_(r.Tarikh).slice(0, 7) === bulan);
  const ada = new Set(rekodKK.map(r => String(r.ID_Kelas) + '|' + fmtTarikh_(r.Tarikh) + '|' + r.Waktu));
  let slot = 0, diisi = 0;
  const [y, m0] = bulan.split('-').map(Number), akhir = new Date(Date.UTC(y, m0, 0)).toISOString().slice(0, 10);
  kelas.forEach(k => {
    const j = jm.get(String(k.ID_Kelas)); if (!j) return;
    for (let d = bulan + '-01'; d <= akhir && d <= hariIni; d = tambahHari_(d, 1)) {
      slotHari_(j, d).forEach(s => { if (s.rehat) return; slot++; if (ada.has(String(k.ID_Kelas) + '|' + d + '|' + s.w)) diisi++; });
    }
  });
  // Guru bertugas
  const lap = adaSheet_(SHEET_LAP) ? read_(SHEET_LAP).rows.filter(r => fmtTarikh_(r.Tarikh).slice(0, 7) === bulan) : [];
  const statusLap = {}; lap.forEach(r => { statusLap[String(r.Status)] = (statusLap[String(r.Status)] || 0) + 1; });
  // Unit HEM: OPR yang disahkan bulan ini
  const opr = oprSemua_().filter(o => fmtTarikh_(o.Tarikh).slice(0, 7) === bulan);
  const unitOpr = unitSemua_().filter(r => String(r.Status) !== 'Arkib').map(r => ({ nama: String(r.Nama),
    disahkan: opr.filter(o => String(o.ID_Unit) === String(r.ID_Unit) && String(o.Status) === 'Disahkan').length,
    program: opr.filter(o => String(o.ID_Unit) === String(r.ID_Unit) && String(o.Status) === 'Disahkan').map(o => String(o.Program)) }));
  const berisiko = senaraiBerisiko_(murid, bulan, tetapan_());
  audit_(u.id, 'LAPORAN_BULANAN', bulan);
  return {
    bulan: bulan, jumlahMurid: murid.length,
    kehadiran: { hari: hari.size, direkod: direkod, peratus: pct(hadir, direkod), lewat: lewat, tidakHadir: tidakHadir, ikutKelas: kehadiranKelas, topLewat: topLewat },
    disiplin: { meritBil: meritBil, meritMata: meritMata, demeritBil: demeritBil, demeritMata: demeritMata, kategoriMerit: susunKat(katM), kategoriDemerit: susunKat(katD), topDemerit: topDemerit },
    kawalan: { slot: slot, diisi: diisi, peratus: pct(diisi, slot), lewat: rekodKK.filter(r => String(r.Status_Masa) === 'Lewat').length, ganti: rekodKK.filter(r => isTrue_(r.Ganti)).length, adaJadual: jm.size > 0 },
    bertugas: { jumlah: lap.length, status: statusLap },
    unit: unitOpr, berisiko: { jumlah: berisiko.length, teratas: berisiko.slice(0, 10).map(x => ({ nama: x.nama, namaKelas: x.namaKelas, sebab: x.sebab.map(s => s.teks).join('; ') })) }
  };
}

/* ---------- Naik kelas tahun baharu (PK HEM / Super Admin) ----------
 * peta = { idKelasAsal: idKelasBaru | 'TAMAT' | 'KEKAL' }. Sandaran dibuat dahulu supaya boleh dipulihkan. */
function apiNaikKelas_(p, u) {
  wajib_(u, 'urus_kelas_semua');
  const tahunIni = today_().slice(0, 4);
  if (String(tetapan_().NAIK_KELAS_TAHUN) === tahunIni && !p.paksa) throw err_('Murid sudah dinaikkan kelas pada tahun ' + tahunIni + '. Tandakan "Jalankan semula" jika anda pasti.', 'SUDAH_NAIK');
  const kelas = new Map(read_('tbl_kelas').rows.map(k => [String(k.ID_Kelas), k])), peta = p.peta || {};
  Object.keys(peta).forEach(asal => {
    const s = String(peta[asal]);
    if (!kelas.has(asal)) throw err_('Kelas asal tidak dijumpai: ' + asal);
    if (s !== 'TAMAT' && s !== 'KEKAL' && !kelas.has(s)) throw err_('Kelas sasaran tidak dijumpai bagi ' + kelas.get(asal).Nama_Kelas + '.');
  });
  const sandaran = buatSandaran_();
  const rows = read_('tbl_murid').rows.filter(m => String(m.Status) === 'Aktif' && peta[String(m.ID_Kelas)] && String(peta[String(m.ID_Kelas)]) !== 'KEKAL');
  let naik = 0, tamat = 0;
  rows.forEach(m => {
    const s = String(peta[String(m.ID_Kelas)]);
    if (s === 'TAMAT') { m.Status = 'Tamat'; tamat++; return; }
    const k = kelas.get(s);
    m.ID_Kelas = s; m.Nama_Kelas = String(k.Nama_Kelas); m.Tahun = String(k.Tahun); naik++;
  });
  updateRows_('tbl_murid', rows);
  simpanTetapanKunci_('NAIK_KELAS_TAHUN', tahunIni);
  CacheService.getScriptCache().remove('dashboard');
  audit_(u.id, 'NAIK_KELAS', `naik=${naik} tamat=${tamat} sandaran=${sandaran.nama}`);
  return { naik: naik, tamat: tamat, sandaran: sandaran.nama };
}

/* ---------- BAP: senarai semak penerimaan ----------
 * Semua murid warganegara layak. Guru kelas menanda murid kelasnya yang sudah menerima; PIC / PK HEM semua kelas.
 * Ditanda = rekod agihan berstatus Selesai Agihan; nyahtanda = status kembali Layak (rekod tidak dipadam). */
function skimBap_() { return read_('tbl_kategori_bantuan').rows.find(k => String(k.Kod).toUpperCase() === 'BAP' && String(k.Status_Aktif) === 'Aktif'); }
function kelasBap_(u, idKelas, k) {
  if (!boleh_(u, 'urus_bantuan') && !picSkim_(u, k) && kelasSaya_(u).indexOf(idKelas) < 0) {
    throw err_('Hanya guru kelas tersebut, PIC BAP atau PK HEM boleh membuka senarai ini.', 'FORBIDDEN');
  }
}

// Tanpa idKelas: ringkasan setiap kelas (PIC / PK HEM). Dengan idKelas: senarai murid kelas itu.
function apiGetBap_(p, u) {
  const k = skimBap_();
  if (!k) throw err_('Skim bantuan berkod BAP tidak dijumpai.');
  const tahun = /^\d{4}$/.test(String(p.tahun || '')) ? String(p.tahun) : today_().slice(0, 4);
  const terima = new Map(read_('tbl_agihan_bantuan').rows
    .filter(a => String(a.ID_Kategori) === String(k.ID_Kategori) && String(a.Tahun) === tahun && String(a.Status) === 'Selesai Agihan')
    .map(a => [String(a.ID_Murid), String(a.Tarikh_Serahan || '')]));
  const aktif = read_('tbl_murid').rows.filter(m => String(m.Status) === 'Aktif');
  const idKelas = String(p.idKelas || '');
  if (!idKelas) {
    wajibSkim_(u, k);
    const g = {};
    aktif.forEach(m => {
      const c = g[String(m.ID_Kelas)] = g[String(m.ID_Kelas)] || { idKelas: String(m.ID_Kelas), layak: 0, terima: 0, bukanWarga: 0 };
      if (!warganegara_(m)) { c.bukanWarga++; return; }
      c.layak++; if (terima.has(String(m.ID_Murid))) c.terima++;
    });
    return { tahun: tahun, kelas: Object.keys(g).map(id => g[id]) };
  }
  kelasBap_(u, idKelas, k);
  const dalam = aktif.filter(m => String(m.ID_Kelas) === idKelas);
  return {
    tahun: tahun, idKelas: idKelas, bolehUbah: tahun === today_().slice(0, 4), bukanWarga: dalam.filter(m => !warganegara_(m)).length,
    murid: dalam.filter(warganegara_).sort((a, b) => String(a.Nama).localeCompare(String(b.Nama))).map(m => ({
      id: String(m.ID_Murid), nama: String(m.Nama), terima: terima.has(String(m.ID_Murid)), tarikh: terima.get(String(m.ID_Murid)) || '' }))
  };
}

// ubah: [{ idMurid, terima }] — hanya murid yang berubah
function apiSimpanBap_(p, u) {
  const k = skimBap_();
  if (!k) throw err_('Skim bantuan berkod BAP tidak dijumpai.');
  const idKelas = String(p.idKelas || '');
  kelasBap_(u, idKelas, k);
  const ubah = Array.isArray(p.ubah) ? p.ubah : [];
  if (ubah.length > 500) throw err_('Terlalu banyak perubahan dalam satu simpanan.');
  const tahun = today_().slice(0, 4), hariIni = today_(), now = new Date();
  const sah = new Set(read_('tbl_murid').rows.filter(m => String(m.ID_Kelas) === idKelas && String(m.Status) === 'Aktif' && warganegara_(m)).map(m => String(m.ID_Murid)));
  const sedia = new Map(read_('tbl_agihan_bantuan').rows.filter(a => String(a.ID_Kategori) === String(k.ID_Kategori) && String(a.Tahun) === tahun).map(a => [String(a.ID_Murid), a]));
  const baru = [], kemas = [];
  ubah.forEach(x => {
    const id = String(x.idMurid), r = sedia.get(id);
    if (!sah.has(id)) throw err_('Murid ' + id + ' bukan dalam senarai BAP kelas ini.');
    if (!!x.terima === (!!r && String(r.Status) === 'Selesai Agihan')) return;
    if (r) { Object.assign(r, { Status: x.terima ? 'Selesai Agihan' : 'Layak', Tarikh_Serahan: x.terima ? hariIni : '', ID_Guru_Rekod: u.id, Timestamp: now }); kemas.push(r); }
    else baru.push({ ID_Agihan: genId_('AG'), ID_Kategori: String(k.ID_Kategori), ID_Murid: id, Tahun: tahun, Status: 'Selesai Agihan', Tarikh_Serahan: hariIni, Amaun: '', Catatan: '', ID_Guru_Rekod: u.id, Timestamp: now });
  });
  updateRows_('tbl_agihan_bantuan', kemas);
  insert_('tbl_agihan_bantuan', baru);
  CacheService.getScriptCache().remove('dashboard');
  audit_(u.id, 'SEMAK_BAP', `${idKelas} ubah=${baru.length + kemas.length}`);
  return { disimpan: baru.length + kemas.length };
}

// Data pendapatan semua murid untuk menilai kelayakan: PIC perlu melihat murid di luar kelasnya, bagi skimnya sahaja
function apiCalonSkim_(p, u) {
  const kat = read_('tbl_kategori_bantuan').rows.find(k => String(k.ID_Kategori) === String(p.idKategori));
  if (!kat) throw err_('Kategori bantuan tidak sah.');
  wajibSkim_(u, kat);
  const t = tetapan_();
  audit_(u.id, 'LIHAT_CALON_SKIM', String(kat.Kod));
  return read_('tbl_murid').rows.filter(m => String(m.Status) === 'Aktif').map(m => {
    const pk = perKapita_(m, t);
    return { id: String(m.ID_Murid), pendapatan: kosongPendapatan_(m) ? null : Number(m.Pendapatan_Isi_Rumah),
      kategoriPendapatan: String(m.Kategori_Pendapatan || ''), perKapita: pk.perKapita, bilIsiRumah: pk.bil, yatim: String(m.Yatim || '') };
  });
}

// Tambah lajur skema yang belum wujud (cth selepas naik taraf versi) tanpa perlu jalankan setup
function pastiLajur_(name) {
  const sh = sheet_(name);
  const ada = headers_(sh);
  if (SCHEMA[name].headers.every(h => ada.indexOf(h) >= 0)) return;
  ensureSheet_(SpreadsheetApp.getActiveSpreadsheet(), name, SCHEMA[name]);
  delete _MEMO['h:' + name]; delete _MEMO['r:' + name];
}

/* ---------- Guru ---------- */
/* ==================================================================
 *  PENDAFTARAN AKAUN GURU (awam, tanpa log masuk)
 * ================================================================== */
function infoDaftar_() {
  const t = tetapan_();
  const mod = ['Kelulusan', 'Automatik', 'Tutup'].indexOf(t.PENDAFTARAN_GURU) >= 0 ? t.PENDAFTARAN_GURU : 'Kelulusan';
  if (mod === 'Tutup') return { dibuka: false, namaSekolah: t.NAMA_SEKOLAH || '' };
  // Hanya maklumat asas kelas; tiada nama guru atau murid didedahkan
  const kelas = read_('tbl_kelas').rows.filter(k => String(k.Status) !== 'Tidak Aktif')
    .sort((a, b) => tahunSort_(a.Tahun) - tahunSort_(b.Tahun) || String(a.Nama_Kelas).localeCompare(String(b.Nama_Kelas)))
    .map(k => ({ id: String(k.ID_Kelas), nama: String(k.Nama_Kelas), tahun: String(k.Tahun), sesi: String(k.Sesi || ''), adaGuru: !!k.ID_Guru_Kelas }));
  return { dibuka: true, mod: mod, domain: String(t.DOMAIN_EMEL || 'moe-dl.edu.my').replace(/^@/, ''), namaSekolah: t.NAMA_SEKOLAH || '', kelas: kelas };
}

function daftarGuru_(p) {
  ensureTetapan_();          // dipanggil di bawah lock (doPost)
  const info = infoDaftar_();
  if (!info.dibuka) throw err_('Pendaftaran akaun ditutup. Hubungi PK HEM.');

  const id = String(p.idEmel || '').trim().toLowerCase().replace(/@.*$/, '');
  if (!/^[a-z0-9._-]{3,64}$/.test(id)) throw err_('ID emel MOE tidak sah. Masukkan bahagian sebelum @ sahaja.');
  const emel = id + '@' + info.domain;

  // Had cubaan: 3 sejam bagi setiap emel + had keseluruhan yang longgar (seorang penyerang tidak boleh
  // menyekat pendaftaran guru lain dengan mudah)
  const cache = CacheService.getScriptCache();
  const kE = 'daftar_e_' + id, bilE = Number(cache.get(kE) || 0);
  const bil = Number(cache.get('daftar_bil') || 0);
  if (bilE >= 3) throw err_('Terlalu banyak cubaan untuk emel ini. Cuba lagi selepas sejam.');
  if (bil >= 150) throw err_('Terlalu banyak pendaftaran dalam masa singkat. Cuba lagi sebentar.');
  cache.put(kE, String(bilE + 1), 3600);
  cache.put('daftar_bil', String(bil + 1), 600);
  const nama = String(p.nama || '').replace(/\s+/g, ' ').trim().toUpperCase();
  if (nama.length < 3) throw err_('Sila isi nama penuh.');
  const kp = noKP_(p.noKP);
  if (!/^\d{12}$/.test(kp)) throw err_('No. KP mesti 12 digit.');
  const tel = String(p.tel || '').replace(/[^\d+]/g, '');
  if (tel.replace(/\D/g, '').length < 9) throw err_('No. telefon tidak sah.');
  const pw = String(p.pin || '');
  if (pw.length < APP.MIN_KATA_LALUAN) throw err_('Kata laluan mesti sekurang-kurangnya ' + APP.MIN_KATA_LALUAN + ' aksara.');
  if (pw.length > 64) throw err_('Kata laluan terlalu panjang.');
  if (pw === kp) throw err_('Kata laluan tidak boleh sama dengan No. KP.');

  const guru = read_('tbl_guru').rows;
  const ada = guru.find(g => String(g.Emel).trim().toLowerCase() === emel && String(g.Status_Aktif) !== 'Dipadam');
  if (ada) {
    if (String(ada.Status_Aktif) === 'Menunggu') throw err_('Emel ini sudah didaftarkan dan sedang menunggu kelulusan.');
    throw err_('Emel ini sudah mempunyai akaun. Sila log masuk, atau minta PK HEM tetapkan semula kata laluan.');
  }
  if (guru.some(g => String(g.No_KP) === kp && ['Aktif', 'Menunggu'].indexOf(String(g.Status_Aktif)) >= 0)) {
    throw err_('No. KP ini sudah didaftarkan dengan emel lain. Hubungi PK HEM.');
  }

  const kelasPilih = Array.from(new Set((Array.isArray(p.kelas) ? p.kelas : []).map(String).filter(Boolean))).slice(0, 2);
  const kInfo = new Map(info.kelas.map(k => [k.id, k]));
  kelasPilih.forEach(k => {
    if (!kInfo.has(k)) throw err_('Kelas tidak dijumpai.');
    if (kInfo.get(k).adaGuru) throw err_('Kelas ' + kInfo.get(k).nama + ' sudah mempunyai guru kelas. Hubungi PK HEM.');
  });

  // Mod Automatik: akaun aktif hanya selepas guru membuka pautan pengesahan yang dihantar ke emel MOE
  // (membuktikan pemilikan emel). Jika emel gagal dihantar, akaun kekal menunggu kelulusan PK HEM.
  const auto = info.mod === 'Automatik';
  const g = {
    ID_Guru: nextId_(guru.map(r => r.ID_Guru), 'G', 3), Nama_Penuh: nama, Emel: emel, No_Telefon: tel,
    Jawatan: 'Guru', Gred: '', Subjek: '', Peranan: 'guru', Sesi: '', Foto_ID: '',
    PIN_Hash: hash_(pw), Tukar_PIN: 'FALSE', Status_Aktif: 'Menunggu',
    Tarikh_Daftar: new Date(), No_KP: kp, Kelas_Mohon: JSON.stringify(kelasPilih), Cara_Daftar: 'Sendiri'
  };
  insert_('tbl_guru', [g]);
  let emelDihantar = false;
  if (auto) emelDihantar = hantarPengesahan_(g);
  audit_(g.ID_Guru, 'DAFTAR_SENDIRI', `${emel} ${emelDihantar ? 'pengesahan emel' : 'menunggu'} kelas=${kelasPilih.join(',')}`);
  return {
    emel: emel, status: g.Status_Aktif, kelas: [],
    mesej: emelDihantar
      ? 'Pendaftaran diterima. Buka pautan pengesahan yang dihantar ke ' + emel + ' (sah 6 jam) untuk mengaktifkan akaun.'
      : 'Pendaftaran diterima. Akaun anda akan aktif selepas diluluskan oleh PK HEM.'
  };
}

function hantarPengesahan_(g) {
  try {
    const url = ScriptApp.getService().getUrl();
    if (!url) return false;
    const kod = Utilities.getUuid().replace(/-/g, '').toLowerCase();
    CacheService.getScriptCache().put('sahkan_' + kod, String(g.ID_Guru), 6 * 3600);
    const pautan = url + '?sahkan=' + kod;
    MailApp.sendEmail({
      to: g.Emel, name: namaApp_(),
      subject: namaApp_() + ': Sahkan emel anda',
      body: 'Salam sejahtera ' + g.Nama_Penuh + ',\n\nBuka pautan di bawah untuk mengaktifkan akaun ' + namaApp_() +
        ' anda (sah selama 6 jam):\n\n' + pautan + '\n\nJika anda tidak mendaftar, abaikan emel ini.'
    });
    return true;
  } catch (e) {
    return false;   // kuota emel habis / kebenaran MailApp belum diberi
  }
}

/* ==================================================================
 *  LUPA KATA LALUAN (awam): kod 6 digit dihantar ke emel guru, sah 15 minit, maksimum 5 cubaan.
 *  Jawapan sentiasa sama sama ada emel wujud atau tidak (tidak mendedahkan akaun berdaftar).
 * ================================================================== */
const LUPA_SAAT = 15 * 60;

function lupaPin_(p) {
  const emel = String(p.emel || '').trim().toLowerCase();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(emel)) throw err_('Masukkan emel yang sah.');
  const cache = CacheService.getScriptCache();
  const kE = 'lupa_bil_' + emel, bilE = Number(cache.get(kE) || 0), bil = Number(cache.get('lupa_bil') || 0);
  if (bilE >= 3) throw err_('Terlalu banyak permintaan untuk emel ini. Cuba lagi selepas sejam.');
  if (bil >= 100) throw err_('Terlalu banyak permintaan. Cuba lagi sebentar.');
  cache.put(kE, String(bilE + 1), 3600);
  cache.put('lupa_bil', String(bil + 1), 600);
  const jawapan = { mesej: 'Jika emel ini berdaftar dan aktif, kod 6 digit telah dihantar ke ' + emel + '. Kod sah selama 15 minit. Semak juga folder Spam.' };

  const g = read_('tbl_guru').rows.find(r => String(r.Emel).trim().toLowerCase() === emel && String(r.Status_Aktif) === 'Aktif');
  if (!g) return jawapan;
  const b = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, Utilities.getUuid() + Utilities.getUuid());
  const kod = String(((b[0] & 0xff) * 65536 + (b[1] & 0xff) * 256 + (b[2] & 0xff)) % 1000000).padStart(6, '0');
  cache.put('lupa_' + emel, JSON.stringify({ h: hashKod_(emel, kod), cuba: 0 }), LUPA_SAAT);
  try {
    MailApp.sendEmail({
      to: emel, name: namaApp_(),
      subject: namaApp_() + ': Kod tetapan semula kata laluan',
      body: 'Salam sejahtera ' + g.Nama_Penuh + ',\n\nKod untuk menetapkan semula kata laluan ' + namaApp_() + ' anda:\n\n    ' + kod +
        '\n\nKod ini sah selama 15 minit. Jika anda tidak meminta kod ini, abaikan emel ini; kata laluan anda tidak berubah.'
    });
    audit_(String(g.ID_Guru), 'LUPA_PIN_KOD', emel);
  } catch (e) {
    cache.remove('lupa_' + emel);
    audit_(String(g.ID_Guru), 'LUPA_PIN_GAGAL_EMEL', String(e.message).slice(0, 200));
  }
  return jawapan;
}

function resetPinKod_(p) {
  const emel = String(p.emel || '').trim().toLowerCase();
  const kod = String(p.kod || '').replace(/\D/g, '');
  const baru = String(p.pinBaru || '');
  const cache = CacheService.getScriptCache();
  const c = parseJSON_(cache.get('lupa_' + emel), null);
  if (!c) throw err_('Kod tidak sah atau telah tamat tempoh. Minta kod baharu.');
  if (c.cuba >= 5) { cache.remove('lupa_' + emel); throw err_('Terlalu banyak cubaan salah. Minta kod baharu.'); }
  if (hashKod_(emel, kod) !== c.h) {
    c.cuba++; cache.put('lupa_' + emel, JSON.stringify(c), LUPA_SAAT);
    throw err_('Kod salah. Baki cubaan: ' + (5 - c.cuba) + '.');
  }
  if (baru.length < APP.MIN_KATA_LALUAN) throw err_('Kata laluan baharu mesti sekurang-kurangnya ' + APP.MIN_KATA_LALUAN + ' aksara.');
  if (baru.length > 64) throw err_('Kata laluan terlalu panjang (maksimum 64 aksara).');
  const g = read_('tbl_guru').rows.find(r => String(r.Emel).trim().toLowerCase() === emel && String(r.Status_Aktif) === 'Aktif');
  if (!g) { cache.remove('lupa_' + emel); throw err_('Akaun tidak aktif. Hubungi PK HEM.'); }
  if (String(g.No_KP || '') && baru === String(g.No_KP)) throw err_('Kata laluan tidak boleh sama dengan No. KP.');
  g.PIN_Hash = hash_(baru);
  g.Tukar_PIN = 'FALSE';
  updateRow_('tbl_guru', g);
  cache.removeAll(['lupa_' + emel, 'cuba_' + emel]);   // kod sekali guna; buka sekatan log masuk
  batalSesi_(g.ID_Guru);                                 // tamatkan sesi lama di semua peranti
  audit_(String(g.ID_Guru), 'RESET_PIN_SENDIRI', emel);
  return { berjaya: true, emel: emel };
}

// Kod disimpan sebagai hash (dengan SALT rahsia skrip) supaya tidak terdedah dalam cache
function hashKod_(emel, kod) {
  const salt = PropertiesService.getScriptProperties().getProperty('SALT') || '';
  return hex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + '|' + emel + '|' + kod, Utilities.Charset.UTF_8));
}

/* ==================================================================
 *  TETAPAN ASAS: nama aplikasi, nama sekolah & logo (logo PNG telus, disimpan dalam folder foto)
 * ================================================================== */
const KUNCI_SISTEM = ['FOLDER_FOTO_ID', 'LOGO_ID', 'SPBT_BUKU', 'RMT_TETAPAN', 'CARTA_HEM', 'FOLDER_SANDARAN_ID', 'SANDARAN_TERAKHIR', 'NAIK_KELAS_TAHUN'];
function namaApp_() { return String(tetapan_().NAMA_APP || APP.NAMA); }

// Awam (skrin log masuk): logo hanya dihantar jika versi di pelayar berbeza
function infoAwam_(p) {
  const t = tetapan_();
  const versi = String(t.LOGO_ID || '');
  const out = { namaApp: t.NAMA_APP || APP.NAMA, namaSekolah: t.NAMA_SEKOLAH || '', logoVersi: versi };
  if (versi && String(p.versi || '') !== versi) out.logo = logoDataUrl_(versi);
  return out;
}

function logoDataUrl_(id) {
  const cache = CacheService.getScriptCache();
  const c = cache.get('logo:' + id);
  if (c) return c;
  try {
    const folder = folderFoto_(false);
    if (!folder) return '';
    const url = fotoDataUrl_(id, folder.getId());
    try { cache.put('logo:' + id, url, 21600); } catch (e) { /* terlalu besar untuk cache */ }
    return url;
  } catch (e) { return ''; }
}

function apiSimpanLogo_(p, u) {
  wajib_(u, 'urus_tetapan');
  ensureTetapan_();
  const r = read_('tbl_tetapan').rows.find(x => String(x.Kunci) === 'LOGO_ID');
  const lama = r ? String(r.Nilai || '') : '';
  let baru = '';
  if (!p.buang) {
    const m = String(p.dataUrl || '').match(/^data:image\/png;base64,(.+)$/);
    if (!m) throw err_('Logo mesti dalam format PNG.');
    const bytes = Utilities.base64Decode(m[1]);
    if (bytes.length > 400 * 1024) throw err_('Saiz logo melebihi 400KB. Gunakan gambar yang lebih kecil.');
    const file = folderFoto_().createFile(Utilities.newBlob(bytes, 'image/png', 'logo_' + genId_('L') + '.png'));
    baru = file.getId();
  }
  r.Nilai = baru;
  updateRow_('tbl_tetapan', r);
  if (lama && lama !== baru) {
    try { DriveApp.getFileById(lama).setTrashed(true); } catch (e) { /* abaikan */ }
    CacheService.getScriptCache().remove('logo:' + lama);
  }
  audit_(u.id, p.buang ? 'BUANG_LOGO' : 'SIMPAN_LOGO', baru);
  return { logoVersi: baru, logo: baru ? logoDataUrl_(baru) : '' };
}

// Tetapkan guru kelas bagi kelas yang masih kosong; pulangkan kelas yang berjaya diberi
function beriKelas_(idGuru, ids) {
  const want = new Set(ids);
  const ubah = read_('tbl_kelas').rows.filter(k => want.has(String(k.ID_Kelas)) && !k.ID_Guru_Kelas && String(k.Status) !== 'Tidak Aktif');
  ubah.forEach(k => { k.ID_Guru_Kelas = idGuru; });
  updateRows_('tbl_kelas', ubah);
  CacheService.getScriptCache().remove('dashboard');
  return ubah.map(k => String(k.ID_Kelas));
}

function apiSenaraiPendaftaran_(p, u) {
  wajib_(u, 'urus_guru');
  const kelas = new Map(read_('tbl_kelas').rows.map(k => [String(k.ID_Kelas), k]));
  return read_('tbl_guru').rows.filter(g => String(g.Status_Aktif) === 'Menunggu').map(g => ({
    id: String(g.ID_Guru), nama: String(g.Nama_Penuh), emel: String(g.Emel), tel: String(g.No_Telefon),
    noKP: String(g.No_KP || ''), tarikh: g.Tarikh_Daftar,
    kelas: parseJSON_(g.Kelas_Mohon, []).map(id => {
      const k = kelas.get(String(id)) || {};
      return { id: String(id), nama: String(k.Nama_Kelas || id), adaGuru: !!k.ID_Guru_Kelas };
    })
  }));
}

function apiPutusPendaftaran_(p, u) {
  wajib_(u, 'urus_guru');
  const g = read_('tbl_guru').rows.find(r => String(r.ID_Guru) === String(p.id) && String(r.Status_Aktif) === 'Menunggu');
  if (!g) throw err_('Pendaftaran tidak dijumpai atau sudah diproses.');
  if (p.lulus) {
    g.Status_Aktif = 'Aktif';
    updateRow_('tbl_guru', g);
    const diberi = beriKelas_(String(g.ID_Guru), parseJSON_(g.Kelas_Mohon, []).map(String));
    audit_(u.id, 'LULUS_DAFTAR', `${g.ID_Guru} kelas=${diberi.join(',')}`);
    return { status: 'Aktif', guru: guruAwam_(g), kelasDiberi: diberi };
  }
  g.Status_Aktif = 'Tidak Aktif';
  updateRow_('tbl_guru', g);
  audit_(u.id, 'TOLAK_DAFTAR', String(g.ID_Guru));
  return { status: 'Ditolak' };
}

/* ---------- Import murid dari fail Excel (dimuat naik dalam Tetapan) ---------- */
function apiImportMuridExcel_(p, u) {
  wajib_(u, 'urus_tetapan');
  const rows = Array.isArray(p.rows) ? p.rows : [];
  if (rows.length < 2) throw err_('Fail tiada data murid.');
  if (rows.length > 6000) throw err_('Fail terlalu besar (maksimum 6000 baris).');
  const vals = rows.map(r => (Array.isArray(r) ? r : []).map(v => v === null || v === undefined ? '' : String(v)));
  const m = prosesImportIDME_(vals);
  audit_(u.id, 'IMPORT_MURID_EXCEL', String(p.namaFail || '').slice(0, 100));
  return { mesej: m.replace(/^✅\s*/, '') };
}

function apiSimpanGuru_(p, u) {
  const rows = read_('tbl_guru').rows;
  const admin = boleh_(u, 'urus_guru');

  if (p.id) {
    const g = rows.find(r => String(r.ID_Guru) === String(p.id));
    if (!g) throw err_('Guru tidak dijumpai.');
    const diriSendiri = String(g.ID_Guru) === u.id;
    if (!admin && !diriSendiri) throw err_('Anda hanya boleh mengemas kini profil sendiri.', 'FORBIDDEN');
    // Akaun Super Admin hanya boleh diubah oleh Super Admin (atau pemiliknya sendiri)
    if (String(g.Peranan) === 'superadmin' && u.peranan !== 'superadmin' && !diriSendiri) {
      throw err_('Hanya Super Admin boleh mengubah akaun Super Admin.', 'FORBIDDEN');
    }

    // Medan yang boleh dikemas kini sendiri
    if (p.tel !== undefined) g.No_Telefon = String(p.tel).slice(0, 20);
    if (p.subjek !== undefined) g.Subjek = String(p.subjek).slice(0, 200);
    if (p.fotoId !== undefined) g.Foto_ID = idFail_(p.fotoId);
    if (admin) {
      ['nama:Nama_Penuh', 'jawatan:Jawatan', 'gred:Gred', 'sesi:Sesi'].forEach(pair => {
        const [k, h] = pair.split(':');
        if (p[k] !== undefined) g[h] = String(p[k]);
      });
      if (p.emel !== undefined) {
        const e = String(p.emel).trim().toLowerCase();
        if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(e)) throw err_('Emel tidak sah.');
        if (rows.some(r => r !== g && String(r.Status_Aktif) !== 'Dipadam' && String(r.Emel).trim().toLowerCase() === e)) throw err_('Emel sudah digunakan.');
        g.Emel = e;
      }
      if (p.status !== undefined) {
        const baru = p.status === 'Aktif' ? 'Aktif' : 'Tidak Aktif';
        if (baru !== 'Aktif' && String(g.ID_Guru) === u.id) throw err_('Anda tidak boleh menyahaktifkan akaun sendiri.');
        if (baru !== 'Aktif' && g.Peranan === 'superadmin') pastikanSuperAdmin_(g);
        if (baru !== String(g.Status_Aktif)) batalSesi_(g.ID_Guru);
        g.Status_Aktif = baru;
      }
      if (p.peranan !== undefined) setPeranan_(g, p.peranan, u);
    }
    updateRow_('tbl_guru', g);
    audit_(u.id, 'KEMASKINI_GURU', String(g.ID_Guru));
    return guruAwam_(g);
  }

  wajib_(u, 'urus_guru');
  const nama = String(p.nama || '').trim();
  const emel = String(p.emel || '').trim().toLowerCase();
  if (!nama || !emel) throw err_('Sila isi nama dan emel guru.');
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(emel)) throw err_('Emel tidak sah.');
  if (rows.some(r => String(r.Status_Aktif) !== 'Dipadam' && String(r.Emel).trim().toLowerCase() === emel)) throw err_('Emel sudah didaftarkan.');

  const pin = pinSementara_();
  const g = {
    ID_Guru: nextId_(rows.map(r => r.ID_Guru), 'G', 3), Nama_Penuh: nama, Emel: emel,
    No_Telefon: String(p.tel || ''), Jawatan: String(p.jawatan || 'Guru'), Gred: String(p.gred || ''),
    Subjek: String(p.subjek || ''), Peranan: 'guru', Sesi: String(p.sesi || ''),
    Foto_ID: idFail_(p.fotoId), PIN_Hash: hash_(pin), Tukar_PIN: 'TRUE',
    Status_Aktif: 'Aktif', Tarikh_Daftar: new Date()
  };
  if (p.peranan) setPeranan_(g, p.peranan, u);
  insert_('tbl_guru', [g]);
  audit_(u.id, 'TAMBAH_GURU', g.ID_Guru);
  return Object.assign(guruAwam_(g), { pinSementara: pin });
}

function setPeranan_(g, peranan, u) {
  if (!PERANAN[peranan]) throw err_('Peranan tidak sah.');
  if ((peranan === 'superadmin' || g.Peranan === 'superadmin') && u.peranan !== 'superadmin') {
    throw err_('Hanya Super Admin boleh memberi atau menarik peranan Super Admin.', 'FORBIDDEN');
  }
  if (g.Peranan === 'superadmin' && peranan !== 'superadmin') pastikanSuperAdmin_(g);
  if (String(g.Peranan) !== peranan && g.ID_Guru) batalSesi_(g.ID_Guru);
  g.Peranan = peranan;
}

// Mesti sentiasa ada sekurang-kurangnya seorang Super Admin aktif
function pastikanSuperAdmin_(g) {
  const lain = read_('tbl_guru').rows.filter(r => r !== g && String(r.ID_Guru) !== String(g.ID_Guru) &&
    String(r.Peranan) === 'superadmin' && String(r.Status_Aktif) === 'Aktif');
  if (!lain.length) throw err_('Mesti ada sekurang-kurangnya seorang Super Admin. Lantik Super Admin lain dahulu.');
}

// Paksa pengguna log masuk semula (peranan ditukar, dinyahaktif atau dipadam)
function batalSesi_(idGuru) {
  CacheService.getScriptCache().put('rev_' + idGuru, String(Date.now()), APP.TOKEN_SAAT + 60);
}

// Semua pengguna (aktif & tidak aktif) untuk Urus Pengguna
function apiSenaraiPengguna_(p, u) {
  wajib_(u, 'urus_guru');
  return read_('tbl_guru').rows.filter(g => ['Aktif', 'Tidak Aktif'].indexOf(String(g.Status_Aktif)) >= 0)
    .map(g => Object.assign(guruAwam_(g), { status: String(g.Status_Aktif), tarikhDaftar: g.Tarikh_Daftar }));
}

/* ---------- Padam pengguna (soft delete: rekod lama masih memaparkan nama) ---------- */
function apiPadamGuru_(p, u) {
  wajib_(u, 'urus_guru');
  const g = read_('tbl_guru').rows.find(r => String(r.ID_Guru) === String(p.id) && String(r.Status_Aktif) !== 'Dipadam');
  if (!g) throw err_('Pengguna tidak dijumpai.');
  if (String(g.ID_Guru) === u.id) throw err_('Anda tidak boleh memadam akaun sendiri.');
  if (g.Peranan === 'superadmin') {
    if (u.peranan !== 'superadmin') throw err_('Hanya Super Admin boleh memadam Super Admin.', 'FORBIDDEN');
    pastikanSuperAdmin_(g);
  }
  g.Status_Aktif = 'Dipadam'; g.PIN_Hash = ''; g.Peranan = 'guru';
  updateRow_('tbl_guru', g);
  const kelas = read_('tbl_kelas').rows.filter(k => String(k.ID_Guru_Kelas) === String(g.ID_Guru));
  kelas.forEach(k => { k.ID_Guru_Kelas = ''; });
  updateRows_('tbl_kelas', kelas);
  const skim = read_('tbl_kategori_bantuan').rows.filter(k => String(k.ID_Guru_PIC || '') === String(g.ID_Guru));
  skim.forEach(k => { k.ID_Guru_PIC = ''; });       // skim kembali tanpa PIC
  updateRows_('tbl_kategori_bantuan', skim);
  batalSesi_(g.ID_Guru);
  CacheService.getScriptCache().remove('dashboard');
  audit_(u.id, 'PADAM_GURU', `${g.ID_Guru} ${g.Emel}`);
  return { id: String(g.ID_Guru), kelasDilepaskan: kelas.map(k => String(k.ID_Kelas)) };
}

function apiResetPinGuru_(p, u) {
  wajib_(u, 'urus_guru');
  const g = read_('tbl_guru').rows.find(r => String(r.ID_Guru) === String(p.id));
  if (!g) throw err_('Guru tidak dijumpai.');
  if (g.Peranan === 'superadmin' && u.peranan !== 'superadmin') throw err_('Tiada kebenaran.', 'FORBIDDEN');
  if (String(g.Status_Aktif) === 'Dipadam') throw err_('Akaun ini telah dipadam.');
  const pin = pinSementara_();
  g.PIN_Hash = hash_(pin);
  g.Tukar_PIN = 'TRUE';
  updateRow_('tbl_guru', g);
  batalSesi_(g.ID_Guru);          // sesi lama ditamatkan serta-merta
  audit_(u.id, 'RESET_PIN', String(g.ID_Guru));
  return { pinSementara: pin };
}

/* ---------- Tetapan ---------- */
function apiSimpanTetapan_(p, u) {
  wajib_(u, 'urus_tetapan');
  const nilai = p.nilai || {};
  const tb = read_('tbl_tetapan');
  const NOMBOR = { MERIT_AWAL: [0, 1000], MERIT_MAKS: [1, 1000], AMBANG_AMARAN: [0, 1000], AMBANG_KRITIKAL: [0, 1000],
    HAD_B40: [0, 1e7], HAD_M40: [0, 1e7] };
  const PILIHAN = { TEMA_LALAI: TEMA_SAH, LATAR_LALAI: LATAR_SAH, PENDAFTARAN_GURU: ['Kelulusan', 'Automatik', 'Tutup'], PERKAPITA_KIRA_PENJAGA: ['Ya', 'Tidak'] };

  // Sahkan SEMUA nilai dahulu supaya tiada simpanan separuh jika satu nilai salah
  const ubah = tb.rows.filter(r => KUNCI_SISTEM.indexOf(String(r.Kunci)) < 0 && nilai[String(r.Kunci)] !== undefined);
  const baru = {};
  ubah.forEach(r => {
    const k = String(r.Kunci), v = String(nilai[k]).trim();
    if (PILIHAN[k] && PILIHAN[k].indexOf(v) < 0) throw err_('Nilai ' + k + ' tidak sah.');
    if ((k === 'NAMA_APP' || k === 'NAMA_SEKOLAH') && (!v || v.length > 120)) throw err_((k === 'NAMA_APP' ? 'Nama aplikasi' : 'Nama sekolah') + ' wajib diisi (maksimum 120 aksara).');
    if (NOMBOR[k]) {
      const n = Number(v);
      if (v === '' || isNaN(n) || n < NOMBOR[k][0] || n > NOMBOR[k][1]) throw err_(k + ' mesti nombor antara ' + NOMBOR[k][0] + ' dan ' + NOMBOR[k][1] + '.');
    }
    if (k === 'SESI_TAHUN') {
      const o = parseJSON_(v, null);
      if (!o || typeof o !== 'object' || Array.isArray(o)) throw err_('SESI_TAHUN mesti JSON yang sah.');
    }
    baru[k] = v.slice(0, 2000);
  });
  const sedia = tetapan_();
  const nilaiAkhir = k => {
    const v = baru[k] !== undefined ? baru[k] : sedia[k], n = Number(v);
    return v === undefined || v === '' || isNaN(n) ? Number((TETAPAN_LALAI.find(r => r[0] === k) || [])[1]) : n;
  };
  if (nilaiAkhir('AMBANG_KRITIKAL') > nilaiAkhir('AMBANG_AMARAN')) throw err_('AMBANG_KRITIKAL mesti tidak melebihi AMBANG_AMARAN.');
  if (nilaiAkhir('HAD_B40') > nilaiAkhir('HAD_M40')) throw err_('HAD_B40 mesti tidak melebihi HAD_M40.');
  if (nilaiAkhir('MERIT_AWAL') > nilaiAkhir('MERIT_MAKS')) throw err_('MERIT_AWAL mesti tidak melebihi MERIT_MAKS.');

  ubah.forEach(r => { r.Nilai = baru[String(r.Kunci)]; });
  updateRows_('tbl_tetapan', ubah);
  const n = ubah.length;
  CacheService.getScriptCache().remove('dashboard');
  audit_(u.id, 'SIMPAN_TETAPAN', n + ' kunci');
  return { dikemaskini: n };
}

/* ---------- Sesi persekolahan ikut tahun (Prasekolah, Tahun 1–6) ---------- */
// Kunci peta dinormalkan (cth "TAHUN 1" → "TAHUN1") supaya padan dengan label IDME
function petaSesiTahun_() {
  const peta = parseJSON_(tetapan_().SESI_TAHUN, {});
  const out = {};
  Object.keys(peta || {}).forEach(k => { if (peta[k] === 'Pagi' || peta[k] === 'Petang') out[namaKunci_(k)] = peta[k]; });
  return out;
}

function apiSimpanSesiTahun_(p, u) {
  wajib_(u, 'urus_tetapan');
  ensureTetapan_();
  const masuk = p.peta || {};
  const peta = {};
  Object.keys(masuk).forEach(t => {
    const v = String(masuk[t] || '');
    if (v && v !== 'Pagi' && v !== 'Petang') throw err_('Sesi mesti Pagi atau Petang.');
    if (v) peta[String(t)] = v;
  });
  const r = read_('tbl_tetapan').rows.find(x => String(x.Kunci) === 'SESI_TAHUN');
  r.Nilai = JSON.stringify(peta);
  updateRow_('tbl_tetapan', r);

  // Guna pada semua kelas aktif dalam tahun berkenaan
  const norm = {}; Object.keys(peta).forEach(t => { norm[namaKunci_(t)] = peta[t]; });
  const ubah = read_('tbl_kelas').rows.filter(k => {
    const v = norm[namaKunci_(k.Tahun)];
    if (!v || String(k.Status) === 'Tidak Aktif' || String(k.Sesi) === v) return false;
    k.Sesi = v; return true;
  });
  updateRows_('tbl_kelas', ubah);
  CacheService.getScriptCache().remove('dashboard');
  audit_(u.id, 'SESI_TAHUN', JSON.stringify(peta));
  return { peta: peta, kelasDikemaskini: ubah.length };
}

function apiGetTetapan_(p, u) {
  wajib_(u, 'urus_tetapan');
  ensureTetapan_();
  return read_('tbl_tetapan').rows
    .filter(r => KUNCI_SISTEM.indexOf(String(r.Kunci)) < 0)
    .map(r => ({ kunci: String(r.Kunci), nilai: String(r.Nilai), keterangan: String(r.Keterangan) }));
}

/* ---------- Murid: foto & tag (guru kelas / PK HEM) ---------- */
function apiKemaskiniMurid_(p, u) {
  const m = read_('tbl_murid').rows.find(r => String(r.ID_Murid) === String(p.idMurid));
  if (!m) throw err_('Murid tidak dijumpai.');
  if (!boleh_(u, 'urus_kelas_semua') && kelasSaya_(u).indexOf(String(m.ID_Kelas)) < 0) {
    throw err_('Hanya guru kelas atau PK HEM boleh mengubah maklumat murid ini.', 'FORBIDDEN');
  }
  if (p.fotoId !== undefined) m.Foto_ID = idFail_(p.fotoId);
  if (p.tag !== undefined) m.Tag = String(p.tag).slice(0, 100);
  updateRow_('tbl_murid', m);
  audit_(u.id, 'KEMASKINI_MURID', String(m.ID_Murid));
  return { idMurid: String(m.ID_Murid), fotoId: String(m.Foto_ID || ''), tag: String(m.Tag || '') };
}

// Muat naik + tetapkan foto murid dalam SATU panggilan (untuk ambil gambar pukal); foto lama dibuang ke Trash
function apiFotoMurid_(p, u) {
  const m = read_('tbl_murid').rows.find(r => String(r.ID_Murid) === String(p.idMurid));
  if (!m) throw err_('Murid tidak dijumpai.');
  if (!boleh_(u, 'urus_kelas_semua') && kelasSaya_(u).indexOf(String(m.ID_Kelas)) < 0) {
    throw err_('Hanya guru kelas atau PK HEM boleh menukar gambar murid ini.', 'FORBIDDEN');
  }
  const r = apiMuatNaikFoto_({ dataUrl: p.dataUrl, jenis: 'murid' }, u);
  const lama = String(m.Foto_ID || '');
  m.Foto_ID = r.fileId;
  updateRow_('tbl_murid', m);
  if (lama && lama !== r.fileId) {
    try {
      const f = DriveApp.getFileById(lama), folder = folderFoto_().getId();
      const it = f.getParents(); let dalam = false;
      while (it.hasNext()) if (it.next().getId() === folder) dalam = true;
      if (dalam) f.setTrashed(true);
    } catch (e) { /* abaikan */ }
  }
  return { idMurid: String(m.ID_Murid), fotoId: r.fileId };
}

/* ---------- Foto (disimpan SULIT dalam Drive, bukan pautan awam) ---------- */
function apiMuatNaikFoto_(p, u) {
  const m = String(p.dataUrl || '').match(/^data:(image\/(png|jpeg|jpg|webp));base64,(.+)$/);
  if (!m) throw err_('Format gambar tidak sah (PNG/JPEG/WEBP sahaja).');
  const bytes = Utilities.base64Decode(m[3]);
  if (bytes.length > APP.MAKS_FOTO_MB * 1024 * 1024) throw err_('Saiz gambar melebihi ' + APP.MAKS_FOTO_MB + 'MB.');
  const jenis = String(p.jenis || 'umum').replace(/[^a-z]/gi, '').slice(0, 20) || 'umum';
  const blob = Utilities.newBlob(bytes, m[1], `${jenis}_${genId_('F')}.${m[2] === 'jpeg' ? 'jpg' : m[2]}`);
  const file = folderFoto_().createFile(blob);
  audit_(u.id, 'MUAT_NAIK_FOTO', file.getId());
  return { fileId: file.getId() };
}

function apiLihatFoto_(p) {
  const id = idFail_(p.fileId);
  if (!id) throw err_('ID fail tiada.');
  const folder = folderFoto_(false);
  if (!folder) throw err_('Fail tidak dijumpai.');
  return { dataUrl: fotoDataUrl_(id, folder.getId()) };
}

// Muat beberapa foto sekali gus (frontend simpan dalam IndexedDB, jadi setiap foto dimuat sekali sahaja).
// Respons dihadkan kepada APP.MAKS_FOTO_PUKAL_MB; ID yang tidak dimuat dipulangkan dalam "baki" untuk diminta semula.
function apiLihatFotoPukal_(p) {
  const ids = (Array.isArray(p.fileIds) ? p.fileIds : []).map(String).filter(v => /^[\w-]{20,}$/.test(v)).slice(0, 25);
  const out = {};
  const folder = folderFoto_(false);
  if (!folder) { ids.forEach(id => { out[id] = null; }); return out; }
  const folderId = folder.getId(), had = APP.MAKS_FOTO_PUKAL_MB * 1024 * 1024;
  let jumlah = 0;
  const baki = [];
  ids.forEach(id => {
    if (jumlah >= had) { baki.push(id); return; }
    try { out[id] = fotoDataUrl_(id, folderId); jumlah += out[id].length; } catch (e) { out[id] = null; }
  });
  if (baki.length) out._baki = baki;
  return out;
}

function fotoDataUrl_(id, folderId) {
  let file;
  try { file = DriveApp.getFileById(id); } catch (e) { throw err_('Fail tidak dijumpai.'); }
  // Pastikan fail memang dalam folder sistem (elak akses fail Drive lain)
  let dalamFolder = false;
  const it = file.getParents();
  while (it.hasNext()) if (it.next().getId() === folderId) dalamFolder = true;
  if (!dalamFolder) throw err_('Akses ditolak.', 'FORBIDDEN');
  const blob = file.getBlob();
  return 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes());
}

// cipta=false (laluan baca tanpa lock): pulangkan null jika folder belum wujud, supaya folder hanya
// dicipta oleh muat naik (yang berjalan di bawah script lock) dan tidak tercipta dua kali
function folderFoto_(cipta) {
  const t = tetapan_();
  if (t.FOLDER_FOTO_ID) {
    try { return DriveApp.getFolderById(t.FOLDER_FOTO_ID); } catch (e) { /* cipta semula */ }
  }
  if (cipta === false) return null;
  const f = DriveApp.createFolder('HEM Pintar - Fail Foto (SULIT)');
  const r = read_('tbl_tetapan').rows.find(x => String(x.Kunci) === 'FOLDER_FOTO_ID');
  if (r) { r.Nilai = f.getId(); updateRow_('tbl_tetapan', r); }
  else insert_('tbl_tetapan', [{ Kunci: 'FOLDER_FOTO_ID', Nilai: f.getId(), Keterangan: 'Dijana automatik — jangan ubah' }]);
  return f;
}

/* ==================================================================
 *  4. PEMBANTU JADUAL (baca/tulis Sheets secara pukal)
 * ================================================================== */
function sheet_(name) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh) throw err_('Jadual ' + name + ' tiada. Jalankan setup().');
  return sh;
}

// Cache dalam satu permintaan: elak baca jadual yang sama berulang kali
const _MEMO = {};

function headers_(sh) {
  const k = 'h:' + sh.getName();
  if (!_MEMO[k]) _MEMO[k] = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  return _MEMO[k];
}

function read_(name) {
  if (_MEMO['r:' + name]) return _MEMO['r:' + name];
  const sh = sheet_(name);
  const values = sh.getDataRange().getValues();
  const headers = values.shift().map(String);
  const rows = values
    .map((r, i) => { const o = { _row: i + 2 }; headers.forEach((h, j) => { o[h] = r[j]; }); return o; })
    .filter(o => headers.some(h => o[h] !== '' && o[h] !== null));
  _MEMO['r:' + name] = { sh: sh, headers: headers, rows: rows };
  return _MEMO['r:' + name];
}

function cell_(v) {
  if (v === undefined || v === null) return '';
  if (typeof v === 'object' && !(v instanceof Date)) return JSON.stringify(v);
  return v;
}

function ensureRows_(sh, lastNeeded) {
  const max = sh.getMaxRows();
  if (lastNeeded > max) sh.insertRowsAfter(max, lastNeeded - max + 50);
}

function applyTextFormat_(sh, name, headers, startRow, n) {
  (SCHEMA[name].text || []).forEach(h => {
    const c = headers.indexOf(h) + 1;
    if (c > 0) sh.getRange(startRow, c, n, 1).setNumberFormat('@');
  });
}

function insert_(name, objs) {
  if (!objs || !objs.length) return;
  const sh = sheet_(name);
  const headers = headers_(sh);
  const start = sh.getLastRow() + 1;
  ensureRows_(sh, start + objs.length - 1);
  applyTextFormat_(sh, name, headers, start, objs.length);
  sh.getRange(start, 1, objs.length, headers.length)
    .setValues(objs.map(o => headers.map(h => cell_(o[h]))));
  delete _MEMO['r:' + name];
}

function updateRow_(name, obj) {
  const sh = sheet_(name);
  const headers = headers_(sh);
  applyTextFormat_(sh, name, headers, obj._row, 1);
  sh.getRange(obj._row, 1, 1, headers.length).setValues([headers.map(h => cell_(obj[h]))]);
  delete _MEMO['r:' + name];
}

// Kemas kini banyak baris dengan SATU tulisan (baris berdekatan, cth kehadiran sekelas)
function updateRows_(name, objs) {
  if (!objs || !objs.length) return;
  const sh = sheet_(name);
  const headers = headers_(sh);
  const nos = objs.map(o => o._row);
  const min = Math.min.apply(null, nos), max = Math.max.apply(null, nos);
  if (max - min + 1 <= 1000) {
    const rng = sh.getRange(min, 1, max - min + 1, headers.length);
    const data = rng.getValues();
    objs.forEach(o => { data[o._row - min] = headers.map(h => cell_(o[h])); });
    applyTextFormat_(sh, name, headers, min, max - min + 1);
    rng.setValues(data);
  } else {
    objs.forEach(o => updateRow_(name, o));
  }
  delete _MEMO['r:' + name];
}

function writeAll_(name, objs) {
  const sh = sheet_(name);
  const headers = headers_(sh);
  const last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, headers.length).clearContent();
  delete _MEMO['r:' + name];
  if (!objs.length) return;
  ensureRows_(sh, objs.length + 1);
  applyTextFormat_(sh, name, headers, 2, objs.length);
  sh.getRange(2, 1, objs.length, headers.length)
    .setValues(objs.map(o => headers.map(h => cell_(o[h]))));
  delete _MEMO['r:' + name];
}

function tetapan_() {
  const t = {};
  read_('tbl_tetapan').rows.forEach(r => { t[String(r.Kunci)] = String(r.Nilai); });
  return t;
}

// appendRow adalah atomik: selamat walaupun dipanggil serentak tanpa lock (cth log masuk, tindakan baca)
function audit_(idGuru, tindakan, butiran) {
  try {
    const sh = sheet_('tbl_log_audit');
    const o = { Timestamp: new Date(), ID_Guru: String(idGuru), Tindakan: tindakan, Butiran: String(butiran || '').slice(0, 500) };
    sh.appendRow(headers_(sh).map(h => cell_(o[h])));
  } catch (e) { /* log audit tidak boleh menggagalkan tindakan utama */ }
}

/* ==================================================================
 *  5. UTILITI
 * ================================================================== */
function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function err_(mesej, kod) {
  const e = new Error(mesej);
  e.kod = kod || '';
  return e;
}

function msg_(m) {
  try { SpreadsheetApp.getUi().alert(m); } catch (e) { Logger.log(m); }
}

const hex_ = b => b.map(x => ('0' + (x & 0xff).toString(16)).slice(-2)).join('');

// Format v2: "v2$<pusingan>$<salt rawak setiap pengguna>$<hex>" — SHA-256 diulang APP.HASH_ULANG kali
function hash_(pin, salt, ulang) {
  salt = salt || Utilities.getUuid().replace(/-/g, '');
  ulang = ulang || APP.HASH_ULANG;
  let b = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + '|' + String(pin), Utilities.Charset.UTF_8);
  for (let i = 1; i < ulang; i++) b = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, b);
  return 'v2$' + ulang + '$' + salt + '$' + hex_(b);
}

// Format lama (v1.0): satu pusingan dengan SALT yang dikongsi — hanya untuk menyemak & naik taraf
function hashLama_(pin) {
  const salt = PropertiesService.getScriptProperties().getProperty('SALT') || '';
  return hex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + '|' + String(pin), Utilities.Charset.UTF_8));
}

function semakPin_(simpan, pin) {
  const s = String(simpan || '');
  if (!s || !pin) return false;
  const v = s.split('$');
  if (v[0] === 'v2' && v.length === 4) return hash_(pin, v[2], Number(v[1]) || APP.HASH_ULANG) === s;
  return hashLama_(pin) === s;
}

// Kata laluan sementara rawak 10 aksara (tanpa aksara mengelirukan seperti 0/O, 1/l/I)
function pinSementara_() {
  const abjad = 'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnpqrstuvwxyz23456789';
  const b = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, Utilities.getUuid() + Utilities.getUuid());
  let s = '';
  for (let i = 0; i < 10; i++) s += abjad.charAt((b[i] & 0xff) % abjad.length);
  return s;
}

// ID fail Drive yang sah (atau '' untuk kosongkan) — elak teks sewenang-wenang disimpan dalam medan foto
function idFail_(v) {
  const s = String(v == null ? '' : v).trim();
  if (!s) return '';
  if (!/^[\w-]{20,100}$/.test(s)) throw err_('ID fail foto tidak sah.');
  return s;
}

function jsonSel_(v, label) {
  const s = JSON.stringify(v == null ? {} : v);
  if (s.length > APP.MAKS_JSON_SEL) throw err_(label + ' terlalu besar untuk disimpan.');
  return s;
}

function genId_(prefix) {
  return prefix + '-' + Utilities.formatDate(new Date(), APP.TZ, 'yyMMdd') + '-' +
    Utilities.getUuid().replace(/-/g, '').slice(0, 8).toUpperCase();
}

function nextId_(ids, prefix, pad) {
  let maks = 0;
  ids.forEach(id => {
    const m = String(id).match(new RegExp('^' + prefix + '(\\d+)$'));
    if (m) maks = Math.max(maks, Number(m[1]));
  });
  return prefix + String(maks + 1).padStart(pad, '0');
}

function today_() { return Utilities.formatDate(new Date(), APP.TZ, 'yyyy-MM-dd'); }

function fmtTarikh_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, APP.TZ, 'yyyy-MM-dd');
  return String(v || '').slice(0, 10);
}

function parseTarikh_(s) {
  const [y, m, d] = String(s).split('-').map(Number);
  return new Date(y, m - 1, d, 12);
}

function semakTarikh_(s) {
  const v = String(s || '');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(v)) throw err_('Format tarikh mesti YYYY-MM-DD.');
  return v;
}

function isTrue_(v) { return v === true || String(v).toUpperCase() === 'TRUE'; }

function parseJSON_(s, lalai) {
  if (!s) return lalai;
  try { return JSON.parse(s); } catch (e) { return lalai; }
}

function statusMerit_(merit, t) {
  const kritikal = Number(t.AMBANG_KRITIKAL) || 50;
  const amaran = Number(t.AMBANG_AMARAN) || 70;
  if (merit < kritikal) return 'Kritikal';
  if (merit <= amaran) return 'Amaran';
  return 'Baik';
}

/* ---- Pembersih data IDME ---- */
function kelasId_(tahun, nama) {
  const slug = v => String(v || '').toUpperCase().replace(/[^A-Z0-9]+/g, '-').replace(/^-|-$/g, '');
  return 'K-' + [slug(tahun), slug(nama)].filter(Boolean).join('-');
}

// "TAHUN 3" + "ARIF" → "3 ARIF"; nama yang sudah mengandungi tahun (cth "PRASEKOLAH ALTI", "3 ARIF") dikekalkan
function namaKelasPaparan_(tahun, nama) {
  const n = String(nama || '').trim(), t = String(tahun || '').trim().toUpperCase();
  if (!n) return t;
  const d = t.match(/\d+/);
  if (d) return new RegExp('^' + d[0] + '\\b').test(n) || n.toUpperCase().indexOf(t) >= 0 ? n : d[0] + ' ' + n;
  return t && n.toUpperCase().indexOf(t) < 0 ? t + ' ' + n : n;
}

function jantina_(s) {
  const v = String(s).toUpperCase();
  if (v.indexOf('LELAKI') === 0 || v === 'L') return 'L';
  if (v.indexOf('PEREMPUAN') === 0 || v === 'P') return 'P';
  return v.charAt(0);
}

// IDME: h/b/tttt → tttt-bb-hh
function tarikh_(s) {
  const v = String(s).trim();
  let m = v.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})$/);
  if (m) return `${m[3]}-${m[2].padStart(2, '0')}-${m[1].padStart(2, '0')}`;
  m = v.match(/^(\d{4})-(\d{2})-(\d{2})/);
  return m ? `${m[1]}-${m[2]}-${m[3]}` : v;
}

function num_(s) {
  const n = parseFloat(String(s).replace(/[^0-9.\-]/g, ''));
  return isNaN(n) ? 0 : n;
}

// Pulihkan sifar yang hilang & tukar ke format antarabangsa untuk WhatsApp (60...)
function normTel_(s) {
  const raw = String(s || '');
  if (/e\+/i.test(raw)) return '';
  const d = raw.replace(/\D/g, '');
  if (d.length < 8) return '';
  if (d.indexOf('60') === 0) return d;
  if (d.indexOf('0') === 0) return '6' + d;
  return '60' + d;
}

function bersih_(s) {
  return String(s || '').replace(/^["'\s,]+|["'\s,]+$/g, '').replace(/\s+/g, ' ').trim();
}

function noKP_(s) {
  const v = String(s || '').trim();
  if (!v || /e\+/i.test(v)) return '';
  const d = v.replace(/[\s-]/g, '');
  return /^\d{11}$/.test(d) ? '0' + d : d;   // pulihkan sifar hadapan yang hilang
}

function poskod_(s) {
  const d = String(s || '').replace(/\D/g, '');
  return d && d.length < 5 ? d.padStart(5, '0') : d;
}

function tahunSort_(t) {
  const s = String(t).toUpperCase();
  if (s.indexOf('PRA') === 0) return 0;
  const m = s.match(/\d+/);
  if (m) return Number(m[0]);
  // IDME menulis tahun dalam perkataan, cth "TAHUN LIMA"
  const kata = ['SATU', 'DUA', 'TIGA', 'EMPAT', 'LIMA', 'ENAM'].findIndex(k => new RegExp('\\b' + k + '\\b').test(s));
  return kata >= 0 ? kata + 1 : 99;
}


