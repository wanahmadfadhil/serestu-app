/**
 * e-Kokurikulum — Backend Google Apps Script
 * Berdasarkan PRD v1.1 Aplikasi Web Pengurusan Kokurikulum
 *
 * CARA PASANG
 * 1. Cipta Google Sheets baharu > Extensions > Apps Script. Tampal fail ini sebagai Code.gs.
 * 2. Emel Super Admin dan nama sekolah sudah ditetapkan dalam CONFIG di bawah.
 * 3. Jalankan fungsi setup() sekali dan benarkan akses (Sheets, Drive, Mail).
 * 4. Deploy > New deployment > Web app
 *      Execute as: Me   |   Who has access: Anyone
 * 5. Pilihan A (hos dalam Apps Script): tambah fail HTML bernama "index", tampal index.html. Buka URL /exec.
 *    Pilihan B (hos luar, cth Netlify): tampal URL /exec ke CONFIG.API_URL dalam index.html.
 * 6. Jalankan pasangTrigger() untuk backup mingguan, emel peringatan harian dan arkib log bulanan.
 * 7. Daftar akaun dalam aplikasi menggunakan emel Super Admin — akaun itu aktif serta-merta.
 */

const CONFIG = {
  SUPER_ADMIN_EMAIL: 'serestu@moe-dl.edu.my',   // Super Admin utama (ditanam dalam kod)
  NAMA_SEKOLAH: 'SK Sungai Tua Baharu',
  TOKEN_HARI: 7,
  MAKS_PERJUMPAAN: 15,
  MAKS_GAMBAR_BAIT: 6 * 1024 * 1024,
  KATA_LALUAN_LALAI: '12345678',   // untuk guru yang diimport dari sheet Data Guru (wajib ditukar semasa log masuk pertama)
  SHEET_DATA_GURU: 'Data Guru',    // nama tab sumber (huruf besar/kecil tidak penting)
  SHEET_DATA_MURID: 'DATA MURID',  // tab sumber maklumat murid: nama, No. KP, tahun, kelas
  SHEET_DATA_KOKO: 'DATA KOKO',    // tab sumber unit murid: beruniform, kelab, sukan, rumah sukan
  TAHUN_KOKO: [4, 5, 6]            // hanya murid tahun ini terlibat dalam kokurikulum
};

const P_COLS = Array.from({ length: CONFIG.MAKS_PERJUMPAAN }, (_, i) => 'P' + (i + 1));

const SCHEMA = {
  Users: ['ID_Guru', 'Nama', 'Emel', 'No_Telefon', 'Hash_Kata_Laluan', 'Salt', 'Peranan_Global', 'Status', 'Cubaan_Gagal', 'Kunci_Hingga', 'Tarikh_Daftar', 'Log_Masuk_Terakhir', 'No_KP', 'Tukar_Kata_Laluan', 'ID_Gambar', 'Kelas_Guru', 'Pengesah'],
  Token: ['Token', 'ID_Guru', 'Jenis', 'Tarikh_Cipta', 'Tarikh_Luput'],
  Kategori: ['ID_Kategori', 'Nama_Kategori', 'Wajib', 'Susunan', 'Status', 'Warna', 'Rekod_Perjumpaan', 'Tab_PAJSK'],
  Unit: ['ID_Unit', 'ID_Kategori', 'Nama_Unit', 'ID_Folder_Drive', 'Status', 'Nama_SPPB'],
  Guru_Unit: ['ID', 'Tahun', 'ID_Guru', 'ID_Unit', 'Peranan_Unit', 'Status'],
  Murid: ['ID_Murid', 'ID_APDM', 'Nama_Murid', 'Kelas', 'Jantina', 'Status', 'Tarikh_Kemaskini', 'Tarikh_Daftar'],
  Enrolmen: ['ID_Enrolmen', 'Tahun', 'ID_Murid', 'ID_Unit', 'ID_Kategori', 'Kelas', 'Status', 'Tarikh_Masuk'],
  Carta_Organisasi: ['ID', 'Tahun', 'ID_Unit', 'Jawatan', 'Jenis', 'ID_Rujukan', 'ID_Fail_Gambar', 'Susunan'],
  Jadual_Perjumpaan: ['ID_Perjumpaan', 'Tahun', 'Bil', 'Tarikh', 'Status', 'Catatan'],
  Kehadiran: ['ID_Rekod', 'Tahun', 'ID_Unit', 'ID_Murid'].concat(P_COLS).concat(['Dikemaskini_Oleh', 'Tarikh_Kemaskini']),
  Laporan: ['ID_Laporan', 'Tahun', 'ID_Unit', 'ID_Perjumpaan', 'Versi_Format', 'Data_JSON', 'Bil_Hadir', 'Status', 'ID_Guru_Pelapor', 'Tarikh_Hantar', 'ID_Pengesah', 'Tarikh_Sah', 'Catatan_Pengesah', 'Tarikh_Kemaskini'],
  Format_Laporan: ['Versi', 'Tarikh_Cipta', 'Medan_JSON', 'Aktif'],
  Pencapaian: ['ID', 'Tahun', 'ID_Murid', 'ID_Unit', 'Nama_Aktiviti', 'Peringkat', 'Pencapaian', 'Tarikh', 'ID_Fail_Bukti', 'Direkod_Oleh', 'Khidmat'],
  Tetapan: ['Kunci', 'Nilai'],
  Ekstra_PAJSK: ['ID', 'Tahun', 'ID_Murid', 'Jenis', 'Perkara', 'Nilai', 'Aktiviti', 'Tarikh', 'Direkod_Oleh', 'Tarikh_Rekod'],
  Pengumuman: ['ID', 'Tajuk', 'Isi', 'Tarikh_Mula', 'Tarikh_Tamat', 'Status', 'Dicipta_Oleh', 'Tarikh_Cipta'],
  Log: ['Masa', 'ID_Guru', 'Tindakan', 'Sheet', 'ID_Rekod', 'Nilai_Lama', 'Nilai_Baharu']
};

const PERINGKAT = ['Sekolah', 'Zon', 'Daerah', 'Negeri', 'Kebangsaan', 'Antarabangsa'];
const LAYANAN_B = ['Dikecualikan', 'Dikira hadir', 'Dikira tidak hadir'];
const JENIS_MEDAN = ['teks', 'panjang', 'tarikh', 'masa', 'pilihan', 'gambar'];

/* =========================================================================
 * PAJSK MOEIS: senarai pilihan SPPB (daripada TEMPLATE_PAJSK_MOEIS.xlsx, Sekolah Rendah)
 * ========================================================================= */
const SPPB = {
  tab: { Kelab: 'Kelab & Persatuan', Beruniform: 'Badan Beruniform', Sukan: 'Sukan & Permainan' },
  aktiviti: {
    Kelab: ["KELAB KITAR SEMULA (SR)", "KELAB DOKTOR MUDA (SR)", "KELAB KERJAYA (SR)", "KELAB PENCEGAHAN JENAYAH (SR)", "PERSATUAN AGAMA ISLAM (SR)", "PERSATUAN BAHASA ARAB (SR)", "PERSATUAN BAHASA INGGERIS (SR)", "PERSATUAN BAHASA MELAYU (SR)", "TVPSS (SR)", "KELAB RUKUN NEGARA (SR)", "KELAB SAINS, TEKNOLOGI, KEJURUTERAAN DAN MATEMATIK (STEM)/SOLAR (SR)", "KELAB SKIM PINJAMAN BUKU TEKS (SR)", "SENI MUZIK (SR)"],
    Beruniform: ["ANGKATAN PERTAHANAN AWAM MALAYSIA (SR)", "BULAN SABIT MERAH MALAYSIA (BSMM) (SR)", "KADET REMAJA SEKOLAH (SR)", "PERGERAKAN PUTERI ISLAM MALAYSIA (PPIM) (SR)", "PERSATUAN PANDU PUTERI MALAYSIA (SR)", "PERSEKUTUAN PENGAKAP MALAYSIA (SR)"],
    Sukan: ["BADMINTON (SR)", "BOLA BALING (SR)", "BOLA JARING (SR)", "BOLA SEPAK (SR)", "BOLA TAMPAR (SR)", "CATUR (SR)", "MEMANAH (SR)", "OLAHRAGA (SR)", "SEPAK TAKRAW (SR)", "SILAT (SR)", "SOFBOL (SR)"]
  },
  tambahan: ["ALAM SEKITAR (SR)", "KELAB DOKTOR MUDA (SR)", "KELAB KERJAYA (SR)", "KELAB KOMPUTER/ICT (SR)", "KELAB PENCEGAHAN JENAYAH (SR)", "PERSATUAN AGAMA ISLAM (SR)", "PERSATUAN BAHASA ARAB (SR)", "PERSATUAN BAHASA INGGERIS (SR)", "PERSATUAN BAHASA MELAYU (SR)", "TVPSS (SR)", "BULAN SABIT MERAH MALAYSIA (BSMM) (SR)", "KADET REMAJA SEKOLAH (SR)", "PERGERAKAN PUTERI ISLAM MALAYSIA (PPIM) (SR)", "PERSATUAN PANDU PUTERI MALAYSIA (SR)", "PERSEKUTUAN PENGAKAP MALAYSIA (SR)", "BADMINTON (SR)", "BOLA BALING (SR)", "BOLA JARING (SR)", "BOLA SEPAK (SR)", "BOLA TAMPAR (SR)", "HOKI (SR)"],
  jawatan: {
    Kelab: ["PENGERUSI", "NAIB PENGERUSI", "SETIAUSAHA", "BENDAHARI", "NAIB SETIAUSAHA", "NAIB BENDAHARI", "AHLI JAWATANKUASA", "AHLI AKTIF", "AHLI BIASA"],
    Beruniform: ["PENGERUSI", "NAIB PENGERUSI", "SETIAUSAHA", "BENDAHARI", "NAIB SETIAUSAHA", "NAIB BENDAHARI", "PEGAWAI WARAN KRS I & II", "STAF SARJAN KRS", "SARJAN KRS", "KOPERAL KRS", "LANS KOPERAL KRS", "SARJAN TKRS", "KOPERAL TKRS", "LANS KOPERAL TKRS", "WARAN II", "STAF SARJAN", "BINTARA KANAN", "FLIGHT SARJAN", "SARJAN", "BINTARA MUDA", "SARJAN UDARA", "KOPERAL", "LASKAR KANAN", "KOPERAL UDARA", "LANS KOPERAL", "LASKAR MUDA", "SUB INSPEKTOR", "SARJAN MEJAR", "PEGAWAI KADET", "KETUA BOMBA JUNIOR", "PENOLONG KETUA BOMBA JUNIOR", "KETUA PASUKAN", "PEGAWAI KANAN KADET", "PEGAWAI WARAN KADET", "SARJAN KADET", "KOPERAL KADET", "LANS KOPERAL KADET", "TUSPA (SARJAN)", "TUSPA (KOPERAL)", "TUSPA (LANS KOPERAL)", "KETUA TRUP", "PENOLONG KETUA TRUP", "KETUA PATROL", "PENOLONG KETUA PATROL", "PENGAKAP (KUATERMASTER)", "KETUA PEK", "PENOLONG KETUA PEK", "KETUA SEKAWAN", "PENOLONG KETUA SEKAWAN", "MEGA PUTERI", "PUTERI MEGASARI", "PUTERI BISTARIA", "PUTERI KEJORA", "PUTERI SAUJANA /KETUA PASUKAN", "PUTERI MESTIKA", "KETUA SEKSYEN UTAMA", "PENOLONG KETUA SEKSYEN", "KETUA PLATUN", "PENOLONG KETUA PLATUN", "SARJAN DEWASA / KETUA KADET", "LANS KOPERAL DEWASA / KOPERAL KADET", "BRASS BAND (DRUM MAJOR)", "BRASS BAND (KETUA SEKSYEN)", "BRASS BAND (PENOLONG KETUA SEKSYEN)", "BRASS BAND (KUATERMASTER)", "AHLI JAWATANKUASA", "AHLI AKTIF", "AHLI BIASA"],
    Sukan: ["PENGERUSI", "KAPTEN PASUKAN", "KETUA RUMAH", "NAIB PENGERUSI", "PENOLONG KETUA RUMAH", "PENOLONG KAPTEN", "SETIAUSAHA", "BENDAHARI", "NAIB SETIAUSAHA", "NAIB BENDAHARI", "AHLI JAWATANKUASA", "AHLI AKTIF", "AHLI BIASA", "BENGKUNG MERAH / TALI PINGGANG HITAM", "BENGKUNG HIJAU 3 / TALI PINGGANG MERAH 2", "BENGKUNG HIJAU 2 / TALI PINGGANG MERAH 1", "BENGKUNG HIJAU 1 / TALI PINGGANG BIRU 2", "CULA 3 / TALI PINGGANG BIRU 1", "CLASS 3 / CULA 2 / TALI PINGGANG HIJAU", "TALI PINGGANG KUNING"]
  },
  komitmen: ["MENUNJUKKAN KEPIMPINAN", "MENGURUS AKTIVITI", "MENYEDIA PERALATAN", "MEMBANTU GURU ATAU RAKAN", "MENGEMAS PERALATAN", "MEMBERSIH KAWASAN", "MENEPATI WAKTU", "MENUNJUKKAN MINAT", "MENUNJUKKAN KESUNGGUHAN", "MENGIKUT ARAHAN", "MENCUBA", "MEMBERI KERJASAMA", "MANA-MANA NILAI MURNI YANG BOLEH DIPERHATIKAN"],
  khidmat: ["MURID YANG DIDAFTARKAN SEBAGAI ATLET ATAU PESERTA PROGRAM/ PERTANDINGAN/ KARNIVAL/ KURSUS", "MELIBATKAN KEMAHIRAN KHUSUS - BERKAITAN ASPEK TEKNIKAL (PENGADIL, JURULATIH PASUKAN, DAN LAIN-LAIN YANG BERKAITAN ASPEK TEKNIKAL)", "PELIBATAN MURID YANG TERLIBAT DALAM AKTIVITI SEPERTI PERSEMBAHAN SELINGAN", "MEMBANTU DARI SEGI MENJAYAKAN AKTIVITI KELAB PERSATUAN SUKAN SEKOLAH SEPERTI MENGAMBIL BAHAGIAN DALAM PERBARISAN ATAU PERSEMBAHAN, KUMPULAN SORAK/PENYOKONG, DAN YANG BERKAITAN"],
  tahap: ["JOHAN", "NAIB JOHAN", "KETIGA", "KEEMPAT", "KELIMA"],
  peringkat: ["Antarabangsa", "Kebangsaan", "Negeri", "Bahagian (Sabah/Sarawak)", "Zon/Daerah", "Sekolah"],
  peringkatCapai: ["Kebangsaan", "Negeri", "Bahagian (Sabah/Sarawak)", "Zon/Daerah", "Sekolah"],
  // Tab Ekstra (helaian Senarai lajur O, P, Q, R dalam TEMPLATE_PAJSK_MOEIS.xlsx)
  anugerah: ["ANUGERAH REMAJA PERDANA", "PROGRAM KEPIMPINAN GENERASI MADANI - TIER 3", "PROGRAM KEPIMPINAN GENERASI MADANI - TIER 2", "ANUGERAH PENGAKAP RAJA / RAMBU PENGAKAP KANAK-KANAK", "ANUGERAH PANDU PUTERI RAJA / PANDU PUTERI TUNAS KETUA PESURUHJAYA", "ANUGERAH PUTERI AZ-ZAHRA / PUTERI BONGSU - PPIM", "ANUGERAH GRAND PRIOR / PERISAI PERKHIDMATAN KHAS 1000 JAM EMAS - SJAM", "ANUGERAH KHAS KOKURIKULUM DAN SUKAN - KEBANGSAAN", "ANUGERAH KHAS KOKURIKULUM DAN SUKAN - NEGERI", "ANUGERAH KHAS KOKURIKULUM DAN SUKAN - ZON/DAERAH/BAHAGIAN", "ANUGERAH KHAS KOKURIKULUM DAN SUKAN - SEKOLAH", "ANUGERAH-ANUGERAH LAIN/PENGHARGAAN KHAS SETARA – KEBANGSAAN", "ANUGERAH-ANUGERAH LAIN/PENGHARGAAN KHAS SETARA – NEGERI", "ANUGERAH-ANUGERAH LAIN/PENGHARGAAN KHAS SETARA - ZON/DAERAH/BAHAGIAN"],
  tahapAnugerah: ["PENERIMA", "EMAS", "PERAK", "GANGSA"],
  perkhidmatan: ["KETUA MURID/PENGERUSI|10", "TIMBALAN/PENOLONG KETUA MURID|8", "PENGAWAS/PENGAWAS PUSAT SUMBER/KETUA ASRAMA/KETUA PRS/KETUA BADAR/KUARTERMASTER/PENGERUSI BADAN PERKHIDMATAN SEKOLAH|7", "PENOLONG KETUA ASRAMA/PENOLONG KETUA BADAR/AHLI LEMBAGA PENGARAH KOPERASI SEKOLAH|6", "AJK UTAMA BADAN-BADAN PERKHIDMATAN ATAU PENGURUSAN/BENDAHARI/SETIAUSAHA/KETUA BIRO-ASRAMA/PRS/PENERIMA SIJIL DAN LENCANA SLAD/PROSTAR/PENGAWAS KOPERASI/KETUA KELAS/KETUA TINGKATAN|5", "AJK KECIL BIRO/KETUA BILIK/KETUA DORM/AJK BADAR/PENERIMA SIJIL SLAD/PENOLONG KETUA KELAS/PENOLONG KETUA BILIK|3", "PENGAWAS PUSAT SUMBER (TINGKATAN 6)|6"],
  nilam: ["5 BINTANG", "4 BINTANG", "3 BINTANG", "2 BINTANG", "1 BINTANG"],
  // Lajur F hingga L tab Ekstra: bilangan penyertaan (1 hingga 5) bagi setiap kategori
  khidmatMasyarakat: ["Kejiranan/Komuniti", "Alam Sekitar", "Bencana Alam", "Kerohanian", "Khidmat Jabatan/Agensi", "Kebudayaan/Kesenian/Integrasi", "Insaniah"]
};
const JENIS_EKSTRA = ['ANUGERAH', 'KHIDMAT', 'NILAM', 'PERKHIDMATAN'];
const PAJSK_LALAI = {
  PAJSK_KOMITMEN: JSON.stringify(['MENEPATI WAKTU', 'MENUNJUKKAN MINAT', 'MENUNJUKKAN KESUNGGUHAN', 'MEMBERI KERJASAMA']),
  PAJSK_KHIDMAT: '',
  PAJSK_JAWATAN: 'IKUT_HADIR',   // AHLI BIASA | AHLI AKTIF | IKUT_HADIR
  PAJSK_AKTIF_PERATUS: '75'
};

/** Tetapan papan pemuka awam (boleh diubah oleh Super Admin). */
const AWAM_LALAI = { AWAM_AKTIF: 'Ya', AWAM_PENCAPAIAN: 'Singkat', AWAM_GALERI: 'Ya', AWAM_SEMAKAN: 'Ya' };
const TZ = Session.getScriptTimeZone() || 'Asia/Kuala_Lumpur';

/* =========================================================================
 * PINTU MASUK
 * ========================================================================= */

function doGet(e) {
  if (e && e.parameter && e.parameter.ping) return json_({ ok: true, data: 'pong' });
  try {
    return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('e-Kokurikulum')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);   // halang aplikasi dibenamkan dalam laman lain (clickjacking)
  } catch (err) {
    return json_({ ok: true, data: 'API e-Kokurikulum aktif. Gunakan POST.' });
  }
}

function doPost(e) {
  let req = {};
  try { req = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'Permintaan tidak sah.' }); }
  return json_(handle_(req));
}

/** Dipanggil melalui google.script.run apabila index.html dihos dalam Apps Script. */
function apiCall(body) {
  let req = {};
  try { req = JSON.parse(body); } catch (e) { /* abaikan */ }
  return JSON.stringify(handle_(req));
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

const ROUTES = {
  publicInfo: { fn: a_publicInfo, pub: true },
  logo: { fn: a_logo, pub: true },
  awam: { fn: a_awam, pub: true },
  imgAwam: { fn: a_imgAwam, pub: true },
  semakMurid: { fn: a_semakMurid, pub: true },
  pengumuman: { fn: a_pengumuman },
  savePengumuman: { fn: a_savePengumuman, w: true },
  deletePengumuman: { fn: a_deletePengumuman, w: true },
  register: { fn: a_register, pub: true, w: true, ns: true },
  login: { fn: a_login, pub: true, w: true, ns: true, then: loginExtra_ },
  forgotRequest: { fn: a_forgotRequest, pub: true, w: true, ns: true },
  forgotReset: { fn: a_forgotReset, pub: true, w: true, ns: true },
  logout: { fn: a_logout, w: true, ns: true, pw: true },
  bootstrap: { fn: a_bootstrap },
  dashboard: { fn: a_dashboard },
  unit: { fn: a_unit },
  saveKehadiran: { fn: a_saveKehadiran, w: true },
  saveLaporan: { fn: a_saveLaporan, w: true },
  statusLaporan: { fn: a_statusLaporan, w: true },
  laporanMenunggu: { fn: a_laporanMenunggu },
  sahkanPukal: { fn: a_sahkanPukal, w: true },
  aiLaporan: { fn: a_aiLaporan },
  aiTetapan: { fn: a_aiTetapan },
  saveAiTetapan: { fn: a_saveAiTetapan, w: true, ns: true },
  uploadImage: { fn: a_uploadImage },
  image: { fn: a_image },
  saveCarta: { fn: a_saveCarta, w: true },
  savePencapaian: { fn: a_savePencapaian, w: true },
  deletePencapaian: { fn: a_deletePencapaian, w: true },
  pelibatanBorang: { fn: a_pelibatanBorang },
  simpanPelibatan: { fn: a_simpanPelibatan, w: true },
  simpanEkstra: { fn: a_simpanEkstra, w: true },
  padamEkstra: { fn: a_padamEkstra, w: true },
  muridPicker: { fn: a_muridPicker },
  enrol: { fn: a_enrol, w: true },
  unenrol: { fn: a_unenrol, w: true },
  guruUnit: { fn: a_guruUnit, w: true, ns: true },
  requestUnit: { fn: a_requestUnit, w: true, ns: true },
  changePassword: { fn: a_changePassword, w: true, ns: true, pw: true },
  saveProfil: { fn: a_saveProfil, w: true, ns: true },
  tukarUnit: { fn: a_tukarUnit, w: true, ns: true },
  kelasList: { fn: a_kelasList },
  bulkMurid: { fn: a_bulkMurid, w: true },
  kelasUnit: { fn: a_kelasUnit },
  simpanKelasUnit: { fn: a_simpanKelasUnit, w: true },
  tutupPratonton: { fn: a_tutupPratonton },
  importGuru: { fn: a_importGuru, w: true, ns: true },
  importKoko: { fn: a_importKoko, w: true },
  belumUnit: { fn: a_belumUnit },
  unitGuna: { fn: a_unitGuna },
  deleteUnit: { fn: a_deleteUnit, w: true },
  daftarMurid: { fn: a_daftarMurid, w: true },
  users: { fn: a_users },
  userAction: { fn: a_userAction, w: true, ns: true },
  saveKategori: { fn: a_saveKategori, w: true },
  saveUnit: { fn: a_saveUnit, w: true },
  murid: { fn: a_murid },
  saveMurid: { fn: a_saveMurid, w: true },
  importMurid: { fn: a_importMurid, w: true },
  saveJadual: { fn: a_saveJadual, w: true },
  saveFormat: { fn: a_saveFormat, w: true },
  saveTetapan: { fn: a_saveTetapan, w: true },
  log: { fn: a_log },
  pajsk: { fn: a_pajsk },
  pajskMoeis: { fn: a_pajskMoeis },
  pajskPeta: { fn: a_pajskPeta },
  savePajskPeta: { fn: a_savePajskPeta, w: true },
  tutupSesi: { fn: a_tutupSesi, w: true }
};

function handle_(req) {
  const t0 = Date.now();
  try {
    const r = ROUTES[req && req.action];
    if (!r) throw fail_('Tindakan tidak dikenali.');
    const ctx = r.pub ? null : auth_(req.token);
    if (ctx && ctx.user.Tukar_Kata_Laluan === 'Ya' && !r.pw) throw fail_('Sila tukar kata laluan lalai anda dahulu.', 'PWCHANGE');
    const run = function () {
      // Tindakan menulis: baca semula data selepas kunci diperoleh. Data yang dibaca sebelum kunci
      // (contohnya Token dan Users semasa auth_) mungkin sudah lapuk dan nombor barisnya boleh beralih.
      if (r.w) {
        Object.keys(T_).forEach(k => { delete T_[k]; });
        if (ctx) {
          const u = rows_('Users').find(x => x.ID_Guru === ctx.id);
          if (!u || u.Status !== 'Aktif') throw fail_('Akaun anda tidak aktif.', 'AUTH');
          ctx.user = u;
        }
      }
      return r.fn(req.data || {}, ctx);
    };
    let data = r.w ? withLock_(run, r.ns) : run();
    if (r.then) data = r.then(data);
    return { ok: true, data: data === undefined ? null : data, ms: Date.now() - t0 };
  } catch (e) {
    return { ok: false, error: String((e && e.message) || e), code: (e && e.code) || 'ERR' };
  }
}

function fail_(msg, code) {
  const e = new Error(msg);
  e.code = code || 'ERR';
  return e;
}

function withLock_(fn, keepCache) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) throw fail_('Sistem sedang sibuk. Cuba sekali lagi.', 'BUSY');
  try {
    const out = fn();
    SpreadsheetApp.flush();
    if (!keepCache) clearCaches_();
    return out;
  } finally {
    lock.releaseLock();
  }
}

/* =========================================================================
 * CACHE PELAYAN (CacheService)
 * Data yang jarang berubah disimpan dalam cache supaya setiap panggilan tidak
 * perlu membaca semua sheet. Cache dikosongkan selepas setiap penulisan data
 * dan apabila sheet disunting secara manual (onEdit).
 * ========================================================================= */

function cacheKeys_() {
  const t = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  return ['c_tet', 'c_common', 'c_awam', 'stats_' + t];
}
function clearCaches_() {
  try { CacheService.getScriptCache().removeAll(cacheKeys_()); } catch (e) { /* abaikan */ }
}
function cached_(key, ttl, fn) {
  const cache = CacheService.getScriptCache();
  const hit = cache.get(key);
  if (hit) { try { return JSON.parse(hit); } catch (e) { /* bina semula */ } }
  const v = fn();
  try { const s = JSON.stringify(v); if (s.length < 95000) cache.put(key, s, ttl); } catch (e) { /* terlalu besar */ }
  return v;
}
/** Simple trigger: suntingan manual dalam Sheets mengosongkan cache. */
function onEdit() { clearCaches_(); }

/* =========================================================================
 * LAPISAN DATA (Google Sheets)
 * Semua nilai dibaca dan ditulis sebagai teks. Satu getValues bagi setiap sheet per panggilan.
 * ========================================================================= */

let SS_ = null;
const T_ = {};

function ss_() {
  if (SS_) return SS_;
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  SS_ = id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
  if (!SS_) throw fail_('Pangkalan data belum disediakan. Jalankan setup() dahulu.');
  return SS_;
}

function sheet_(name) {
  return ss_().getSheetByName(name) || ensureSheet_(ss_(), name);
}

function ensureSheet_(ss, name) {
  const head = SCHEMA[name];
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, sh.getMaxRows(), head.length).setNumberFormat('@');
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground('#1D4F5E').setFontColor('#FFFFFF');
    sh.setFrozenRows(1);
    return sh;
  }
  // Migrasi: tambah lajur yang belum wujud
  const lastCol = Math.max(sh.getLastColumn(), 1);
  const cur = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  let used = cur.length;
  while (used > 0 && !cur[used - 1]) used--;
  const missing = head.filter(h => cur.indexOf(h) < 0);
  if (missing.length) {
    sh.getRange(1, used + 1, sh.getMaxRows(), missing.length).setNumberFormat('@');
    sh.getRange(1, used + 1, 1, missing.length).setValues([missing]).setFontWeight('bold').setBackground('#1D4F5E').setFontColor('#FFFFFF');
  }
  return sh;
}

function table_(name) {
  if (T_[name]) return T_[name];
  const sh = sheet_(name);
  const lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  let head = [], values = [];
  if (lastRow >= 1 && lastCol >= 1) {
    const all = sh.getRange(1, 1, lastRow, lastCol).getValues();
    head = all.shift().map(String);
    values = all;
  }
  const rows = values.map((r, i) => {
    const o = { _row: i + 2 };
    head.forEach((h, j) => { if (h) o[h] = norm_(r[j]); });
    return o;
  });
  T_[name] = { sh: sh, head: head, rows: rows };
  return T_[name];
}

function rows_(name) { return table_(name).rows; }

function norm_(v) {
  if (v === null || v === undefined) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    if (isNaN(v.getTime())) return '';
    const hms = Utilities.formatDate(v, TZ, 'HH:mm:ss');
    return Utilities.formatDate(v, TZ, hms === '00:00:00' ? 'yyyy-MM-dd' : 'yyyy-MM-dd HH:mm:ss');
  }
  return String(v);
}

/** Elak formula injection: teks bermula dengan = + - @ disimpan sebagai teks literal. */
function cell_(v) {
  if (v === null || v === undefined) return '';
  const s = String(v);
  if (/^[=+\-@]/.test(s) && !/^-?\d+(\.\d+)?$/.test(s)) return "'" + s;
  return s;
}

function toArr_(t, o) { return t.head.map(h => cell_(o[h])); }

function insert_(name, o) { return insertMany_(name, [o])[0]; }

/**
 * Tambah baris di hujung sheet.
 * - Jika jadual belum dibaca dalam panggilan ini, hanya baris tajuk dibaca (tidak membaca semua
 *   baris Log/Token setiap kali menulis satu baris).
 * - Julat baharu ditetapkan sebagai teks (@) supaya baris selepas baris 1000 tidak ditukar oleh
 *   Sheets kepada nombor/tarikh (contoh: 0123456789 -> 123456789).
 */
function insertMany_(name, objs) {
  if (!objs.length) return objs;
  const loaded = T_[name];
  const t = loaded || headOnly_(name);
  const start = t.sh.getLastRow() + 1;
  const need = start + objs.length - 1 - t.sh.getMaxRows();
  if (need > 0) t.sh.insertRowsAfter(t.sh.getMaxRows(), need);
  t.sh.getRange(start, 1, objs.length, t.head.length).setNumberFormat('@').setValues(objs.map(o => toArr_(t, o)));
  objs.forEach((o, i) => {
    o._row = start + i;
    t.head.forEach(h => { o[h] = (o[h] === undefined || o[h] === null) ? '' : String(o[h]); });
    if (loaded) loaded.rows.push(o);
  });
  return objs;
}

function headOnly_(name) {
  const sh = sheet_(name);
  const lastCol = sh.getLastColumn();
  const head = lastCol ? sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String) : [];
  return { sh: sh, head: head, rows: [] };
}

function update_(name, o, patch) {
  const t = table_(name);
  Object.keys(patch || {}).forEach(k => { o[k] = (patch[k] === null || patch[k] === undefined) ? '' : String(patch[k]); });
  t.sh.getRange(o._row, 1, 1, t.head.length).setValues([toArr_(t, o)]);
  return o;
}

/** Tulis semula semua baris sheet dalam satu panggilan (untuk kemas kini pukal). */
function writeAll_(name) {
  const t = table_(name);
  if (!t.rows.length) return;
  t.sh.getRange(2, 1, t.rows.length, t.head.length).setValues(t.rows.map(o => toArr_(t, o)));
}

/** Buang baris dengan menulis semula baki jadual (lebih pantas daripada deleteRow berulang). */
function removeMany_(name, objs) {
  if (!objs.length) return;
  const t = table_(name);
  const kill = {};
  objs.forEach(o => { kill[o._row] = 1; });
  const keep = t.rows.filter(o => !kill[o._row]);
  if (t.rows.length) t.sh.getRange(2, 1, t.rows.length, t.head.length).clearContent();
  if (keep.length) t.sh.getRange(2, 1, keep.length, t.head.length).setValues(keep.map(o => toArr_(t, o)));
  delete T_[name];
}

/* =========================================================================
 * UTILITI
 * ========================================================================= */

function now_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'); }
function today_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
function plusMin_(m) { return Utilities.formatDate(new Date(Date.now() + m * 60000), TZ, 'yyyy-MM-dd HH:mm:ss'); }
function uid_(p) { return (p || '') + Utilities.getUuid().replace(/-/g, '').slice(0, 12); }
function clean_(v, max) { return String(v === null || v === undefined ? '' : v).trim().slice(0, max || 200); }
function parse_(s, def) { try { return s ? JSON.parse(s) : def; } catch (e) { return def; } }
function isYmd_(s) { return /^\d{4}-\d{2}-\d{2}$/.test(String(s || '')); }

function addDays_(ymd, n) {
  const p = String(ymd).split('-').map(Number);
  return new Date(Date.UTC(p[0], p[1] - 1, p[2] + n)).toISOString().slice(0, 10);
}

function diffDays_(a, b) {
  const pa = String(a).split('-').map(Number), pb = String(b).split('-').map(Number);
  return Math.round((Date.UTC(pb[0], pb[1] - 1, pb[2]) - Date.UTC(pa[0], pa[1] - 1, pa[2])) / 86400000);
}

function sha_(s) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(s), Utilities.Charset.UTF_8)
    .map(b => ('0' + (b & 255).toString(16)).slice(-2)).join('');
}

function hashPw_(pw, salt) {
  let h = salt + '|' + pw;
  for (let i = 0; i < 100; i++) h = sha_(h + salt);
  return h;
}

function log_(who, tindakan, sheet, id, lama, baru) {
  insert_('Log', { Masa: now_(), ID_Guru: who || '', Tindakan: tindakan, Sheet: sheet || '', ID_Rekod: id || '', Nilai_Lama: clean_(lama, 500), Nilai_Baharu: clean_(baru, 500) });
}

function mapBy_(list, key) {
  const o = {};
  list.forEach(r => { o[r[key]] = r; });
  return o;
}

/* =========================================================================
 * TETAPAN, SESI, PERANAN
 * ========================================================================= */

function tetapan_() {
  const build = () => {
    const o = {};
    rows_('Tetapan').forEach(r => { if (r.Kunci) o[r.Kunci] = r.Nilai; });
    if (!o.NAMA_SEKOLAH || o.NAMA_SEKOLAH === 'Nama Sekolah Anda') o.NAMA_SEKOLAH = CONFIG.NAMA_SEKOLAH;
    return o;
  };
  if (T_.Tetapan) return build();
  return cached_('c_tet', 1800, build);
}

function setTetapan_(k, v) {
  const r = rows_('Tetapan').find(x => x.Kunci === k);
  if (r) update_('Tetapan', r, { Nilai: v });
  else insert_('Tetapan', { Kunci: k, Nilai: v });
}

function sesi_() { return tetapan_().SESI_AKTIF || Utilities.formatDate(new Date(), TZ, 'yyyy'); }

function publicTetapan_(t) {
  const keys = ['NAMA_SEKOLAH', 'ALAMAT_SEKOLAH', 'ID_FAIL_LOGO', 'TEMPOH_KUNCI_HARI', 'HARI_PERINGATAN', 'LAYANAN_B', 'NAMA_PENGESAH', 'JAWATAN_PENGESAH'];
  const o = {};
  keys.forEach(k => { o[k] = t[k] || ''; });
  Object.keys(AWAM_LALAI).forEach(k => { o[k] = t[k] || AWAM_LALAI[k]; });
  o.AI_AKTIF = t.AI_AKTIF === 'Tidak' ? 'Tidak' : 'Ya';
  o.EMEL_LAPORAN = t.EMEL_LAPORAN === 'Tidak' ? 'Tidak' : 'Ya';
  o.AI_SEDIA = aiConfig_().key ? 'Ya' : 'Tidak';
  return o;
}

function superEmail_() {
  return String(CONFIG.SUPER_ADMIN_EMAIL).toLowerCase().trim();
}

function auth_(token) {
  if (!token) throw fail_('Sila log masuk.', 'AUTH');
  const cache = CacheService.getScriptCache();
  let gid = cache.get('tk_' + token);
  if (!gid) {
    const t = rows_('Token').find(r => r.Token === token && r.Jenis === 'Sesi');
    if (!t || t.Tarikh_Luput < now_()) throw fail_('Sesi anda telah tamat. Sila log masuk semula.', 'AUTH');
    gid = t.ID_Guru;
    // Cache tidak melebihi baki hayat token (maksimum 6 jam)
    let baki = 21600;
    try { baki = Math.floor((Utilities.parseDate(t.Tarikh_Luput, TZ, 'yyyy-MM-dd HH:mm:ss').getTime() - Date.now()) / 1000); } catch (e) { baki = 300; }
    if (baki > 0) cache.put('tk_' + token, gid, Math.min(21600, baki));
  }
  const u = rows_('Users').find(r => r.ID_Guru === gid);
  if (!u || u.Status !== 'Aktif') throw fail_('Akaun anda tidak aktif.', 'AUTH');
  const sup = u.Peranan_Global === 'SuperAdmin';
  return { id: gid, user: u, super: sup, pengesah: sup || !!u.Pengesah, sesi: sesi_(), token: token };
}

function killTokens_(list) {
  if (!list.length) return;
  const cache = CacheService.getScriptCache();
  list.forEach(t => { if (t.Jenis === 'Sesi') cache.remove('tk_' + t.Token); });
  removeMany_('Token', list);
}

function unitRole_(ctx, idUnit) {
  if (ctx.super) return 'Super';
  const g = rows_('Guru_Unit').find(r => r.ID_Guru === ctx.id && r.ID_Unit === idUnit && r.Tahun === ctx.sesi && r.Status === 'Aktif');
  return g ? g.Peranan_Unit : null;
}

function needUnit_(ctx, idUnit, admin) {
  const r = unitRole_(ctx, idUnit);
  if (!r || (admin && r === 'Penasihat')) throw fail_('Anda tiada akses untuk tindakan ini dalam unit ini.', 'FORBIDDEN');
  return r;
}

/** Super Admin atau guru yang dilantik sebagai pengesah laporan (contoh PK Kokurikulum, Setiausaha Kokurikulum). */
function needPengesah_(ctx) {
  if (!ctx.pengesah) throw fail_('Hanya pengesah laporan boleh melakukan tindakan ini.', 'FORBIDDEN');
}

function needSuper_(ctx) {
  if (!ctx.super) throw fail_('Hanya Super Admin boleh melakukan tindakan ini.', 'FORBIDDEN');
}

function myUnitIds_(ctx) {
  const o = {};
  rows_('Guru_Unit').forEach(g => { if (g.ID_Guru === ctx.id && g.Tahun === ctx.sesi && g.Status === 'Aktif') o[g.ID_Unit] = g.Peranan_Unit; });
  return o;
}

function jadual_(sesi) {
  return rows_('Jadual_Perjumpaan').filter(j => j.Tahun === sesi).sort((a, b) => Number(a.Bil) - Number(b.Bil));
}

function jadualOut_(sesi) {
  return jadual_(sesi).map(j => ({ id: j.ID_Perjumpaan, bil: Number(j.Bil), tarikh: j.Tarikh, status: j.Status || 'Dirancang', catatan: j.Catatan }));
}

function pct_(row, bils, mode) {
  let h = 0, n = 0;
  bils.forEach(b => {
    const v = row['P' + b];
    if (!v) return;
    if (v === 'H') { h++; n++; }
    else if (v === 'T') n++;
    else if (v === 'B') {
      if (mode === 'Dikira hadir') { h++; n++; }
      else if (mode === 'Dikira tidak hadir') n++;
    }
  });
  return { h: h, n: n };
}

/* =========================================================================
 * AWAM & PENGESAHAN
 * ========================================================================= */

function a_publicInfo() {
  const t = tetapan_();
  const c = common_(sesi_());
  const kat = c.kategori;
  const unit = c.unit.filter(u => u.status !== 'Tidak Aktif').map(u => ({ id: u.id, idKategori: u.idKategori, nama: u.nama }));
  return { namaSekolah: t.NAMA_SEKOLAH || CONFIG.NAMA_SEKOLAH, sesi: sesi_(), idLogo: t.ID_FAIL_LOGO || '', kategori: kat, unit: unit, awam: (t.AWAM_AKTIF || 'Ya') !== 'Tidak' };
}

function a_logo() {
  const id = tetapan_().ID_FAIL_LOGO;
  if (!id) return { id: '', data: '' };
  const cache = CacheService.getScriptCache();
  const hit = cache.get('logo_' + id);
  if (hit) return { id: id, data: hit };
  try {
    const blob = DriveApp.getFileById(id).getBlob();
    const data = 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes());
    if (data.length < 95000) cache.put('logo_' + id, data, 21600);
    return { id: id, data: data };
  } catch (e) {
    return { id: '', data: '' };
  }
}

function katOut_(k) {
  return { id: k.ID_Kategori, nama: k.Nama_Kategori, wajib: k.Wajib !== 'Tidak', susunan: Number(k.Susunan || 0), status: k.Status || 'Aktif', warna: k.Warna || '', perjumpaan: k.Rekod_Perjumpaan !== 'Tidak' };
}

/** Tahun murid daripada nama kelas ("4 UIAM" -> 4). */
function tahunKelas_(kelas) {
  const m = String(kelas || '').match(/^\s*(\d)/);
  return m ? Number(m[1]) : 0;
}
/** Murid aktif yang terlibat dalam kokurikulum (Tahun 4-6). Kelas tanpa nombor tahun tetap dikira. */
function layak_(m) {
  if (m.Status !== 'Aktif') return false;
  const t = tahunKelas_(m.Kelas);
  return !t || CONFIG.TAHUN_KOKO.indexOf(t) >= 0;
}
/** Ahli wajib direkod bagi perjumpaan pada tarikh ini? Tarikh_Masuk kosong = sejak awal sesi. */
function sudahMasuk_(masuk, tarikh) {
  return !masuk || !tarikh || String(masuk) <= String(tarikh);
}
/** Murid didaftarkan ke sistem dalam 30 hari terakhir (murid pindah masuk). */
function muridBaharu_(m) {
  return !!m.Tarikh_Daftar && addDays_(String(m.Tarikh_Daftar).slice(0, 10), 30) >= today_();
}
/** Halang kehadiran/laporan bagi unit dalam kategori yang tidak merekod perjumpaan (contoh Rumah Sukan). */
function needRekodPerjumpaan_(idUnit) {
  const u = rows_('Unit').find(x => x.ID_Unit === idUnit);
  if (!u) throw fail_('Unit tidak dijumpai.');
  if (katPerjumpaan_()[u.ID_Kategori] === false) throw fail_('Kategori unit ini tidak merekod perjumpaan, kehadiran atau laporan.');
}
/** Kategori yang merekod perjumpaan (Rumah Sukan tidak). */
function katPerjumpaan_() {
  const o = {};
  rows_('Kategori').forEach(k => { o[k.ID_Kategori] = k.Rekod_Perjumpaan !== 'Tidak'; });
  return o;
}

function a_register(d) {
  const nama = clean_(d.nama, 120);
  const emel = String(d.emel || '').toLowerCase().trim();
  const tel = clean_(d.telefon, 20);
  const pw = String(d.kataLaluan || '');
  if (!nama) throw fail_('Sila isi nama penuh.');
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(emel)) throw fail_('Format emel tidak sah.');
  if (pw.length < 8) throw fail_('Kata laluan mesti sekurang-kurangnya 8 aksara.');
  if (rows_('Users').some(u => u.Emel === emel)) throw fail_('Emel ini sudah didaftarkan. Log masuk atau gunakan Lupa kata laluan.');
  const isSuper = emel === superEmail_();
  const salt = Utilities.getUuid();
  const u = insert_('Users', {
    ID_Guru: uid_('G'), Nama: nama, Emel: emel, No_Telefon: tel, Hash_Kata_Laluan: hashPw_(pw, salt), Salt: salt,
    Peranan_Global: isSuper ? 'SuperAdmin' : 'Guru', Status: isSuper ? 'Aktif' : 'Menunggu',
    Cubaan_Gagal: 0, Kunci_Hingga: '', Tarikh_Daftar: now_(), Log_Masuk_Terakhir: ''
  });
  const sesi = sesi_();
  const aktif = mapBy_(rows_('Unit').filter(x => x.Status !== 'Tidak Aktif'), 'ID_Unit');
  const units = (d.units || []).filter((id, i, a) => aktif[id] && a.indexOf(id) === i).slice(0, 20);
  insertMany_('Guru_Unit', units.map(id => ({ ID: uid_('GU'), Tahun: sesi, ID_Guru: u.ID_Guru, ID_Unit: id, Peranan_Unit: 'Penasihat', Status: isSuper ? 'Aktif' : 'Menunggu' })));
  log_(u.ID_Guru, 'Daftar akaun', 'Users', u.ID_Guru, '', emel);
  return { aktif: isSuper };
}

function a_login(d) {
  const emel = String(d.emel || '').toLowerCase().trim();
  const u = rows_('Users').find(r => r.Emel === emel);
  if (!u) throw fail_('Emel atau kata laluan tidak betul.');
  if (u.Kunci_Hingga && u.Kunci_Hingga > now_()) throw fail_('Akaun dikunci sementara. Cuba lagi selepas jam ' + u.Kunci_Hingga.slice(11, 16) + '.');
  if (hashPw_(String(d.kataLaluan || ''), u.Salt) !== u.Hash_Kata_Laluan) {
    const n = Number(u.Cubaan_Gagal || 0) + 1;
    update_('Users', u, n >= 5 ? { Cubaan_Gagal: 0, Kunci_Hingga: plusMin_(15) } : { Cubaan_Gagal: n });
    throw fail_(n >= 5 ? 'Terlalu banyak cubaan. Akaun dikunci selama 15 minit.' : 'Emel atau kata laluan tidak betul.');
  }
  if (emel === superEmail_() && (u.Peranan_Global !== 'SuperAdmin' || u.Status !== 'Aktif')) {
    update_('Users', u, { Peranan_Global: 'SuperAdmin', Status: 'Aktif' });
  }
  if (u.Status === 'Menunggu') throw fail_('Akaun anda sedang menunggu kelulusan Super Admin.');
  if (u.Status !== 'Aktif') throw fail_('Akaun anda tidak aktif. Hubungi Super Admin.');
  const tok = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  const now = now_();
  const expired = rows_('Token').filter(t => t.Tarikh_Luput < now);
  if (expired.length) killTokens_(expired);
  insert_('Token', { Token: tok, ID_Guru: u.ID_Guru, Jenis: 'Sesi', Tarikh_Cipta: now, Tarikh_Luput: Utilities.formatDate(new Date(Date.now() + CONFIG.TOKEN_HARI * 86400000), TZ, 'yyyy-MM-dd HH:mm:ss') });
  update_('Users', u, { Cubaan_Gagal: 0, Kunci_Hingga: '', Log_Masuk_Terakhir: now });
  return { token: tok, tukar: u.Tukar_Kata_Laluan === 'Ya' };
}

/** Selepas log masuk: sertakan bootstrap dan papan pemuka supaya tiada panggilan tambahan. */
function loginExtra_(data) {
  if (data.tukar) return data;
  try {
    const ctx = auth_(data.token);
    data.boot = a_bootstrap({}, ctx);
    data.dash = a_dashboard({}, ctx);
  } catch (e) { /* log masuk tetap berjaya; aplikasi akan memuat data seperti biasa */ }
  return data;
}

function a_forgotRequest(d) {
  const emel = String(d.emel || '').toLowerCase().trim();
  const u = rows_('Users').find(r => r.Emel === emel && r.Status === 'Aktif');
  if (u) {
    const old = rows_('Token').filter(t => t.Jenis === 'Reset' && t.ID_Guru === u.ID_Guru);
    // Elak emel berulang (spam / kuota MailApp habis): satu kod setiap 2 minit
    const had = Utilities.formatDate(new Date(Date.now() - 2 * 60000), TZ, 'yyyy-MM-dd HH:mm:ss');
    if (old.some(t => t.Tarikh_Cipta > had)) return { ok: true };
    const kod = String(100000 + parseInt(Utilities.getUuid().replace(/-/g, '').slice(0, 8), 16) % 900000);
    if (old.length) removeMany_('Token', old);
    insert_('Token', { Token: sha_(kod + u.Salt), ID_Guru: u.ID_Guru, Jenis: 'Reset', Tarikh_Cipta: now_(), Tarikh_Luput: plusMin_(15) });
    MailApp.sendEmail({
      to: emel,
      subject: 'Kod set semula kata laluan e-Kokurikulum',
      body: 'Kod anda: ' + kod + '\nKod ini sah selama 15 minit.\n\nJika anda tidak meminta kod ini, abaikan emel ini.'
    });
  }
  return { ok: true };
}

function a_forgotReset(d) {
  const emel = String(d.emel || '').toLowerCase().trim();
  const kod = String(d.kod || '').trim();
  const pw = String(d.kataLaluan || '');
  if (pw.length < 8) throw fail_('Kata laluan mesti sekurang-kurangnya 8 aksara.');
  if (pw === CONFIG.KATA_LALUAN_LALAI) throw fail_('Kata laluan lalai tidak boleh digunakan. Pilih kata laluan lain.');
  const u = rows_('Users').find(r => r.Emel === emel && r.Status === 'Aktif');
  if (!u) throw fail_('Kod tidak sah atau telah tamat tempoh.');
  const resets = rows_('Token').filter(x => x.Jenis === 'Reset' && x.ID_Guru === u.ID_Guru);
  const t = resets.find(x => x.Token === sha_(kod + u.Salt));
  if (!t || t.Tarikh_Luput < now_()) {
    const n = Number(u.Cubaan_Gagal || 0) + 1;
    if (n >= 5) { update_('Users', u, { Cubaan_Gagal: 0 }); killTokens_(resets); throw fail_('Terlalu banyak cubaan. Minta kod baharu.'); }
    update_('Users', u, { Cubaan_Gagal: n });
    throw fail_('Kod tidak sah atau telah tamat tempoh.');
  }
  const salt = Utilities.getUuid();
  update_('Users', u, { Salt: salt, Hash_Kata_Laluan: hashPw_(pw, salt), Cubaan_Gagal: 0, Kunci_Hingga: '', Tukar_Kata_Laluan: '' });
  killTokens_(rows_('Token').filter(x => x.ID_Guru === u.ID_Guru));
  log_(u.ID_Guru, 'Set semula kata laluan', 'Users', u.ID_Guru, '', '');
  return { ok: true };
}

function a_logout(d, ctx) {
  killTokens_(rows_('Token').filter(t => t.Token === ctx.token));
  return { ok: true };
}

function a_changePassword(d, ctx) {
  const u = ctx.user;
  if (hashPw_(String(d.lama || ''), u.Salt) !== u.Hash_Kata_Laluan) throw fail_('Kata laluan semasa tidak betul.');
  const pw = String(d.baru || '');
  if (pw.length < 8) throw fail_('Kata laluan baharu mesti sekurang-kurangnya 8 aksara.');
  if (pw === CONFIG.KATA_LALUAN_LALAI) throw fail_('Kata laluan lalai tidak boleh digunakan. Pilih kata laluan lain.');
  if (pw === String(d.lama || '')) throw fail_('Kata laluan baharu mesti berbeza daripada kata laluan semasa.');
  const salt = Utilities.getUuid();
  update_('Users', u, { Salt: salt, Hash_Kata_Laluan: hashPw_(pw, salt), Tukar_Kata_Laluan: '' });
  // Tamatkan sesi lain (peranti lain) selepas kata laluan ditukar; sesi semasa dikekalkan
  killTokens_(rows_('Token').filter(t => t.ID_Guru === ctx.id && t.Token !== ctx.token));
  log_(ctx.id, 'Tukar kata laluan', 'Users', ctx.id, '', '');
  return { ok: true };
}

/** Guru mengemas kini profil sendiri: nama, telefon dan gambar. */
function a_saveProfil(d, ctx) {
  ensureSheet_(ss_(), 'Users');
  delete T_.Users;
  const u = rows_('Users').find(x => x.ID_Guru === ctx.id);
  if (!u) throw fail_('Akaun tidak dijumpai.');
  const patch = {};
  if (d.nama !== undefined) {
    const nama = clean_(d.nama, 120);
    if (nama.length < 3) throw fail_('Sila isi nama penuh.');
    patch.Nama = nama;
  }
  if (d.telefon !== undefined) {
    const tel = clean_(d.telefon, 20);
    if (tel && !/^[0-9+\-\s]{9,20}$/.test(tel)) throw fail_('Format nombor telefon tidak sah.');
    patch.No_Telefon = tel;
  }
  if (d.kelas !== undefined) {
    const k = clean_(d.kelas, 40);
    if (k && senaraiKelas_().indexOf(k) < 0) throw fail_('Kelas tidak dijumpai dalam senarai murid.');
    patch.Kelas_Guru = k;
  }
  if (d.idGambar !== undefined) {
    const id = clean_(d.idGambar, 80);
    if (id) {
      let f;
      try { f = DriveApp.getFileById(id); } catch (e) { throw fail_('Gambar tidak dijumpai.'); }
      if (!underRoot_(f)) throw fail_('Gambar tidak sah.', 'FORBIDDEN');
    }
    const lama = u.ID_Gambar;
    if (lama && lama !== id) { try { DriveApp.getFileById(lama).setTrashed(true); } catch (e) { /* abaikan */ } }
    patch.ID_Gambar = id;
  }
  update_('Users', u, patch);
  log_(ctx.id, 'Kemas kini profil', 'Users', ctx.id, '', Object.keys(patch).join(', '));
  return { user: { id: u.ID_Guru, nama: u.Nama, emel: u.Emel, telefon: u.No_Telefon, peranan: u.Peranan_Global, super: ctx.super, gambar: u.ID_Gambar || '', kelas: u.Kelas_Guru || '', pengesah: ctx.super ? (u.Pengesah || 'Super Admin') : (u.Pengesah || '') } };
}

/* =========================================================================
 * BOOTSTRAP & DASHBOARD
 * ========================================================================= */

function common_(sesi) {
  return cached_('c_common', 1800, () => {
    const kat = rows_('Kategori').filter(k => k.Status !== 'Tidak Aktif').map(katOut_).sort((a, b) => a.susunan - b.susunan);
    const count = {};
    rows_('Enrolmen').forEach(e => { if (e.Tahun === sesi && e.Status === 'Aktif') count[e.ID_Unit] = (count[e.ID_Unit] || 0) + 1; });
    const unit = rows_('Unit').map(u => ({ id: u.ID_Unit, idKategori: u.ID_Kategori, nama: u.Nama_Unit, status: u.Status || 'Aktif', ahli: count[u.ID_Unit] || 0 }))
      .sort((a, b) => a.nama.localeCompare(b.nama));
    const format = rows_('Format_Laporan').map(f => ({ versi: f.Versi, aktif: f.Aktif === 'Ya', medan: parse_(f.Medan_JSON, []) }));
    return { sesi: sesi, kategori: kat, unit: unit, jadual: jadualOut_(sesi), format: format };
  });
}

function a_bootstrap(d, ctx) {
  const sesi = ctx.sesi;
  let c = common_(sesi);
  if (c.sesi !== sesi) { CacheService.getScriptCache().remove('c_common'); c = common_(sesi); }
  const kat = c.kategori, unit = c.unit;
  const gu = rows_('Guru_Unit').filter(g => g.Tahun === sesi);
  const saya = gu.filter(g => g.ID_Guru === ctx.id && g.Status !== 'Tidak Aktif')
    .map(g => ({ id: g.ID, idUnit: g.ID_Unit, peranan: g.Peranan_Unit, status: g.Status }));
  const adminOf = {};
  saya.forEach(s => { if (s.status === 'Aktif' && s.peranan === 'Admin') adminOf[s.idUnit] = 1; });
  const menunggu = {
    users: ctx.super ? rows_('Users').filter(u => u.Status === 'Menunggu').length : 0,
    guruUnit: (function () {
      const aktif = {};
      rows_('Users').forEach(x => { if (x.Status === 'Aktif') aktif[x.ID_Guru] = 1; });
      return gu.filter(g => g.Status === 'Menunggu' && aktif[g.ID_Guru] && (ctx.super || adminOf[g.ID_Unit])).length;
    })(),
    laporan: ctx.pengesah ? rows_('Laporan').filter(l => l.Tahun === sesi && l.Status === 'Dihantar').length : 0
  };
  const format = c.format;
  const u = ctx.user;
  return {
    user: { id: u.ID_Guru, nama: u.Nama, emel: u.Emel, telefon: u.No_Telefon, peranan: u.Peranan_Global, super: ctx.super, gambar: u.ID_Gambar || '', kelas: u.Kelas_Guru || '', pengesah: ctx.super ? (u.Pengesah || 'Super Admin') : (u.Pengesah || '') },
    sesi: sesi, today: today_(), tetapan: publicTetapan_(tetapan_()),
    kategori: kat, unit: unit, saya: saya, jadual: c.jadual, format: format, menunggu: menunggu,
    dash: d && d.dash ? a_dashboard({}, ctx) : undefined
  };
}

function computeStats_() {
  const cache = CacheService.getScriptCache();
  const key = 'stats_' + today_();
  const hit = cache.get(key);
  if (hit) return JSON.parse(hit);
  const sesi = sesi_(), mode = tetapan_().LAYANAN_B || 'Dikecualikan', today = today_();
  const jad = jadual_(sesi);
  const held = jad.filter(j => j.Status !== 'Dibatalkan' && j.Tarikh && j.Tarikh <= today);
  const heldBils = held.map(j => Number(j.Bil));
  const enrol = rows_('Enrolmen').filter(e => e.Tahun === sesi && e.Status === 'Aktif');
  const byUnit = {}, masukOf = {};
  enrol.forEach(e => { (byUnit[e.ID_Unit] = byUnit[e.ID_Unit] || []).push(e.ID_Murid); masukOf[e.ID_Unit + '|' + e.ID_Murid] = e.Tarikh_Masuk || ''; });
  const kh = {};
  rows_('Kehadiran').forEach(k => { if (k.Tahun === sesi) kh[k.ID_Unit + '|' + k.ID_Murid] = k; });
  const lap = {};
  rows_('Laporan').forEach(l => { if (l.Tahun === sesi) lap[l.ID_Unit + '|' + l.ID_Perjumpaan] = l.Status; });
  let H = 0, N = 0;
  const katP = katPerjumpaan_();
  const units = rows_('Unit').filter(u => u.Status !== 'Tidak Aktif').map(u => {
    const ahli = byUnit[u.ID_Unit] || [];
    if (katP[u.ID_Kategori] === false) {
      return { id: u.ID_Unit, nama: u.Nama_Unit, idKategori: u.ID_Kategori, ahli: ahli.length, peratus: null, status: ahli.length ? 'tiada' : 'kosong', tertunggak: [], menungguSah: 0 };
    }
    let h = 0, n = 0;
    const tertunggak = [];
    held.forEach(j => {
      let marked = 0;
      const wajib = ahli.filter(m => sudahMasuk_(masukOf[u.ID_Unit + '|' + m], j.Tarikh));
      wajib.forEach(m => { const r = kh[u.ID_Unit + '|' + m]; if (r && r['P' + j.Bil]) marked++; });
      const ls = lap[u.ID_Unit + '|' + j.ID_Perjumpaan] || '';
      const kOk = !wajib.length || marked >= wajib.length;
      const lOk = ls === 'Dihantar' || ls === 'Disahkan';
      if (ahli.length && (!kOk || !lOk)) tertunggak.push({ bil: Number(j.Bil), tarikh: j.Tarikh, id: j.ID_Perjumpaan, kehadiran: kOk, laporan: ls });
    });
    ahli.forEach(m => {
      const r = kh[u.ID_Unit + '|' + m];
      if (r) { const p = pct_(r, heldBils, mode); h += p.h; n += p.n; }
    });
    H += h; N += n;
    return {
      id: u.ID_Unit, nama: u.Nama_Unit, idKategori: u.ID_Kategori, ahli: ahli.length,
      peratus: n ? Math.round(h / n * 100) : null,
      status: !ahli.length ? 'kosong' : (tertunggak.length ? 'merah' : 'hijau'), tertunggak: tertunggak,
      menungguSah: jad.filter(j => lap[u.ID_Unit + '|' + j.ID_Perjumpaan] === 'Dihantar').length
    };
  });
  const murid = rows_('Murid').filter(layak_);
  const aktif = {};
  murid.forEach(m => { aktif[m.ID_Murid] = 1; });
  const kategori = rows_('Kategori').filter(k => k.Status !== 'Tidak Aktif').map(k => {
    const set = {};
    enrol.forEach(e => { if (e.ID_Kategori === k.ID_Kategori && aktif[e.ID_Murid]) set[e.ID_Murid] = 1; });
    const c = Object.keys(set).length;
    const o = katOut_(k);
    o.berdaftar = c;
    o.belum = Math.max(0, murid.length - c);
    o.baharu = o.wajib ? murid.filter(m => !set[m.ID_Murid] && muridBaharu_(m)).length : 0;
    return o;
  }).sort((a, b) => a.susunan - b.susunan);
  // Murid (bukan pendaftaran) yang belum lengkap berunit dalam kategori wajib
  const wajibK = kategori.filter(k => k.wajib).map(k => k.id);
  const adaK = {};
  enrol.forEach(e => { (adaK[e.ID_Murid] = adaK[e.ID_Murid] || {})[e.ID_Kategori] = 1; });
  let bm = 0, bb = 0;
  const belumKelas = {};
  murid.forEach(m => { const a = adaK[m.ID_Murid] || {}; if (wajibK.some(k => !a[k])) { bm++; if (muridBaharu_(m)) bb++; belumKelas[m.Kelas] = (belumKelas[m.Kelas] || 0) + 1; } });
  const next = jad.find(j => j.Status !== 'Dibatalkan' && j.Tarikh && j.Tarikh >= today);
  const out = {
    sesi: sesi, today: today, jumlahMurid: murid.length, kategori: kategori, unit: units, belumUnit: { murid: bm, baharu: bb }, belumKelas: belumKelas,
    peratus: N ? Math.round(H / N * 100) : null, diadakan: held.length,
    jumlahPerjumpaan: jad.filter(j => j.Status !== 'Dibatalkan').length,
    seterusnya: next ? { bil: Number(next.Bil), tarikh: next.Tarikh, id: next.ID_Perjumpaan } : null
  };
  try { cache.put(key, JSON.stringify(out), 1800); } catch (e) { /* melebihi 100KB: abaikan cache */ }
  return out;
}

function a_dashboard(d, ctx) {
  const s0 = computeStats_();
  const kelas = String(ctx.user.Kelas_Guru || '');
  // Guru kelas: bilangan murid kelasnya yang belum lengkap berunit
  const s = Object.assign({}, s0, { kelasSaya: kelas ? { kelas: kelas, belum: (s0.belumKelas || {})[kelas] || 0 } : null });
  delete s.belumKelas;
  if (ctx.super) return s;
  const mine = myUnitIds_(ctx);
  const unit = s.unit.filter(u => mine[u.id]);
  const katSaya = {};
  unit.forEach(u => { katSaya[u.idKategori] = 1; });
  // Guru penasihat hanya menerima bilangan murid baharu yang belum berunit dalam kategori unit mereka
  const kategori = s.kategori.filter(k => katSaya[k.id] && k.baharu).map(k => ({ id: k.id, nama: k.nama, baharu: k.baharu }));
  return Object.assign({}, s, { unit: unit, kategori: kategori, jumlahMurid: null, peratus: null });
}

/* =========================================================================
 * HALAMAN UNIT (satu panggilan untuk semua tab)
 * ========================================================================= */

function cartaOut_(sesi, idUnit, users, murid) {
  users = users || mapBy_(rows_('Users'), 'ID_Guru');
  murid = murid || mapBy_(rows_('Murid'), 'ID_Murid');
  return rows_('Carta_Organisasi').filter(c => c.Tahun === sesi && c.ID_Unit === idUnit).map(c => {
    const p = c.Jenis === 'Guru' ? users[c.ID_Rujukan] : murid[c.ID_Rujukan];
    return {
      id: c.ID, jawatan: c.Jawatan, jenis: c.Jenis, idRujukan: c.ID_Rujukan, idGambar: c.ID_Fail_Gambar, susunan: Number(c.Susunan || 0),
      nama: p ? (c.Jenis === 'Guru' ? p.Nama : p.Nama_Murid) : '',
      gambarGuru: p && c.Jenis === 'Guru' ? (p.ID_Gambar || '') : ''
    };
  }).sort((a, b) => a.susunan - b.susunan);
}

function laporanOut_(l, users) {
  const pel = users[l.ID_Guru_Pelapor], pen = users[l.ID_Pengesah];
  return {
    id: l.ID_Laporan, idPerjumpaan: l.ID_Perjumpaan, versi: l.Versi_Format, data: parse_(l.Data_JSON, {}),
    bilHadir: l.Bil_Hadir, status: l.Status, pelapor: pel ? pel.Nama : '', tarikhHantar: l.Tarikh_Hantar,
    pengesah: pen ? pen.Nama : '', tarikhSah: l.Tarikh_Sah, catatan: l.Catatan_Pengesah, kemaskini: l.Tarikh_Kemaskini
  };
}

function a_unit(d, ctx) {
  const idUnit = String(d.idUnit || '');
  // Pengesah boleh membaca unit (untuk menyemak laporan) walaupun bukan guru unit tersebut
  const role = ctx.pengesah && !unitRole_(ctx, idUnit) ? '' : needUnit_(ctx, idUnit);
  const u = rows_('Unit').find(x => x.ID_Unit === idUnit);
  if (!u) throw fail_('Unit tidak dijumpai.');
  const sesi = ctx.sesi;
  const users = mapBy_(rows_('Users'), 'ID_Guru');
  const murid = mapBy_(rows_('Murid'), 'ID_Murid');
  const ahli = rows_('Enrolmen').filter(e => e.Tahun === sesi && e.ID_Unit === idUnit && e.Status === 'Aktif').map(e => {
    const m = murid[e.ID_Murid] || {};
    return { idEnrolmen: e.ID_Enrolmen, id: e.ID_Murid, nama: m.Nama_Murid || '(rekod murid tiada)', kelas: m.Kelas || e.Kelas, jantina: m.Jantina || '', masuk: e.Tarikh_Masuk || '' };
  }).sort((a, b) => a.nama.localeCompare(b.nama));
  const isAdmin = role === 'Admin' || role === 'Super';
  const guru = rows_('Guru_Unit').filter(g => g.Tahun === sesi && g.ID_Unit === idUnit && (g.Status === 'Aktif' || (isAdmin && g.Status === 'Menunggu'))).map(g => {
    const x = users[g.ID_Guru] || {};
    return { id: g.ID, idGuru: g.ID_Guru, nama: x.Nama || '', emel: x.Emel || '', telefon: x.No_Telefon || '', gambar: x.ID_Gambar || '', peranan: g.Peranan_Unit, status: g.Status };
  }).sort((a, b) => (a.peranan === 'Admin' ? 0 : 1) - (b.peranan === 'Admin' ? 0 : 1) || a.nama.localeCompare(b.nama));
  const kehadiran = {};
  rows_('Kehadiran').forEach(k => {
    if (k.Tahun !== sesi || k.ID_Unit !== idUnit) return;
    const o = {};
    P_COLS.forEach((c, i) => { if (k[c]) o[i + 1] = k[c]; });
    kehadiran[k.ID_Murid] = o;
  });
  const laporan = rows_('Laporan').filter(l => l.Tahun === sesi && l.ID_Unit === idUnit).map(l => laporanOut_(l, users));
  const pencapaian = rows_('Pencapaian').filter(p => p.Tahun === sesi && p.ID_Unit === idUnit).map(p => ({
    id: p.ID, idMurid: p.ID_Murid, nama: (murid[p.ID_Murid] || {}).Nama_Murid || '', aktiviti: p.Nama_Aktiviti, peringkat: p.Peringkat,
    pencapaian: p.Pencapaian, tarikh: p.Tarikh, idBukti: p.ID_Fail_Bukti, khidmat: p.Khidmat || '', olehSaya: p.Direkod_Oleh === ctx.id
  })).sort((a, b) => String(b.tarikh).localeCompare(String(a.tarikh)));
  return {
    unit: { id: u.ID_Unit, nama: u.Nama_Unit, idKategori: u.ID_Kategori, status: u.Status || 'Aktif' },
    peranan: role, ahli: ahli, guru: guru, kehadiran: kehadiran, laporan: laporan,
    carta: cartaOut_(sesi, idUnit, users, murid), pencapaian: pencapaian, at: now_()
  };
}

/* =========================================================================
 * KEHADIRAN
 * ========================================================================= */

function a_saveKehadiran(d, ctx) {
  const idUnit = String(d.idUnit || '');
  const bil = Number(d.bil);
  needUnit_(ctx, idUnit);
  needRekodPerjumpaan_(idUnit);
  const j = jadual_(ctx.sesi).find(x => Number(x.Bil) === bil);
  if (!j) throw fail_('Perjumpaan ini tiada dalam jadual.');
  if (j.Status === 'Dibatalkan') throw fail_('Perjumpaan ini telah dibatalkan.');
  const today = today_();
  if (j.Tarikh && j.Tarikh > today) throw fail_('Kehadiran hanya boleh direkod pada atau selepas tarikh perjumpaan.');
  if (!ctx.super && j.Tarikh) {
    const had = Number(tetapan_().TEMPOH_KUNCI_HARI || 7);
    if (addDays_(j.Tarikh, had) < today) throw fail_('Tempoh merekod kehadiran perjumpaan ini telah tamat. Hubungi Super Admin untuk pindaan.', 'LOCKED');
  }
  const members = {};
  rows_('Enrolmen').forEach(e => { if (e.Tahun === ctx.sesi && e.ID_Unit === idUnit && e.Status === 'Aktif' && sudahMasuk_(e.Tarikh_Masuk, j.Tarikh)) members[e.ID_Murid] = 1; });
  const col = 'P' + bil;
  const t = table_('Kehadiran');
  const existing = {};
  t.rows.forEach(r => { if (r.Tahun === ctx.sesi && r.ID_Unit === idUnit) existing[r.ID_Murid] = r; });
  const now = now_();
  const marks = d.marks || {};
  const changed = [], baru = [], logs = [];
  Object.keys(marks).forEach(idm => {
    const v = String(marks[idm] || '');
    if (['H', 'T', 'B', ''].indexOf(v) < 0 || !members[idm]) return;
    const r = existing[idm];
    if (r) {
      if (r[col] === v) return;
      if (r[col]) logs.push({ Masa: now, ID_Guru: ctx.id, Tindakan: 'Pinda kehadiran P' + bil, Sheet: 'Kehadiran', ID_Rekod: r.ID_Rekod, Nilai_Lama: r[col], Nilai_Baharu: v });
      r[col] = v; r.Dikemaskini_Oleh = ctx.id; r.Tarikh_Kemaskini = now;
      changed.push(r);
    } else if (v) {
      const o = { ID_Rekod: uid_('K'), Tahun: ctx.sesi, ID_Unit: idUnit, ID_Murid: idm, Dikemaskini_Oleh: ctx.id, Tarikh_Kemaskini: now };
      o[col] = v;
      baru.push(o);
    }
  });
  if (changed.length) {
    // Tulis tiga lajur sahaja (Pn, Dikemaskini_Oleh, Tarikh_Kemaskini) untuk semua baris — 2 panggilan tanpa mengira bilangan murid
    const n = t.rows.length;
    const ci = t.head.indexOf(col), di = t.head.indexOf('Dikemaskini_Oleh'), ti = t.head.indexOf('Tarikh_Kemaskini');
    t.sh.getRange(2, ci + 1, n, 1).setValues(t.rows.map(r => [cell_(r[col])]));
    if (ti === di + 1) t.sh.getRange(2, di + 1, n, 2).setValues(t.rows.map(r => [cell_(r.Dikemaskini_Oleh), cell_(r.Tarikh_Kemaskini)]));
    else changed.forEach(r => update_('Kehadiran', r, {}));
  }
  insertMany_('Kehadiran', baru);
  if (logs.length) insertMany_('Log', logs);
  return { bil: bil, disimpan: changed.length + baru.length };
}

/* =========================================================================
 * LAPORAN PERJUMPAAN
 * ========================================================================= */

/** Medan laporan berkesan: Tajuk sentiasa ada di atas, walaupun tiada dalam format. */
function medanLaporan_(fields) {
  const f = (fields || []).slice();
  if (!f.some(x => x.id === 'tajuk')) f.unshift({ id: 'tajuk', label: 'Tajuk', jenis: 'teks', wajib: true });
  return f;
}

function a_saveLaporan(d, ctx) {
  const idUnit = String(d.idUnit || '');
  needUnit_(ctx, idUnit);
  needRekodPerjumpaan_(idUnit);
  const j = jadual_(ctx.sesi).find(x => x.ID_Perjumpaan === d.idPerjumpaan);
  if (!j) throw fail_('Perjumpaan tidak dijumpai.');
  if (j.Status === 'Dibatalkan') throw fail_('Perjumpaan ini telah dibatalkan.');
  let l = rows_('Laporan').find(x => x.Tahun === ctx.sesi && x.ID_Unit === idUnit && x.ID_Perjumpaan === d.idPerjumpaan);
  if (l && l.Status === 'Disahkan') throw fail_('Laporan ini telah disahkan dan dikunci.');
  if (l && l.Status === 'Dihantar' && !ctx.super) throw fail_('Laporan sedang menunggu pengesahan dan tidak boleh disunting.');
  const formats = rows_('Format_Laporan');
  const fmt = l ? formats.find(f => f.Versi === l.Versi_Format) : formats.find(f => f.Aktif === 'Ya');
  if (!fmt) throw fail_('Format laporan belum ditetapkan.');
  const fields = medanLaporan_(parse_(fmt.Medan_JSON, []));
  const src = d.data || {};
  const data = {};
  fields.forEach(f => {
    const v = src[f.id];
    data[f.id] = f.jenis === 'gambar' ? (Array.isArray(v) ? v.slice(0, 6).map(x => clean_(x, 80)) : []) : clean_(v, 5000);
  });
  if (d.hantar) {
    const miss = fields.filter(f => f.wajib && (f.jenis === 'gambar' ? !data[f.id].length : !data[f.id]));
    if (miss.length) throw fail_('Sila lengkapkan: ' + miss.map(f => f.label).join(', ') + '.');
  }
  let hadir = 0;
  rows_('Kehadiran').forEach(k => { if (k.Tahun === ctx.sesi && k.ID_Unit === idUnit && k['P' + j.Bil] === 'H') hadir++; });
  const now = now_();
  const patch = {
    Data_JSON: JSON.stringify(data), Bil_Hadir: hadir, Status: d.hantar ? 'Dihantar' : 'Draf', ID_Guru_Pelapor: ctx.id,
    Tarikh_Hantar: d.hantar ? now : (l ? l.Tarikh_Hantar : ''), Tarikh_Kemaskini: now
  };
  if (l) update_('Laporan', l, patch);
  else l = insert_('Laporan', Object.assign({ ID_Laporan: uid_('L'), Tahun: ctx.sesi, ID_Unit: idUnit, ID_Perjumpaan: j.ID_Perjumpaan, Versi_Format: fmt.Versi, ID_Pengesah: '', Tarikh_Sah: '', Catatan_Pengesah: '' }, patch));
  if (d.hantar) {
    log_(ctx.id, 'Hantar laporan P' + j.Bil, 'Laporan', l.ID_Laporan, '', '');
    emelLaporanDihantar_(ctx, idUnit, j, data);
  }
  return { laporan: laporanOut_(l, mapBy_(rows_('Users'), 'ID_Guru')) };
}


/* =========================================================================
 * PEMBANTU AI: jana objektif, aktiviti dan refleksi laporan perjumpaan
 * Kunci API disimpan dalam Script Properties (tidak dalam Sheets, tidak dihantar ke aplikasi).
 * Tiada nama murid dihantar kepada perkhidmatan AI, hanya bilangan.
 * ========================================================================= */
const AI_MODEL_LALAI = { gemini: 'gemini-3.8-flash', claude: 'claude-haiku-5-5' };

function aiConfig_() {
  const p = PropertiesService.getScriptProperties();
  const provider = p.getProperty('AI_PROVIDER') === 'claude' ? 'claude' : 'gemini';
  return {
    provider: provider, key: p.getProperty('AI_KEY_' + provider) || '',
    model: p.getProperty('AI_MODEL') || AI_MODEL_LALAI[provider],
    had: Math.max(1, Math.min(200, parseInt(p.getProperty('AI_HAD_HARIAN') || '30', 10) || 30))
  };
}

function a_aiTetapan(d, ctx) {
  needSuper_(ctx);
  const c = aiConfig_();
  return { provider: c.provider, model: c.model, had: c.had, aktif: tetapan_().AI_AKTIF !== 'Tidak', adaKunci: !!c.key,
    kunci: c.key ? '••••' + c.key.slice(-4) : '', lalai: AI_MODEL_LALAI };
}

function a_saveAiTetapan(d, ctx) {
  needSuper_(ctx);
  const p = PropertiesService.getScriptProperties();
  const provider = d.provider === 'claude' ? 'claude' : 'gemini';
  p.setProperty('AI_PROVIDER', provider);
  const model = clean_(d.model, 80);
  if (model && !/^[a-zA-Z0-9._:\-\/]+$/.test(model)) throw fail_('Nama model tidak sah.');
  if (model && model !== AI_MODEL_LALAI[provider]) p.setProperty('AI_MODEL', model); else p.deleteProperty('AI_MODEL');
  if (d.had !== undefined) p.setProperty('AI_HAD_HARIAN', String(Math.max(1, Math.min(200, parseInt(d.had, 10) || 30))));
  if (d.buangKunci) p.deleteProperty('AI_KEY_' + provider);
  else if (d.kunci) {
    const k = String(d.kunci).trim();
    if (k.length < 20 || /\s/.test(k)) throw fail_('Kunci API tidak sah.');
    p.setProperty('AI_KEY_' + provider, k);
  }
  setTetapan_('AI_AKTIF', d.aktif === false ? 'Tidak' : 'Ya');
  clearCaches_();
  log_(ctx.id, 'Tetapan pembantu AI', 'Tetapan', 'AI', '', provider + (d.kunci ? ', kunci baharu' : '') + (d.buangKunci ? ', kunci dibuang' : ''));
  let uji = null;
  if (d.uji) {
    try { const out = aiPanggil_(aiConfig_(), 'Balas dengan JSON sahaja.', 'Pulangkan {"ok":"ya"}', 200); uji = { ok: true, teks: String(out).slice(0, 80) }; }
    catch (e) { uji = { ok: false, ralat: String(e.message || e).slice(0, 300) }; }
  }
  return Object.assign(a_aiTetapan(d, ctx), { uji: uji });
}

/** Model sandaran jika model utama sesak (503/overloaded). Alias "-latest" sentiasa menunjuk ke model Flash semasa. */
const AI_MODEL_SANDARAN = { gemini: ['gemini-flash-latest', 'gemini-flash-lite-latest'], claude: [] };
/** Kod HTTP yang bermaksud pelayan AI sesak atau gangguan sementara: wajar dicuba semula. */
const AI_KOD_SEMENTARA = [429, 500, 502, 503, 504, 529];

/**
 * Panggil penyedia AI dan pulangkan teks jawapan.
 * Jika pelayan sesak: cuba semula model utama hingga 3 kali (jeda 1s, 2s), kemudian cuba model sandaran.
 * Ralat kekal (kunci salah, permintaan tidak sah) tidak dicuba semula.
 */
function aiPanggil_(c, sistem, mesej, maks) {
  if (!c.key) throw fail_('Pembantu AI belum disediakan. Super Admin perlu memasukkan kunci API.');
  const mula = Date.now(), HAD_MS = 55000;
  const models = [c.model].concat((AI_MODEL_SANDARAN[c.provider] || []).filter(m => m !== c.model));
  let utama = null;
  for (let mi = 0; mi < models.length; mi++) {
    const cubaan = mi === 0 ? 3 : 1;
    for (let i = 0; i < cubaan; i++) {
      if (i > 0) Utilities.sleep(1000 * Math.pow(2, i - 1) + Math.floor(Math.random() * 400));
      const r = aiSekali_(c, models[mi], sistem, mesej, maks);
      if (r.teks) {
        if (mi > 0) Logger.log('AI: model ' + c.model + ' sesak, jawapan daripada model sandaran ' + models[mi]);
        return r.teks;
      }
      if (!r.sementara) {
        if (mi === 0) throw r.ralat;   // ralat kekal pada model utama: laporkan terus
        break;                          // model sandaran tidak tersedia: abaikan, laporkan ralat model utama
      }
      if (mi === 0) utama = r;
      if (Date.now() - mula > HAD_MS) break;
    }
    if (Date.now() - mula > HAD_MS) break;
  }
  if (utama && utama.kod === 429) throw fail_('Had penggunaan perkhidmatan AI telah dicapai (429). Cuba lagi dalam beberapa minit.', 'AISESAK');
  throw fail_('Perkhidmatan AI sedang sesak' + (utama && utama.kod ? ' (ralat ' + utama.kod + ')' : '') + '. Sistem telah mencuba beberapa kali. Cuba lagi dalam 1 hingga 2 minit.', 'AISESAK');
}

/** Satu panggilan HTTP. Pulangkan { teks } atau { ralat, sementara, kod }. */
function aiSekali_(c, model, sistem, mesej, maks) {
  let url, opt, ulang = null;
  if (c.provider === 'claude') {
    url = 'https://api.anthropic.com/v1/messages';
    opt = { method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      headers: { 'x-api-key': c.key, 'anthropic-version': '2023-06-01' },
      payload: JSON.stringify({ model: model, max_tokens: maks || 1500, system: sistem, messages: [{ role: 'user', content: mesej }] }) };
  } else {
    url = 'https://generativelanguage.googleapis.com/v1beta/models/' + encodeURIComponent(model) + ':generateContent';
    const gen = { temperature: 0.7, maxOutputTokens: Math.max(2048, maks || 1500), responseMimeType: 'application/json' };
    // Mod "berfikir" (thinking) dimatikan/diminimumkan: laporan ringkas tidak memerlukannya dan ia melambatkan jawapan
    const think = thinkingGemini_(model);
    if (think) gen.thinkingConfig = think;
    const bina = g => ({ method: 'post', contentType: 'application/json', muteHttpExceptions: true, headers: { 'x-goog-api-key': c.key },
      payload: JSON.stringify({ systemInstruction: { parts: [{ text: sistem }] }, contents: [{ role: 'user', parts: [{ text: mesej }] }], generationConfig: g }) });
    opt = bina(gen);
    if (think) { const g2 = Object.assign({}, gen); delete g2.thinkingConfig; ulang = bina(g2); }
  }
  const panggil = o => {
    try { return UrlFetchApp.fetch(url, o); } catch (e) {
      const m = String(e.message || e);
      if (/permission|external_request|kebenaran/i.test(m)) throw fail_('Apps Script belum dibenarkan menghubungi perkhidmatan luar. Pemilik skrip perlu membuka editor Apps Script, pilih fungsi benarkanAI, tekan Run dan luluskan kebenaran, kemudian Deploy versi baharu.');
      return null;   // gangguan rangkaian/tamat masa: dikira sementara
    }
  };
  let res = panggil(opt);
  if (!res) return { sementara: true, kod: 0, ralat: fail_('Perkhidmatan AI tidak dapat dihubungi.') };
  // Model tidak menyokong tetapan thinking: cuba sekali lagi tanpanya
  if (ulang && res.getResponseCode() === 400 && /think/i.test(res.getContentText())) {
    res = panggil(ulang);
    if (!res) return { sementara: true, kod: 0, ralat: fail_('Perkhidmatan AI tidak dapat dihubungi.') };
  }
  const code = res.getResponseCode(), body = parse_(res.getContentText(), {});
  const pesan = String((body.error && body.error.message) || '').slice(0, 200);
  if (AI_KOD_SEMENTARA.indexOf(code) >= 0) return { sementara: true, kod: code, ralat: fail_('Perkhidmatan AI sesak (' + code + ').') };
  if (code === 401 || code === 403) return { sementara: false, kod: code, ralat: fail_('Kunci API ditolak oleh penyedia AI. Super Admin perlu menyemak kunci.') };
  if (code === 404) return { sementara: false, kod: code, ralat: fail_('Model AI "' + model + '" tidak dijumpai. Super Admin perlu menyemak nama model di Pentadbiran, Pembantu AI.') };
  if (code < 200 || code >= 300) return { sementara: false, kod: code, ralat: fail_('Perkhidmatan AI memulangkan ralat ' + code + ': ' + pesan) };
  let teks = '';
  if (c.provider === 'claude') teks = ((body.content || []).find(x => x.type === 'text') || {}).text || '';
  else teks = (((((body.candidates || [])[0] || {}).content || {}).parts || []).filter(x => x.text && !x.thought).map(x => x.text).join(''));
  // Jawapan kosong (contoh disekat atau terputus) juga dicuba semula
  if (!teks) return { sementara: true, kod: code, ralat: fail_('Perkhidmatan AI tidak memulangkan jawapan. Cuba lagi.') };
  return { teks: teks };
}

/** Tetapan thinking paling rendah mengikut keluarga model Gemini (null = jangan hantar). */
function thinkingGemini_(model) {
  const m = String(model || '').toLowerCase();
  if (/gemini-2\.5-flash/.test(m)) return { thinkingBudget: 0 };          // 2.5 Flash / Flash-Lite: matikan
  if (/gemini-2\.5-pro/.test(m)) return { thinkingBudget: 128 };          // 2.5 Pro: minimum yang dibenarkan
  if (/gemini-[3-9]/.test(m)) return { thinkingLevel: /flash/.test(m) ? 'minimal' : 'low' };
  return null;
}

function aiJson_(teks) {
  let t = String(teks).trim().replace(/^```(?:json)?\s*/i, '').replace(/```\s*$/, '');
  const a = t.indexOf('{'), b = t.lastIndexOf('}');
  if (a >= 0 && b > a) t = t.slice(a, b + 1);
  const o = parse_(t, null);
  if (!o || typeof o !== 'object') throw fail_('Jawapan AI tidak dapat dibaca. Cuba jana semula.');
  return o;
}

function panduanMedan_(f) {
  const s = (f.id + ' ' + f.label).toLowerCase();
  if (/objektif|tujuan|matlamat/.test(s)) return 'Senarai 3 hingga 4 objektif bernombor (1., 2., ...). Setiap objektif bermula dengan "Murid dapat" dan boleh diukur.';
  if (/aktiviti|perjalanan|pelaksanaan|langkah/.test(s)) return 'Senarai 4 hingga 6 langkah aktiviti bernombor mengikut urutan: pembukaan dan taklimat, aktiviti utama, latihan atau pertandingan kecil, dan penutup. Ayat ringkas, kala lampau.';
  if (/refleksi|rumusan|ulasan|penilaian/.test(s)) return 'Satu perenggan 3 hingga 5 ayat dari sudut guru penasihat: perkara yang berjalan baik, kelemahan atau cabaran, dan satu cadangan penambahbaikan untuk perjumpaan seterusnya.';
  if (/cadangan|tindakan/.test(s)) return 'Dua hingga tiga cadangan penambahbaikan bernombor.';
  return 'Teks ringkas dan sesuai untuk laporan rasmi.';
}

function a_aiLaporan(d, ctx) {
  const idUnit = String(d.idUnit || '');
  needUnit_(ctx, idUnit);
  const t = tetapan_();
  if (t.AI_AKTIF === 'Tidak') throw fail_('Pembantu AI telah dimatikan oleh Super Admin.');
  const c = aiConfig_();
  if (!c.key) throw fail_('Pembantu AI belum disediakan. Super Admin perlu memasukkan kunci API di Pentadbiran, Pembantu AI.');
  const tajuk = clean_(d.tajuk, 150);
  if (tajuk.length < 3) throw fail_('Isi Tajuk dahulu supaya AI tahu aktiviti perjumpaan ini.');
  // Guna data dalam cache (common_) supaya tidak membaca sheet Jadual, Format, Unit dan Kategori setiap kali
  let c0 = common_(ctx.sesi);
  if (c0.sesi !== ctx.sesi) { CacheService.getScriptCache().remove('c_common'); c0 = common_(ctx.sesi); }
  const j = (c0.jadual.find(x => x.id === d.idPerjumpaan) || null);
  if (!j) throw fail_('Perjumpaan tidak dijumpai.');
  const fmt = c0.format.find(f => String(f.versi) === String(d.versi)) || c0.format.find(f => f.aktif);
  const semua = medanLaporan_(fmt ? fmt.medan : []);
  const ingin = Array.isArray(d.medan) ? d.medan.map(String) : [];
  const sasaran = semua.filter(f => ingin.indexOf(f.id) >= 0 && (f.jenis === 'teks' || f.jenis === 'panjang') && f.id !== 'tajuk');
  if (!sasaran.length) throw fail_('Tiada medan untuk dijana.');

  // Had penggunaan harian setiap guru
  const cache = CacheService.getScriptCache();
  const kunciHad = 'ai_' + ctx.id + '_' + today_();
  const guna = Number(cache.get(kunciHad) || 0);
  if (guna >= c.had && !ctx.super) throw fail_('Had harian ' + c.had + ' kali penjanaan AI telah dicapai. Cuba lagi esok.', 'HAD');

  const cu = c0.unit.find(x => x.id === idUnit) || {};
  const ck = c0.kategori.find(x => x.id === cu.idKategori) || {};
  const u = { Nama_Unit: cu.nama || '' }, k = { Nama_Kategori: ck.nama || '' };
  // Bilangan hadir/ahli dihantar oleh aplikasi (sudah ada dalam peranti). Hanya untuk konteks AI, jadi tidak perlu membaca sheet Kehadiran yang besar.
  const nOk = v => v !== undefined && v !== null && v !== '' && isFinite(v) && Number(v) >= 0 && Number(v) <= 2000;
  let hadir, bilAhli;
  if (nOk(d.hadir) && nOk(d.ahli)) { hadir = Math.floor(Number(d.hadir)); bilAhli = Math.floor(Number(d.ahli)); }
  else {
    bilAhli = rows_('Enrolmen').filter(e => e.Tahun === ctx.sesi && e.ID_Unit === idUnit && e.Status === 'Aktif').length;
    hadir = 0;
    rows_('Kehadiran').forEach(r => { if (r.Tahun === ctx.sesi && r.ID_Unit === idUnit && r['P' + j.bil] === 'H') hadir++; });
  }
  const lain = d.data || {};
  const konteks = semua.filter(f => sasaran.indexOf(f) < 0 && f.jenis !== 'gambar' && lain[f.id])
    .map(f => '- ' + f.label + ': ' + clean_(lain[f.id], 600)).join('\n');

  const sistem = 'Anda membantu guru penasihat kokurikulum sekolah rendah di Malaysia menulis laporan perjumpaan. ' +
    'Tulis dalam Bahasa Melayu baku yang ringkas, jelas dan sesuai untuk laporan rasmi sekolah. ' +
    'Murid ialah murid Tahun 4 hingga 6. Jangan reka nama orang, nombor, tarikh atau pencapaian yang tidak diberi. ' +
    'Balas dengan satu objek JSON sahaja tanpa teks lain.';
  const mesej = [
    'Maklumat perjumpaan:',
    '- Sekolah: ' + (t.NAMA_SEKOLAH || CONFIG.NAMA_SEKOLAH),
    '- Unit: ' + (u.Nama_Unit || '') + (k.Nama_Kategori ? ' (' + k.Nama_Kategori + ')' : ''),
    '- Perjumpaan ke-' + j.bil + (j.tarikh ? ', ' + j.tarikh : ''),
    '- Tajuk: ' + tajuk,
    bilAhli ? '- Kehadiran: ' + hadir + ' daripada ' + bilAhli + ' ahli' : '',
    d.nota ? '- Catatan guru tentang perjumpaan ini: ' + clean_(d.nota, 800) : '',
    konteks ? 'Maklumat lain dalam laporan:\n' + konteks : '',
    '',
    'Hasilkan kandungan bagi medan berikut. Kunci JSON mesti tepat seperti di bawah:',
    sasaran.map(f => '- "' + f.id + '" (' + f.label + '): ' + panduanMedan_(f)).join('\n'),
    '',
    'Gunakan "\\n" untuk baris baharu dalam senarai bernombor.'
  ].filter(x => x !== '').join('\n');

  const o = aiJson_(aiPanggil_(c, sistem, mesej, Math.min(2500, 500 + sasaran.length * 400)));
  const hasil = {};
  sasaran.forEach(f => {
    let v = o[f.id];
    if (Array.isArray(v)) v = v.map((x, i) => (/^\d+[.)]/.test(String(x).trim()) ? '' : (i + 1) + '. ') + String(x).trim()).join('\n');
    if (v != null && v !== '') hasil[f.id] = clean_(String(v).replace(/\r/g, ''), 5000);
  });
  if (!Object.keys(hasil).length) throw fail_('Jawapan AI kosong. Cuba jana semula.');
  cache.put(kunciHad, String(guna + 1), 86400);
  return { hasil: hasil, baki: ctx.super ? null : Math.max(0, c.had - guna - 1) };
}

/**
 * JALANKAN SEKALI dari editor Apps Script (pilih benarkanAI, tekan Run).
 * Ini memaparkan tetingkap kebenaran "Connect to an external service" supaya pembantu AI boleh berfungsi.
 */
function benarkanAI() {
  const r = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/', { muteHttpExceptions: true });
  Logger.log('Kebenaran sambungan luar diberikan (kod ' + r.getResponseCode() + '). Sekarang Deploy > Manage deployments > Edit > New version.');
}

/** Jalankan sekali dari editor Apps Script untuk memberi kebenaran UrlFetchApp dan menguji kunci. */
function ujiAI() {
  const c = aiConfig_();
  Logger.log('Penyedia: ' + c.provider + ', model: ' + c.model + ', kunci: ' + (c.key ? 'ada' : 'TIADA'));
  if (c.key) Logger.log(aiPanggil_(c, 'Balas dengan JSON sahaja.', 'Pulangkan {"ok":"ya"}', 200));
}

function a_statusLaporan(d, ctx) {
  needPengesah_(ctx);
  const l = rows_('Laporan').find(x => x.ID_Laporan === d.id);
  if (!l) throw fail_('Laporan tidak dijumpai.');
  if (d.status !== 'Disahkan' && d.status !== 'Dikembalikan') throw fail_('Status tidak sah.');
  if (d.status === 'Dikembalikan' && !clean_(d.catatan, 10)) throw fail_('Nyatakan pembetulan yang diperlukan.');
  // Hanya laporan yang dihantar boleh disahkan/dikembalikan. Super Admin boleh membuka semula laporan yang telah disahkan.
  const bolehUbah = l.Status === 'Dihantar' || (ctx.super && l.Status === 'Disahkan' && d.status === 'Dikembalikan');
  if (!bolehUbah) throw fail_('Laporan ini berstatus ' + (l.Status || 'Draf') + ' dan tidak boleh di' + (d.status === 'Disahkan' ? 'sahkan' : 'kembalikan') + '. Muat semula senarai.');
  update_('Laporan', l, {
    Status: d.status, ID_Pengesah: ctx.id, Tarikh_Sah: d.status === 'Disahkan' ? now_() : '', Catatan_Pengesah: clean_(d.catatan, 500)
  });
  log_(ctx.id, (d.status === 'Disahkan' ? 'Sahkan' : 'Kembalikan') + ' laporan', 'Laporan', l.ID_Laporan, '', d.catatan || '');
  if (d.status === 'Dikembalikan') emelLaporanDikembalikan_(l, d.catatan);
  return { laporan: laporanOut_(l, mapBy_(rows_('Users'), 'ID_Guru')) };
}

/* ---------- Pengesahan laporan: senarai menunggu dan notifikasi ---------- */
function urlAplikasi_() {
  try { return ScriptApp.getService().getUrl() || ''; } catch (e) { return ''; }
}
function emelLaporanDihantar_(ctx, idUnit, j, data) {
  try {
    if (tetapan_().EMEL_LAPORAN === 'Tidak') return;
    // Emel kepada pengesah yang dilantik; jika tiada, kepada Super Admin
    const aktif = rows_('Users').filter(u => u.Status === 'Aktif' && u.Emel && u.ID_Guru !== ctx.id);
    const dilantik = aktif.filter(u => u.Pengesah);
    const supers = dilantik.length ? dilantik : aktif.filter(u => u.Peranan_Global === 'SuperAdmin');
    if (!supers.length || MailApp.getRemainingDailyQuota() < supers.length) return;
    const unit = (rows_('Unit').find(u => u.ID_Unit === idUnit) || {}).Nama_Unit || '';
    const n = rows_('Laporan').filter(l => l.Tahun === ctx.sesi && l.Status === 'Dihantar').length;
    const url = urlAplikasi_();
    const badan = 'Assalamualaikum dan salam sejahtera,\n\n' + ctx.user.Nama + ' telah menghantar laporan perjumpaan untuk disahkan.\n\n' +
      'Unit: ' + unit + '\nPerjumpaan: ke-' + j.Bil + (j.Tarikh ? ' (' + j.Tarikh + ')' : '') + '\nTajuk: ' + (data.tajuk || '-') + '\n\n' +
      'Jumlah laporan menunggu pengesahan: ' + n + '.\n' + (url ? '\nSemak dan sahkan di: ' + url + '#/admin/sahkan\n' : '\nSemak di aplikasi: menu Pengesahan laporan.\n') +
      '\nEmel ini dihantar secara automatik oleh sistem e-Kokurikulum.';
    supers.forEach(u => MailApp.sendEmail(u.Emel, 'Laporan untuk disahkan: ' + unit + ', perjumpaan ' + j.Bil, badan));
  } catch (e) { /* emel tidak menghalang penghantaran laporan */ }
}
function emelLaporanDikembalikan_(l, catatan) {
  try {
    const u = rows_('Users').find(x => x.ID_Guru === l.ID_Guru_Pelapor);
    if (!u || !u.Emel || MailApp.getRemainingDailyQuota() < 1) return;
    const unit = (rows_('Unit').find(x => x.ID_Unit === l.ID_Unit) || {}).Nama_Unit || '';
    const j = jadual_(l.Tahun).find(x => x.ID_Perjumpaan === l.ID_Perjumpaan) || {};
    const url = urlAplikasi_();
    MailApp.sendEmail(u.Emel, 'Laporan dikembalikan untuk pembetulan: ' + unit + ', perjumpaan ' + (j.Bil || ''),
      'Assalamualaikum dan salam sejahtera ' + u.Nama + ',\n\nLaporan perjumpaan ke-' + (j.Bil || '') + ' bagi unit ' + unit + ' telah dikembalikan untuk pembetulan.\n\n' +
      'Catatan pengesah:\n' + (catatan || '-') + '\n\n' + (url ? 'Betulkan di: ' + url + '#/laporan/' + l.ID_Unit + '/' + l.ID_Perjumpaan + '\n\n' : '') +
      'Emel ini dihantar secara automatik oleh sistem e-Kokurikulum.');
  } catch (e) { /* abaikan */ }
}
function a_laporanMenunggu(d, ctx) {
  needPengesah_(ctx);
  const sesi = ctx.sesi;
  const units = mapBy_(rows_('Unit'), 'ID_Unit'), users = mapBy_(rows_('Users'), 'ID_Guru'), jad = mapBy_(jadual_(sesi), 'ID_Perjumpaan');
  const laporan = rows_('Laporan').filter(l => l.Tahun === sesi && l.Status === 'Dihantar').map(l => {
    const u = units[l.ID_Unit] || {}, j = jad[l.ID_Perjumpaan] || {}, data = parse_(l.Data_JSON, {});
    return { id: l.ID_Laporan, idUnit: l.ID_Unit, unit: u.Nama_Unit || '', idKategori: u.ID_Kategori || '', idPerjumpaan: l.ID_Perjumpaan,
      bil: Number(j.Bil || 0), tarikh: j.Tarikh || '', tajuk: data.tajuk || '', pelapor: (users[l.ID_Guru_Pelapor] || {}).Nama || '', tarikhHantar: l.Tarikh_Hantar, bilHadir: l.Bil_Hadir };
  }).sort((a, b) => String(a.tarikhHantar).localeCompare(String(b.tarikhHantar)));
  const disahkan = rows_('Laporan').filter(l => l.Tahun === sesi && l.Status === 'Disahkan').length;
  const pengesah = rows_('Users').filter(u => u.Status === 'Aktif' && u.Pengesah).map(u => ({ nama: u.Nama, jawatan: u.Pengesah }));
  return { laporan: laporan, disahkan: disahkan, pengesah: pengesah };
}
function a_sahkanPukal(d, ctx) {
  needPengesah_(ctx);
  const ids = {};
  (d.ids || []).forEach(x => { ids[String(x)] = 1; });
  if (!Object.keys(ids).length) throw fail_('Pilih sekurang-kurangnya satu laporan.');
  const t = table_('Laporan');
  let n = 0;
  t.rows.forEach(l => { if (ids[l.ID_Laporan] && l.Status === 'Dihantar') { l.Status = 'Disahkan'; l.ID_Pengesah = ctx.id; l.Tarikh_Sah = now_(); l.Catatan_Pengesah = ''; n++; } });
  if (n) writeAll_('Laporan');
  log_(ctx.id, 'Sahkan laporan pukal', 'Laporan', '', '', n + ' laporan');
  return { bil: n };
}

/* =========================================================================
 * GAMBAR (Drive peribadi + proxy)
 * ========================================================================= */

function rootFolder_() {
  const p = PropertiesService.getScriptProperties();
  const id = p.getProperty('ROOT_FOLDER_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) { /* cipta semula */ } }
  const f = DriveApp.createFolder('e-Kokurikulum (fail sistem)');
  p.setProperty('ROOT_FOLDER_ID', f.getId());
  return f;
}

function subFolder_(nama) {
  const root = rootFolder_();
  const it = root.getFoldersByName(nama);
  return it.hasNext() ? it.next() : root.createFolder(nama);
}

function unitFolder_(idUnit, sesi) {
  const u = rows_('Unit').find(x => x.ID_Unit === idUnit);
  if (!u) throw fail_('Unit tidak dijumpai.');
  let folder = null;
  if (u.ID_Folder_Drive) { try { folder = DriveApp.getFolderById(u.ID_Folder_Drive); } catch (e) { folder = null; } }
  if (!folder) {
    folder = rootFolder_().createFolder(u.Nama_Unit);
    const fid = folder.getId();
    withLock_(() => { delete T_.Unit; const r = rows_('Unit').find(x => x.ID_Unit === idUnit); if (r) update_('Unit', r, { ID_Folder_Drive: fid }); });
  }
  const it = folder.getFoldersByName(String(sesi));
  return it.hasNext() ? it.next() : folder.createFolder(String(sesi));
}

function a_uploadImage(d, ctx) {
  const scope = ['logo', 'guru'].indexOf(d.scope) >= 0 ? d.scope : 'unit';
  if (scope === 'logo') needSuper_(ctx); else if (scope === 'unit') needUnit_(ctx, d.idUnit);
  const b64 = String(d.base64 || '');
  if (!b64) throw fail_('Tiada gambar diterima.');
  if (b64.length > CONFIG.MAKS_GAMBAR_BAIT * 1.37) throw fail_('Saiz gambar melebihi had 6 MB.');
  const mime = /^image\/(jpeg|png|webp)$/.test(d.mime) ? d.mime : 'image/jpeg';
  const folder = scope === 'logo' ? rootFolder_() : scope === 'guru' ? subFolder_('Gambar Guru') : unitFolder_(d.idUnit, ctx.sesi);
  const nama = clean_(d.nama, 80) || ('gambar-' + Date.now() + '.jpg');
  const f = folder.createFile(Utilities.newBlob(Utilities.base64Decode(b64), mime, nama));
  return { id: f.getId() };
}

function underRoot_(file) {
  const root = PropertiesService.getScriptProperties().getProperty('ROOT_FOLDER_ID');
  if (!root) return false;
  let level = [file.getParents()];
  for (let depth = 0; depth < 5; depth++) {
    const next = [];
    for (const it of level) {
      while (it.hasNext()) {
        const f = it.next();
        if (f.getId() === root) return true;
        next.push(f.getParents());
      }
    }
    if (!next.length) return false;
    level = next;
  }
  return false;
}

function a_image(d) {
  let file;
  try { file = DriveApp.getFileById(String(d.id || '')); } catch (e) { throw fail_('Gambar tidak dijumpai.'); }
  if (!underRoot_(file)) throw fail_('Akses gambar tidak dibenarkan.', 'FORBIDDEN');
  const blob = file.getBlob();
  return { data: 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes()) };
}

/* =========================================================================
 * CARTA ORGANISASI & PENCAPAIAN
 * ========================================================================= */

function a_saveCarta(d, ctx) {
  const idUnit = String(d.idUnit || '');
  needUnit_(ctx, idUnit, true);
  const old = rows_('Carta_Organisasi').filter(c => c.Tahun === ctx.sesi && c.ID_Unit === idUnit);
  const rows = (d.items || []).slice(0, 60).map((it, i) => ({
    ID: clean_(it.id, 40) || uid_('C'), Tahun: ctx.sesi, ID_Unit: idUnit, Jawatan: clean_(it.jawatan, 60),
    Jenis: it.jenis === 'Guru' ? 'Guru' : 'Murid', ID_Rujukan: clean_(it.idRujukan, 40), ID_Fail_Gambar: clean_(it.idGambar, 80), Susunan: i + 1
  })).filter(r => r.Jawatan);
  removeMany_('Carta_Organisasi', old);
  insertMany_('Carta_Organisasi', rows);
  log_(ctx.id, 'Kemas kini carta organisasi', 'Carta_Organisasi', idUnit, old.length + ' baris', rows.length + ' baris');
  return { carta: cartaOut_(ctx.sesi, idUnit) };
}

function a_savePencapaian(d, ctx) {
  const idUnit = String(d.idUnit || '');
  const role = needUnit_(ctx, idUnit);
  const ahli = {};
  rows_('Enrolmen').forEach(e => { if (e.Tahun === ctx.sesi && e.ID_Unit === idUnit && e.Status === 'Aktif') ahli[e.ID_Murid] = 1; });
  // Rekod baharu boleh untuk ramai murid sekali gus (contoh satu pasukan)
  const ids = !d.id && Array.isArray(d.idMurids) ? d.idMurids.map(String).filter((x, i, a) => a.indexOf(x) === i) : [String(d.idMurid || '')];
  if (!ids.length || !ids[0]) throw fail_('Pilih sekurang-kurangnya seorang murid.');
  if (ids.length > 60) throw fail_('Maksimum 60 murid bagi satu rekod.');
  if (ids.some(x => !ahli[x])) throw fail_('Terdapat murid yang bukan ahli unit ini.');
  if (PERINGKAT.indexOf(d.peringkat) < 0) throw fail_('Pilih peringkat pencapaian.');
  if (!clean_(d.aktiviti, 5)) throw fail_('Isi nama aktiviti atau pertandingan.');
  const patch = {
    ID_Murid: ids[0], Nama_Aktiviti: clean_(d.aktiviti, 150), Peringkat: d.peringkat, Pencapaian: clean_(d.pencapaian, 80),
    Tarikh: isYmd_(d.tarikh) ? d.tarikh : '', ID_Fail_Bukti: clean_(d.idBukti, 80)
  };
  if (d.khidmat !== undefined) patch.Khidmat = khidmatSah_(d.khidmat);
  ensureSheet_(ss_(), 'Pencapaian'); delete T_.Pencapaian;
  if (d.id) {
    const p = rows_('Pencapaian').find(x => x.ID === d.id && x.ID_Unit === idUnit);
    if (!p) throw fail_('Rekod tidak dijumpai.');
    if (role === 'Penasihat' && p.Direkod_Oleh !== ctx.id) throw fail_('Hanya perekod atau ketua unit boleh menyunting rekod ini.', 'FORBIDDEN');
    update_('Pencapaian', p, patch);
  } else {
    insertMany_('Pencapaian', ids.map(idm => Object.assign({ ID: uid_('PC'), Tahun: ctx.sesi, ID_Unit: idUnit, Direkod_Oleh: ctx.id }, patch, { ID_Murid: idm })));
  }
  return { ok: true, bil: ids.length };
}

function a_deletePencapaian(d, ctx) {
  const p = rows_('Pencapaian').find(x => x.ID === d.id);
  if (!p) throw fail_('Rekod tidak dijumpai.');
  // Perekod sentiasa boleh memadam rekod sendiri (termasuk guru bukan guru unit melalui halaman Pelibatan)
  if (p.Direkod_Oleh !== ctx.id) {
    const role = needUnit_(ctx, p.ID_Unit);
    if (role === 'Penasihat') throw fail_('Hanya perekod atau ketua unit boleh memadam rekod ini.', 'FORBIDDEN');
  }
  removeMany_('Pencapaian', [p]);
  log_(ctx.id, 'Padam pencapaian', 'Pencapaian', p.ID, p.Nama_Aktiviti, '');
  return { ok: true };
}

/* =========================================================================
 * PELIBATAN PAJSK OLEH SEMUA GURU
 * Mana-mana guru aktif (bukan hanya guru unit) boleh merekod penyertaan, pencapaian dan
 * khidmat sumbangan murid. Rekod terus masuk dan dilekatkan pada unit murid dalam kategori
 * yang dipilih, supaya templat PAJSK MOEIS terus terisi.
 * ========================================================================= */
function khidmatSah_(v) { return SPPB.khidmat.indexOf(String(v || '')) >= 0 ? String(v) : ''; }

function a_pelibatanBorang(d, ctx) {
  const sesi = ctx.sesi;
  const kats = rows_('Kategori').filter(k => k.Status !== 'Tidak Aktif' && tabKategori_(k))
    .sort((a, b) => Number(a.Susunan) - Number(b.Susunan));
  const kSet = mapBy_(kats, 'ID_Kategori');
  const unit = {}, katUnit = {};
  rows_('Unit').forEach(u => { unit[u.ID_Unit] = u.Nama_Unit; katUnit[u.ID_Unit] = u.ID_Kategori; });
  const ada = {};
  rows_('Enrolmen').forEach(e => {
    if (e.Tahun === sesi && e.Status === 'Aktif' && kSet[e.ID_Kategori]) (ada[e.ID_Murid] = ada[e.ID_Murid] || {})[e.ID_Kategori] = e.ID_Unit;
  });
  const semua = mapBy_(rows_('Murid'), 'ID_Murid');
  const murid = rows_('Murid').filter(layak_)
    .sort((a, b) => String(a.Kelas).localeCompare(String(b.Kelas), 'ms', { numeric: true }) || a.Nama_Murid.localeCompare(b.Nama_Murid))
    .map(m => [m.ID_Murid, m.Nama_Murid, m.Kelas, ada[m.ID_Murid] || {}]);
  const rekod = rows_('Pencapaian').filter(p => p.Tahun === sesi && p.Direkod_Oleh === ctx.id).map(p => {
    const m = semua[p.ID_Murid] || {};
    return { id: p.ID, idMurid: p.ID_Murid, nama: m.Nama_Murid || '(rekod murid tiada)', kelas: m.Kelas || '', idUnit: p.ID_Unit, unit: unit[p.ID_Unit] || '',
      idKategori: katUnit[p.ID_Unit] || '', aktiviti: p.Nama_Aktiviti, peringkat: p.Peringkat, pencapaian: p.Pencapaian, tarikh: p.Tarikh, khidmat: p.Khidmat || '' };
  }).sort((a, b) => String(b.tarikh).localeCompare(String(a.tarikh)) || a.aktiviti.localeCompare(b.aktiviti) || a.nama.localeCompare(b.nama));
  const ekstra = rows_('Ekstra_PAJSK').filter(x => x.Tahun === sesi && x.Direkod_Oleh === ctx.id).map(x => {
    const m = semua[x.ID_Murid] || {};
    return { id: x.ID, idMurid: x.ID_Murid, nama: m.Nama_Murid || '(rekod murid tiada)', kelas: m.Kelas || '', jenis: x.Jenis, perkara: x.Perkara, nilai: x.Nilai, aktiviti: x.Aktiviti, tarikh: x.Tarikh };
  }).sort((a, b) => String(b.tarikh).localeCompare(String(a.tarikh)) || a.nama.localeCompare(b.nama));
  return {
    kategori: kats.map(k => ({ id: k.ID_Kategori, nama: k.Nama_Kategori, tab: tabKategori_(k) })),
    unit: unit, murid: murid, rekod: rekod, khidmat: SPPB.khidmat, ekstra: ekstra,
    senarai: { anugerah: SPPB.anugerah, tahapAnugerah: SPPB.tahapAnugerah, perkhidmatan: SPPB.perkhidmatan, nilam: SPPB.nilam, khidmatMasyarakat: SPPB.khidmatMasyarakat }
  };
}

function a_simpanPelibatan(d, ctx) {
  ensureSheet_(ss_(), 'Pencapaian'); delete T_.Pencapaian;
  const sesi = ctx.sesi;
  const aktiviti = clean_(d.aktiviti, 150);
  if (aktiviti.length < 5) throw fail_('Isi nama aktiviti atau pertandingan.');
  if (PERINGKAT.indexOf(d.peringkat) < 0) throw fail_('Pilih peringkat.');
  const tarikh = isYmd_(d.tarikh) ? d.tarikh : '';
  if (tarikh && tarikh > today_()) throw fail_('Tarikh tidak boleh selepas hari ini.');

  // Sunting satu rekod: perekod, ketua unit atau Super Admin
  if (d.id) {
    const p = rows_('Pencapaian').find(x => x.ID === d.id && x.Tahun === sesi);
    if (!p) throw fail_('Rekod tidak dijumpai.');
    if (p.Direkod_Oleh !== ctx.id && !ctx.super && unitRole_(ctx, p.ID_Unit) !== 'Admin') throw fail_('Hanya perekod atau ketua unit boleh menyunting rekod ini.', 'FORBIDDEN');
    update_('Pencapaian', p, { Nama_Aktiviti: aktiviti, Peringkat: d.peringkat, Tarikh: tarikh, Pencapaian: clean_(d.pencapaian, 80), Khidmat: khidmatSah_(d.khidmat) });
    log_(ctx.id, 'Sunting pelibatan', 'Pencapaian', p.ID, '', aktiviti);
    return { bil: 1, tanpaUnit: [], sudahAda: 0 };
  }

  // Rekod baharu untuk ramai murid
  const kat = rows_('Kategori').find(k => k.ID_Kategori === d.idKategori && k.Status !== 'Tidak Aktif');
  if (!kat || !tabKategori_(kat)) throw fail_('Pilih kategori (Kelab, Beruniform atau Sukan).');
  const list = Array.isArray(d.murid) ? d.murid : [];
  if (!list.length) throw fail_('Pilih sekurang-kurangnya seorang murid.');
  if (list.length > 300) throw fail_('Maksimum 300 murid bagi satu rekod.');
  const murid = mapBy_(rows_('Murid').filter(layak_), 'ID_Murid');
  const unitMurid = {};
  rows_('Enrolmen').forEach(e => { if (e.Tahun === sesi && e.Status === 'Aktif' && e.ID_Kategori === kat.ID_Kategori) unitMurid[e.ID_Murid] = e.ID_Unit; });
  const wujud = {};
  rows_('Pencapaian').forEach(p => { if (p.Tahun === sesi) wujud[p.ID_Murid + '|' + p.ID_Unit + '|' + p.Nama_Aktiviti.toLowerCase() + '|' + p.Tarikh] = 1; });
  const add = [], tanpaUnit = [], seen = {};
  let sudahAda = 0;
  list.forEach(x => {
    const id = String((x && x.idMurid) || '');
    if (!id || seen[id]) return;
    seen[id] = 1;
    const m = murid[id];
    if (!m) return;
    const idU = unitMurid[id];
    if (!idU) { tanpaUnit.push(m.Nama_Murid); return; }
    if (wujud[id + '|' + idU + '|' + aktiviti.toLowerCase() + '|' + tarikh]) { sudahAda++; return; }
    add.push({ ID: uid_('PC'), Tahun: sesi, ID_Murid: id, ID_Unit: idU, Nama_Aktiviti: aktiviti, Peringkat: d.peringkat,
      Pencapaian: clean_(x.pencapaian, 80), Tarikh: tarikh, ID_Fail_Bukti: '', Direkod_Oleh: ctx.id, Khidmat: khidmatSah_(x.khidmat) });
  });
  if (!add.length) {
    if (sudahAda) throw fail_('Semua murid yang dipilih sudah direkod untuk aktiviti dan tarikh ini.');
    throw fail_('Tiada murid yang dipilih mempunyai unit ' + kat.Nama_Kategori + '. Daftarkan unit murid dahulu.');
  }
  insertMany_('Pencapaian', add);
  log_(ctx.id, 'Rekod pelibatan', 'Pencapaian', '', '', aktiviti + ' (' + d.peringkat + '): ' + add.length + ' murid');
  return { bil: add.length, tanpaUnit: tanpaUnit, sudahAda: sudahAda };
}

/* ---------- Tab Ekstra PAJSK: Anugerah Khas, Khidmat Masyarakat, NILAM, Perkhidmatan ---------- */
/** Sahkan nilai bagi satu rekod Ekstra. Pulangkan { Perkara, Nilai } atau baling ralat. */
function ekstraNilai_(jenis, perkara, nilai) {
  perkara = String(perkara || ''); nilai = String(nilai || '');
  if (jenis === 'ANUGERAH') {
    if (SPPB.anugerah.indexOf(perkara) < 0) throw fail_('Pilih jenis anugerah.');
    if (SPPB.tahapAnugerah.indexOf(nilai) < 0) throw fail_('Pilih tahap anugerah.');
    return { Perkara: perkara, Nilai: nilai };
  }
  if (jenis === 'KHIDMAT') {
    if (SPPB.khidmatMasyarakat.indexOf(perkara) < 0) throw fail_('Pilih kategori khidmat masyarakat.');
    return { Perkara: perkara, Nilai: '' };
  }
  if (jenis === 'NILAM') {
    if (SPPB.nilam.indexOf(nilai) < 0) throw fail_('Pilih tahap NILAM.');
    return { Perkara: '', Nilai: nilai };
  }
  if (jenis === 'PERKHIDMATAN') {
    if (SPPB.perkhidmatan.indexOf(perkara) < 0) throw fail_('Pilih jawatan perkhidmatan.');
    return { Perkara: perkara, Nilai: '' };
  }
  throw fail_('Jenis rekod tidak sah.');
}

function a_simpanEkstra(d, ctx) {
  ensureSheet_(ss_(), 'Ekstra_PAJSK'); delete T_.Ekstra_PAJSK;
  const sesi = ctx.sesi;
  const aktiviti = clean_(d.aktiviti, 150);
  const tarikh = isYmd_(d.tarikh) ? d.tarikh : '';
  if (tarikh && tarikh > today_()) throw fail_('Tarikh tidak boleh selepas hari ini.');

  if (d.id) {
    const x = rows_('Ekstra_PAJSK').find(r => r.ID === d.id && r.Tahun === sesi);
    if (!x) throw fail_('Rekod tidak dijumpai.');
    if (x.Direkod_Oleh !== ctx.id && !ctx.super) throw fail_('Hanya guru yang merekod boleh menyunting rekod ini.', 'FORBIDDEN');
    if (x.Jenis === 'KHIDMAT' && aktiviti.length < 3) throw fail_('Isi nama aktiviti khidmat masyarakat.');
    update_('Ekstra_PAJSK', x, Object.assign(ekstraNilai_(x.Jenis, d.perkara, d.nilai), { Aktiviti: aktiviti, Tarikh: tarikh }));
    log_(ctx.id, 'Sunting ekstra PAJSK', 'Ekstra_PAJSK', x.ID, '', x.Jenis);
    return { bil: 1, sudahAda: 0 };
  }

  const jenis = String(d.jenis || '');
  if (JENIS_EKSTRA.indexOf(jenis) < 0) throw fail_('Jenis rekod tidak sah.');
  if (jenis === 'KHIDMAT' && aktiviti.length < 3) throw fail_('Isi nama aktiviti khidmat masyarakat.');
  const list = Array.isArray(d.murid) ? d.murid : [];
  if (!list.length) throw fail_('Pilih sekurang-kurangnya seorang murid.');
  if (list.length > 300) throw fail_('Maksimum 300 murid bagi satu rekod.');
  const murid = mapBy_(rows_('Murid').filter(layak_), 'ID_Murid');
  const kunci = r => [r.ID_Murid, r.Jenis, r.Perkara, r.Nilai, String(r.Aktiviti).toLowerCase(), r.Tarikh].join('|');
  const wujud = {};
  rows_('Ekstra_PAJSK').forEach(r => { if (r.Tahun === sesi) wujud[kunci(r)] = 1; });
  const add = [], seen = {};
  let sudahAda = 0;
  const now = now_();
  list.forEach(x => {
    const id = String((x && x.idMurid) || '');
    if (!id || seen[id] || !murid[id]) return;
    seen[id] = 1;
    // Nilai bagi setiap murid: tahap anugerah, bintang NILAM atau jawatan perkhidmatan
    const v = (x.nilai !== undefined && x.nilai !== '') ? x.nilai : d.nilai;
    const pn = jenis === 'PERKHIDMATAN' ? ekstraNilai_(jenis, v, '') : ekstraNilai_(jenis, d.perkara, v);
    const o = Object.assign({ ID: uid_('EX'), Tahun: sesi, ID_Murid: id, Jenis: jenis, Aktiviti: aktiviti, Tarikh: tarikh, Direkod_Oleh: ctx.id, Tarikh_Rekod: now }, pn);
    if (wujud[kunci(o)]) { sudahAda++; return; }
    wujud[kunci(o)] = 1;
    add.push(o);
  });
  if (!add.length) throw fail_(sudahAda ? 'Semua murid yang dipilih sudah direkod untuk perkara ini.' : 'Tiada murid sah dipilih.');
  insertMany_('Ekstra_PAJSK', add);
  log_(ctx.id, 'Rekod ekstra PAJSK', 'Ekstra_PAJSK', '', '', jenis + ': ' + add.length + ' murid');
  return { bil: add.length, sudahAda: sudahAda };
}

function a_padamEkstra(d, ctx) {
  const x = rows_('Ekstra_PAJSK').find(r => r.ID === d.id);
  if (!x) throw fail_('Rekod tidak dijumpai.');
  if (x.Direkod_Oleh !== ctx.id && !ctx.super) throw fail_('Hanya guru yang merekod boleh memadam rekod ini.', 'FORBIDDEN');
  removeMany_('Ekstra_PAJSK', [x]);
  log_(ctx.id, 'Padam ekstra PAJSK', 'Ekstra_PAJSK', x.ID, x.Jenis, '');
  return { ok: true };
}

/** Baris tab Ekstra bagi seorang murid: satu nilai terbaik bagi setiap lajur (SPPB hanya menerima satu). */
function barisEkstra_(id, nama, rec) {
  rec = rec || [];
  const ang = rec.filter(x => x.Jenis === 'ANUGERAH').sort((a, b) =>
    (SPPB.anugerah.indexOf(a.Perkara) - SPPB.anugerah.indexOf(b.Perkara)) ||
    (['EMAS', 'PERAK', 'GANGSA', 'PENERIMA'].indexOf(a.Nilai) - ['EMAS', 'PERAK', 'GANGSA', 'PENERIMA'].indexOf(b.Nilai)))[0];
  const km = SPPB.khidmatMasyarakat.map(k => { const n = rec.filter(x => x.Jenis === 'KHIDMAT' && x.Perkara === k).length; return n ? Math.min(5, n) : ''; });
  const nilam = rec.filter(x => x.Jenis === 'NILAM').map(x => x.Nilai).sort((a, b) => SPPB.nilam.indexOf(a) - SPPB.nilam.indexOf(b))[0] || '';
  const mata = v => Number(String(v).split('|')[1] || 0);
  const khid = rec.filter(x => x.Jenis === 'PERKHIDMATAN').map(x => x.Perkara).sort((a, b) => mata(b) - mata(a))[0] || '';
  return [id, nama, 'Ya', ang ? ang.Perkara : '', ang ? ang.Nilai : ''].concat(km, [nilam, khid]);
}

/* =========================================================================
 * KEAHLIAN MURID & GURU DALAM UNIT
 * ========================================================================= */

function a_muridPicker(d, ctx) {
  const idUnit = String(d.idUnit || '');
  needUnit_(ctx, idUnit);
  const u = rows_('Unit').find(x => x.ID_Unit === idUnit);
  if (!u) throw fail_('Unit tidak dijumpai.');
  const taken = {};
  rows_('Enrolmen').forEach(e => { if (e.Tahun === ctx.sesi && e.ID_Kategori === u.ID_Kategori && e.Status === 'Aktif') taken[e.ID_Murid] = 1; });
  const list = rows_('Murid').filter(m => layak_(m) && !taken[m.ID_Murid])
    .map(m => [m.ID_Murid, m.Nama_Murid, m.Kelas, muridBaharu_(m) ? 1 : 0])
    .sort((a, b) => String(a[2]).localeCompare(String(b[2]), 'ms', { numeric: true }) || String(a[1]).localeCompare(String(b[1])));
  return { murid: list };
}

function a_enrol(d, ctx) {
  const idUnit = String(d.idUnit || '');
  needUnit_(ctx, idUnit);
  const masuk = isYmd_(d.tarikh) && d.tarikh <= today_() ? d.tarikh : today_();
  const u = rows_('Unit').find(x => x.ID_Unit === idUnit);
  if (!u) throw fail_('Unit tidak dijumpai.');
  const murid = mapBy_(rows_('Murid'), 'ID_Murid');
  const taken = {};
  rows_('Enrolmen').forEach(e => { if (e.Tahun === ctx.sesi && e.ID_Kategori === u.ID_Kategori && e.Status === 'Aktif') taken[e.ID_Murid] = e.ID_Unit; });
  const add = [], konflik = [];
  (d.ids || []).slice(0, 500).forEach(id => {
    const m = murid[id];
    if (!m || m.Status !== 'Aktif') return;
    if (taken[id]) { if (taken[id] !== idUnit) konflik.push(m.Nama_Murid); return; }
    taken[id] = idUnit;
    add.push({ ID_Enrolmen: uid_('E'), Tahun: ctx.sesi, ID_Murid: id, ID_Unit: idUnit, ID_Kategori: u.ID_Kategori, Kelas: m.Kelas, Status: 'Aktif', Tarikh_Masuk: masuk });
  });
  insertMany_('Enrolmen', add);
  if (add.length) log_(ctx.id, 'Tambah ahli', 'Enrolmen', idUnit, '', add.length + ' murid');
  return { ditambah: add.length, konflik: konflik };
}

function a_unenrol(d, ctx) {
  const e = rows_('Enrolmen').find(x => x.ID_Enrolmen === d.idEnrolmen && x.Tahun === ctx.sesi);
  if (!e) throw fail_('Rekod keahlian tidak dijumpai.');
  needUnit_(ctx, e.ID_Unit, true);
  update_('Enrolmen', e, { Status: 'Keluar' });
  log_(ctx.id, 'Keluarkan ahli', 'Enrolmen', e.ID_Enrolmen, 'Aktif', 'Keluar');
  return { ok: true };
}

function a_guruUnit(d, ctx) {
  const g = rows_('Guru_Unit').find(x => x.ID === d.id && x.Tahun === ctx.sesi);
  if (!g) throw fail_('Rekod guru tidak dijumpai.');
  const op = d.op;
  if (op === 'admin' || op === 'penasihat') {
    needSuper_(ctx);
    if (g.Status !== 'Aktif') throw fail_('Luluskan guru ini dahulu.');
    // Satu ketua (Guru Penasihat unit) bagi setiap unit: ketua sedia ada menjadi guru penasihat biasa
    let ketuaLama = '';
    if (op === 'admin') {
      const users = mapBy_(rows_('Users'), 'ID_Guru');
      rows_('Guru_Unit').filter(x => x.Tahun === g.Tahun && x.ID_Unit === g.ID_Unit && x.ID !== g.ID && x.Peranan_Unit === 'Admin')
        .forEach(x => { update_('Guru_Unit', x, { Peranan_Unit: 'Penasihat' }); ketuaLama = (users[x.ID_Guru] || {}).Nama || ''; });
    }
    update_('Guru_Unit', g, { Peranan_Unit: op === 'admin' ? 'Admin' : 'Penasihat' });
    log_(ctx.id, 'Guru unit: ' + op, 'Guru_Unit', g.ID, ketuaLama, g.ID_Guru);
    return { ok: true, ketuaLama: ketuaLama };
  } else if (op === 'lulus' || op === 'tolak' || op === 'buang') {
    needUnit_(ctx, g.ID_Unit, true);
    if (g.ID_Guru === ctx.id && op !== 'lulus') throw fail_('Anda tidak boleh membuang diri sendiri daripada unit.');
    update_('Guru_Unit', g, { Status: op === 'lulus' ? 'Aktif' : (op === 'tolak' ? 'Ditolak' : 'Tidak Aktif') });
  } else {
    throw fail_('Tindakan tidak sah.');
  }
  log_(ctx.id, 'Guru unit: ' + op, 'Guru_Unit', g.ID, '', g.ID_Guru);
  return { ok: true };
}

function a_requestUnit(d, ctx) {
  const idUnit = String(d.idUnit || '');
  const u = rows_('Unit').find(x => x.ID_Unit === idUnit && x.Status !== 'Tidak Aktif');
  if (!u) throw fail_('Unit tidak dijumpai.');
  const g = rows_('Guru_Unit').find(x => x.Tahun === ctx.sesi && x.ID_Unit === idUnit && x.ID_Guru === ctx.id);
  if (g && g.Status === 'Aktif') throw fail_('Anda sudah menjadi guru unit ini.');
  const st = 'Aktif';  // guru menyertai unit terus tanpa kelulusan
  if (g) update_('Guru_Unit', g, { Status: st, Peranan_Unit: 'Penasihat' });
  else insert_('Guru_Unit', { ID: uid_('GU'), Tahun: ctx.sesi, ID_Guru: ctx.id, ID_Unit: idUnit, Peranan_Unit: 'Penasihat', Status: st });
  return { status: st };
}

/* =========================================================================
 * SUPER ADMIN: PENGGUNA, KATEGORI, UNIT
 * ========================================================================= */

function a_users(d, ctx) {
  needSuper_(ctx);
  const units = mapBy_(rows_('Unit'), 'ID_Unit');
  const gu = {};
  rows_('Guru_Unit').forEach(g => {
    if (g.Tahun !== ctx.sesi || g.Status === 'Tidak Aktif' || g.Status === 'Ditolak') return;
    (gu[g.ID_Guru] = gu[g.ID_Guru] || []).push({ id: g.ID, idUnit: g.ID_Unit, nama: (units[g.ID_Unit] || {}).Nama_Unit || '', peranan: g.Peranan_Unit, status: g.Status });
  });
  return {
    users: rows_('Users').map(u => ({
      id: u.ID_Guru, nama: u.Nama, emel: u.Emel, telefon: u.No_Telefon, peranan: u.Peranan_Global, status: u.Status,
      daftar: u.Tarikh_Daftar, login: u.Log_Masuk_Terakhir, unit: gu[u.ID_Guru] || [],
      kp: u.No_KP ? '••••••••' + String(u.No_KP).slice(-4) : '', tukar: u.Tukar_Kata_Laluan === 'Ya', gambar: u.ID_Gambar || '', kelas: u.Kelas_Guru || '', pengesah: u.Pengesah || ''
    })).sort((a, b) => (a.status === 'Menunggu' ? 0 : 1) - (b.status === 'Menunggu' ? 0 : 1) || a.nama.localeCompare(b.nama)),
    lalai: CONFIG.KATA_LALUAN_LALAI
  };
}

function a_userAction(d, ctx) {
  needSuper_(ctx);
  const u = rows_('Users').find(x => x.ID_Guru === d.idGuru);
  if (!u) throw fail_('Pengguna tidak dijumpai.');
  const op = d.op;
  if (u.ID_Guru === ctx.id && (op === 'nyahaktif' || op === 'guru')) throw fail_('Anda tidak boleh menukar status atau peranan akaun sendiri.');
  if (u.Emel === superEmail_() && (op === 'nyahaktif' || op === 'guru')) throw fail_('Akaun Super Admin utama sekolah tidak boleh dinyahaktifkan atau diturunkan.');
  if (op === 'lulus' || op === 'aktifkan') {
    update_('Users', u, { Status: 'Aktif' });
    if (op === 'lulus') rows_('Guru_Unit').filter(g => g.ID_Guru === u.ID_Guru && g.Tahun === ctx.sesi && g.Status === 'Menunggu')
      .forEach(g => update_('Guru_Unit', g, { Status: 'Aktif' }));
  } else if (op === 'nyahaktif') {
    update_('Users', u, { Status: 'Tidak Aktif' });
    killTokens_(rows_('Token').filter(t => t.ID_Guru === u.ID_Guru));
  } else if (op === 'super') {
    update_('Users', u, { Peranan_Global: 'SuperAdmin' });
  } else if (op === 'guru') {
    update_('Users', u, { Peranan_Global: 'Guru' });
  } else if (op === 'pengesah') {
    const jaw = clean_(d.jawatan, 60);
    if (jaw.length < 2) throw fail_('Pilih jawatan pengesah.');
    if (u.Status !== 'Aktif') throw fail_('Akaun ini belum aktif.');
    ensureSheet_(ss_(), 'Users'); delete T_.Users;
    update_('Users', rows_('Users').find(x => x.ID_Guru === u.ID_Guru), { Pengesah: jaw });
  } else if (op === 'bukanPengesah') {
    ensureSheet_(ss_(), 'Users'); delete T_.Users;
    update_('Users', rows_('Users').find(x => x.ID_Guru === u.ID_Guru), { Pengesah: '' });
  } else if (op === 'reset') {
    if (u.ID_Guru === ctx.id) throw fail_('Gunakan Tukar kata laluan di halaman Profil untuk akaun sendiri.');
    const salt = Utilities.getUuid();
    update_('Users', u, { Salt: salt, Hash_Kata_Laluan: hashPw_(CONFIG.KATA_LALUAN_LALAI, salt), Tukar_Kata_Laluan: 'Ya', Cubaan_Gagal: 0, Kunci_Hingga: '' });
    killTokens_(rows_('Token').filter(t => t.ID_Guru === u.ID_Guru));
  } else {
    throw fail_('Tindakan tidak sah.');
  }
  log_(ctx.id, 'Pengguna: ' + op, 'Users', u.ID_Guru, '', u.Emel);
  return { ok: true };
}

function a_saveKategori(d, ctx) {
  needSuper_(ctx);
  const nama = clean_(d.nama, 60);
  if (!nama) throw fail_('Isi nama kategori.');
  const warna = /^#[0-9a-fA-F]{6}$/.test(d.warna || '') ? d.warna : '';
  const all = rows_('Kategori');
  if (all.some(k => k.Nama_Kategori.toLowerCase() === nama.toLowerCase() && k.ID_Kategori !== d.id)) throw fail_('Nama kategori sudah wujud.');
  const patch = { Nama_Kategori: nama, Wajib: d.wajib === false ? 'Tidak' : 'Ya', Status: d.status === 'Tidak Aktif' ? 'Tidak Aktif' : 'Aktif', Warna: warna, Rekod_Perjumpaan: d.perjumpaan === false ? 'Tidak' : 'Ya' };
  ensureSheet_(ss_(), 'Kategori');
  delete T_.Kategori;
  if (d.id) {
    const k = all.find(x => x.ID_Kategori === d.id);
    if (!k) throw fail_('Kategori tidak dijumpai.');
    update_('Kategori', k, patch);
  } else {
    const max = all.reduce((m, k) => Math.max(m, Number(k.Susunan || 0)), 0);
    insert_('Kategori', Object.assign({ ID_Kategori: uid_('KT'), Susunan: max + 1 }, patch));
  }
  log_(ctx.id, 'Simpan kategori', 'Kategori', d.id || '', '', nama);
  return { ok: true };
}

function a_saveUnit(d, ctx) {
  needSuper_(ctx);
  const nama = clean_(d.nama, 80);
  if (!nama) throw fail_('Isi nama unit.');
  if (!rows_('Kategori').some(k => k.ID_Kategori === d.idKategori)) throw fail_('Pilih kategori.');
  const all = rows_('Unit');
  if (all.some(u => u.Nama_Unit.toLowerCase() === nama.toLowerCase() && u.ID_Unit !== d.id)) throw fail_('Nama unit sudah wujud.');
  const patch = { Nama_Unit: nama, ID_Kategori: d.idKategori, Status: d.status === 'Tidak Aktif' ? 'Tidak Aktif' : 'Aktif' };
  if (d.id) {
    const u = all.find(x => x.ID_Unit === d.id);
    if (!u) throw fail_('Unit tidak dijumpai.');
    if (u.ID_Kategori !== d.idKategori && rows_('Enrolmen').some(e => e.ID_Unit === u.ID_Unit && e.Tahun === ctx.sesi && e.Status === 'Aktif')) {
      throw fail_('Unit ini sudah ada ahli. Kategori tidak boleh ditukar pada sesi ini.');
    }
    update_('Unit', u, patch);
  } else {
    insert_('Unit', Object.assign({ ID_Unit: uid_('U'), ID_Folder_Drive: '' }, patch));
  }
  log_(ctx.id, 'Simpan unit', 'Unit', d.id || '', '', nama);
  return { ok: true };
}

/* =========================================================================
 * DATA MURID
 * ========================================================================= */

function a_murid(d, ctx) {
  needSuper_(ctx);
  return { murid: rows_('Murid').map(m => [m.ID_Murid, m.Nama_Murid, m.Kelas, m.Jantina, m.Status, m.ID_APDM ? 1 : 0]) };
}

function normJantina_(v) {
  const s = String(v || '').trim().toUpperCase();
  if (/^(L|LELAKI|M|MALE)$/.test(s)) return 'L';
  if (/^(P|PEREMPUAN|F|FEMALE)$/.test(s)) return 'P';
  return '';
}

function a_saveMurid(d, ctx) {
  needSuper_(ctx);
  const nama = clean_(d.nama, 150);
  if (!nama) throw fail_('Isi nama murid.');
  const patch = { Nama_Murid: nama, Kelas: clean_(d.kelas, 40), Jantina: normJantina_(d.jantina), Status: ['Aktif', 'Berpindah', 'Tamat'].indexOf(d.status) >= 0 ? d.status : 'Aktif', Tarikh_Kemaskini: now_() };
  if (d.id) {
    const m = rows_('Murid').find(x => x.ID_Murid === d.id);
    if (!m) throw fail_('Murid tidak dijumpai.');
    update_('Murid', m, patch);
  } else {
    ensureSheet_(ss_(), 'Murid'); delete T_.Murid;
  insert_('Murid', Object.assign({ ID_Murid: uid_('M'), ID_APDM: '', Tarikh_Daftar: now_() }, patch));
  }
  return { ok: true };
}

function a_importMurid(d, ctx) {
  needSuper_(ctx);
  return importCore_(d.rows || [], !!d.tandaTiada, ctx.id);
}

function importCore_(list, tandaTiada, who) {
  ensureSheet_(ss_(), 'Murid'); delete T_.Murid;
  const t = table_('Murid');
  const now = now_();
  const byKey = {};
  t.rows.forEach(r => { if (r.ID_APDM) byKey[r.ID_APDM] = r; });
  let baru = 0, kemas = 0, sama = 0, tiada = 0;
  const seen = {}, add = [];
  list.slice(0, 6000).forEach(x => {
    const k = clean_(x.k, 64), nama = clean_(x.nama, 150), kelas = clean_(x.kelas, 40), jan = normJantina_(x.jantina);
    if (!k || !nama || seen[k]) return;
    seen[k] = 1;
    const r = byKey[k];
    if (r) {
      if (r.Nama_Murid !== nama || r.Kelas !== kelas || r.Jantina !== jan || r.Status !== 'Aktif') {
        r.Nama_Murid = nama; r.Kelas = kelas; r.Jantina = jan; r.Status = 'Aktif'; r.Tarikh_Kemaskini = now;
        kemas++;
      } else sama++;
    } else {
      add.push({ ID_Murid: uid_('M'), ID_APDM: k, Nama_Murid: nama, Kelas: kelas, Jantina: jan, Status: 'Aktif', Tarikh_Kemaskini: now, Tarikh_Daftar: now });
      baru++;
    }
  });
  const pindah = {};
  if (tandaTiada) {
    t.rows.forEach(r => {
      if (r.ID_APDM && !seen[r.ID_APDM] && r.Status === 'Aktif') { r.Status = 'Berpindah'; r.Tarikh_Kemaskini = now; tiada++; pindah[r.ID_Murid] = 1; }
    });
  }
  if (kemas || tiada) writeAll_('Murid');
  // Sama seperti a_bulkMurid: murid berpindah dikeluarkan daripada unit sesi semasa
  if (tiada) {
    const sesi = sesi_();
    const te = table_('Enrolmen');
    let keluar = 0;
    te.rows.forEach(e => { if (pindah[e.ID_Murid] && e.Tahun === sesi && e.Status === 'Aktif') { e.Status = 'Keluar'; keluar++; } });
    if (keluar) writeAll_('Enrolmen');
  }
  insertMany_('Murid', add);
  log_(who, 'Import murid', 'Murid', '', '', 'baru ' + baru + ', dikemas kini ' + kemas + ', ditanda berpindah ' + tiada);
  return { baru: baru, kemas: kemas, sama: sama, tiada: tiada };
}

/* =========================================================================
 * JADUAL, FORMAT LAPORAN, TETAPAN
 * ========================================================================= */

function a_saveJadual(d, ctx) {
  needSuper_(ctx);
  const sesi = ctx.sesi;
  const items = (d.items || []).map(x => ({
    bil: Number(x.bil), tarikh: isYmd_(x.tarikh) ? x.tarikh : '',
    status: x.status === 'Dibatalkan' ? 'Dibatalkan' : 'Dirancang', catatan: clean_(x.catatan, 200)
  })).sort((a, b) => a.bil - b.bil);
  if (!items.length) throw fail_('Jadual mesti mempunyai sekurang-kurangnya satu perjumpaan.');
  if (items.length > CONFIG.MAKS_PERJUMPAAN) throw fail_('Maksimum ' + CONFIG.MAKS_PERJUMPAAN + ' perjumpaan setahun.');
  items.forEach((x, i) => { if (x.bil !== i + 1) throw fail_('Nombor perjumpaan mesti berturutan bermula dari 1.'); });
  const cur = jadual_(sesi);
  const byBil = {};
  cur.forEach(j => { byBil[Number(j.Bil)] = j; });
  const removed = cur.filter(j => Number(j.Bil) > items.length);
  removed.forEach(j => {
    const col = 'P' + j.Bil;
    const ada = rows_('Kehadiran').some(k => k.Tahun === sesi && k[col]) || rows_('Laporan').some(l => l.ID_Perjumpaan === j.ID_Perjumpaan);
    if (ada) throw fail_('Perjumpaan ' + j.Bil + ' sudah mempunyai data. Tandakan sebagai Dibatalkan, jangan dibuang.');
  });
  items.forEach(x => {
    const j = byBil[x.bil];
    if (j) {
      if (j.Tarikh !== x.tarikh || j.Status !== x.status || j.Catatan !== x.catatan) update_('Jadual_Perjumpaan', j, { Tarikh: x.tarikh, Status: x.status, Catatan: x.catatan });
    } else {
      insert_('Jadual_Perjumpaan', { ID_Perjumpaan: uid_('J'), Tahun: sesi, Bil: x.bil, Tarikh: x.tarikh, Status: x.status, Catatan: x.catatan });
    }
  });
  removeMany_('Jadual_Perjumpaan', removed);
  log_(ctx.id, 'Simpan jadual perjumpaan', 'Jadual_Perjumpaan', sesi, '', items.length + ' perjumpaan');
  return { jadual: jadualOut_(sesi) };
}

function a_saveFormat(d, ctx) {
  needSuper_(ctx);
  const seen = {};
  const medan = (d.medan || []).slice(0, 25).map((f, i) => {
    const label = clean_(f.label, 80);
    if (!label) throw fail_('Setiap medan mesti ada label.');
    const jenis = JENIS_MEDAN.indexOf(f.jenis) >= 0 ? f.jenis : 'teks';
    let id = clean_(f.id, 30).replace(/[^a-zA-Z0-9_]/g, '') || ('m' + (i + 1));
    while (seen[id]) id += '_';
    seen[id] = 1;
    const o = { id: id, label: label, jenis: jenis, wajib: !!f.wajib };
    if (jenis === 'pilihan') {
      o.pilihan = (f.pilihan || []).map(p => clean_(p, 60)).filter(String).slice(0, 20);
      if (!o.pilihan.length) throw fail_('Medan "' + label + '" memerlukan sekurang-kurangnya satu pilihan.');
    }
    return o;
  });
  if (!medan.length) throw fail_('Format mesti mempunyai sekurang-kurangnya satu medan.');
  const t = table_('Format_Laporan');
  const max = t.rows.reduce((m, f) => Math.max(m, Number(f.Versi || 0)), 0);
  if (t.rows.some(f => f.Aktif === 'Ya')) {
    t.rows.forEach(f => { f.Aktif = 'Tidak'; });
    writeAll_('Format_Laporan');
  }
  insert_('Format_Laporan', { Versi: max + 1, Tarikh_Cipta: now_(), Medan_JSON: JSON.stringify(medan), Aktif: 'Ya' });
  log_(ctx.id, 'Simpan format laporan v' + (max + 1), 'Format_Laporan', String(max + 1), '', medan.length + ' medan');
  return { versi: String(max + 1) };
}

function a_saveTetapan(d, ctx) {
  needSuper_(ctx);
  const kv = d.kv || {};
  const allowed = {
    NAMA_SEKOLAH: v => clean_(v, 120),
    ALAMAT_SEKOLAH: v => clean_(v, 250),
    ID_FAIL_LOGO: v => clean_(v, 80),
    NAMA_PENGESAH: v => clean_(v, 120),
    JAWATAN_PENGESAH: v => clean_(v, 120),
    TEMPOH_KUNCI_HARI: v => String(Math.max(0, Math.min(60, parseInt(v, 10) || 0))),
    HARI_PERINGATAN: v => String(Math.max(0, Math.min(30, parseInt(v, 10) || 0))),
    LAYANAN_B: v => (LAYANAN_B.indexOf(v) >= 0 ? v : 'Dikecualikan'),
    AWAM_AKTIF: v => (v === 'Tidak' ? 'Tidak' : 'Ya'),
    AWAM_PENCAPAIAN: v => (['Penuh', 'Singkat', 'Sembunyi'].indexOf(v) >= 0 ? v : 'Singkat'),
    AWAM_GALERI: v => (v === 'Tidak' ? 'Tidak' : 'Ya'),
    AWAM_SEMAKAN: v => (v === 'Tidak' ? 'Tidak' : 'Ya'),
    EMEL_LAPORAN: v => (v === 'Tidak' ? 'Tidak' : 'Ya')
  };
  Object.keys(kv).forEach(k => { if (allowed[k]) setTetapan_(k, allowed[k](kv[k])); });
  log_(ctx.id, 'Kemas kini tetapan', 'Tetapan', '', '', Object.keys(kv).join(', '));
  return { tetapan: publicTetapan_(tetapan_()) };
}

function a_log(d, ctx) {
  needSuper_(ctx);
  const users = mapBy_(rows_('Users'), 'ID_Guru');
  const all = rows_('Log');
  const n = Math.min(Number(d.had || 300), 1000);
  return {
    log: all.slice(-n).reverse().map(l => ({
      masa: l.Masa, oleh: (users[l.ID_Guru] || {}).Nama || l.ID_Guru, tindakan: l.Tindakan, sheet: l.Sheet, rekod: l.ID_Rekod, lama: l.Nilai_Lama, baru: l.Nilai_Baharu
    }))
  };
}

/* =========================================================================
 * PAJSK & TUTUP SESI
 * ========================================================================= */

function a_pajsk(d, ctx) {
  // Super Admin: seluruh sekolah atau satu kelas. Guru kelas: kelas sendiri sahaja.
  const kelas = ctx.super ? clean_(d.kelas, 40) : String(ctx.user.Kelas_Guru || '');
  if (!ctx.super && !kelas) throw fail_('Tetapkan kelas anda di halaman Profil untuk melihat data PAJSK.', 'FORBIDDEN');
  const sesi = ctx.sesi, mode = tetapan_().LAYANAN_B || 'Dikecualikan', today = today_();
  const kat = rows_('Kategori').filter(k => k.Status !== 'Tidak Aktif').sort((a, b) => Number(a.Susunan) - Number(b.Susunan));
  const units = mapBy_(rows_('Unit'), 'ID_Unit');
  const heldBils = jadual_(sesi).filter(j => j.Status !== 'Dibatalkan' && j.Tarikh && j.Tarikh <= today).map(j => Number(j.Bil));
  const enrol = {};
  rows_('Enrolmen').forEach(e => { if (e.Tahun === sesi && e.Status === 'Aktif') enrol[e.ID_Murid + '|' + e.ID_Kategori] = e; });
  const kh = {};
  rows_('Kehadiran').forEach(k => { if (k.Tahun === sesi) kh[k.ID_Unit + '|' + k.ID_Murid] = k; });
  const jaw = {};
  rows_('Carta_Organisasi').forEach(c => { if (c.Tahun === sesi && c.Jenis === 'Murid') jaw[c.ID_Unit + '|' + c.ID_Rujukan] = c.Jawatan; });
  const best = {};
  rows_('Pencapaian').forEach(p => {
    if (p.Tahun !== sesi) return;
    const key = p.ID_Unit + '|' + p.ID_Murid;
    const r = PERINGKAT.indexOf(p.Peringkat);
    if (!best[key] || r > best[key].r) best[key] = { r: r, t: p.Peringkat + (p.Pencapaian ? ' (' + p.Pencapaian + ')' : '') };
  });
  const rows = rows_('Murid').filter(m => layak_(m) && (!kelas || m.Kelas === kelas))
    .sort((a, b) => String(a.Kelas).localeCompare(String(b.Kelas), 'ms', { numeric: true }) || a.Nama_Murid.localeCompare(b.Nama_Murid))
    .map(m => {
      const cols = kat.map(k => {
        const e = enrol[m.ID_Murid + '|' + k.ID_Kategori];
        if (!e) return ['', '', '', ''];
        const key = e.ID_Unit + '|' + m.ID_Murid;
        const r = kh[key];
        const p = r ? pct_(r, heldBils, mode) : { h: 0, n: 0 };
        return [(units[e.ID_Unit] || {}).Nama_Unit || '', jaw[key] || 'Ahli', p.n ? Math.round(p.h / p.n * 100) : '', best[key] ? best[key].t : ''];
      });
      return [m.Nama_Murid, m.Kelas, m.Jantina].concat([].concat.apply([], cols));
    });
  return { kategori: kat.map(k => k.Nama_Kategori), rows: rows, diadakan: heldBils.length, kelas: kelas };
}


/* ---------- PAJSK MOEIS: pemetaan dan data mengikut templat SPPB ---------- */
function tabKategori_(k) {
  if (k.Tab_PAJSK) return k.Tab_PAJSK === 'Tiada' ? '' : k.Tab_PAJSK;
  const n = String(k.Nama_Kategori || '').toLowerCase();
  if (/rumah/.test(n)) return '';
  if (/unif/.test(n)) return 'Beruniform';
  if (/sukan|permainan/.test(n)) return 'Sukan';
  if (/kelab|persatuan/.test(n)) return 'Kelab';
  return '';
}
const SPPB_SINONIM_ = {
  PBSM: 'BULAN SABIT MERAH', BSMM: 'BULAN SABIT MERAH', 'BULAN SABIT': 'BULAN SABIT MERAH', TKRS: 'KADET REMAJA SEKOLAH', KRS: 'KADET REMAJA SEKOLAH',
  'PUTERI ISLAM': 'PUTERI ISLAM', PPIM: 'PUTERI ISLAM', APM: 'ANGKATAN PERTAHANAN AWAM', 'PERTAHANAN AWAM': 'ANGKATAN PERTAHANAN AWAM', 'PANDU PUTERI': 'PANDU PUTERI',
  STEM: 'SAINS TEKNOLOGI', SAINS: 'SAINS TEKNOLOGI', MATEMATIK: 'SAINS TEKNOLOGI', SOLAR: 'SAINS TEKNOLOGI', SPBT: 'PINJAMAN BUKU TEKS', 'BUKU TEKS': 'PINJAMAN BUKU TEKS',
  MUZIK: 'SENI MUZIK', 'DOKTOR MUDA': 'DOKTOR MUDA', 'AGAMA ISLAM': 'AGAMA ISLAM', PAI: 'AGAMA ISLAM', 'BAHASA ARAB': 'BAHASA ARAB', 'BAHASA INGGERIS': 'BAHASA INGGERIS', ENGLISH: 'BAHASA INGGERIS',
  'BAHASA MELAYU': 'BAHASA MELAYU', 'RUKUN NEGARA': 'RUKUN NEGARA', 'KITAR SEMULA': 'KITAR SEMULA', 'PENCEGAHAN JENAYAH': 'PENCEGAHAN JENAYAH', KERJAYA: 'KERJAYA', TVPSS: 'TVPSS',
  'OLAHRAGA': 'OLAHRAGA', 'BALAPAN': 'OLAHRAGA', TAKRAW: 'SEPAK TAKRAW', 'SEPAK RAGA': 'SEPAK TAKRAW', PANAH: 'MEMANAH', 'SILAT': 'SILAT'
};
/** Cadangan nama aktiviti SPPB bagi unit sistem. */
function cadanganSPPB_(namaUnit, tab) {
  const list = (SPPB.aktiviti[tab] || []).concat(tab ? [] : []);
  if (!list.length) return '';
  const norm = x => ' ' + String(x || '').toUpperCase().replace(/\(SR\)|\(SM\)|\(T6\)/g, ' ').replace(/&/g, ' DAN ').replace(/[^A-Z0-9]+/g, ' ').replace(/\s+/g, ' ').trim() + ' ';
  let u = norm(namaUnit);
  Object.keys(SPPB_SINONIM_).forEach(k => { if (u.indexOf(' ' + k + ' ') >= 0) u += SPPB_SINONIM_[k] + ' '; });
  const buang = /^(KELAB|PERSATUAN|MALAYSIA|DAN|UNIT|PASUKAN|SR|SEKOLAH)$/;
  const tok = u.trim().split(' ').filter(w => w && !buang.test(w));
  if (!tok.length) return '';
  let best = '', skor = 0;
  list.forEach(o => {
    const on = norm(o);
    const n = tok.filter(w => on.indexOf(' ' + w + ' ') >= 0).length;
    const sk = n / tok.length;
    if (sk > skor) { skor = sk; best = o; }
  });
  return skor >= 0.5 ? best : '';
}
function namaSPPB_(u, tab) {
  return u.Nama_SPPB || cadanganSPPB_(u.Nama_Unit, tab) || String(u.Nama_Unit || '').toUpperCase() + ' (SR)';
}
function pajskTetapan_() {
  const t = tetapan_(), o = {};
  Object.keys(PAJSK_LALAI).forEach(k => { o[k] = t[k] !== undefined && t[k] !== '' ? t[k] : PAJSK_LALAI[k]; });
  if (t.PAJSK_KHIDMAT !== undefined) o.PAJSK_KHIDMAT = t.PAJSK_KHIDMAT;
  o.komitmen = parse_(o.PAJSK_KOMITMEN, []).filter(x => SPPB.komitmen.indexOf(x) >= 0).slice(0, 4);
  return o;
}

function a_pajskPeta(d, ctx) {
  needSuper_(ctx);
  const kat = rows_('Kategori').filter(k => k.Status !== 'Tidak Aktif').sort((a, b) => Number(a.Susunan) - Number(b.Susunan));
  const tabOf = {};
  kat.forEach(k => { tabOf[k.ID_Kategori] = tabKategori_(k); });
  const units = rows_('Unit').filter(u => u.Status !== 'Tidak Aktif').sort((a, b) => a.Nama_Unit.localeCompare(b.Nama_Unit)).map(u => ({
    id: u.ID_Unit, nama: u.Nama_Unit, idKategori: u.ID_Kategori, sppb: u.Nama_SPPB || '', cadangan: cadanganSPPB_(u.Nama_Unit, tabOf[u.ID_Kategori])
  }));
  const p = pajskTetapan_();
  return {
    kategori: kat.map(k => ({ id: k.ID_Kategori, nama: k.Nama_Kategori, tab: tabOf[k.ID_Kategori], disimpan: !!k.Tab_PAJSK })),
    unit: units, sppb: SPPB,
    tetapan: { komitmen: p.komitmen, khidmat: p.PAJSK_KHIDMAT, jawatan: p.PAJSK_JAWATAN, aktifPeratus: Number(p.PAJSK_AKTIF_PERATUS) || 75 }
  };
}

function a_savePajskPeta(d, ctx) {
  needSuper_(ctx);
  ensureSheet_(ss_(), 'Kategori'); delete T_.Kategori;
  ensureSheet_(ss_(), 'Unit'); delete T_.Unit;
  const tabs = ['Kelab', 'Beruniform', 'Sukan', 'Tiada'];
  const tk = table_('Kategori');
  let n = 0;
  tk.rows.forEach(k => { const v = (d.kategori || {})[k.ID_Kategori]; if (v !== undefined && tabs.indexOf(v) >= 0 && k.Tab_PAJSK !== v) { k.Tab_PAJSK = v; n++; } });
  if (n) writeAll_('Kategori');
  const tu = table_('Unit');
  let m = 0;
  tu.rows.forEach(u => { const v = (d.unit || {})[u.ID_Unit]; if (v !== undefined) { const c = clean_(v, 150).toUpperCase(); if (u.Nama_SPPB !== c) { u.Nama_SPPB = c; m++; } } });
  if (m) writeAll_('Unit');
  const t = d.tetapan || {};
  if (t.komitmen) setTetapan_('PAJSK_KOMITMEN', JSON.stringify((t.komitmen || []).filter(x => SPPB.komitmen.indexOf(x) >= 0).slice(0, 4)));
  if (t.khidmat !== undefined) setTetapan_('PAJSK_KHIDMAT', SPPB.khidmat.indexOf(t.khidmat) >= 0 ? t.khidmat : '');
  if (t.jawatan) setTetapan_('PAJSK_JAWATAN', ['AHLI BIASA', 'AHLI AKTIF', 'IKUT_HADIR'].indexOf(t.jawatan) >= 0 ? t.jawatan : 'IKUT_HADIR');
  if (t.aktifPeratus !== undefined) setTetapan_('PAJSK_AKTIF_PERATUS', String(Math.max(0, Math.min(100, parseInt(t.aktifPeratus, 10) || 75))));
  log_(ctx.id, 'Pemetaan PAJSK MOEIS', 'Tetapan', '', '', n + ' kategori, ' + m + ' unit');
  return a_pajskPeta(d, ctx);
}

/** Peta kunci murid (K + cincangan No. KP) kepada No. KP sebenar daripada tab DATA MURID. No. KP tidak disimpan dalam sistem. */
function petaKP_(perlu) {
  const out = {};
  const tM = cariTab_(CONFIG.SHEET_DATA_MURID);
  if (!tM || !Object.keys(perlu).length) return out;
  const v = tM.getDataRange().getDisplayValues();
  const h = headerRow_(v, /KELAS|TAHUN|TINGKATAN|PENGENALAN|K\/?P/);
  if (h < 0) return out;
  const H = v[h].map(c => String(c).toUpperCase().replace(/\s+/g, ' ').trim());
  const cKP = H.findIndex(c => /PENGENALAN|NO\.?\s*K\.?\s*\/?\s*P|NOKP|MYKID|MYKAD|NO\.?\s*IC|^IC$|SIJIL/.test(c));
  if (cKP < 0) return out;
  for (let i = h + 1; i < v.length; i++) {
    let kp = String(v[i][cKP] || '').replace(/\D/g, '');
    if (kp.length === 11) kp = '0' + kp;
    if (kp.length < 6) continue;
    const k = 'K' + sha_(kp).slice(0, 20);
    if (perlu[k]) out[k] = kp;
  }
  return out;
}
const PERINGKAT_SPPB_ = { Sekolah: 'Sekolah', Zon: 'Zon/Daerah', Daerah: 'Zon/Daerah', Negeri: 'Negeri', Kebangsaan: 'Kebangsaan', Antarabangsa: 'Antarabangsa' };
function tahapSPPB_(teks) {
  const t = String(teks || '').toLowerCase();
  if (/naib\s*johan|kedua|ke-?\s*2\b|tempat\s*2|perak/.test(t)) return 'NAIB JOHAN';
  if (/johan|juara|pertama|ke-?\s*1\b|tempat\s*1|emas/.test(t)) return 'JOHAN';
  if (/ketiga|ke-?\s*3\b|tempat\s*3|gangsa/.test(t)) return 'KETIGA';
  if (/keempat|ke-?\s*4\b|tempat\s*4/.test(t)) return 'KEEMPAT';
  if (/kelima|ke-?\s*5\b|tempat\s*5/.test(t)) return 'KELIMA';
  return '';
}
function jawatanSPPB_(jaw, tab, pct, p) {
  const opt = SPPB.jawatan[tab] || SPPB.jawatan.Kelab;
  const j = String(jaw || '').toUpperCase().replace(/\s+/g, ' ').trim();
  if (j) {
    if (opt.indexOf(j) >= 0) return j;
    const ada = x => (opt.indexOf(x) >= 0 ? x : '');
    if (/^(PENOLONG|NAIB|TIMBALAN)\s+KAPTEN/.test(j)) return ada('PENOLONG KAPTEN') || 'NAIB PENGERUSI';
    if (/^KAPTEN/.test(j)) return ada('KAPTEN PASUKAN') || 'PENGERUSI';
    if (/^(PENOLONG|NAIB|TIMBALAN)\s+(KETUA|PENGERUSI|PRESIDEN)/.test(j)) return 'NAIB PENGERUSI';
    if (/^(PENOLONG|NAIB)\s+SETIAUSAHA/.test(j)) return 'NAIB SETIAUSAHA';
    if (/^(PENOLONG|NAIB)\s+BENDAHARI/.test(j)) return 'NAIB BENDAHARI';
    if (/^(KETUA|PENGERUSI|PRESIDEN)\b/.test(j) && !/GURU/.test(j)) return 'PENGERUSI';
    if (/SETIAUSAHA/.test(j)) return 'SETIAUSAHA';
    if (/BENDAHARI/.test(j)) return 'BENDAHARI';
    if (/^AHLI( BIASA)?$/.test(j)) return 'AHLI BIASA';
    if (/AHLI AKTIF/.test(j)) return 'AHLI AKTIF';
    return 'AHLI JAWATANKUASA';
  }
  if (p.PAJSK_JAWATAN === 'AHLI AKTIF' || p.PAJSK_JAWATAN === 'AHLI BIASA') return p.PAJSK_JAWATAN;
  return pct != null && pct >= (Number(p.PAJSK_AKTIF_PERATUS) || 75) ? 'AHLI AKTIF' : 'AHLI BIASA';
}

function a_pajskMoeis(d, ctx) {
  const kelas = ctx.super ? clean_(d.kelas, 40) : String(ctx.user.Kelas_Guru || '');
  if (!ctx.super && !kelas) throw fail_('Tetapkan kelas anda di halaman Profil untuk memuat turun templat PAJSK.', 'FORBIDDEN');
  const sesi = ctx.sesi, mode = tetapan_().LAYANAN_B || 'Dikecualikan', today = today_();
  const p = pajskTetapan_();
  const kat = rows_('Kategori').filter(k => k.Status !== 'Tidak Aktif').sort((a, b) => Number(a.Susunan) - Number(b.Susunan));
  const tabOf = {};
  kat.forEach(k => { tabOf[k.ID_Kategori] = tabKategori_(k); });
  const units = mapBy_(rows_('Unit'), 'ID_Unit');
  const heldBils = jadual_(sesi).filter(j => j.Status !== 'Dibatalkan' && j.Tarikh && j.Tarikh <= today).map(j => Number(j.Bil));
  const kh = {};
  rows_('Kehadiran').forEach(k => { if (k.Tahun === sesi) kh[k.ID_Unit + '|' + k.ID_Murid] = k; });
  const jaw = {};
  rows_('Carta_Organisasi').forEach(c => { if (c.Tahun === sesi && c.Jenis === 'Murid') jaw[c.ID_Unit + '|' + c.ID_Rujukan] = c.Jawatan; });
  const capai = {};
  rows_('Pencapaian').forEach(x => { if (x.Tahun === sesi) (capai[x.ID_Unit + '|' + x.ID_Murid] = capai[x.ID_Unit + '|' + x.ID_Murid] || []).push(x); });
  const enrol = {};
  rows_('Enrolmen').forEach(e => { if (e.Tahun === sesi && e.Status === 'Aktif') (enrol[e.ID_Murid] = enrol[e.ID_Murid] || []).push(e); });
  const ekstra = {};
  rows_('Ekstra_PAJSK').forEach(x => { if (x.Tahun === sesi) (ekstra[x.ID_Murid] = ekstra[x.ID_Murid] || []).push(x); });
  const murid = rows_('Murid').filter(m => layak_(m) && (!kelas || m.Kelas === kelas))
    .sort((a, b) => String(a.Kelas).localeCompare(String(b.Kelas), 'ms', { numeric: true }) || a.Nama_Murid.localeCompare(b.Nama_Murid));
  const perlu = {};
  murid.forEach(m => { if (/^K/.test(m.ID_APDM || '')) perlu[m.ID_APDM] = 1; });
  const kp = petaKP_(perlu);
  const katSusun = {};
  kat.forEach((k, i) => { katSusun[k.ID_Kategori] = i; });
  const rankP = SPPB.peringkat.slice().reverse(); // Sekolah .. Antarabangsa
  const out = { Kelab: [], Beruniform: [], Sukan: [], Ekstra: [], Tambahan: [] };
  const tiadaUnit = [], tanpaKP = [], tekaan = {};
  murid.forEach(m => {
    const id = kp[m.ID_APDM] || '';
    if (!id) tanpaKP.push(m.Nama_Murid);
    out.Ekstra.push(barisEkstra_(id, m.Nama_Murid, ekstra[m.ID_Murid]));
    const ada = {};
    (enrol[m.ID_Murid] || []).sort((a, b) => (katSusun[a.ID_Kategori] || 0) - (katSusun[b.ID_Kategori] || 0)).forEach(e => {
      const tab = tabOf[e.ID_Kategori];
      if (!tab) return;
      const u = units[e.ID_Unit] || {};
      const akt = namaSPPB_(u, tab);
      if (!u.Nama_SPPB) tekaan[u.Nama_Unit] = akt;
      const key = e.ID_Unit + '|' + m.ID_Murid;
      const r = kh[key];
      let hadir = 0;
      if (r) heldBils.forEach(b => { const v = r['P' + b]; if (v === 'H' || (v === 'B' && mode === 'Dikira hadir')) hadir++; });
      const pc = r ? pct_(r, heldBils, mode) : { h: 0, n: 0 };
      const pct = pc.n ? Math.round(pc.h / pc.n * 100) : null;
      // Pelibatan: peringkat tertinggi dan bilangan penyertaan pada peringkat itu (maksimum 3)
      const rec = capai[key] || [];
      let pelP = '', pelN = '', capP = '', capT = '', capSkor = -1;
      if (rec.length) {
        const pk = rec.map(x => PERINGKAT_SPPB_[x.Peringkat] || 'Sekolah');
        pelP = pk.reduce((a, b) => (rankP.indexOf(b) > rankP.indexOf(a) ? b : a), pk[0]);
        pelN = String(Math.min(3, pk.filter(x => x === pelP).length));
        rec.forEach((x, i) => {
          const t = tahapSPPB_(x.Pencapaian);
          if (!t) return;
          let pr = pk[i] === 'Antarabangsa' ? 'Kebangsaan' : pk[i];
          const sk = rankP.indexOf(pr) * 10 + (5 - SPPB.tahap.indexOf(t));
          if (sk > capSkor) { capSkor = sk; capP = pr; capT = t; }
        });
      }
      // Khidmat sumbangan: rekod terkini murid bagi unit ini; jika tiada, tetapan lalai sekolah
      const kd = rec.filter(x => x.Khidmat).sort((a, b) => String(b.Tarikh).localeCompare(String(a.Tarikh)));
      const khidmat = kd.length ? kd[0].Khidmat : (p.PAJSK_KHIDMAT || '');
      if (ada[tab]) {
        out.Tambahan.push([id, m.Nama_Murid, SPPB.tab[tab], akt, pelP, pelN, capP, capT]);
        return;
      }
      ada[tab] = 1;
      const nK = pct == null ? 0 : pct >= 80 ? 4 : pct >= 60 ? 3 : pct >= 40 ? 2 : pct > 0 ? 1 : 0;
      const kom = p.komitmen.slice(0, nK);
      while (kom.length < 4) kom.push('');
      out[tab].push([id, m.Nama_Murid, 'Ya', akt, jawatanSPPB_(jaw[key], tab, pct, p), pelP, pelN, capP, capT].concat(kom, [khidmat, Math.min(12, hadir)]));
    });
    const kurang = ['Kelab', 'Beruniform', 'Sukan'].filter(t => !ada[t] && kat.some(k => tabOf[k.ID_Kategori] === t));
    if (kurang.length) tiadaUnit.push({ nama: m.Nama_Murid, kelas: m.Kelas, tab: kurang });
  });
  return {
    kelas: kelas, sesi: sesi, diadakan: heldBils.length, sekolah: tetapan_().NAMA_SEKOLAH || CONFIG.NAMA_SEKOLAH,
    sheets: out, tiadaUnit: tiadaUnit, tanpaKP: tanpaKP, adaTabDataMurid: !!cariTab_(CONFIG.SHEET_DATA_MURID),
    tekaan: Object.keys(tekaan).map(n => ({ unit: n, sppb: tekaan[n] }))
  };
}

function a_tutupSesi(d, ctx) {
  needSuper_(ctx);
  const lama = ctx.sesi, baru = String(d.tahunBaru || '').trim();
  if (!/^\d{4}$/.test(baru) || Number(baru) <= Number(lama)) throw fail_('Tahun sesi baharu mesti selepas ' + lama + '.');
  if (String(d.sahkan || '') !== baru) throw fail_('Taip tahun sesi baharu untuk pengesahan.');
  // 1. Salinan arkib penuh sebelum sebarang perubahan
  let arkib = '';
  try {
    const f = DriveApp.getFileById(ss_().getId()).makeCopy('Arkib e-Kokurikulum sesi ' + lama + ' (' + now_() + ')', rootFolder_());
    arkib = f.getUrl ? f.getUrl() : f.getId();
  } catch (e) { if (!d.tanpaArkib) throw fail_('Salinan arkib gagal dibuat: ' + (e.message || e) + '. Sesi tidak ditutup.'); }
  // 2. Tahun 6 tamat, 3. murid lain naik tahun (4 UIAM -> 5 UIAM)
  let tamat = 0, naik = 0;
  const t = table_('Murid');
  t.rows.forEach(m => {
    if (m.Status !== 'Aktif') return;
    const k = String(m.Kelas || '').trim(), th = tahunKelas_(k);
    if (d.tamatTahun6 && th === 6) { m.Status = 'Tamat'; m.Tarikh_Kemaskini = now_(); tamat++; }
    else if (d.naikTahun && th >= 1 && th <= 5) { m.Kelas = k.replace(/^\s*\d/, String(th + 1)); m.Tarikh_Kemaskini = now_(); naik++; }
  });
  if (tamat || naik) writeAll_('Murid');
  // 4. Guru kelas dikosongkan kerana kelas bertukar
  let guruKelas = 0;
  const tu = table_('Users');
  if (tu.head.indexOf('Kelas_Guru') >= 0) {
    tu.rows.forEach(u => { if (u.Kelas_Guru) { u.Kelas_Guru = ''; guruKelas++; } });
    if (guruKelas) writeAll_('Users');
  }
  if (!jadual_(baru).length) {
    const rows = [];
    for (let i = 1; i <= 12; i++) rows.push({ ID_Perjumpaan: uid_('J'), Tahun: baru, Bil: i, Tarikh: '', Status: 'Dirancang', Catatan: '' });
    insertMany_('Jadual_Perjumpaan', rows);
  }
  let guru = 0, ahli = 0;
  if (d.salinGuru) {
    const ada = {};
    rows_('Guru_Unit').forEach(g => { if (g.Tahun === baru) ada[g.ID_Guru + '|' + g.ID_Unit] = 1; });
    const add = rows_('Guru_Unit').filter(g => g.Tahun === lama && g.Status === 'Aktif' && !ada[g.ID_Guru + '|' + g.ID_Unit])
      .map(g => ({ ID: uid_('GU'), Tahun: baru, ID_Guru: g.ID_Guru, ID_Unit: g.ID_Unit, Peranan_Unit: g.Peranan_Unit, Status: 'Aktif' }));
    insertMany_('Guru_Unit', add);
    guru = add.length;
  }
  if (d.salinAhli) {
    const murid = mapBy_(rows_('Murid'), 'ID_Murid');
    const ada = {};
    rows_('Enrolmen').forEach(e => { if (e.Tahun === baru) ada[e.ID_Murid + '|' + e.ID_Kategori] = 1; });
    const add = rows_('Enrolmen').filter(e => e.Tahun === lama && e.Status === 'Aktif' && murid[e.ID_Murid] && murid[e.ID_Murid].Status === 'Aktif' && !ada[e.ID_Murid + '|' + e.ID_Kategori])
      .map(e => ({ ID_Enrolmen: uid_('E'), Tahun: baru, ID_Murid: e.ID_Murid, ID_Unit: e.ID_Unit, ID_Kategori: e.ID_Kategori, Kelas: murid[e.ID_Murid].Kelas, Status: 'Aktif', Tarikh_Masuk: '' }));
    insertMany_('Enrolmen', add);
    ahli = add.length;
  }
  setTetapan_('SESI_AKTIF', baru);
  log_(ctx.id, 'Tutup sesi', 'Tetapan', 'SESI_AKTIF', lama, baru + ' (tamat ' + tamat + ', naik ' + naik + ', ahli ' + ahli + ')');
  return { sesi: baru, tamat: tamat, naik: naik, guru: guru, ahli: ahli, guruKelas: guruKelas, arkib: arkib };
}

/** Pratonton kesan tutup sesi sebelum dilaksanakan. */
function a_tutupPratonton(d, ctx) {
  needSuper_(ctx);
  const sesi = ctx.sesi, c = { 4: 0, 5: 0, 6: 0, lain: 0 };
  const aktif = {};
  rows_('Murid').forEach(m => {
    if (m.Status !== 'Aktif') return;
    const th = tahunKelas_(m.Kelas);
    if (c[th] !== undefined) c[th]++; else c.lain++;
    aktif[m.ID_Murid] = th;
  });
  const en = rows_('Enrolmen').filter(e => e.Tahun === sesi && e.Status === 'Aktif' && aktif[e.ID_Murid] !== undefined);
  const lapBelum = rows_('Laporan').filter(l => l.Tahun === sesi && l.Status === 'Dihantar').length;
  return {
    sesi: sesi, tahun4: c[4], tahun5: c[5], tahun6: c[6], lain: c.lain,
    ahliBawa: en.filter(e => aktif[e.ID_Murid] !== 6).length, ahliTamat: en.filter(e => aktif[e.ID_Murid] === 6).length,
    guru: rows_('Guru_Unit').filter(g => g.Tahun === sesi && g.Status === 'Aktif').length,
    guruKelas: rows_('Users').filter(u => u.Kelas_Guru).length, laporanBelumSah: lapBelum
  };
}

/* ---------- Data murid: tindakan pukal ---------- */
function a_bulkMurid(d, ctx) {
  needSuper_(ctx);
  const ids = {};
  (d.ids || []).forEach(x => { ids[String(x)] = 1; });
  if (!Object.keys(ids).length) throw fail_('Pilih sekurang-kurangnya seorang murid.');
  const status = d.status ? String(d.status) : '', kelas = d.kelas !== undefined ? clean_(d.kelas, 40) : null;
  if (status && ['Aktif', 'Berpindah', 'Tamat'].indexOf(status) < 0) throw fail_('Status tidak sah.');
  if (!status && !kelas) throw fail_('Pilih status atau kelas baharu.');
  const t = table_('Murid');
  let n = 0;
  t.rows.forEach(m => {
    if (!ids[m.ID_Murid]) return;
    if (status) m.Status = status;
    if (kelas) m.Kelas = kelas;
    m.Tarikh_Kemaskini = now_(); n++;
  });
  if (n) writeAll_('Murid');
  // Murid yang berpindah atau tamat dikeluarkan daripada unit sesi semasa
  let keluar = 0;
  if (status && status !== 'Aktif') {
    const te = table_('Enrolmen');
    te.rows.forEach(e => { if (ids[e.ID_Murid] && e.Tahun === ctx.sesi && e.Status === 'Aktif') { e.Status = 'Keluar'; keluar++; } });
    if (keluar) writeAll_('Enrolmen');
  }
  log_(ctx.id, 'Kemas kini pukal murid', 'Murid', '', '', n + ' murid' + (status ? ', status ' + status : '') + (kelas ? ', kelas ' + kelas : '') + (keluar ? ', ' + keluar + ' keluar unit' : ''));
  return { bil: n, keluar: keluar };
}

/* =========================================================================
 * SETUP, MENU, TRIGGER
 * ========================================================================= */

function setup() {
  const props = PropertiesService.getScriptProperties();
  let ss = null;
  const id = props.getProperty('SPREADSHEET_ID');
  if (id) { try { ss = SpreadsheetApp.openById(id); } catch (e) { ss = null; } }
  if (!ss) { try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) { ss = null; } }
  if (!ss) ss = SpreadsheetApp.create('Pangkalan Data e-Kokurikulum');
  props.setProperty('SPREADSHEET_ID', ss.getId());
  props.setProperty('SUPER_ADMIN_EMAIL', CONFIG.SUPER_ADMIN_EMAIL.toLowerCase().trim());
  SS_ = ss;
  try { ss.setSpreadsheetTimeZone(TZ); } catch (e) { /* abaikan */ }
  Object.keys(SCHEMA).forEach(name => ensureSheet_(ss, name));
  ss.getSheets().forEach(sh => {
    if (!SCHEMA[sh.getName()] && sh.getLastRow() === 0 && ss.getSheets().length > Object.keys(SCHEMA).length) ss.deleteSheet(sh);
  });

  if (!rows_('Kategori').length) {
    insertMany_('Kategori', [
      { ID_Kategori: uid_('KT'), Nama_Kategori: 'Badan Beruniform', Wajib: 'Ya', Susunan: 1, Status: 'Aktif', Warna: '#7B6B2C' },
      { ID_Kategori: uid_('KT'), Nama_Kategori: 'Kelab dan Persatuan', Wajib: 'Ya', Susunan: 2, Status: 'Aktif', Warna: '#4152A6' },
      { ID_Kategori: uid_('KT'), Nama_Kategori: 'Sukan dan Permainan', Wajib: 'Ya', Susunan: 3, Status: 'Aktif', Warna: '#1B7F55' }
    ]);
  }
  const tet = tetapan_();
  const defaults = {
    SESI_AKTIF: Utilities.formatDate(new Date(), TZ, 'yyyy'), NAMA_SEKOLAH: CONFIG.NAMA_SEKOLAH, ALAMAT_SEKOLAH: '',
    TEMPOH_KUNCI_HARI: '7', HARI_PERINGATAN: '2', LAYANAN_B: 'Dikecualikan', ID_FAIL_LOGO: '',
    NAMA_PENGESAH: '', JAWATAN_PENGESAH: 'Penolong Kanan Kokurikulum'
  };
  Object.keys(defaults).forEach(k => { if (tet[k] === undefined) setTetapan_(k, defaults[k]); });
  if (!rows_('Tetapan').some(r => r.Kunci === 'NAMA_SEKOLAH' && r.Nilai && r.Nilai !== 'Nama Sekolah Anda')) setTetapan_('NAMA_SEKOLAH', CONFIG.NAMA_SEKOLAH);
  if (!rows_('Format_Laporan').length) {
    insert_('Format_Laporan', {
      Versi: 1, Tarikh_Cipta: now_(), Aktif: 'Ya', Medan_JSON: JSON.stringify([
        { id: 'masa', label: 'Masa', jenis: 'masa', wajib: false },
        { id: 'tempat', label: 'Tempat', jenis: 'teks', wajib: false },
        { id: 'objektif', label: 'Objektif', jenis: 'panjang', wajib: true },
        { id: 'aktiviti', label: 'Aktiviti', jenis: 'panjang', wajib: true },
        { id: 'refleksi', label: 'Refleksi', jenis: 'panjang', wajib: false },
        { id: 'gambar', label: 'Gambar aktiviti', jenis: 'gambar', wajib: false }
      ])
    });
  }
  const sesi = sesi_();
  if (!jadual_(sesi).length) {
    const rows = [];
    for (let i = 1; i <= 12; i++) rows.push({ ID_Perjumpaan: uid_('J'), Tahun: sesi, Bil: i, Tarikh: '', Status: 'Dirancang', Catatan: '' });
    insertMany_('Jadual_Perjumpaan', rows);
  }
  rootFolder_();
  log_('', 'Setup', '', '', '', 'OK');
  clearCaches_();
  Logger.log('Setup selesai. Spreadsheet: ' + ss.getUrl());
}

function onOpen() {
  try {
    SpreadsheetApp.getUi().createMenu('e-Kokurikulum')
      .addItem('Setup awal', 'setup')
      .addItem('Import murid dari tab APDM_Mentah', 'importDariTab')
      .addItem('Daftar guru dari tab Data Guru', 'importGuruDariTab')
      .addItem('Daftar murid dari DATA MURID dan DATA KOKO', 'importKokoDariTab')
      .addItem('Pasang trigger automatik', 'pasangTrigger')
      .addItem('Backup sekarang', 'backupMingguan')
      .addItem('Arkib log lama sekarang', 'arkibLog')
      .addToUi();
  } catch (e) { /* bukan konteks Sheets */ }
}

function pasangTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (['backupMingguan', 'peringatanHarian', 'arkibLog'].indexOf(t.getHandlerFunction()) >= 0) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('arkibLog').timeBased().onMonthDay(1).atHour(3).create();
  ScriptApp.newTrigger('backupMingguan').timeBased().onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(2).create();
  ScriptApp.newTrigger('peringatanHarian').timeBased().everyDays(1).atHour(7).create();
}

function backupMingguan() {
  const root = rootFolder_();
  const it = root.getFoldersByName('Backup');
  const folder = it.hasNext() ? it.next() : root.createFolder('Backup');
  DriveApp.getFileById(ss_().getId()).makeCopy('Backup e-Kokurikulum ' + today_(), folder);
  const files = [];
  const fi = folder.getFiles();
  while (fi.hasNext()) files.push(fi.next());
  files.sort((a, b) => b.getDateCreated() - a.getDateCreated()).slice(8).forEach(f => f.setTrashed(true));
}

/* ---------- Arkib Log dan pembersihan Token (trigger bulanan) ---------- */
const LOG_SIMPAN_HARI = 90;   // log lebih lama daripada ini dipindahkan ke fail arkib

function arkibLogSS_() {
  const p = PropertiesService.getScriptProperties();
  const id = p.getProperty('ARKIB_LOG_ID');
  if (id) { try { return SpreadsheetApp.openById(id); } catch (e) { /* cipta semula */ } }
  const ss = SpreadsheetApp.create('Arkib Log e-Kokurikulum');
  try { DriveApp.getFileById(ss.getId()).moveTo(rootFolder_()); } catch (e) { /* kekal di My Drive */ }
  p.setProperty('ARKIB_LOG_ID', ss.getId());
  return ss;
}

/** Pindahkan log lama ke fail arkib dan buang token yang telah luput. Dipanggil oleh trigger bulanan atau menu. */
function arkibLog() {
  const r = withLock_(() => {
    Object.keys(T_).forEach(k => { delete T_[k]; });
    const t = table_('Log');
    const had = addDays_(today_(), -LOG_SIMPAN_HARI);
    const lama = t.rows.filter(x => x.Masa && String(x.Masa).slice(0, 10) < had);
    if (lama.length) {
      const sh = arkibLogSS_().getSheets()[0];
      if (sh.getLastRow() === 0) sh.getRange(1, 1, 1, t.head.length).setValues([t.head]).setFontWeight('bold');
      const start = sh.getLastRow() + 1;
      const need = start + lama.length - 1 - sh.getMaxRows();
      if (need > 0) sh.insertRowsAfter(sh.getMaxRows(), need);
      if (sh.getMaxColumns() < t.head.length) sh.insertColumnsAfter(sh.getMaxColumns(), t.head.length - sh.getMaxColumns());
      sh.getRange(start, 1, lama.length, t.head.length).setNumberFormat('@').setValues(lama.map(o => toArr_(t, o)));
      removeMany_('Log', lama);
    }
    const now = now_();
    const luput = rows_('Token').filter(x => x.Tarikh_Luput < now);
    if (luput.length) killTokens_(luput);
    if (lama.length) log_('', 'Arkib log', 'Log', '', '', lama.length + ' baris dipindahkan, ' + luput.length + ' token luput dibuang');
    return { log: lama.length, token: luput.length };
  }, true);
  Logger.log('Arkib log: ' + r.log + ' baris, token luput: ' + r.token);
  try { SpreadsheetApp.getUi().alert('Arkib selesai: ' + r.log + ' baris log dipindahkan, ' + r.token + ' token luput dibuang.'); } catch (e) { /* trigger: tiada UI */ }
  return r;
}

function peringatanHarian() {
  const s = computeStats_();
  const hari = Number(tetapan_().HARI_PERINGATAN || 2);
  const today = today_();
  const users = mapBy_(rows_('Users'), 'ID_Guru');
  const guruOf = {};
  rows_('Guru_Unit').forEach(g => { if (g.Tahun === s.sesi && g.Status === 'Aktif') (guruOf[g.ID_Unit] = guruOf[g.ID_Unit] || []).push(g.ID_Guru); });
  const perGuru = {};
  s.unit.forEach(u => u.tertunggak.forEach(t => {
    const days = diffDays_(t.tarikh, today);
    if (days < hari || (days - hari) % 7 !== 0) return;
    const apa = [];
    if (!t.kehadiran) apa.push('kehadiran belum lengkap');
    if (t.laporan !== 'Dihantar' && t.laporan !== 'Disahkan') apa.push(t.laporan ? 'laporan masih ' + t.laporan.toLowerCase() : 'laporan belum ditulis');
    (guruOf[u.id] || []).forEach(g => { (perGuru[g] = perGuru[g] || []).push(u.nama + ', perjumpaan ' + t.bil + ' (' + t.tarikh + '): ' + apa.join(' dan ')); });
  }));
  Object.keys(perGuru).forEach(gid => {
    const u = users[gid];
    if (!u || u.Status !== 'Aktif' || MailApp.getRemainingDailyQuota() < 1) return;
    MailApp.sendEmail(u.Emel, 'Peringatan kokurikulum: tindakan belum lengkap',
      'Salam sejahtera ' + u.Nama + ',\n\nPerkara berikut masih belum lengkap:\n- ' + perGuru[gid].join('\n- ') +
      '\n\nSila kemas kini dalam aplikasi e-Kokurikulum.\n\nEmel ini dijana secara automatik.');
  });
}


/* =========================================================================
 * PAPAN PEMUKA AWAM (tanpa log masuk)
 * Hanya data ringkasan. Tiada No. KP, telefon, emel atau kehadiran individu.
 * Dicache 10 minit dan dikosongkan setiap kali data berubah.
 * ========================================================================= */

function namaSingkat_(n) {
  const w = String(n || '').replace(/\s+(bin|binti|bt|bte|a\/l|a\/p|anak)\s.*$/i, '').trim().split(/\s+/).filter(String);
  if (!w.length) return '';
  const pre = /^(mohd|muhammad|mohamad|mohammad|muhamad|nur|nurul|siti|wan|nik|ahmad|abdul|ab|che|megat|puteri|tengku|syed|sharifah)$/i;
  if (pre.test(w[0]) && w[1]) return w[0] + ' ' + w[1] + (w[2] ? ' ' + w[2][0].toUpperCase() + '.' : '');
  return w[0] + (w[1] ? ' ' + w[1][0].toUpperCase() + '.' : '');
}

function awamData_() {
  return cached_('c_awam', 600, () => {
    const t = tetapan_(), sesi = sesi_(), today = today_();
    const mode = t.LAYANAN_B || 'Dikecualikan';
    const opt = {}; Object.keys(AWAM_LALAI).forEach(k => { opt[k] = t[k] || AWAM_LALAI[k]; });
    const st = computeStats_();
    const jad = jadual_(sesi);
    const held = jad.filter(j => j.Status !== 'Dibatalkan' && j.Tarikh && j.Tarikh <= today).sort((a, b) => Number(a.Bil) - Number(b.Bil));
    const murid = mapBy_(rows_('Murid').filter(layak_), 'ID_Murid');
    const units = rows_('Unit').filter(u => u.Status !== 'Tidak Aktif');
    const unitMap = mapBy_(units, 'ID_Unit');
    const kat = rows_('Kategori').filter(k => k.Status !== 'Tidak Aktif').sort((a, b) => Number(a.Susunan) - Number(b.Susunan));
    const enrol = rows_('Enrolmen').filter(e => e.Tahun === sesi && e.Status === 'Aktif' && murid[e.ID_Murid]);

    // Pendaftaran kategori dengan pecahan jantina
    const jumlahMurid = Object.keys(murid).length;
    const kategori = kat.map(k => {
      const set = {};
      enrol.forEach(e => { if (e.ID_Kategori === k.ID_Kategori) set[e.ID_Murid] = murid[e.ID_Murid].Jantina || '-'; });
      const ids = Object.keys(set);
      return {
        id: k.ID_Kategori, nama: k.Nama_Kategori, warna: k.Warna || '', wajib: k.Wajib !== 'Tidak', berdaftar: ids.length,
        L: ids.filter(i => set[i] === 'L').length, P: ids.filter(i => set[i] === 'P').length, belum: Math.max(0, jumlahMurid - ids.length)
      };
    });

    // Trend kehadiran seluruh sekolah bagi setiap perjumpaan
    const aktifKey = {};
    enrol.forEach(e => { aktifKey[e.ID_Unit + '|' + e.ID_Murid] = 1; });
    const kh = rows_('Kehadiran').filter(k => k.Tahun === sesi && aktifKey[k.ID_Unit + '|' + k.ID_Murid]);
    const trend = held.map(j => {
      let h = 0, n = 0;
      kh.forEach(k => {
        const v = k['P' + j.Bil];
        if (v === 'H') { h++; n++; } else if (v === 'T') n++;
        else if (v === 'B') { if (mode === 'Dikira hadir') { h++; n++; } else if (mode === 'Dikira tidak hadir') n++; }
      });
      return { bil: Number(j.Bil), tarikh: j.Tarikh, peratus: n ? Math.round(h / n * 100) : null, direkod: n };
    });

    // Laporan lengkap setiap unit
    const lap = {};
    rows_('Laporan').forEach(l => { if (l.Tahun === sesi && (l.Status === 'Dihantar' || l.Status === 'Disahkan')) lap[l.ID_Unit + '|' + l.ID_Perjumpaan] = 1; });
    const users = mapBy_(rows_('Users').filter(u => u.Status === 'Aktif'), 'ID_Guru');
    const guruOf = {}, guruSet = {};
    rows_('Guru_Unit').filter(g => g.Tahun === sesi && g.Status === 'Aktif' && users[g.ID_Guru])
      .sort((a, b) => (a.Peranan_Unit === 'Admin' ? 0 : 1) - (b.Peranan_Unit === 'Admin' ? 0 : 1))
      .forEach(g => { (guruOf[g.ID_Unit] = guruOf[g.ID_Unit] || []).push(users[g.ID_Guru].Nama); guruSet[g.ID_Guru] = 1; });
    const stU = {};
    st.unit.forEach(u => { stU[u.id] = u; });
    let lapAda = 0, lapPerlu = 0;
    const katP = katPerjumpaan_();
    const direktori = units.map(u => {
      const s = stU[u.ID_Unit] || { ahli: 0, peratus: null };
      const rekod = katP[u.ID_Kategori] !== false;
      let ada = 0;
      held.forEach(j => { if (lap[u.ID_Unit + '|' + j.ID_Perjumpaan]) ada++; });
      if (s.ahli && rekod) { lapAda += ada; lapPerlu += held.length; }
      return {
        id: u.ID_Unit, nama: u.Nama_Unit, idKategori: u.ID_Kategori, ahli: s.ahli, peratus: s.peratus,
        lengkap: rekod && s.ahli && held.length ? Math.round(ada / held.length * 100) : null, guru: guruOf[u.ID_Unit] || [], rekod: rekod
      };
    }).sort((a, b) => a.nama.localeCompare(b.nama));
    const top = direktori.filter(u => u.ahli && u.rekod && u.peratus != null)
      .sort((a, b) => (b.lengkap || 0) - (a.lengkap || 0) || b.peratus - a.peratus || b.ahli - a.ahli).slice(0, 5)
      .map(u => ({ id: u.id, nama: u.nama, idKategori: u.idKategori, peratus: u.peratus, lengkap: u.lengkap, ahli: u.ahli }));

    // Papan pencapaian (Zon ke atas)
    let pencapaian = [];
    if (opt.AWAM_PENCAPAIAN !== 'Sembunyi') {
      pencapaian = rows_('Pencapaian').filter(p => p.Tahun === sesi && PERINGKAT.indexOf(p.Peringkat) >= 1)
        .sort((a, b) => String(b.Tarikh).localeCompare(String(a.Tarikh))).slice(0, 20).map(p => {
          const m = murid[p.ID_Murid] || rows_('Murid').find(x => x.ID_Murid === p.ID_Murid) || {};
          return {
            nama: opt.AWAM_PENCAPAIAN === 'Penuh' ? (m.Nama_Murid || '') : namaSingkat_(m.Nama_Murid), kelas: m.Kelas || '',
            unit: (unitMap[p.ID_Unit] || {}).Nama_Unit || '', idKategori: (unitMap[p.ID_Unit] || {}).ID_Kategori || '',
            aktiviti: p.Nama_Aktiviti, peringkat: p.Peringkat, pencapaian: p.Pencapaian, tarikh: p.Tarikh
          };
        });
    }

    // Galeri: gambar daripada laporan yang telah disahkan sahaja
    const galeri = [];
    if (opt.AWAM_GALERI !== 'Tidak') {
      const fmts = {};
      rows_('Format_Laporan').forEach(f => { fmts[f.Versi] = parse_(f.Medan_JSON, []).filter(x => x.jenis === 'gambar').map(x => x.id); });
      const jById = mapBy_(jad, 'ID_Perjumpaan');
      rows_('Laporan').filter(l => l.Tahun === sesi && l.Status === 'Disahkan' && unitMap[l.ID_Unit])
        .sort((a, b) => String(b.Tarikh_Sah).localeCompare(String(a.Tarikh_Sah))).some(l => {
          const data = parse_(l.Data_JSON, {});
          (fmts[l.Versi_Format] || []).forEach(fid => (data[fid] || []).forEach(id => {
            if (galeri.length < 12 && id) galeri.push({ id: id, tajuk: clean_(data.tajuk, 150), unit: unitMap[l.ID_Unit].Nama_Unit, idKategori: unitMap[l.ID_Unit].ID_Kategori, bil: Number((jById[l.ID_Perjumpaan] || {}).Bil || 0), tarikh: (jById[l.ID_Perjumpaan] || {}).Tarikh || '' });
          }));
          return galeri.length >= 12;
        });
    }

    const pengumuman = rows_('Pengumuman').filter(p => p.Status === 'Aktif' && (!p.Tarikh_Mula || p.Tarikh_Mula <= today) && (!p.Tarikh_Tamat || p.Tarikh_Tamat >= today))
      .sort((a, b) => String(b.Tarikh_Mula || b.Tarikh_Cipta).localeCompare(String(a.Tarikh_Mula || a.Tarikh_Cipta))).slice(0, 6)
      .map(p => ({ id: p.ID, tajuk: p.Tajuk, isi: p.Isi, mula: p.Tarikh_Mula, tamat: p.Tarikh_Tamat }));

    return {
      namaSekolah: t.NAMA_SEKOLAH || CONFIG.NAMA_SEKOLAH, sesi: sesi, today: today, dikemaskini: now_(), opt: opt, idLogo: t.ID_FAIL_LOGO || '',
      angka: {
        murid: jumlahMurid, unit: direktori.filter(u => u.ahli).length, unitSemua: direktori.length, guru: Object.keys(guruSet).length,
        diadakan: st.diadakan, dirancang: st.jumlahPerjumpaan, kehadiran: st.peratus, laporan: lapPerlu ? Math.round(lapAda / lapPerlu * 100) : null
      },
      seterusnya: st.seterusnya,
      kategori: kategori,
      trend: trend, jadual: jadualOut_(sesi), direktori: direktori, top: top, pencapaian: pencapaian, galeri: galeri, pengumuman: pengumuman
    };
  });
}

function a_awam() {
  const d = awamData_();
  if (d.opt.AWAM_AKTIF === 'Tidak') return { tutup: true, namaSekolah: d.namaSekolah, sesi: d.sesi };
  return d;
}

/** Gambar awam: hanya gambar yang tersenarai dalam galeri laporan yang disahkan. */
function a_imgAwam(d) {
  const aw = awamData_();
  if (aw.opt.AWAM_AKTIF === 'Tidak' || !aw.galeri.some(g => g.id === d.id)) throw fail_('Gambar tidak tersedia.', 'FORBIDDEN');
  return a_image(d);
}

/** Semakan ibu bapa: kunci dicincang dalam peranti daripada No. MyKid; nombor asal tidak pernah dihantar. */
function a_semakMurid(d) {
  const t = tetapan_();
  if ((t.AWAM_SEMAKAN || 'Ya') === 'Tidak' || (t.AWAM_AKTIF || 'Ya') === 'Tidak') throw fail_('Semakan tidak dibuka buat masa ini.');
  const k = String(d.k || '');
  if (!/^K[0-9a-f]{20}$/.test(k)) throw fail_('Nombor MyKid tidak sah.');
  const cache = CacheService.getScriptCache();
  const slot = 'sm_' + Utilities.formatDate(new Date(), TZ, 'HHmm');
  const n = Number(cache.get(slot) || 0) + 1;
  cache.put(slot, String(n), 120);
  if (n > 120) throw fail_('Terlalu banyak semakan. Cuba lagi sebentar.', 'BUSY');
  const m = rows_('Murid').find(x => x.ID_APDM === k);
  if (!m) throw fail_('Tiada rekod dijumpai. Pastikan nombor MyKid betul, atau hubungi guru kelas.');
  const sesi = sesi_(), today = today_(), mode = t.LAYANAN_B || 'Dikecualikan';
  const heldBils = jadual_(sesi).filter(j => j.Status !== 'Dibatalkan' && j.Tarikh && j.Tarikh <= today).map(j => Number(j.Bil));
  const units = mapBy_(rows_('Unit'), 'ID_Unit'), kats = mapBy_(rows_('Kategori'), 'ID_Kategori');
  const jaw = {};
  rows_('Carta_Organisasi').forEach(c => { if (c.Tahun === sesi && c.Jenis === 'Murid' && c.ID_Rujukan === m.ID_Murid) jaw[c.ID_Unit] = c.Jawatan; });
  const kh = {};
  rows_('Kehadiran').forEach(r => { if (r.Tahun === sesi && r.ID_Murid === m.ID_Murid) kh[r.ID_Unit] = r; });
  const unit = rows_('Enrolmen').filter(e => e.Tahun === sesi && e.ID_Murid === m.ID_Murid && e.Status === 'Aktif').map(e => {
    const r = kh[e.ID_Unit];
    const c = { H: 0, T: 0, B: 0 };
    if (r) heldBils.forEach(b => { const v = r['P' + b]; if (c[v] !== undefined) c[v]++; });
    const p = r ? pct_(r, heldBils, mode) : { h: 0, n: 0 };
    return {
      unit: (units[e.ID_Unit] || {}).Nama_Unit || '', idKategori: e.ID_Kategori, kategori: (kats[e.ID_Kategori] || {}).Nama_Kategori || '',
      jawatan: jaw[e.ID_Unit] || 'Ahli', peratus: p.n ? Math.round(p.h / p.n * 100) : null, H: c.H, T: c.T, B: c.B
    };
  });
  const pencapaian = rows_('Pencapaian').filter(p => p.Tahun === sesi && p.ID_Murid === m.ID_Murid)
    .map(p => ({ aktiviti: p.Nama_Aktiviti, peringkat: p.Peringkat, pencapaian: p.Pencapaian, tarikh: p.Tarikh, unit: (units[p.ID_Unit] || {}).Nama_Unit || '' }));
  return { nama: m.Nama_Murid, kelas: m.Kelas, status: m.Status, sesi: sesi, diadakan: heldBils.length, unit: unit, pencapaian: pencapaian };
}

/* ---------- Pengumuman (Super Admin) ---------- */
function pengumumanOut_(p) {
  return { id: p.ID, tajuk: p.Tajuk, isi: p.Isi, mula: p.Tarikh_Mula, tamat: p.Tarikh_Tamat, status: p.Status, cipta: p.Tarikh_Cipta };
}
function a_pengumuman(d, ctx) {
  needSuper_(ctx);
  return { pengumuman: rows_('Pengumuman').map(pengumumanOut_).sort((a, b) => String(b.mula || b.cipta).localeCompare(String(a.mula || a.cipta))) };
}
function a_savePengumuman(d, ctx) {
  needSuper_(ctx);
  const tajuk = clean_(d.tajuk, 120);
  if (!tajuk) throw fail_('Isi tajuk pengumuman.');
  const patch = {
    Tajuk: tajuk, Isi: clean_(d.isi, 1500), Tarikh_Mula: isYmd_(d.mula) ? d.mula : '', Tarikh_Tamat: isYmd_(d.tamat) ? d.tamat : '',
    Status: d.status === 'Draf' ? 'Draf' : 'Aktif'
  };
  if (patch.Tarikh_Mula && patch.Tarikh_Tamat && patch.Tarikh_Tamat < patch.Tarikh_Mula) throw fail_('Tarikh tamat mesti selepas tarikh mula.');
  ensureSheet_(ss_(), 'Pengumuman');
  if (d.id) {
    const p = rows_('Pengumuman').find(x => x.ID === d.id);
    if (!p) throw fail_('Pengumuman tidak dijumpai.');
    update_('Pengumuman', p, patch);
  } else {
    insert_('Pengumuman', Object.assign({ ID: uid_('PG'), Dicipta_Oleh: ctx.id, Tarikh_Cipta: now_() }, patch));
  }
  log_(ctx.id, 'Simpan pengumuman', 'Pengumuman', d.id || '', '', tajuk);
  return { ok: true };
}
function a_deletePengumuman(d, ctx) {
  needSuper_(ctx);
  const p = rows_('Pengumuman').find(x => x.ID === d.id);
  if (!p) throw fail_('Pengumuman tidak dijumpai.');
  removeMany_('Pengumuman', [p]);
  log_(ctx.id, 'Padam pengumuman', 'Pengumuman', d.id, p.Tajuk, '');
  return { ok: true };
}

/* ---------- Import akaun guru dari tab Data Guru ---------- */

function cariSheetDataGuru_() {
  const want = String(CONFIG.SHEET_DATA_GURU).toLowerCase().replace(/[^a-z0-9]/g, '');
  const all = ss_().getSheets();
  return all.find(sh => sh.getName().toLowerCase().replace(/[^a-z0-9]/g, '') === want)
    || all.find(sh => /data/i.test(sh.getName()) && /guru/i.test(sh.getName()) && !SCHEMA[sh.getName()]) || null;
}

function importGuruCore_(who) {
  const src = cariSheetDataGuru_();
  if (!src) throw fail_('Tab "' + CONFIG.SHEET_DATA_GURU + '" tidak dijumpai dalam Google Sheets.');
  const v = src.getDataRange().getDisplayValues();
  let hi = -1;
  for (let i = 0; i < Math.min(10, v.length); i++) { if (v[i].some(c => /E-?MEL|EMAIL/i.test(c))) { hi = i; break; } }
  if (hi < 0) throw fail_('Baris tajuk dengan lajur emel tidak dijumpai dalam tab ' + src.getName() + '.');
  const h = v[hi].map(c => String(c).toUpperCase().replace(/\s+/g, ' ').trim());
  const find = re => h.findIndex(c => re.test(c));
  const cE = find(/E-?MEL|EMAIL/), cN = find(/NAMA/);
  const cK = find(/NO\.?\s*K\.?\s*\/?\s*P|NOKP|KAD PENGENALAN|MYKAD|NO\.?\s*IC|^IC$/);
  const cT = find(/TELEFON|NO\.?\s*TEL|H\/P|^HP$/);
  if (cN < 0) throw fail_('Lajur nama tidak dijumpai dalam tab ' + src.getName() + '.');

  // Pastikan lajur No_KP dan Tukar_Kata_Laluan wujud dalam sheet Users
  ensureSheet_(ss_(), 'Users');
  delete T_.Users;
  const t = table_('Users');
  const byEmel = {};
  t.rows.forEach(u => { byEmel[String(u.Emel).toLowerCase()] = u; });
  const superE = superEmail_(), now = now_();
  const add = [], langkau = [], seen = {};
  let kemas = 0, sama = 0;
  for (let i = hi + 1; i < v.length; i++) {
    const r = v[i];
    const emel = String(r[cE] || '').toLowerCase().trim();
    const nama = clean_(r[cN], 120);
    if (!emel && !nama) continue;
    const baris = i + 1;
    if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(emel)) { langkau.push({ baris: baris, nama: nama, sebab: 'emel tidak sah' }); continue; }
    if (!nama) { langkau.push({ baris: baris, nama: emel, sebab: 'tiada nama' }); continue; }
    if (seen[emel]) { langkau.push({ baris: baris, nama: nama, sebab: 'emel berulang' }); continue; }
    seen[emel] = 1;
    let kp = cK >= 0 ? String(r[cK] || '').replace(/\D/g, '') : '';
    if (kp.length === 11) kp = '0' + kp;
    const tel = cT >= 0 ? clean_(r[cT], 20) : '';
    const u = byEmel[emel];
    if (u) {
      let ubah = false;
      if (kp && u.No_KP !== kp) { u.No_KP = kp; ubah = true; }
      if (tel && !u.No_Telefon) { u.No_Telefon = tel; ubah = true; }
      if (ubah) kemas++; else sama++;
      continue;
    }
    const salt = Utilities.getUuid();
    const isSuper = emel === superE;
    add.push({
      ID_Guru: uid_('G'), Nama: nama, Emel: emel, No_Telefon: tel, Hash_Kata_Laluan: hashPw_(CONFIG.KATA_LALUAN_LALAI, salt), Salt: salt,
      Peranan_Global: isSuper ? 'SuperAdmin' : 'Guru', Status: 'Aktif', Cubaan_Gagal: 0, Kunci_Hingga: '', Tarikh_Daftar: now,
      Log_Masuk_Terakhir: '', No_KP: kp, Tukar_Kata_Laluan: 'Ya'
    });
  }
  if (kemas) writeAll_('Users');
  insertMany_('Users', add);
  log_(who, 'Import guru', 'Users', '', '', 'baru ' + add.length + ', dikemas kini ' + kemas + ', dilangkau ' + langkau.length);
  return { baru: add.length, kemas: kemas, sama: sama, langkau: langkau.slice(0, 50), jumlahLangkau: langkau.length, tab: src.getName() };
}

function a_importGuru(d, ctx) {
  needSuper_(ctx);
  return importGuruCore_(ctx.id);
}

/** Menu Sheets: import akaun guru tanpa membuka aplikasi. */
function importGuruDariTab() {
  const ui = SpreadsheetApp.getUi();
  try {
    const r = withLock_(() => importGuruCore_('menu'), true);
    ui.alert('Import guru selesai (tab ' + r.tab + ')\n\n' + r.baru + ' akaun baharu\n' + r.kemas + ' akaun sedia ada dikemas kini\n' + r.sama + ' tiada perubahan\n' + r.jumlahLangkau + ' baris dilangkau' +
      (r.langkau.length ? '\n\n' + r.langkau.slice(0, 10).map(x => 'Baris ' + x.baris + ' (' + x.nama + '): ' + x.sebab).join('\n') : '') +
      '\n\nKata laluan lalai: ' + CONFIG.KATA_LALUAN_LALAI + '. Guru wajib menukarnya semasa log masuk pertama.');
  } catch (e) { ui.alert(String(e.message || e)); }
}






/* ---------- Guru kelas: tetapkan unit bagi setiap murid kelas ---------- */
function kelasSaya_(d, ctx) {
  const k = ctx.super && d.kelas ? clean_(d.kelas, 40) : String(ctx.user.Kelas_Guru || '');
  if (!k) throw fail_('Tetapkan kelas anda di halaman Profil dahulu.', 'FORBIDDEN');
  return k;
}
function a_kelasUnit(d, ctx) {
  const kelas = kelasSaya_(d, ctx), sesi = ctx.sesi;
  const kats = rows_('Kategori').filter(k => k.Status !== 'Tidak Aktif').sort((a, b) => Number(a.Susunan) - Number(b.Susunan));
  const count = {}, ada = {};
  rows_('Enrolmen').forEach(e => {
    if (e.Tahun !== sesi || e.Status !== 'Aktif') return;
    count[e.ID_Unit] = (count[e.ID_Unit] || 0) + 1;
    (ada[e.ID_Murid] = ada[e.ID_Murid] || {})[e.ID_Kategori] = e.ID_Unit;
  });
  const unit = rows_('Unit').filter(u => u.Status !== 'Tidak Aktif').map(u => ({ id: u.ID_Unit, idKategori: u.ID_Kategori, nama: u.Nama_Unit, ahli: count[u.ID_Unit] || 0 }))
    .sort((a, b) => a.nama.localeCompare(b.nama));
  const murid = rows_('Murid').filter(m => layak_(m) && m.Kelas === kelas).map(m => ({ id: m.ID_Murid, nama: m.Nama_Murid, jantina: m.Jantina || '', unit: ada[m.ID_Murid] || {} }))
    .sort((a, b) => a.nama.localeCompare(b.nama));
  return { kelas: kelas, kategori: kats.map(katOut_), unit: unit, murid: murid };
}
function a_simpanKelasUnit(d, ctx) {
  const kelas = kelasSaya_(d, ctx), sesi = ctx.sesi;
  const murid = mapBy_(rows_('Murid').filter(m => layak_(m) && m.Kelas === kelas), 'ID_Murid');
  const units = mapBy_(rows_('Unit'), 'ID_Unit');
  const list = Array.isArray(d.perubahan) ? d.perubahan : [];
  if (!list.length) throw fail_('Tiada perubahan untuk disimpan.');
  if (list.length > 600) throw fail_('Terlalu banyak perubahan sekali gus.');
  // Sebelum perjumpaan pertama sesi: dikira sejak awal sesi. Selepas itu: sejak hari ini.
  const today = today_();
  const bermula = jadual_(sesi).some(j => j.Status !== 'Dibatalkan' && j.Tarikh && j.Tarikh <= today);
  const masuk = bermula ? today : '';
  const t = table_('Enrolmen');
  const semasa = {};
  t.rows.forEach(e => { if (e.Tahun === sesi && e.Status === 'Aktif') semasa[e.ID_Murid + '|' + e.ID_Kategori] = e; });
  const add = [];
  let baru = 0, tukar = 0, keluar = 0, ubah = false;
  list.forEach(c => {
    const m = murid[c.idMurid];
    if (!m) throw fail_('Murid bukan dalam kelas ' + kelas + '.', 'FORBIDDEN');
    const idK = String(c.idKategori || ''), idU = String(c.idUnit || '');
    const e = semasa[m.ID_Murid + '|' + idK];
    if (!idU) { if (e) { e.Status = 'Keluar'; ubah = true; keluar++; } return; }
    const u = units[idU];
    if (!u || u.ID_Kategori !== idK || u.Status === 'Tidak Aktif') throw fail_('Unit tidak sah bagi kategori ini.');
    if (e && e.ID_Unit === idU) return;
    // Tukar unit: rekod lama ditutup (Keluar) supaya sejarah keahlian kekal, rekod baharu dicipta
    if (e) { e.Status = 'Keluar'; ubah = true; tukar++; } else baru++;
    const ne = { ID_Enrolmen: uid_('E'), Tahun: sesi, ID_Murid: m.ID_Murid, ID_Unit: idU, ID_Kategori: idK, Kelas: m.Kelas, Status: 'Aktif', Tarikh_Masuk: masuk };
    semasa[m.ID_Murid + '|' + idK] = ne; add.push(ne);
  });
  if (ubah) writeAll_('Enrolmen');
  insertMany_('Enrolmen', add);
  log_(ctx.id, 'Tetapan unit kelas ' + kelas, 'Enrolmen', '', '', baru + ' baharu, ' + tukar + ' bertukar, ' + keluar + ' dikosongkan');
  return { baru: baru, tukar: tukar, keluar: keluar };
}

/* ---------- Guru: tukar atau keluar unit, senarai kelas ---------- */
function senaraiKelas_() {
  const set = {};
  rows_('Murid').forEach(m => { if (layak_(m) && m.Kelas) set[m.Kelas] = 1; });
  return Object.keys(set).sort((a, b) => a.localeCompare(b, 'ms', { numeric: true }));
}
function a_kelasList() {
  return { kelas: senaraiKelas_() };
}
/** Tukar unit (dari -> ke), keluar unit (dari sahaja) atau sertai (ke sahaja). Terus tanpa kelulusan. */
function a_tukarUnit(d, ctx) {
  const sesi = ctx.sesi, dari = String(d.dari || ''), ke = String(d.ke || '');
  if (!dari && !ke) throw fail_('Tiada perubahan unit.');
  if (dari && dari === ke) throw fail_('Pilih unit yang berbeza.');
  const units = mapBy_(rows_('Unit'), 'ID_Unit');
  if (ke && (!units[ke] || units[ke].Status === 'Tidak Aktif')) throw fail_('Unit baharu tidak dijumpai.');
  const gu = rows_('Guru_Unit').filter(x => x.Tahun === sesi && x.ID_Guru === ctx.id);
  let tanpaAdmin = false;
  if (dari) {
    const g = gu.find(x => x.ID_Unit === dari && (x.Status === 'Aktif' || x.Status === 'Menunggu'));
    if (!g) throw fail_('Anda bukan guru unit ini.');
    update_('Guru_Unit', g, { Status: 'Tidak Aktif' });
    if (g.Peranan_Unit === 'Admin') {
      tanpaAdmin = !rows_('Guru_Unit').some(x => x.Tahun === sesi && x.ID_Unit === dari && x.Status === 'Aktif' && x.Peranan_Unit === 'Admin');
    }
  }
  if (ke) {
    const g = gu.find(x => x.ID_Unit === ke);
    if (g && g.Status === 'Aktif') throw fail_('Anda sudah menjadi guru unit ' + units[ke].Nama_Unit + '.');
    if (g) update_('Guru_Unit', g, { Status: 'Aktif', Peranan_Unit: 'Penasihat' });
    else insert_('Guru_Unit', { ID: uid_('GU'), Tahun: sesi, ID_Guru: ctx.id, ID_Unit: ke, Peranan_Unit: 'Penasihat', Status: 'Aktif' });
  }
  log_(ctx.id, dari && ke ? 'Tukar unit' : dari ? 'Keluar unit' : 'Sertai unit', 'Guru_Unit', ctx.id,
    dari ? (units[dari] || {}).Nama_Unit || dari : '', ke ? units[ke].Nama_Unit : '');
  return { tanpaAdmin: tanpaAdmin, namaDari: dari ? (units[dari] || {}).Nama_Unit || '' : '' };
}

/* ---------- Padam atau gabung unit yang tersalah daftar (Super Admin) ---------- */
const DATA_UNIT_ = ['Enrolmen', 'Kehadiran', 'Laporan', 'Guru_Unit', 'Carta_Organisasi', 'Pencapaian'];

function unitGuna_(id) {
  const c = {};
  DATA_UNIT_.forEach(n => { c[n] = rows_(n).filter(r => r.ID_Unit === id).length; });
  const sesi = sesi_();
  c.ahli = rows_('Enrolmen').filter(e => e.ID_Unit === id && e.Tahun === sesi && e.Status === 'Aktif').length;
  c.jumlah = DATA_UNIT_.reduce((a, n) => a + c[n], 0);
  return c;
}
function a_unitGuna(d, ctx) {
  needSuper_(ctx);
  const u = rows_('Unit').find(x => x.ID_Unit === d.id);
  if (!u) throw fail_('Unit tidak dijumpai.');
  return { guna: unitGuna_(u.ID_Unit) };
}

function a_deleteUnit(d, ctx) {
  needSuper_(ctx);
  const src = rows_('Unit').find(x => x.ID_Unit === d.id);
  if (!src) throw fail_('Unit tidak dijumpai.');
  const guna = unitGuna_(src.ID_Unit);
  const hasil = { dipindah: {}, dibuang: {} };

  if (d.mode === 'gabung') {
    const dst = rows_('Unit').find(x => x.ID_Unit === d.gabungKe);
    if (!dst || dst.ID_Unit === src.ID_Unit) throw fail_('Pilih unit sasaran yang lain.');
    if (dst.ID_Kategori !== src.ID_Kategori) throw fail_('Unit hanya boleh digabung dalam kategori yang sama.');
    const S = src.ID_Unit, D = dst.ID_Unit;
    const kerja = (nama, kunci, gabung) => {
      const t = table_(nama);
      const sasaran = {};
      if (kunci) t.rows.forEach(r => { if (r.ID_Unit === D) sasaran[kunci(r)] = r; });
      const buang = [];
      let pindah = 0;
      t.rows.forEach(r => {
        if (r.ID_Unit !== S) return;
        const ada = kunci ? sasaran[kunci(r)] : null;
        if (ada) { if (gabung) gabung(ada, r); buang.push(r); }
        else { r.ID_Unit = D; pindah++; }
      });
      if (buang.length) removeMany_(nama, buang); else if (pindah) writeAll_(nama);
      hasil.dipindah[nama] = pindah; hasil.dibuang[nama] = buang.length;
    };
    // Kehadiran: satu baris setiap murid; gabungkan tanda yang kosong pada unit sasaran
    kerja('Kehadiran', r => r.Tahun + '|' + r.ID_Murid, (a, b) => { P_COLS.forEach(c => { if (!a[c] && b[c]) a[c] = b[c]; }); });
    // Laporan: jika kedua-dua unit ada laporan bagi perjumpaan sama, kekalkan status yang lebih tinggi
    const rank = { Disahkan: 4, Dihantar: 3, Dikembalikan: 2, Draf: 1 };
    kerja('Laporan', r => r.Tahun + '|' + r.ID_Perjumpaan, (a, b) => {
      if ((rank[b.Status] || 0) > (rank[a.Status] || 0)) SCHEMA.Laporan.forEach(c => { if (c !== 'ID_Laporan' && c !== 'ID_Unit') a[c] = b[c]; });
    });
    kerja('Guru_Unit', r => r.Tahun + '|' + r.ID_Guru, (a, b) => { if (b.Peranan_Unit === 'Admin') a.Peranan_Unit = 'Admin'; if (b.Status === 'Aktif') a.Status = 'Aktif'; });
    kerja('Enrolmen', r => r.Tahun + '|' + r.ID_Murid + '|' + r.Status, null);
    kerja('Carta_Organisasi', null, null);
    kerja('Pencapaian', null, null);
    // Kira semula bilangan hadir laporan unit sasaran selepas kehadiran digabung
    const jb = mapBy_(rows_('Jadual_Perjumpaan'), 'ID_Perjumpaan');
    const khD = rows_('Kehadiran').filter(r => r.ID_Unit === D);
    const tL = table_('Laporan');
    let ubahL = false;
    tL.rows.forEach(l => {
      const j = l.ID_Unit === D && jb[l.ID_Perjumpaan];
      if (!j) return;
      const n = khD.filter(r => r.Tahun === l.Tahun && r['P' + j.Bil] === 'H').length;
      if (String(l.Bil_Hadir) !== String(n)) { l.Bil_Hadir = n; ubahL = true; }
    });
    if (ubahL) writeAll_('Laporan');
    removeMany_('Unit', [rows_('Unit').find(x => x.ID_Unit === S)]);
    log_(ctx.id, 'Gabung unit', 'Unit', S, src.Nama_Unit, 'digabung ke ' + dst.Nama_Unit);
    return { mode: 'gabung', ke: dst.Nama_Unit, hasil: hasil };
  }

  // Padam: unit tanpa data terus dipadam; unit berdata memerlukan pengesahan nama unit
  if (guna.jumlah && String(d.sahkan || '').trim().toUpperCase() !== String(src.Nama_Unit).trim().toUpperCase()) {
    throw fail_('Taip nama unit dengan tepat untuk mengesahkan pemadaman semua data unit ini.');
  }
  DATA_UNIT_.forEach(n => {
    const buang = rows_(n).filter(r => r.ID_Unit === src.ID_Unit);
    if (buang.length) removeMany_(n, buang);
    hasil.dibuang[n] = buang.length;
  });
  if (src.ID_Folder_Drive) { try { DriveApp.getFolderById(src.ID_Folder_Drive).setTrashed(true); } catch (e) { /* abaikan */ } }
  removeMany_('Unit', [rows_('Unit').find(x => x.ID_Unit === src.ID_Unit)]);
  log_(ctx.id, 'Padam unit', 'Unit', src.ID_Unit, src.Nama_Unit, JSON.stringify(hasil.dibuang));
  return { mode: 'padam', hasil: hasil };
}

/* ---------- Murid belum berunit (Super Admin) ---------- */
function a_belumUnit(d, ctx) {
  needSuper_(ctx);
  const sesi = ctx.sesi;
  const kats = rows_('Kategori').filter(k => k.Status !== 'Tidak Aktif').sort((a, b) => Number(a.Susunan) - Number(b.Susunan));
  const wajib = kats.filter(k => k.Wajib !== 'Tidak');
  const ada = {};
  rows_('Enrolmen').forEach(e => { if (e.Tahun === sesi && e.Status === 'Aktif') (ada[e.ID_Murid] = ada[e.ID_Murid] || {})[e.ID_Kategori] = e.ID_Unit; });
  const murid = rows_('Murid').filter(layak_).map(m => {
    const u = ada[m.ID_Murid] || {};
    return { id: m.ID_Murid, nama: m.Nama_Murid, kelas: m.Kelas, jantina: m.Jantina, baru: muridBaharu_(m), unit: u, tiada: wajib.filter(k => !u[k.ID_Kategori]).map(k => k.ID_Kategori) };
  }).filter(m => m.tiada.length)
    .sort((a, b) => (b.baru ? 1 : 0) - (a.baru ? 1 : 0) || String(a.kelas).localeCompare(String(b.kelas), 'ms', { numeric: true }) || a.nama.localeCompare(b.nama));
  return { kategori: kats.map(katOut_), murid: murid };
}

/** Daftar seorang murid ke unit bagi beberapa kategori sekali gus (Super Admin). */
function a_daftarMurid(d, ctx) {
  needSuper_(ctx);
  const sesi = ctx.sesi;
  const m = rows_('Murid').find(x => x.ID_Murid === d.idMurid);
  if (!m || m.Status !== 'Aktif') throw fail_('Murid tidak dijumpai atau tidak aktif.');
  const masuk = isYmd_(d.tarikh) && d.tarikh <= today_() ? d.tarikh : today_();
  const units = mapBy_(rows_('Unit'), 'ID_Unit');
  const t = table_('Enrolmen');
  const semasa = {};
  t.rows.forEach(e => { if (e.Tahun === sesi && e.ID_Murid === m.ID_Murid && e.Status === 'Aktif') semasa[e.ID_Kategori] = e; });
  const add = [];
  let ubah = false, baru = 0, pindah = 0;
  Object.keys(d.units || {}).forEach(idK => {
    const idU = String(d.units[idK] || '');
    if (!idU) return;
    const u = units[idU];
    if (!u || u.ID_Kategori !== idK || u.Status === 'Tidak Aktif') throw fail_('Unit tidak sah bagi kategori ini.');
    const e = semasa[idK];
    if (e && e.ID_Unit === idU) return;
    // Tukar unit: tutup rekod lama (Keluar), cipta rekod baharu
    if (e) { e.Status = 'Keluar'; ubah = true; pindah++; } else baru++;
    add.push({ ID_Enrolmen: uid_('E'), Tahun: sesi, ID_Murid: m.ID_Murid, ID_Unit: idU, ID_Kategori: idK, Kelas: m.Kelas, Status: 'Aktif', Tarikh_Masuk: masuk });
  });
  if (ubah) writeAll_('Enrolmen');
  insertMany_('Enrolmen', add);
  log_(ctx.id, 'Daftar murid ke unit', 'Enrolmen', m.ID_Murid, '', m.Nama_Murid + ': ' + baru + ' baharu, ' + pindah + ' bertukar');
  return { baru: baru, pindah: pindah };
}

/* ---------- Pendaftaran murid ke unit dari tab DATA MURID + DATA KOKO ---------- */

function cariTab_(nama) {
  const want = String(nama).toLowerCase().replace(/[^a-z0-9]/g, '');
  const all = ss_().getSheets();
  return all.find(sh => sh.getName().toLowerCase().replace(/[^a-z0-9]/g, '') === want) || null;
}
/** Seragamkan nama untuk padanan: huruf besar, tanpa tanda baca, BIN/BINTI diseragamkan. */
function normNama_(s) {
  let x = ' ' + String(s || '').toUpperCase().replace(/[`'’‘".,()\[\]@]/g, ' ').replace(/\s+/g, ' ') + ' ';
  x = x.replace(/ A\s*\/\s*L /g, ' AL ').replace(/ A\s*\/\s*P /g, ' AP ').replace(/\//g, ' ');
  x = x.replace(/ (BINTI|BINTE|BTE|BT|BTI|BINT) /g, ' BINTI ').replace(/ (BIN|B) /g, ' BIN ');
  x = x.replace(/ (MOHD|MHD|MOH|MUHD) /g, ' MOHD ').replace(/ (ABD) /g, ' ABDUL ');
  return x.replace(/\s+/g, ' ').trim();
}
function normUnit_(s) {
  return String(s || '').toUpperCase().replace(/&/g, ' DAN ').replace(/[^A-Z0-9]+/g, '');
}
function namaUnitCantik_(s) {
  const kekal = /^(PBSM|TKRS|KRS|STEM|ICT|PKBM|PPIM|SPBT|NILAM|RIMUP|KRSM|TKD|TKRSM)$/;
  return String(s || '').trim().replace(/\s+/g, ' ').split(' ').map(w => {
    const u = w.toUpperCase();
    if (kekal.test(u.replace(/[^A-Z]/g, '')) || !/[AEIOU]/.test(u)) return u;
    if (/^(DAN|&)$/.test(u)) return u === '&' ? '&' : 'dan';
    return u.charAt(0) + u.slice(1).toLowerCase();
  }).join(' ');
}
function levenshtein_(a, b) {
  if (a === b) return 0;
  const m = a.length, n = b.length;
  if (!m || !n) return m || n;
  let prev = [];
  for (let j = 0; j <= n; j++) prev[j] = j;
  for (let i = 1; i <= m; i++) {
    const cur = [i];
    for (let j = 1; j <= n; j++) cur[j] = Math.min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
    prev = cur;
  }
  return prev[n];
}
function headerRow_(v, need) {
  for (let i = 0; i < Math.min(15, v.length); i++) {
    const h = v[i].map(c => String(c).toUpperCase());
    if (h.some(c => /NAMA/.test(c)) && h.some(c => need.test(c))) return i;
  }
  return -1;
}

function importKokoCore_(who) {
  const tM = cariTab_(CONFIG.SHEET_DATA_MURID), tK = cariTab_(CONFIG.SHEET_DATA_KOKO);
  if (!tM) throw fail_('Tab "' + CONFIG.SHEET_DATA_MURID + '" tidak dijumpai.');
  if (!tK) throw fail_('Tab "' + CONFIG.SHEET_DATA_KOKO + '" tidak dijumpai.');
  const sesi = sesi_(), now = now_(), today = today_();

  /* 1. DATA MURID */
  const vm = tM.getDataRange().getDisplayValues();
  const hm = headerRow_(vm, /KELAS|TAHUN|TINGKATAN|PENGENALAN|K\/?P/);
  if (hm < 0) throw fail_('Baris tajuk (NAMA, KELAS) tidak dijumpai dalam tab ' + tM.getName() + '.');
  const HM = vm[hm].map(c => String(c).toUpperCase().replace(/\s+/g, ' ').trim());
  const fm = (re, not) => HM.findIndex(c => re.test(c) && !(not && not.test(c)));
  const mN = fm(/NAMA/, /KELAS|IBU|BAPA|PENJAGA|GURU|SEKOLAH|RUMAH/), mKP = fm(/PENGENALAN|NO\.?\s*K\.?\s*\/?\s*P|NOKP|MYKID|MYKAD|NO\.?\s*IC|^IC$|SIJIL/);
  const mT = fm(/TAHUN|TINGKATAN/, /LAHIR|MASUK|DAFTAR/), mK = fm(/KELAS/, /GURU/), mJ = fm(/JANTINA/);
  if (mN < 0 || mK < 0) throw fail_('Lajur NAMA dan KELAS diperlukan dalam tab ' + tM.getName() + '.');
  const kata = { SATU: 1, DUA: 2, TIGA: 3, EMPAT: 4, LIMA: 5, ENAM: 6 };
  const recs = [];
  for (let i = hm + 1; i < vm.length; i++) {
    const r = vm[i];
    const nama = clean_(r[mN], 150);
    if (!nama) continue;
    let kp = mKP >= 0 ? String(r[mKP] || '').replace(/\D/g, '') : '';
    if (kp.length === 11) kp = '0' + kp;
    let kelas = clean_(r[mK], 40);
    let tahun = 0;
    const ts = mT >= 0 ? String(r[mT] || '').toUpperCase() : '';
    const dg = ts.match(/\d/);
    if (dg) tahun = Number(dg[0]); else Object.keys(kata).some(w => { if (ts.indexOf(w) >= 0) { tahun = kata[w]; return true; } return false; });
    if (!tahun) tahun = tahunKelas_(kelas);
    if (kelas && !/^\d/.test(kelas) && tahun) kelas = tahun + ' ' + kelas;
    recs.push({ baris: i + 1, nama: nama, norm: normNama_(nama), kp: kp, tahun: tahun, kelas: kelas, jantina: mJ >= 0 ? normJantina_(r[mJ]) : '' });
  }
  const idx = {}, idxC = {}, byFirst = {};
  recs.forEach(x => {
    (idx[x.norm] = idx[x.norm] || []).push(x);
    const c = x.norm.replace(/ /g, '');
    (idxC[c] = idxC[c] || []).push(x);
    const f = x.norm.charAt(0);
    (byFirst[f] = byFirst[f] || []).push(x);
  });
  const layakT = x => CONFIG.TAHUN_KOKO.indexOf(x.tahun) >= 0;

  /* 2. Kemas kini jadual Murid bagi murid Tahun 4-6 */
  ensureSheet_(ss_(), 'Kategori'); delete T_.Kategori;
  ensureSheet_(ss_(), 'Murid'); delete T_.Murid;
  ensureSheet_(ss_(), 'Enrolmen'); delete T_.Enrolmen;
  const tMur = table_('Murid');
  const byKey = {};
  tMur.rows.forEach(r => { if (r.ID_APDM) byKey[r.ID_APDM] = r; });
  const addM = [];
  let muridBaru = 0, muridKemas = 0, muridUbah = false;
  recs.filter(layakT).forEach(x => {
    const kN = 'N' + sha_(x.nama.toUpperCase().replace(/\s+/g, ' ').trim()).slice(0, 20);
    const k = x.kp.length >= 6 ? 'K' + sha_(x.kp).slice(0, 20) : kN;
    let r = byKey[k];
    if (!r && k !== kN && byKey[kN]) { r = byKey[kN]; delete byKey[kN]; r.ID_APDM = k; byKey[k] = r; muridUbah = true; }
    if (r) {
      const jan = x.jantina || r.Jantina;
      if (r.Nama_Murid !== x.nama || r.Kelas !== x.kelas || r.Jantina !== jan || r.Status !== 'Aktif') {
        r.Nama_Murid = x.nama; r.Kelas = x.kelas; r.Jantina = jan; r.Status = 'Aktif'; r.Tarikh_Kemaskini = now;
        muridKemas++; muridUbah = true;
      }
    } else {
      r = { ID_Murid: uid_('M'), ID_APDM: k, Nama_Murid: x.nama, Kelas: x.kelas, Jantina: x.jantina, Status: 'Aktif', Tarikh_Kemaskini: now, Tarikh_Daftar: now };
      byKey[k] = r; addM.push(r); muridBaru++;
    }
    x.id = r.ID_Murid;
  });
  if (muridUbah) writeAll_('Murid');
  insertMany_('Murid', addM);

  /* 3. Kategori (cipta jika tiada) */
  const vk = tK.getDataRange().getDisplayValues();
  const hk = headerRow_(vk, /UNIF|KELAB|PERSATUAN|SUKAN|RUMAH/);
  if (hk < 0) throw fail_('Baris tajuk (NAMA dan lajur unit) tidak dijumpai dalam tab ' + tK.getName() + '.');
  const HK = vk[hk].map(c => String(c).toUpperCase().replace(/\s+/g, ' ').trim());
  const fk = (re, not) => HK.findIndex(c => re.test(c) && !(not && not.test(c)));
  const kN = fk(/NAMA/, /KELAS|IBU|BAPA|GURU|UNIT|KELAB|RUMAH/), kK = fk(/KELAS/, /KELAB/);
  if (kN < 0) throw fail_('Lajur NAMA tidak dijumpai dalam tab ' + tK.getName() + '.');
  const kats = rows_('Kategori');
  const addK = [];
  let susun = kats.reduce((m, k) => Math.max(m, Number(k.Susunan || 0)), 0);
  const katFor = (re, not, nama, warna, rekod) => {
    let k = kats.concat(addK).find(x => re.test(x.Nama_Kategori) && !(not && not.test(x.Nama_Kategori)));
    if (!k) {
      k = { ID_Kategori: uid_('KT'), Nama_Kategori: nama, Wajib: 'Ya', Susunan: ++susun, Status: 'Aktif', Warna: warna, Rekod_Perjumpaan: rekod ? 'Ya' : 'Tidak' };
      addK.push(k);
    }
    return k;
  };
  // Ejaan tajuk dibenarkan berbeza sedikit, contohnya BERUNIFOM, BERUNIFORM atau UNIFORM
  const semua = [
    ['Badan Beruniform', fk(/UNIF|BERUNI|UNIFOM/), () => katFor(/unif/i, null, 'Badan Beruniform', '#7B6B2C', true)],
    ['Kelab dan Persatuan', fk(/KELAB|PERSATUAN/), () => katFor(/kelab|persatuan/i, null, 'Kelab dan Persatuan', '#4152A6', true)],
    ['Sukan dan Permainan', fk(/SUKAN|PERMAINAN/, /RUMAH/), () => katFor(/sukan|permainan/i, /rumah/i, 'Sukan dan Permainan', '#1B7F55', true)],
    ['Rumah Sukan', fk(/RUMAH/), () => katFor(/rumah/i, null, 'Rumah Sukan', '#B23A48', false)]
  ];
  const lajurTiada = semua.filter(x => x[1] < 0).map(x => x[0]);
  const lajur = semua.filter(x => x[1] >= 0).map(x => ({ col: x[1], kat: x[2](), tajuk: HK[x[1]] }));
  if (!lajur.length) throw fail_('Lajur unit (Beruniform, Kelab, Sukan, Rumah) tidak dijumpai dalam tab ' + tK.getName() + '.');
  // Kategori Rumah Sukan sedia ada: pastikan tidak menuntut rekod perjumpaan
  const kRumah = lajur.find(l => /rumah/i.test(l.kat.Nama_Kategori));
  if (kRumah && kRumah.kat.Rekod_Perjumpaan !== 'Tidak' && kats.indexOf(kRumah.kat) >= 0) { kRumah.kat.Rekod_Perjumpaan = 'Tidak'; writeAll_('Kategori'); }
  insertMany_('Kategori', addK);

  /* 4. Unit (cipta jika tiada) */
  const tUnit = table_('Unit');
  const unitIdx = {};
  tUnit.rows.forEach(u => { unitIdx[u.ID_Kategori + '|' + normUnit_(u.Nama_Unit)] = u; });
  const addU = [];
  let unitUbah = false;
  const unitFor = (kat, nama) => {
    if (/rumah/i.test(kat.Nama_Kategori) && !/^rumah/i.test(nama)) nama = 'Rumah ' + nama;
    const key = kat.ID_Kategori + '|' + normUnit_(nama);
    let u = unitIdx[key];
    if (!u) {
      u = { ID_Unit: uid_('U'), ID_Kategori: kat.ID_Kategori, Nama_Unit: namaUnitCantik_(nama), ID_Folder_Drive: '', Status: 'Aktif' };
      unitIdx[key] = u; addU.push(u);
    } else if (u.Status === 'Tidak Aktif') { u.Status = 'Aktif'; unitUbah = true; }
    return u;
  };

  /* 5. Padanan nama dan pendaftaran */
  const tEn = table_('Enrolmen');
  const enIdx = {};
  tEn.rows.forEach(e => { if (e.Tahun === sesi && e.Status === 'Aktif') enIdx[e.ID_Murid + '|' + e.ID_Kategori] = e; });
  const addE = [];
  let didaftar = 0, dipindah = 0, sama = 0, enUbah = false, dipadan = 0;
  const tidakJumpa = [], berulang = [], bukanTahun = [], hampir = [];
  const kosong = /^(-|–|TIADA|NIL|N\/A|NA|0)?$/i;
  const seenMurid = {};
  for (let i = hk + 1; i < vk.length; i++) {
    const r = vk[i];
    const nama = clean_(r[kN], 150);
    if (!nama) continue;
    const baris = i + 1, n = normNama_(nama);
    let c = idx[n] || idxC[n.replace(/ /g, '')] || [];
    let fuzzy = false;
    if (!c.length) {
      const nc = n.replace(/ /g, ''), lim = nc.length >= 15 ? 2 : 1;
      let best = 99, list = [];
      (byFirst[n.charAt(0)] || []).forEach(x => {
        const xc = x.norm.replace(/ /g, '');
        if (Math.abs(xc.length - nc.length) > lim) return;
        const d = levenshtein_(nc, xc);
        if (d < best) { best = d; list = [x]; } else if (d === best) list.push(x);
      });
      if (best <= lim) { c = list; fuzzy = true; }
    }
    if (c.length > 1 && kK >= 0 && r[kK]) {
      const kl = normUnit_(r[kK]);
      const f = c.filter(x => normUnit_(x.kelas) === kl || normUnit_(x.kelas).slice(-kl.length) === kl);
      if (f.length) c = f;
    }
    if (!c.length) { tidakJumpa.push({ baris: baris, nama: nama }); continue; }
    if (c.length > 1) { berulang.push({ baris: baris, nama: nama, calon: c.map(x => x.kelas).join(', ') }); continue; }
    const m = c[0];
    if (!layakT(m)) { bukanTahun.push({ baris: baris, nama: nama, kelas: m.kelas }); continue; }
    if (fuzzy) hampir.push({ baris: baris, koko: nama, murid: m.nama });
    if (seenMurid[m.id]) { berulang.push({ baris: baris, nama: nama, calon: 'nama sama pada baris ' + seenMurid[m.id] }); continue; }
    seenMurid[m.id] = baris;
    dipadan++;
    lajur.forEach(l => {
      const v = clean_(r[l.col], 80);
      if (kosong.test(v)) return;
      const u = unitFor(l.kat, v);
      const key = m.id + '|' + l.kat.ID_Kategori;
      const e = enIdx[key];
      if (e) {
        if (e.ID_Unit === u.ID_Unit) { sama++; if (e.Kelas !== m.kelas || e.Tarikh_Masuk) { e.Kelas = m.kelas; e.Tarikh_Masuk = ''; enUbah = true; } }
        else {
          // Tukar unit: rekod lama ditutup, rekod baharu dicipta (sejarah keahlian kekal)
          e.Status = 'Keluar'; enUbah = true; dipindah++;
          const ne = { ID_Enrolmen: uid_('E'), Tahun: sesi, ID_Murid: m.id, ID_Unit: u.ID_Unit, ID_Kategori: l.kat.ID_Kategori, Kelas: m.kelas, Status: 'Aktif', Tarikh_Masuk: '' };
          enIdx[key] = ne; addE.push(ne);
        }
      } else {
        const ne = { ID_Enrolmen: uid_('E'), Tahun: sesi, ID_Murid: m.id, ID_Unit: u.ID_Unit, ID_Kategori: l.kat.ID_Kategori, Kelas: m.kelas, Status: 'Aktif', Tarikh_Masuk: '' };
        enIdx[key] = ne; addE.push(ne); didaftar++;
      }
    });
  }
  if (unitUbah) writeAll_('Unit');
  insertMany_('Unit', addU);
  if (enUbah) writeAll_('Enrolmen');
  insertMany_('Enrolmen', addE);
  log_(who, 'Daftar murid DATA KOKO', 'Enrolmen', '', '', 'dipadan ' + dipadan + ', baru ' + didaftar + ', dipindah ' + dipindah + ', unit baru ' + addU.length);
  const cap = a => a.slice(0, 150);
  return {
    tabMurid: tM.getName(), tabKoko: tK.getName(), sesi: sesi, tahun: CONFIG.TAHUN_KOKO,
    lajur: lajur.map(l => ({ tajuk: l.tajuk, kategori: l.kat.Nama_Kategori })), lajurTiada: lajurTiada,
    muridLayak: recs.filter(layakT).length, muridBaru: muridBaru, muridKemas: muridKemas,
    dipadan: dipadan, didaftar: didaftar, dipindah: dipindah, sama: sama,
    kategoriBaru: addK.map(k => k.Nama_Kategori), unitBaru: addU.map(u => u.Nama_Unit),
    tidakJumpa: cap(tidakJumpa), berulang: cap(berulang), bukanTahun: cap(bukanTahun), hampir: cap(hampir),
    jumlah: { tidakJumpa: tidakJumpa.length, berulang: berulang.length, bukanTahun: bukanTahun.length, hampir: hampir.length }
  };
}

function a_importKoko(d, ctx) {
  needSuper_(ctx);
  return importKokoCore_(ctx.id);
}

/** Menu Sheets: daftar murid tanpa membuka aplikasi. */
function importKokoDariTab() {
  const ui = SpreadsheetApp.getUi();
  try {
    const r = withLock_(() => importKokoCore_('menu'));
    const senarai = (t, a, f) => a.length ? '\n\n' + t + ' (' + a.length + '):\n' + a.slice(0, 12).map(f).join('\n') + (a.length > 12 ? '\n…' : '') : '';
    ui.alert('Pendaftaran murid selesai (sesi ' + r.sesi + ')\n\n' +
      (r.lajurTiada.length ? 'AMARAN: lajur ' + r.lajurTiada.join(', ') + ' tidak dijumpai dalam ' + r.tabKoko + '. Semak tajuk lajur.\n\n' : '') +
      'Lajur dibaca: ' + r.lajur.map(l => l.tajuk + ' = ' + l.kategori).join('; ') + '\n\n' +
      r.muridLayak + ' murid Tahun ' + r.tahun.join(', ') + ' dalam ' + r.tabMurid + ' (' + r.muridBaru + ' baharu, ' + r.muridKemas + ' dikemas kini)\n' +
      r.dipadan + ' murid dalam ' + r.tabKoko + ' berjaya dipadankan\n' + r.didaftar + ' pendaftaran unit baharu, ' + r.dipindah + ' bertukar unit, ' + r.sama + ' tiada perubahan' +
      (r.unitBaru.length ? '\n\nUnit baharu dicipta: ' + r.unitBaru.join(', ') : '') +
      senarai('Tidak dijumpai dalam ' + r.tabMurid, r.tidakJumpa, x => 'Baris ' + x.baris + ': ' + x.nama) +
      senarai('Nama berulang, perlu didaftar manual', r.berulang, x => 'Baris ' + x.baris + ': ' + x.nama + ' (' + x.calon + ')') +
      senarai('Bukan Tahun ' + r.tahun.join('/'), r.bukanTahun, x => 'Baris ' + x.baris + ': ' + x.nama + ' (' + x.kelas + ')') +
      senarai('Dipadankan secara hampir, sila semak', r.hampir, x => 'Baris ' + x.baris + ': ' + x.koko + ' = ' + x.murid));
  } catch (e) { ui.alert(String(e.message || e)); }
}

/* ---------- Import pantas terus dalam Sheets (tanpa web app) ---------- */

const TAHUN_PERKATAAN = { SATU: 1, DUA: 2, TIGA: 3, EMPAT: 4, LIMA: 5, ENAM: 6 };

function parseTahun_(v) {
  const s = String(v || '').toUpperCase();
  const m = s.match(/\d/);
  if (m) return m[0];
  const w = Object.keys(TAHUN_PERKATAAN).find(k => s.indexOf(k) >= 0);
  return w ? String(TAHUN_PERKATAAN[w]) : '';
}

function namaKunci_(nama) { return 'N' + sha_(String(nama).toUpperCase().replace(/\s+/g, ' ').trim()).slice(0, 20); }

function importDariTab() {
  const ui = SpreadsheetApp.getUi();
  const sh = ss_().getSheetByName('APDM_Mentah');
  if (!sh) { ui.alert('Cipta tab bernama APDM_Mentah dan tampal data eksport APDM/iDME (termasuk baris tajuk).'); return; }
  const v = sh.getDataRange().getDisplayValues();
  let hi = -1;
  for (let i = 0; i < Math.min(15, v.length); i++) { if (v[i].some(c => /NAMA/i.test(c))) { hi = i; break; } }
  if (hi < 0) { ui.alert('Baris tajuk yang mengandungi lajur NAMA tidak dijumpai.'); return; }
  const h = v[hi].map(c => String(c).toUpperCase().trim());
  const find = (re, not) => h.findIndex(c => re.test(c) && !(not && not.test(c)));
  const cNama = find(/NAMA/, /(IBU|BAPA|PENJAGA|SEKOLAH|KELAS|GURU)/);
  const cKelas = find(/KELAS/, /(GURU)/);
  const cTahun = find(/(TAHUN|TINGKATAN)/, /(LAHIR|MASUK|DAFTAR)/);
  const cJan = find(/JANTINA/);
  const cIc = find(/(MYKID|NO\.?\s*K\/?P|KAD PENGENALAN|SIJIL LAHIR|NO\.?\s*PENGENALAN)/);
  if (cNama < 0 || cKelas < 0) { ui.alert('Lajur NAMA dan KELAS diperlukan.'); return; }
  const list = [];
  for (let i = hi + 1; i < v.length; i++) {
    const r = v[i];
    const nama = String(r[cNama] || '').trim();
    if (!nama) continue;
    const tahun = cTahun >= 0 ? parseTahun_(r[cTahun]) : '';
    let kelas = String(r[cKelas] || '').trim();
    if (kelas && !/^\d/.test(kelas) && tahun) kelas = tahun + ' ' + kelas;
    let ic = cIc >= 0 ? String(r[cIc] || '').replace(/\D/g, '') : '';
    if (ic.length === 11) ic = '0' + ic;   // sifar di hadapan hilang jika sel berformat nombor
    list.push({ k: ic.length >= 6 ? 'K' + sha_(ic).slice(0, 20) : namaKunci_(nama), nama: nama, kelas: kelas, jantina: cJan >= 0 ? r[cJan] : '' });
  }
  const res = withLock_(() => importCore_(list, false, 'menu'));
  ui.alert('Import selesai: ' + res.baru + ' murid baharu, ' + res.kemas + ' dikemas kini, ' + res.sama + ' tiada perubahan.\n\nNombor kad pengenalan tidak disimpan; hanya cincangan (hash) digunakan sebagai kunci.');
}
