/**
 * =====================================================================
 *  SISTEM SPL & SPD — PT CITRA BANJAR ABADI
 *  File : Inti.gs   (TAHAP 1 — fungsi inti, dipakai semua tahap)
 *  Isi  : baca master & parameter, hitung lembur, hitung SPD,
 *         format nomor surat, log aksi.
 *  Catatan: JANGAN ubah nama sheet / judul kolom di spreadsheet,
 *           script mencari kolom berdasarkan judulnya.
 * =====================================================================
 */

const SH = {
  KARYAWAN: 'Karyawan', PARAM: 'Parameter', PILIHAN: 'Pilihan', LIBUR: 'Libur',
  SPL: 'Data SPL', SPD: 'Data SPD', PERSONEL: 'Personel SPD', BIAYA: 'Biaya SPD', LOG: 'Log'
};
const TZ = 'Asia/Jakarta';
const PERAN = ['KARYAWAN', 'ADMIN'];   // ADMIN = boleh login form & membuat surat
const NAMA_HARI = ['Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu', 'Minggu'];

/* ---------- Akses spreadsheet (aman juga saat dipanggil dari Web App) ---------- */
function ss_() {
  const aktif = SpreadsheetApp.getActiveSpreadsheet();
  if (aktif) return aktif;
  const id = PropertiesService.getScriptProperties().getProperty('SS_ID');
  if (!id) throw new Error('SS_ID belum tersimpan. Jalankan menu "Setup awal" dari spreadsheet.');
  return SpreadsheetApp.openById(id);
}
/** Nama lain yang juga diterima untuk sheet tertentu */
const SH_ALTERNATIF = { 'Karyawan': ['Data Karyawan', 'Master Karyawan'] };
function cariSheet_(ss, nama) {
  return [nama].concat(SH_ALTERNATIF[nama] || []).map(n => ss.getSheetByName(n)).find(Boolean) || null;
}
function sheet_(nama) {
  const sh = cariSheet_(ss_(), nama);
  if (!sh) throw new Error('Sheet "' + nama + '" tidak ditemukan.');
  return sh;
}

/** Baca sheet berjudul -> array objek {judulKolom: nilai}. Baris tanpa kolom kunci dilewati. */
function bacaTabel_(namaSheet, kolomKunci) {
  const v = sheet_(namaSheet).getDataRange().getValues();
  const head = v.shift().map(String);
  const iKunci = kolomKunci ? head.indexOf(kolomKunci) : 0;
  return v.filter(r => String(r[iKunci]).trim() !== '')
          .map(r => Object.fromEntries(head.map((h, i) => [h, r[i]])));
}
/** Peta judul kolom -> nomor kolom (1-based) */
function kolom_(namaSheet) {
  const sh = sheet_(namaSheet);
  const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const m = {}; head.forEach((h, i) => { if (h) m[String(h)] = i + 1; });
  return m;
}

/* ---------- Parameter ---------- */
let _P = null;
function param_() {
  if (_P) return _P;
  _P = {};
  sheet_(SH.PARAM).getDataRange().getValues().slice(1).forEach(r => {
    const k = String(r[0]).trim(); if (!k) return;
    let v = r[1];
    if (v instanceof Date) v = Utilities.formatDate(v, TZ, 'HH:mm'); // jam yang terbaca sebagai tanggal
    _P[k] = v;
  });
  return _P;
}

/* ---------- Master karyawan ---------- */
let _K = null;
/** Departemen dari Organization: teks sebelum " GROUP", ejaan DEPARTMENT diseragamkan jadi DEPARTEMENT */
function departemenDari_(org) {
  const o = String(org || '').toUpperCase().trim().replace(/DEPARTMENT/g, 'DEPARTEMENT');
  const i = o.indexOf(' GROUP');
  return i > 0 ? o.slice(0, i).trim() : o;
}
/** Peran default bila kolom Peran kosong: semua KARYAWAN (peran lain HANYA bila diisi manual) */
function peranDefault_(jabatan) {
  return 'KARYAWAN';
}

/** Master karyawan. Kolom wajib: Employee ID, Full Name, Job Position, Organization, Email.
 *  Opsional: Basic Salary, Departemen, Peran (kosong = KARYAWAN), Atasan (Employee ID atasan langsung),
 *            Aktif (YA/TIDAK) atau Status Employee (Resign/Keluar/Non Aktif = tidak aktif). */
function karyawan_() {
  if (_K) return _K;
  _K = {};
  bacaTabel_(SH.KARYAWAN, 'Employee ID').forEach(r => {
    const id = String(r['Employee ID']).trim().toUpperCase();
    const jab = String(r['Job Position'] || '').trim();
    const statusEmp = String(r['Status Employee'] || '').trim();
    const aktifKol = String(r['Aktif'] || '').trim().toUpperCase();
    _K[id] = {
      id: id,
      nama: String(r['Full Name'] || '').trim(),
      jabatan: jab,
      organisasi: String(r['Organization'] || '').trim(),
      departemen: String(r['Departemen'] || '').trim() || departemenDari_(r['Organization']),
      gaji: Number(r['Basic Salary']) || 0,           // RAHASIA — jangan dikirim ke form/email karyawan
      email: String(r['Email'] || '').trim().toLowerCase(),
      peran: String(r['Peran'] || '').trim().toUpperCase() || peranDefault_(jab),
      atasan: String(r['Atasan'] || '').trim().toUpperCase(),   // Employee ID atasan langsung
      aktif: aktifKol ? aktifKol !== 'TIDAK' : !/resign|keluar|berhenti|non ?aktif|tidak aktif|phk/i.test(statusEmp)
    };
  });
  return _K;
}
function cariKaryawanByEmail_(email) {
  email = String(email || '').trim().toLowerCase();
  return Object.values(karyawan_()).find(k => k.email && k.email === email) || null;
}
/* ---------- Libur & hari ---------- */
let _L = null;
function liburSet_() {
  if (_L) return _L;
  _L = {};
  sheet_(SH.LIBUR).getDataRange().getValues().slice(1).forEach(r => {
    if (r[0] instanceof Date) _L[ymd_(r[0])] = String(r[2] || 'Libur');
  });
  return _L;
}

/* ===================== FUNGSI MURNI (bisa diuji tanpa spreadsheet) ===================== */

function ymd_(d) {
  if (typeof Utilities !== 'undefined') return Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
  return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
}
/** 1 = Senin ... 7 = Minggu (sama dengan WEEKDAY(x,2) di Excel) */
function hariKe_(d) {
  if (typeof Utilities !== 'undefined') return Number(Utilities.formatDate(d, TZ, 'u'));
  return ((d.getDay() + 6) % 7) + 1;
}
function keTanggal_(v) {
  if (v instanceof Date) return v;
  const m = String(v || '').match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) {
    // jam 12 siang WIB: aman dari geser tanggal walau zona waktu project Apps Script berbeda
    if (typeof Utilities !== 'undefined') return Utilities.parseDate(m[1] + '-' + m[2] + '-' + m[3] + ' 12:00', TZ, 'yyyy-MM-dd HH:mm');
    return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), 12);
  }
  return null;
}

/** "16:00" / "16.00" / "16" / 0.6667 / Date -> angka jam (16). "24:00" -> 24. */
function jamKeAngka_(v) {
  if (v === '' || v === null || v === undefined) return null;
  if (v instanceof Date) {
    const s = (typeof Utilities !== 'undefined') ? Utilities.formatDate(v, TZ, 'H:mm') : v.getHours() + ':' + v.getMinutes();
    return jamKeAngka_(s);
  }
  if (typeof v === 'number') return v <= 1 ? Math.round(v * 24 * 60) / 60 : v;
  const m = String(v).trim().match(/^(\d{1,2})(?:[:.](\d{2}))?$/);
  if (!m) return null;
  const h = Number(m[1]) + Number(m[2] || 0) / 60;
  return h >= 0 && h <= 24 ? h : null;
}
function angkaKeJam_(h) {
  if (h === null || h === undefined) return '';
  const j = Math.floor(h), mnt = Math.round((h - j) * 60);
  return String(j).padStart(2, '0') + ':' + String(mnt).padStart(2, '0');
}
function rentang_(s) { // "18:00-19:00" -> [18,19]
  const p = String(s).split('-').map(x => jamKeAngka_(x));
  return [p[0], p[1]];
}

function statusHari_(tgl, libur) {
  const t = keTanggal_(tgl);
  if (!t) return '';
  return (hariKe_(t) === 7 || libur[ymd_(t)]) ? 'Hari Libur' : 'Hari Kerja';
}

/**
 * Hitung 1 baris lembur — logika sama dengan rumus sheet SPL di Excel:
 *  - Istirahat dipotong 1 jam bila lembur mencakup penuh 18–19 (hari kerja) / 12–13 (hari libur)
 *  - Jam efektif = selisih jam - istirahat (min 0)
 *  - Kompensasi  = Basic Salary / 173 x (1,5 hari kerja | 2 hari libur) x jam efektif
 * @param {{tanggal, mulai, selesai, gaji}} x
 * @return {{hari, statusHari, istirahat, jamEfektif, kompensasi, error:string[], peringatan:string[]}}
 */
function hitungLembur_(x, P, libur) {
  const out = { hari: '', statusHari: '', istirahat: '-', jamEfektif: 0, kompensasi: 0, error: [], peringatan: [] };
  const t = keTanggal_(x.tanggal);
  if (!t) { out.error.push('Tanggal lembur wajib diisi.'); return out; }
  out.hari = NAMA_HARI[hariKe_(t) - 1];
  out.statusHari = statusHari_(t, libur);
  const libur_ = out.statusHari === 'Hari Libur';

  const m = jamKeAngka_(x.mulai), s = jamKeAngka_(x.selesai);
  if (m === null || s === null) { out.error.push('Jam mulai & jam selesai wajib diisi (format JJ:MM).'); return out; }

  const minMulai = jamKeAngka_(libur_ ? P.JAM_MULAI_MIN_LIBUR : (hariKe_(t) === 6 ? P.JAM_MULAI_MIN_SABTU : P.JAM_MULAI_MIN_SENJUM));
  const maksSelesai = jamKeAngka_(P.JAM_SELESAI_MAKS);
  if (m < minMulai) out.error.push('Jam mulai paling awal ' + angkaKeJam_(minMulai) + ' untuk ' + (libur_ ? 'hari libur' : out.hari) + '.');
  if (s > maksSelesai) out.error.push('Jam selesai paling akhir ' + angkaKeJam_(maksSelesai) + '.');
  if (s <= m) out.error.push('Jam selesai harus setelah jam mulai.');

  const ist = rentang_(libur_ ? P.ISTIRAHAT_HARI_LIBUR : P.ISTIRAHAT_HARI_KERJA);
  const kenaIst = m <= ist[0] && s >= ist[1];
  out.istirahat = kenaIst ? angkaKeJam_(ist[0]).replace(':', '.') + ' - ' + angkaKeJam_(ist[1]).replace(':', '.') : '-';
  const durasi = ((s - m) % 24 + 24) % 24;
  out.jamEfektif = Math.max(durasi - (kenaIst ? (ist[1] - ist[0]) : 0), 0);

  if (libur_ && s > jamKeAngka_(P.BATAS_SELESAI_LIBUR))
    out.peringatan.push('Hari libur selesai lewat ' + P.BATAS_SELESAI_LIBUR + ' — melebihi batas UU Ketenagakerjaan.');

  const faktor = Number(libur_ ? P.FAKTOR_HARI_LIBUR : P.FAKTOR_HARI_KERJA);
  out.kompensasi = x.gaji ? Math.round(x.gaji / Number(P.PEMBAGI_TARIF) * faktor * out.jamEfektif) : 0;
  return out;
}

/**
 * Hitung SPD — sama dengan sheet SPD Excel.
 * @param {{berangkat, kembali, personel:[{lama}], biaya:[{jumlah}]}} x
 */
function hitungSPD_(x, P) {
  const out = { lama: 0, totalUangSaku: 0, totalBiayaLain: 0, totalBiaya: 0, personel: [], error: [], peringatan: [] };
  const b = keTanggal_(x.berangkat), k = keTanggal_(x.kembali);
  if (!b || !k) out.error.push('Tanggal berangkat & kembali wajib diisi.');
  else {
    out.lama = Math.max(Math.round((k - b) / 86400000) + 1, 0);
    if (k < b) out.error.push('Tanggal kembali lebih awal dari tanggal berangkat.');
    if (out.lama > Number(P.BATAS_LAMA_DINAS)) out.peringatan.push('Lama dinas melebihi batas ' + P.BATAS_LAMA_DINAS + ' hari — perlu persetujuan khusus.');
  }
  const per = (x.personel || []);
  if (per.length === 0) out.error.push('Minimal 1 personel.');
  if (per.length > Number(P.MAKS_PERSONEL_SPD)) out.error.push('Maksimal ' + P.MAKS_PERSONEL_SPD + ' personel per SPD — buat SPD baru.');
  const ids = {};
  per.forEach((p, i) => {
    const lama = Number(p.lama);
    if (ids[p.id]) out.error.push('Employee ID ' + p.id + ' tercatat lebih dari sekali.');
    ids[p.id] = true;
    if (!(lama >= 1 && lama === Math.floor(lama))) out.error.push('Lama (hari) personel ke-' + (i + 1) + ' belum diisi / tidak valid.');
    else if (out.lama && lama > out.lama) out.error.push('Lama (hari) personel ke-' + (i + 1) + ' melebihi lama dinas.');
    const uang = (lama || 0) * Number(P.UANG_SAKU_HARI);
    out.personel.push(Object.assign({}, p, { uangSaku: uang }));
    out.totalUangSaku += uang;
  });
  (x.biaya || []).forEach(bi => out.totalBiayaLain += Number(bi.jumlah) || 0);
  out.totalBiaya = out.totalUangSaku + out.totalBiayaLain;
  return out;
}

function romawi_(n) {
  return ['', 'I', 'II', 'III', 'IV', 'V', 'VI', 'VII', 'VIII', 'IX', 'X', 'XI', 'XII'][n];
}
/** "007 / SPL / CBA / X / 2026" */
function formatNomor_(urut, jenis, kode, tgl) {
  const t = keTanggal_(tgl);
  const [th, bl] = ymd_(t).split('-').map(Number);
  return String(urut).padStart(3, '0') + ' / ' + jenis + ' / ' + kode + ' / ' + romawi_(bl) + ' / ' + th;
}
/** No. urut berikutnya dalam tahun yang sama (ambil MAX + 1, bukan jumlah — aman bila ada surat dihapus). */
function urutBerikutnya_(daftarNomor, tahun) {
  let maks = 0;
  daftarNomor.forEach(n => {
    const p = String(n).split('/').map(s => s.trim());
    if (p.length === 5 && Number(p[4]) === Number(tahun)) maks = Math.max(maks, Number(p[0]) || 0);
  });
  return maks + 1;
}
/* ===================== akhir fungsi murni ===================== */

/** Nomor surat berikutnya dari database. jenis: 'SPL' | 'SPD'.
 *  (Tahap 2 akan memanggil ini di dalam LockService agar tidak dobel.) */
function nomorBerikutnya_(jenis, tgl) {
  const sh = sheet_(jenis === 'SPL' ? SH.SPL : SH.SPD);
  const n = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().map(r => r[0]) : [];
  const t = keTanggal_(tgl);
  return formatNomor_(urutBerikutnya_(n, Number(ymd_(t).slice(0, 4))), jenis, param_().KODE_PERUSAHAAN, t);
}

/** Catat aksi ke sheet Log */
function logAksi_(jenis, nomor, aksi, oleh, ket) {
  try {
    sheet_(SH.LOG).appendRow([new Date(), jenis, nomor, aksi, oleh || Session.getActiveUser().getEmail(), ket || '']);
  } catch (e) { console.error('Gagal tulis log: ' + e); }
}

// agar bisa diuji di Node.js (diabaikan oleh Apps Script)
if (typeof module !== 'undefined') module.exports = { hitungLembur_, hitungSPD_, formatNomor_, urutBerikutnya_, jamKeAngka_, angkaKeJam_, statusHari_ };
/**
 * =====================================================================
 *  SISTEM SPL & SPD — File : Menu.gs   (TAHAP 1)
 *  Menu "⚙️ SPL / SPD" di spreadsheet + setup awal + cek data.
 * =====================================================================
 */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('⚙️ SPL / SPD')
    .addItem('1. Setup awal (jalankan sekali)', 'setupAwal')
    .addItem('2. Cek data karyawan', 'cekDataKaryawan')
    .addItem('3. Tes hitung lembur & SPD', 'tesHitung')
    .addSeparator()
    .addItem('🔗 Link form SPL & SPD', 'bukaForm')
    .addSeparator()
    .addItem('✉ Kirim ulang email konfirmasi SPL', 'menuKirimUlangKonfirmasi')
    .addItem('📄 Buat & kirim PDF SPL sekarang', 'menuPdfSPL')
    .addItem('Lihat nomor surat berikutnya', 'lihatNomorBerikutnya')
    .addToUi();
}

/** Setup awal: zona waktu, simpan ID file, kunci sheet sistem, format kolom jam. */
function setupAwal() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone(TZ);
  PropertiesService.getScriptProperties().setProperty('SS_ID', ss.getId());

  // Cek semua sheet yang dibutuhkan ada
  const hilang = Object.values(SH).filter(n => !cariSheet_(ss, n));
  if (hilang.length) {
    SpreadsheetApp.getUi().alert('Sheet berikut tidak ditemukan:\n• ' + hilang.join('\n• ') +
      '\n\nPastikan nama sheet sesuai (sheet karyawan boleh bernama "Karyawan" atau "Data Karyawan").');
    return;
  }

  // Kolom jam di Data SPL disimpan sebagai teks (mis. "16:00") agar tidak bergeser zona waktu
  const c = kolom_(SH.SPL);
  [c['Jam Mulai'], c['Jam Selesai']].forEach(k => ss.getSheetByName(SH.SPL).getRange(2, k, 999, 1).setNumberFormat('@'));

  // Sheet sistem: beri peringatan bila diedit manual (script tetap bisa menulis)
  [SH.SPL, SH.SPD, SH.PERSONEL, SH.BIAYA, SH.LOG].forEach(n => {
    const sh = ss.getSheetByName(n);
    sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => p.remove());
    sh.protect().setDescription('Diisi otomatis oleh sistem SPL/SPD').setWarningOnly(true);
  });
  // Cek kolom wajib di sheet karyawan
  const kk = kolom_(SH.KARYAWAN);
  const kurang = ['Employee ID', 'Full Name', 'Job Position', 'Organization', 'Email'].filter(j => !kk[j]);
  if (kurang.length) SpreadsheetApp.getUi().alert('⚠ Kolom berikut belum ada di sheet karyawan:\n• ' + kurang.join('\n• ') +
    '\n\nJudul kolom harus persis seperti di atas.');
  // Kolom Peran & Atasan di sheet karyawan (ditambahkan bila belum ada)
  const shK = sheet_(SH.KARYAWAN);
  // Kolom Peran: dropdown ADMIN / KARYAWAN (kosong = KARYAWAN)
  let colPeran = kk['Peran'];
  if (!colPeran) { colPeran = shK.getLastColumn() + 1; shK.getRange(1, colPeran).setValue('Peran').setFontWeight('bold'); }
  shK.getRange(2, colPeran, shK.getMaxRows() - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(PERAN, true).setAllowInvalid(false).build());
  if (!kk['Atasan']) {
    const col = shK.getLastColumn() + 1;
    shK.getRange(1, col).setValue('Atasan').setFontWeight('bold')
       .setNote('Isi Employee ID atasan langsung (mis. D085). Atasan inilah yang menerima email persetujuan lembur.');
  }
  // Kolom tambahan Data SPL untuk alur email
  pastikanKolom_(SH.SPL, KOLOM_SPL_TAMBAHAN);
  rahasia_();

  logAksi_('-', '-', 'SETUP AWAL', Session.getActiveUser().getEmail(), 'Zona waktu ' + TZ + ', proteksi sheet sistem.');
  SpreadsheetApp.getUi().alert('✅ Setup awal selesai.\n\n• Zona waktu: ' + TZ +
    '\n• Sheet sistem dikunci (peringatan bila diedit manual)\n• Kolom Peran & Atasan tersedia di sheet karyawan:\n   – Peran: isi ADMIN untuk user yang boleh membuat surat (kosong = KARYAWAN)\n   – Atasan: isi Employee ID atasan langsung\n\nLangkah berikut: menu ⚙️ SPL / SPD → 2. Cek data karyawan.');
}

/** Periksa kelengkapan master karyawan sebelum notifikasi email diaktifkan. */
function cekDataKaryawan() {
  _K = null; _P = null;
  const K = Object.values(karyawan_());
  const P = param_();
  const aktif = K.filter(k => k.aktif);
  const reEmail = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  const masalah = [];   // {jenis, isi}

  // Duplikat ID (karyawan_() menimpa ID dobel, jadi hitung dari sheet langsung)
  const ids = bacaTabel_(SH.KARYAWAN, 'Employee ID').map(r => String(r['Employee ID']).trim().toUpperCase());
  const dobelId = ids.filter((x, i) => ids.indexOf(x) !== i);
  if (dobelId.length) masalah.push({ j: '❌ Employee ID dobel', isi: [...new Set(dobelId)] });

  const tanpaEmail = aktif.filter(k => !k.email);
  if (tanpaEmail.length) masalah.push({ j: '⚠ Belum ada email (' + tanpaEmail.length + ' orang)', isi: tanpaEmail.map(k => k.id + ' — ' + k.nama) });

  const emailSalah = aktif.filter(k => k.email && !reEmail.test(k.email));
  if (emailSalah.length) masalah.push({ j: '❌ Format email salah', isi: emailSalah.map(k => k.id + ' — ' + k.email) });

  const em = aktif.filter(k => k.email).map(k => k.email);
  const dobelEmail = em.filter((x, i) => em.indexOf(x) !== i);
  if (dobelEmail.length) masalah.push({ j: '❌ Email dipakai lebih dari 1 karyawan', isi: [...new Set(dobelEmail)] });

  const peranSalah = aktif.filter(k => PERAN.indexOf(k.peran) < 0);
  if (peranSalah.length) masalah.push({ j: '❌ Peran tidak dikenal', isi: peranSalah.map(k => k.id + ' — ' + k.peran) });

  const admin = aktif.filter(k => k.peran === 'ADMIN');
  if (!admin.length) masalah.push({ j: '❌ Belum ada user ADMIN', isi: ['Isi Peran = ADMIN untuk user yang membuat SPL / SPD.'] });
  const adminTanpaEmail = admin.filter(k => !k.email);
  if (adminTanpaEmail.length) masalah.push({ j: '❌ ADMIN tanpa email (tidak bisa login form)', isi: adminTanpaEmail.map(k => k.id + ' — ' + k.nama) });

  // Atasan langsung
  const Kmap = karyawan_();
  const tanpaAtasan = aktif.filter(k => !k.atasan);
  if (tanpaAtasan.length) masalah.push({ j: '⚠ Atasan belum diisi (' + tanpaAtasan.length + ' orang) — persetujuan lembur mereka dikirim ke admin', isi: tanpaAtasan.map(k => k.id + ' — ' + k.nama) });
  const atasanSalah = aktif.filter(k => k.atasan && !Kmap[k.atasan]);
  if (atasanSalah.length) masalah.push({ j: '❌ Employee ID atasan tidak ditemukan', isi: atasanSalah.map(k => k.id + ' — ' + k.nama + ' → ' + k.atasan) });
  const atasanTanpaEmail = aktif.filter(k => k.atasan && Kmap[k.atasan] && !Kmap[k.atasan].email);
  if (atasanTanpaEmail.length) masalah.push({ j: '⚠ Atasan belum punya email — persetujuan dikirim ke admin', isi: atasanTanpaEmail.map(k => k.id + ' — ' + k.nama + ' → ' + k.atasan + ' ' + Kmap[k.atasan].nama) });
  const diriSendiri = aktif.filter(k => k.atasan && k.atasan === k.id);
  if (diriSendiri.length) masalah.push({ j: '❌ Atasan diisi diri sendiri', isi: diriSendiri.map(k => k.id + ' — ' + k.nama) });

  if (!P.EMAIL_ADMIN) masalah.push({ j: 'ℹ Parameter EMAIL_ADMIN kosong (opsional — salinan PDF SPL)', isi: ['Isi di sheet Parameter bila PDF SPL perlu dikirim juga ke email lain.'] });
  try { masalah.push({ j: 'ℹ Sisa kuota email hari ini: ' + MailApp.getRemainingDailyQuota(), isi: ['Gmail biasa ±100/hari, Google Workspace ±1.500/hari.'] }); } catch (e) {}

  const ringkas = {};
  aktif.forEach(k => ringkas[k.peran] = (ringkas[k.peran] || 0) + 1);

  let html = '<div style="font-family:Arial;font-size:13px">' +
    '<p><b>Karyawan aktif:</b> ' + aktif.length + ' &nbsp;|&nbsp; ' +
    Object.keys(ringkas).map(p => p + ': ' + ringkas[p]).join(' · ') + '</p>';
  if (!masalah.some(m => m.j.indexOf('❌') === 0)) html += '<p style="color:green;font-size:15px">✅ Tidak ada masalah kritis.</p>';
  masalah.forEach(m => {
    html += '<p style="margin:10px 0 4px"><b>' + m.j + '</b></p><div style="max-height:140px;overflow:auto;background:#f6f6f6;padding:6px;border-radius:4px">' +
      m.isi.map(x => '• ' + x).join('<br>') + '</div>';
  });
  html += '<p style="color:#777;margin-top:12px">Perbaiki di sheet Karyawan / Parameter, lalu jalankan cek ini lagi.</p></div>';
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(560).setHeight(520), 'Cek Data Karyawan');
}

/** Uji rumus lembur & SPD dengan contoh — hasil harus sama dengan Excel. */
function tesHitung() {
  _P = null; _L = null;
  const P = param_(), L = liburSet_();
  const kasus = [
    { ket: 'Rabu 16:00–22:00 (contoh di Excel)', tanggal: '2026-10-21', mulai: '16:00', selesai: '22:00', gaji: 4642000, harap: 5 },
    { ket: 'Rabu 16:00–19:00 (contoh di Excel)', tanggal: '2026-10-21', mulai: '16:00', selesai: '19:00', gaji: 4642000, harap: 2 },
    { ket: 'Sabtu 13:00–17:00', tanggal: '2026-10-24', mulai: '13:00', selesai: '17:00', gaji: 4642000, harap: 4 },
    { ket: 'Minggu 08:00–17:00 (potong 12–13)', tanggal: '2026-10-25', mulai: '08:00', selesai: '17:00', gaji: 4642000, harap: 8 },
    { ket: 'Minggu 08:00–20:00 (peringatan UU)', tanggal: '2026-10-25', mulai: '08:00', selesai: '20:00', gaji: 4642000, harap: 11 },
    { ket: 'Rabu mulai 14:00 (harus ditolak)', tanggal: '2026-10-21', mulai: '14:00', selesai: '18:00', gaji: 4642000, harap: 'error' }
  ];
  let html = '<table style="font-family:Arial;font-size:12px;border-collapse:collapse" border="1" cellpadding="4">' +
    '<tr style="background:#1F3864;color:#fff"><th>Kasus</th><th>Status Hari</th><th>Istirahat</th><th>Jam Efektif</th><th>Kompensasi</th><th>Catatan</th><th>OK?</th></tr>';
  kasus.forEach(k => {
    const h = hitungLembur_(k, P, L);
    const ok = k.harap === 'error' ? h.error.length > 0 : (h.error.length === 0 && h.jamEfektif === k.harap);
    html += '<tr><td>' + k.ket + '</td><td>' + h.statusHari + '</td><td>' + h.istirahat + '</td><td>' + h.jamEfektif +
      '</td><td style="text-align:right">' + h.kompensasi.toLocaleString('id-ID') + '</td><td>' + h.error.concat(h.peringatan).join('<br>') +
      '</td><td>' + (ok ? '✅' : '❌') + '</td></tr>';
  });
  html += '</table>';
  const spd = hitungSPD_({ berangkat: '2026-10-21', kembali: '2026-10-23', personel: [{ id: 'D085', lama: 3 }, { id: 'D074', lama: 2 }], biaya: [{ jumlah: 100000 }] }, P);
  html += '<p style="font-family:Arial;font-size:12px"><b>SPD</b> 21–23 Okt 2026, 2 orang (3 & 2 hari) + biaya lain 100.000 → lama ' + spd.lama +
    ' hari, uang saku ' + spd.totalUangSaku.toLocaleString('id-ID') + ', total ' + spd.totalBiaya.toLocaleString('id-ID') +
    (spd.lama === 3 && spd.totalBiaya === 600000 ? ' ✅' : ' ❌') + '</p>';
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(820).setHeight(400), 'Tes Hitung');
}

function lihatNomorBerikutnya() {
  const hari = new Date();
  SpreadsheetApp.getUi().alert('Nomor berikutnya (tanggal hari ini):\n\nSPL : ' + nomorBerikutnya_('SPL', hari) +
    '\nSPD : ' + nomorBerikutnya_('SPD', hari));
}
/**
 * =====================================================================
 *  SISTEM SPL & SPD — File : Web.gs
 *  Web App (form SPL & SPD): tampilan, login OTP email, data form,
 *  alur approval. Karyawan TIDAK perlu akses ke spreadsheet.
 *  Butuh: Inti.gs, Logo.gs, Spl.gs, Spd.gs, Form.html
 * =====================================================================
 */

const PERAN_PEMBUAT = ['ADMIN'];
const STATUS_BATAL = ['DITOLAK', 'DIBATALKAN'];
const SESI_DETIK = 6 * 3600;   // login berlaku 6 jam

/* ------------------------------ Tampilan ------------------------------ */
function doGet(e) {
  if (e && e.parameter && e.parameter.a) return halamanAksi_(e.parameter);   // link dari email (Notif.gs)
  const t = HtmlService.createTemplateFromFile('Form');
  // rumus yang SAMA dengan server dipakai juga di browser (satu sumber rumus)
  t.rumusBersama = 'const NAMA_HARI = ' + JSON.stringify(NAMA_HARI) + ';\n' +
    [ymd_, hariKe_, keTanggal_, jamKeAngka_, angkaKeJam_, rentang_, statusHari_, hitungLembur_, hitungSPD_]
      .map(f => f.toString()).join('\n');
  t.logo = LOGO_DATA_URI;
  return t.evaluate()
    .setTitle('SPL & SPD — ' + (param_().NAMA_PERUSAHAAN || ''))
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Menu spreadsheet: tampilkan link form */
function bukaForm() {
  const url = ScriptApp.getService().getUrl();
  const ui = SpreadsheetApp.getUi();
  if (!url) { ui.alert('Form belum di-deploy.\n\nDeploy → New deployment → Web app (lihat panduan).'); return; }
  ui.showModalDialog(HtmlService.createHtmlOutput(
    '<div style="font-family:Arial;font-size:13px">Link form SPL & SPD (bagikan ke user ADMIN):<br><br>' +
    '<input style="width:100%;padding:6px" value="' + url + '" onclick="this.select()" readonly><br><br>' +
    '<a href="' + url + '" target="_blank">Buka form ↗</a></div>').setWidth(520).setHeight(170), 'Form SPL & SPD');
}

/* ------------------------------ Login ------------------------------ */
function pembuatByEmail_(email) {
  const k = cariKaryawanByEmail_(email);
  return (k && k.aktif && PERAN_PEMBUAT.indexOf(k.peran) >= 0) ? k : null;
}
function infoUser_(k) { return { nama: k.nama, jabatan: k.jabatan, departemen: k.departemen, peran: k.peran, email: k.email }; }
function sesiBuat_(email) {
  const tok = Utilities.getUuid();
  CacheService.getScriptCache().put('S_' + tok, email, SESI_DETIK);
  return tok;
}
function sesiCek_(tok) {
  const email = tok ? CacheService.getScriptCache().get('S_' + tok) : null;
  if (!email) throw new Error('SESI_HABIS');
  const k = pembuatByEmail_(email);
  if (!k) throw new Error('Akun ini tidak berhak membuat SPL / SPD.');
  return k;
}

/** Jika form dibuka dengan akun Google yang emailnya terdaftar (umumnya Google Workspace), langsung masuk. */
function loginOtomatis() {
  let email = '';
  try { email = Session.getActiveUser().getEmail(); } catch (e) {}
  const k = email ? pembuatByEmail_(email) : null;
  return k ? { token: sesiBuat_(k.email), user: infoUser_(k) } : { token: null };
}

function kirimKode(email) {
  email = String(email || '').trim().toLowerCase();
  const k = pembuatByEmail_(email);
  if (!k) throw new Error('Email tidak terdaftar sebagai ADMIN. Hubungi pengelola sistem.');
  const c = CacheService.getScriptCache();
  const n = Number(c.get('N_' + email) || 0);
  if (n >= 5) throw new Error('Terlalu banyak permintaan kode. Coba lagi 1 jam lagi.');
  c.put('N_' + email, String(n + 1), 3600);
  const kode = String(Math.floor(100000 + Math.random() * 900000));
  c.put('OTP_' + email, kode, 600);
  c.remove('G_' + email);
  MailApp.sendEmail({
    to: email,
    subject: 'Kode masuk Form SPL & SPD: ' + kode,
    htmlBody: '<div style="font-family:Arial;font-size:14px">Halo ' + k.nama + ',<br><br>Kode masuk Form SPL & SPD Anda:<br>' +
      '<div style="font-size:28px;font-weight:bold;letter-spacing:6px;margin:12px 0;color:#BE353F">' + kode + '</div>' +
      'Berlaku 10 menit. Abaikan email ini jika Anda tidak meminta kode.<br><br>— ' + (param_().NAMA_PERUSAHAAN || '') + '</div>'
  });
  return 'Kode dikirim ke ' + email.replace(/^(.{2}).*(@.*)$/, '$1•••$2');
}

function verifikasiKode(email, kode) {
  email = String(email || '').trim().toLowerCase();
  const c = CacheService.getScriptCache();
  const salah = Number(c.get('G_' + email) || 0);
  if (salah >= 5) throw new Error('Kode salah 5 kali. Minta kode baru.');
  const benar = c.get('OTP_' + email);
  if (!benar || String(kode).trim() !== benar) {
    c.put('G_' + email, String(salah + 1), 600);
    throw new Error(benar ? 'Kode salah.' : 'Kode kedaluwarsa. Minta kode baru.');
  }
  c.remove('OTP_' + email); c.remove('G_' + email);
  const k = pembuatByEmail_(email);
  logAksi_('-', '-', 'LOGIN', email, 'Login form SPL & SPD');
  return { token: sesiBuat_(email), user: infoUser_(k) };
}

/* ------------------------------ Data form ------------------------------ */
/** Isi dropdown dari sheet Pilihan → {judulKolom: [nilai...]} */
function pilihan_() {
  const v = sheet_(SH.PILIHAN).getDataRange().getValues();
  const head = v.shift();
  const out = {};
  head.forEach((h, i) => { if (h) out[String(h)] = v.map(r => String(r[i]).trim()).filter(Boolean); });
  return out;
}

function dataForm(token) {
  const u = sesiCek_(token);
  const P = param_(), pil = pilihan_();
  const kar = Object.values(karyawan_()).filter(k => k.aktif)
    .map(k => ({ id: k.id, nama: k.nama, jabatan: k.jabatan, departemen: k.departemen }))   // TANPA gaji
    .sort((a, b) => (a.departemen + a.nama).localeCompare(b.departemen + b.nama));
  const jam = []; for (let h = 0; h <= 24; h++) jam.push(angkaKeJam_(h));
  const pk = ['JAM_MULAI_MIN_SENJUM', 'JAM_MULAI_MIN_SABTU', 'JAM_MULAI_MIN_LIBUR', 'JAM_SELESAI_MAKS',
    'ISTIRAHAT_HARI_KERJA', 'ISTIRAHAT_HARI_LIBUR', 'BATAS_SELESAI_LIBUR', 'MAKS_KARYAWAN_PER_SPL',
    'UANG_SAKU_HARI', 'BATAS_LAMA_DINAS', 'MAKS_PERSONEL_SPD', 'NAMA_PERUSAHAAN'];
  return {
    user: infoUser_(u),
    karyawan: kar,
    libur: liburSet_(),
    P: Object.fromEntries(pk.map(k => [k, P[k]])),
    jam: jam,
    pilihan: {
      transportasi: pil['Transportasi'] || [],
      sumberBiaya: pil['Sumber Biaya'] || [],
      dasarPenugasan: pil['Dasar Penugasan'] || [],
      jenisBiaya: pil['Jenis Biaya'] || [],
      bukti: pil['Bukti'] || ['Ada', 'Tidak Ada']
    }
  };
}

/** Tulis beberapa baris objek {judul: nilai} ke sheet sesuai judul kolom. */
function tulisBaris_(namaSheet, daftar, kolomTeks) {
  if (!daftar.length) return;
  const sh = sheet_(namaSheet), c = kolom_(namaSheet), lebar = sh.getLastColumn();
  const rows = daftar.map(o => {
    const r = new Array(lebar).fill('');
    Object.keys(o).forEach(j => { if (c[j]) r[c[j] - 1] = o[j]; });
    return r;
  });
  const r0 = sh.getLastRow() + 1;
  (kolomTeks || []).forEach(j => { if (c[j]) sh.getRange(r0, c[j], rows.length, 1).setNumberFormat('@'); });
  sh.getRange(r0, 1, rows.length, lebar).setValues(rows);
}

/** Karyawan yang terikat surat aktif pada rentang tanggal [dari, sampai] (yyyy-MM-dd).
 *  → {spl:{id:nomor}, spd:{id:nomor}} */
function suratAktif_(dari, sampai) {
  const out = { spl: {}, spd: {} };
  bacaTabel_(SH.SPL, 'Nomor SPL').forEach(r => {
    const t = r['Tanggal Lembur'] instanceof Date ? ymd_(r['Tanggal Lembur']) : '';
    if (t && dari <= t && t <= sampai && STATUS_BATAL.indexOf(String(r['Status'])) < 0)
      out.spl[String(r['Employee ID']).toUpperCase()] = r['Nomor SPL'] + ' (' + t + ')';
  });
  const spdAktif = {};
  bacaTabel_(SH.SPD, 'Nomor SPD').forEach(r => {
    const b = r['Tgl Berangkat'], k = r['Tgl Kembali'];
    if (b instanceof Date && k instanceof Date && STATUS_BATAL.indexOf(String(r['Status'])) < 0 &&
        ymd_(b) <= sampai && dari <= ymd_(k)) spdAktif[r['Nomor SPD']] = true;
  });
  bacaTabel_(SH.PERSONEL, 'Nomor SPD').forEach(r => {
    if (spdAktif[r['Nomor SPD']]) out.spd[String(r['Employee ID']).toUpperCase()] = r['Nomor SPD'];
  });
  return out;
}
/**
 * =====================================================================
 *  SISTEM SPL & SPD — File : Spl.gs
 *  Simpan & riwayat Surat Perintah Lembur (dipanggil dari Form.html)
 *  Butuh: Inti.gs, Web.gs
 * =====================================================================
 */

/**
 * @param {string} token
 * @param {{tanggal:'yyyy-mm-dd', baris:[{id, mulai, selesai, uraian}]}} data
 */
function simpanSPL(token, data) {
  const u = sesiCek_(token);
  pastikanKolom_(SH.SPL, KOLOM_SPL_TAMBAHAN);
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('Server sibuk, coba simpan lagi beberapa detik lagi.');
  let hasilSimpan;
  try {
    _K = null; _L = null; _P = null;
    const P = param_(), K = karyawan_(), L = liburSet_();
    const err = [];
    const tglStr = String(data && data.tanggal || '');
    if (!/^\d{4}-\d{2}-\d{2}$/.test(tglStr)) throw new Error('Tanggal lembur wajib diisi.');
    const baris = (data.baris || []).filter(b => b && String(b.id || '').trim());
    if (!baris.length) throw new Error('Isi minimal 1 karyawan.');
    if (baris.length > Number(P.MAKS_KARYAWAN_PER_SPL)) throw new Error('Maksimal ' + P.MAKS_KARYAWAN_PER_SPL + ' karyawan per SPL.');

    const aktif = suratAktif_(tglStr, tglStr);
    const sudah = {};
    const hasil = baris.map((b, i) => {
      const id = String(b.id).trim().toUpperCase(), no = 'Baris ' + (i + 1) + ': ';
      const k = K[id];
      if (!k || !k.aktif) { err.push(no + 'Employee ID ' + id + ' tidak ditemukan / tidak aktif.'); return null; }
      if (sudah[id]) err.push(no + k.nama + ' tercantum lebih dari sekali.');
      sudah[id] = true;
      if (!String(b.uraian || '').trim()) err.push(no + 'Uraian pekerjaan ' + k.nama + ' wajib diisi.');
      if (aktif.spl[id]) err.push(no + k.nama + ' sudah ada di SPL ' + aktif.spl[id] + '.');
      if (aktif.spd[id]) err.push(no + k.nama + ' sedang dinas luar (' + aktif.spd[id] + ') — lembur tidak dapat diajukan.');
      const h = hitungLembur_({ tanggal: tglStr, mulai: b.mulai, selesai: b.selesai, gaji: k.gaji }, P, L);
      h.error.forEach(e => err.push(no + k.nama + ' — ' + e));
      return { k: k, b: b, h: h };
    });
    if (err.length) return { ok: false, error: err };

    const tgl = Utilities.parseDate(tglStr + ' 00:00', TZ, 'yyyy-MM-dd HH:mm');
    const nomor = nomorBerikutnya_('SPL', tgl);
    const now = new Date();

    tulisBaris_(SH.SPL, hasil.map(x => ({
      'Nomor SPL': nomor, 'Tanggal Lembur': tgl, 'Hari': x.h.hari, 'Status Hari': x.h.statusHari,
      'Employee ID': x.k.id, 'Full Name': x.k.nama, 'Departemen': x.k.departemen,
      'Jam Mulai': angkaKeJam_(jamKeAngka_(x.b.mulai)), 'Jam Selesai': angkaKeJam_(jamKeAngka_(x.b.selesai)),
      'Istirahat': x.h.istirahat, 'Jam Efektif': x.h.jamEfektif, 'Estimasi Kompensasi': x.h.kompensasi,
      'Uraian Pekerjaan': String(x.b.uraian).trim(),
      'Dibuat Oleh': u.nama, 'Email Pembuat': u.email, 'Dibuat Pada': now,
      'Status': ST.KONFIRMASI,
      'Catatan': x.h.peringatan.join(' | '),
      'Status Email': 'Mengirim…',
      'ID Baris': Utilities.getUuid().slice(0, 8)
    })), ['Jam Mulai', 'Jam Selesai']);
    SpreadsheetApp.flush();

    const totalJam = hasil.reduce((s, x) => s + x.h.jamEfektif, 0);
    logAksi_('SPL', nomor, 'DIBUAT', u.email, hasil.length + ' karyawan, ' + totalJam + ' jam, tgl ' + tglStr);
    hasilSimpan = { ok: true, nomor: nomor, jumlah: hasil.length, totalJam: totalJam, status: ST.KONFIRMASI };
  } finally {
    lock.releaseLock();
  }
  // Email konfirmasi ke tiap karyawan (di luar kunci agar pengguna lain tidak menunggu)
  try { kirimKonfirmasiSPL_(hasilSimpan.nomor); } catch (e) { logAksi_('SPL', hasilSimpan.nomor, 'EMAIL GAGAL', u.email, String(e)); }
  return hasilSimpan;
}

/** Riwayat SPL buatan user (30 terakhir), dikelompokkan per nomor, dengan status per karyawan. */
function riwayatSPL(token) {
  const u = sesiCek_(token);
  const g = {};
  bacaTabel_(SH.SPL, 'Nomor SPL').forEach(r => {
    if (String(r['Email Pembuat']).toLowerCase() !== u.email) return;
    const n = r['Nomor SPL'];
    if (!g[n]) g[n] = {
      jenis: 'SPL', nomor: n, tanggal: r['Tanggal Lembur'] instanceof Date ? ymd_(r['Tanggal Lembur']) : '',
      hari: r['Hari'], statusHari: r['Status Hari'], pdf: r['Link PDF'] || '',
      dibuat: r['Dibuat Pada'] instanceof Date ? r['Dibuat Pada'].getTime() : 0, orang: [], totalJam: 0
    };
    g[n].orang.push({ nama: r['Full Name'], mulai: String(r['Jam Mulai']), selesai: String(r['Jam Selesai']),
      jam: Number(r['Jam Efektif']) || 0, uraian: r['Uraian Pekerjaan'], status: r['Status'] });
    g[n].totalJam += Number(r['Jam Efektif']) || 0;
  });
  return Object.values(g).map(x => {
    const selesai = x.orang.filter(o => ST_FINAL.indexOf(o.status) >= 0).length;
    x.status = selesai === x.orang.length ? 'SELESAI' : 'PROSES ' + selesai + '/' + x.orang.length;
    return x;
  }).sort((a, b) => b.dibuat - a.dibuat).slice(0, 30);
}
/**
 * =====================================================================
 *  SISTEM SPL & SPD — File : Spd.gs
 *  Simpan & riwayat Surat Perintah Dinas (dipanggil dari Form.html)
 *  Struktur sama dengan formulir SPD Excel: A. Referensi, B. Detail,
 *  C. Personel (maks 5), D. Biaya lainnya.
 *  Butuh: Inti.gs, Web.gs
 * =====================================================================
 */

/**
 * @param {string} token
 * @param {{tanggalSurat, noReferensi, customer, dasarPenugasan, maksud, tempat, berangkat, kembali,
 *          transportasi, sumberBiaya, personel:[{id, lama, keterangan}], biaya:[{jenis, uraian, jumlah, bukti}]}} d
 */
function simpanSPD(token, d) {
  const u = sesiCek_(token);
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('Server sibuk, coba simpan lagi beberapa detik lagi.');
  try {
    _K = null; _P = null;
    const P = param_(), K = karyawan_(), pil = pilihan_();
    const err = [], warn = [];
    const tglOk = s => /^\d{4}-\d{2}-\d{2}$/.test(String(s || ''));
    const teks = s => String(s || '').trim();

    if (!tglOk(d.tanggalSurat)) err.push('Tanggal surat wajib diisi.');
    if (!teks(d.maksud)) err.push('Maksud / Tujuan wajib diisi.');
    if (!teks(d.tempat)) err.push('Tempat tujuan wajib diisi.');
    if (!tglOk(d.berangkat) || !tglOk(d.kembali)) err.push('Tanggal berangkat & kembali wajib diisi.');
    const cekPilihan = (v, daftar, nama) => { if (teks(v) && (daftar || []).indexOf(teks(v)) < 0) err.push(nama + ' tidak valid.'); };
    cekPilihan(d.dasarPenugasan, pil['Dasar Penugasan'], 'Dasar penugasan');
    cekPilihan(d.transportasi, pil['Transportasi'], 'Transportasi');
    cekPilihan(d.sumberBiaya, pil['Sumber Biaya'], 'Sumber biaya');
    if (!teks(d.transportasi)) err.push('Transportasi wajib dipilih.');
    if (!teks(d.sumberBiaya)) err.push('Sumber biaya wajib dipilih.');

    const personel = (d.personel || []).filter(p => p && teks(p.id)).map(p => Object.assign({}, p, { id: teks(p.id).toUpperCase() }));
    const biaya = (d.biaya || []).filter(b => b && (teks(b.jenis) || Number(b.jumlah)));
    personel.forEach((p, i) => { if (!K[p.id] || !K[p.id].aktif) err.push('Personel ke-' + (i + 1) + ': Employee ID ' + p.id + ' tidak ditemukan / tidak aktif.'); });
    biaya.forEach((b, i) => {
      if (!teks(b.jenis)) err.push('Biaya ke-' + (i + 1) + ': jenis biaya wajib dipilih.');
      if (!(Number(b.jumlah) >= 0)) err.push('Biaya ke-' + (i + 1) + ': jumlah tidak valid.');
    });

    const h = hitungSPD_({ berangkat: d.berangkat, kembali: d.kembali, personel: personel, biaya: biaya.map(b => ({ jumlah: Number(b.jumlah) || 0 })) }, P);
    err.push.apply(err, h.error);
    warn.push.apply(warn, h.peringatan);

    if (tglOk(d.berangkat) && tglOk(d.kembali)) {
      const aktif = suratAktif_(d.berangkat, d.kembali);
      personel.forEach(p => {
        const k = K[p.id]; if (!k) return;
        if (aktif.spd[p.id]) err.push(k.nama + ' sudah ditugaskan di ' + aktif.spd[p.id] + ' pada tanggal yang beririsan.');
        if (aktif.spl[p.id]) warn.push(k.nama + ' punya SPL ' + aktif.spl[p.id] + ' di rentang dinas — SPL tersebut perlu ditinjau.');
      });
    }
    if (err.length) return { ok: false, error: err };

    // SPD: belum ada alur email/persetujuan (menyusul) — status awal TERSIMPAN
    const alur = { status: 'TERSIMPAN' };
    const tp = s => Utilities.parseDate(s + ' 00:00', TZ, 'yyyy-MM-dd HH:mm');
    const tglSurat = tp(d.tanggalSurat);
    const nomor = nomorBerikutnya_('SPD', tglSurat);
    const now = new Date();

    tulisBaris_(SH.SPD, [{
      'Nomor SPD': nomor, 'Tanggal Surat': tglSurat, 'Tahun': Number(d.tanggalSurat.slice(0, 4)),
      'No Urut': Number(nomor.split('/')[0]), 'No Surat Referensi': teks(d.noReferensi), 'Nama Customer': teks(d.customer),
      'Dasar Penugasan': teks(d.dasarPenugasan), 'Maksud / Tujuan': teks(d.maksud), 'Tempat Tujuan': teks(d.tempat),
      'Tgl Berangkat': tp(d.berangkat), 'Tgl Kembali': tp(d.kembali), 'Lama (hari)': h.lama,
      'Transportasi': teks(d.transportasi), 'Sumber Biaya': teks(d.sumberBiaya),
      'Total Uang Saku': h.totalUangSaku, 'Total Biaya Lain': h.totalBiayaLain, 'Total Biaya': h.totalBiaya,
      'Status Dinas': 'Aktif', 'Departemen': u.departemen,
      'Dibuat Oleh': u.nama, 'Email Pembuat': u.email, 'Dibuat Pada': now,
      'Status': alur.status,
      'Diketahui Oleh': alur.diketahuiOleh || '', 'Diketahui Pada': alur.diketahuiOleh ? now : '',
      'Catatan': [].concat(warn, alur.catatan ? [alur.catatan] : []).join(' | '),
      'Status Email': 'Belum dikirim'
    }]);
    tulisBaris_(SH.PERSONEL, h.personel.map((p, i) => ({
      'Nomor SPD': nomor, 'No': i + 1, 'Employee ID': p.id, 'Full Name': K[p.id].nama, 'Job Position': K[p.id].jabatan,
      'Lama (hari)': Number(p.lama), 'Uang Saku': p.uangSaku, 'Keterangan': teks(p.keterangan)
    })));
    tulisBaris_(SH.BIAYA, biaya.map((b, i) => ({
      'Nomor SPD': nomor, 'No': i + 1, 'Jenis Biaya': teks(b.jenis), 'Uraian': teks(b.uraian),
      'Jumlah (Rp)': Number(b.jumlah) || 0, 'Bukti': teks(b.bukti)
    })));
    SpreadsheetApp.flush();

    logAksi_('SPD', nomor, 'DIBUAT', u.email, personel.length + ' personel, ' + h.lama + ' hari ke ' + teks(d.tempat) + ', total Rp' + h.totalBiaya + ' → ' + alur.status);
    return { ok: true, nomor: nomor, jumlah: personel.length, lama: h.lama, totalBiaya: h.totalBiaya, status: alur.status, peringatan: warn };
  } finally {
    lock.releaseLock();
  }
}

/** Riwayat SPD buatan user (30 terakhir). */
function riwayatSPD(token) {
  const u = sesiCek_(token);
  const per = {};
  bacaTabel_(SH.PERSONEL, 'Nomor SPD').forEach(r => {
    (per[r['Nomor SPD']] = per[r['Nomor SPD']] || []).push({ nama: r['Full Name'], lama: Number(r['Lama (hari)']) || 0, keterangan: r['Keterangan'] });
  });
  return bacaTabel_(SH.SPD, 'Nomor SPD')
    .filter(r => String(r['Email Pembuat']).toLowerCase() === u.email)
    .map(r => ({
      jenis: 'SPD', nomor: r['Nomor SPD'], status: r['Status'],
      tanggal: r['Tanggal Surat'] instanceof Date ? ymd_(r['Tanggal Surat']) : '',
      berangkat: r['Tgl Berangkat'] instanceof Date ? ymd_(r['Tgl Berangkat']) : '',
      kembali: r['Tgl Kembali'] instanceof Date ? ymd_(r['Tgl Kembali']) : '',
      lama: Number(r['Lama (hari)']) || 0, tempat: r['Tempat Tujuan'], maksud: r['Maksud / Tujuan'],
      totalBiaya: Number(r['Total Biaya']) || 0,
      dibuat: r['Dibuat Pada'] instanceof Date ? r['Dibuat Pada'].getTime() : 0,
      orang: per[r['Nomor SPD']] || []
    }))
    .sort((a, b) => b.dibuat - a.dibuat).slice(0, 30);
}
/**
 * =====================================================================
 *  SISTEM SPL & SPD — File : Notif.gs
 *  Alur email SPL:
 *   1. Admin buat SPL → email ke tiap karyawan: [Bersedia] / [Tidak Bersedia]
 *      (karyawan tanpa email → email konfirmasi dikirim ke admin pembuat)
 *   2. Tidak Bersedia → email ke admin pembuat + karyawan ybs
 *   3. Bersedia → email ke ATASAN LANGSUNG (kolom "Atasan" di Data Karyawan):
 *      [Setujui] / [Tolak]   (atasan kosong → dikirim ke admin pembuat)
 *   4. Semua karyawan di SPL sudah final → PDF SPL (kop surat) dikirim ke admin
 *      & disimpan di folder "PDF SPL" (satu folder dengan spreadsheet)
 *  Status per baris: MENUNGGU KONFIRMASI → TIDAK BERSEDIA | MENUNGGU ATASAN → DISETUJUI | DITOLAK
 * =====================================================================
 */

const ST = {
  KONFIRMASI: 'MENUNGGU KONFIRMASI', TIDAK: 'TIDAK BERSEDIA', ATASAN: 'MENUNGGU ATASAN',
  SETUJU: 'DISETUJUI', TOLAK: 'DITOLAK', BATAL: 'DIBATALKAN'
};
const ST_FINAL = [ST.TIDAK, ST.SETUJU, ST.TOLAK, ST.BATAL];
/** Kolom tambahan Data SPL untuk alur email (dibuat otomatis bila belum ada) */
const KOLOM_SPL_TAMBAHAN = ['Atasan', 'Email Atasan', 'Konfirmasi Pada', 'Alasan Tidak Bersedia', 'Catatan Atasan', 'Link PDF'];

/* ------------------------------ utilitas ------------------------------ */
function pastikanKolom_(namaSheet, daftar) {
  const sh = sheet_(namaSheet), c = kolom_(namaSheet);
  const kurang = daftar.filter(j => !c[j]);
  if (!kurang.length) return;
  const col = sh.getLastColumn() + 1;
  if (sh.getMaxColumns() < col + kurang.length - 1) sh.insertColumnsAfter(sh.getMaxColumns(), col + kurang.length - 1 - sh.getMaxColumns());
  sh.getRange(1, col, 1, kurang.length).setValues([kurang]).setFontWeight('bold').setBackground('#1F3864').setFontColor('#FFFFFF');
}
function rahasia_() {
  const p = PropertiesService.getScriptProperties();
  let s = p.getProperty('RAHASIA');
  if (!s) { s = Utilities.getUuid() + Utilities.getUuid(); p.setProperty('RAHASIA', s); }
  return s;
}
function tandaTangan_(aksi, id) {
  return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(aksi + ':' + id, rahasia_())).replace(/=+$/, '').slice(0, 24);
}
function urlWeb_() {
  return String(param_().URL_WEB_APP || '').trim() || ScriptApp.getService().getUrl();
}
function linkAksi_(aksi, id, pilih) {
  return urlWeb_() + '?a=' + aksi + '&id=' + encodeURIComponent(id) + '&s=' + tandaTangan_(aksi, id) + (pilih ? '&p=' + pilih : '');
}
function logoBlob_() {
  return Utilities.newBlob(Utilities.base64Decode(LOGO_DATA_URI.split(',')[1]), 'image/jpeg', 'logo.jpg');
}
const escH_ = s => String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
function tglIndo_(d) {
  if (!(d instanceof Date)) return String(d || '');
  const bln = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
  const [y, m, dd] = Utilities.formatDate(d, TZ, 'yyyy-M-d').split('-').map(Number);
  return dd + ' ' + bln[m - 1] + ' ' + y;
}
function waktuIndo_(d) { return d instanceof Date ? tglIndo_(d) + ' ' + Utilities.formatDate(d, TZ, 'HH:mm') : ''; }

/** Baca semua baris SPL → {sh, c, rows:[{_row, ...kolom}]} */
function bacaSPL_() {
  const sh = sheet_(SH.SPL), c = kolom_(SH.SPL);
  const n = sh.getLastRow() - 1;
  const v = n > 0 ? sh.getRange(2, 1, n, sh.getLastColumn()).getValues() : [];
  const head = Object.keys(c);
  const rows = [];
  v.forEach((r, i) => {
    if (!String(r[c['Nomor SPL'] - 1]).trim()) return;
    const o = { _row: i + 2 }; head.forEach(h => o[h] = r[c[h] - 1]); rows.push(o);
  });
  return { sh: sh, c: c, rows: rows };
}
function ubahBaris_(S, row, data) {
  Object.keys(data).forEach(j => { if (S.c[j]) S.sh.getRange(row, S.c[j]).setValue(data[j]); });
}

/** Kirim email (MailApp) dengan logo inline. Mengembalikan '' bila sukses, pesan error bila gagal. */
function kirimEmail_(ke, subjek, isiHtml, cc, lampiran) {
  try {
    if (!ke) return 'tanpa email';
    if (MailApp.getRemainingDailyQuota() < 1) return 'kuota email harian habis';
    const opt = {
      to: ke, subject: subjek, name: 'Sistem SPL ' + (param_().NAMA_PERUSAHAAN || ''),
      htmlBody: bungkusEmail_(isiHtml), inlineImages: { logo: logoBlob_() }
    };
    if (cc) opt.cc = cc;
    if (lampiran) opt.attachments = [lampiran];
    MailApp.sendEmail(opt);
    return '';
  } catch (e) { return String(e.message || e); }
}
function bungkusEmail_(isi) {
  return '<div style="font-family:Arial,sans-serif;font-size:14px;color:#222;max-width:600px">' +
    '<div style="border-bottom:4px solid #BE353F;padding:8px 0 10px"><img src="cid:logo" alt="PT Citra Banjar Abadi" style="height:44px"></div>' +
    '<div style="padding:16px 0">' + isi + '</div>' +
    '<div style="font-size:11px;color:#888;border-top:1px solid #eee;padding-top:8px">Email otomatis dari Sistem SPL & SPD. Mohon tidak membalas email ini.</div></div>';
}
function tombol_(href, teks, warna) {
  return '<a href="' + href + '" style="display:inline-block;background:' + warna + ';color:#fff;text-decoration:none;font-weight:bold;padding:12px 22px;border-radius:8px;margin:4px 8px 4px 0">' + teks + '</a>';
}
function tabelDetail_(r) {
  const baris = [
    ['Nomor SPL', r['Nomor SPL']],
    ['Tanggal Lembur', r['Hari'] + ', ' + tglIndo_(r['Tanggal Lembur']) + ' (' + r['Status Hari'] + ')'],
    ['Karyawan', r['Full Name'] + ' (' + r['Employee ID'] + ')'],
    ['Jam Lembur', r['Jam Mulai'] + ' – ' + r['Jam Selesai'] + (r['Istirahat'] && r['Istirahat'] !== '-' ? ' (istirahat ' + r['Istirahat'] + ')' : '')],
    ['Jam Efektif', String(r['Jam Efektif']).replace('.', ',') + ' jam'],
    ['Uraian Pekerjaan', r['Uraian Pekerjaan']],
    ['Diperintahkan oleh', r['Dibuat Oleh']]
  ];
  return '<table style="border-collapse:collapse;width:100%;font-size:13px">' + baris.map(b =>
    '<tr><td style="padding:6px 8px;background:#F4F5F7;width:150px;border:1px solid #E2E4E8;font-weight:bold">' + escH_(b[0]) +
    '</td><td style="padding:6px 8px;border:1px solid #E2E4E8">' + escH_(b[1]) + '</td></tr>').join('') + '</table>';
}

/* ======================= 1. Email konfirmasi ke karyawan ======================= */
/** Dipanggil setelah SPL tersimpan. */
function kirimKonfirmasiSPL_(nomor) {
  const S = bacaSPL_(), K = karyawan_();
  S.rows.filter(r => r['Nomor SPL'] === nomor && r['Status'] === ST.KONFIRMASI).forEach(r => {
    const k = K[String(r['Employee ID']).toUpperCase()] || {};
    const id = String(r['ID Baris']);
    const keAdmin = !k.email;
    const isi = (keAdmin
      ? '<p><b>' + escH_(r['Full Name']) + '</b> belum memiliki email di sistem. Mohon konfirmasi kesediaan lembur <b>atas nama karyawan tersebut</b>:</p>'
      : '<p>Yth. <b>' + escH_(r['Full Name']) + '</b>,</p><p>Anda mendapat <b>Surat Perintah Lembur</b> dengan rincian berikut. Mohon konfirmasi kesediaan Anda:</p>') +
      tabelDetail_(r) + '<p style="margin-top:18px">' +
      tombol_(linkAksi_('k', id, 'ya'), '✅ Bersedia', '#1E7B4A') + tombol_(linkAksi_('k', id, 'tidak'), '❌ Tidak Bersedia', '#B42318') + '</p>';
    const err = kirimEmail_(keAdmin ? r['Email Pembuat'] : k.email,
      'Konfirmasi Lembur ' + r['Hari'] + ', ' + tglIndo_(r['Tanggal Lembur']) + (keAdmin ? ' — atas nama ' + r['Full Name'] : ''), isi);
    ubahBaris_(S, r._row, { 'Status Email': err ? 'GAGAL: ' + err : 'Konfirmasi terkirim ' + (keAdmin ? 'ke admin ' : '') + Utilities.formatDate(new Date(), TZ, 'dd/MM HH:mm') });
  });
}

/* ======================= Halaman aksi (dibuka dari link email) ======================= */
function halamanAksi_(p) {
  const aksi = p.a, id = p.id, sig = p.s;
  let isi;
  try {
    if (['k', 't'].indexOf(aksi) < 0 || !id || sig !== tandaTangan_(aksi, id)) throw new Error('Link tidak valid atau sudah diubah.');
    const S = bacaSPL_(), r = S.rows.find(x => String(x['ID Baris']) === String(id));
    if (!r) throw new Error('Data lembur tidak ditemukan (mungkin sudah dihapus).');
    const sudah = aksi === 'k' ? r['Status'] !== ST.KONFIRMASI : r['Status'] !== ST.ATASAN;
    const judul = aksi === 'k' ? 'Konfirmasi Kesediaan Lembur' : 'Persetujuan Lembur';
    isi = '<h2>' + judul + '</h2>' + tabelDetail_(r).replace(/width:150px/g, 'width:130px');
    if (aksi === 't') isi += '<p class="sub">Karyawan telah menyatakan <b>BERSEDIA</b>' + (r['Konfirmasi Pada'] instanceof Date ? ' pada ' + waktuIndo_(r['Konfirmasi Pada']) : '') + '.</p>';
    if (sudah) {
      isi += '<div class="msg info">Sudah diproses — status saat ini: <b>' + escH_(r['Status']) + '</b>.</div>';
    } else {
      const ya = aksi === 'k' ? 'Bersedia' : 'Setujui', tidak = aksi === 'k' ? 'Tidak Bersedia' : 'Tolak';
      const pil = p.p === 'tidak' ? 'tidak' : 'ya';
      isi += '<div class="pilih">' +
        '<label class="op ya"><input type="radio" name="j" value="ya"' + (pil === 'ya' ? ' checked' : '') + ' onchange="ubah()"> ✅ ' + ya + '</label>' +
        '<label class="op tidak"><input type="radio" name="j" value="tidak"' + (pil === 'tidak' ? ' checked' : '') + ' onchange="ubah()"> ❌ ' + tidak + '</label></div>' +
        '<div id="al"><label>' + (aksi === 'k' ? 'Alasan tidak bersedia' : 'Catatan') + ' <span class="sub">(opsional)</span></label><textarea id="alasan" rows="3"></textarea></div>' +
        '<button id="b" onclick="kirim()">Kirim</button><div id="o"></div>' +
        '<script>function ubah(){var t=document.querySelector("input[name=j]:checked").value;document.getElementById("al").style.display=(t==="tidak"||' + (aksi === 't') + ')?"block":"none"}ubah();' +
        'function kirim(){var b=document.getElementById("b"),o=document.getElementById("o");b.disabled=true;b.textContent="Mengirim…";' +
        'google.script.run.withSuccessHandler(function(m){document.querySelector(".pilih").remove();document.getElementById("al").remove();b.remove();o.innerHTML="<div class=\\"msg ok\\">"+m+"</div>"})' +
        '.withFailureHandler(function(e){b.disabled=false;b.textContent="Kirim";o.innerHTML="<div class=\\"msg err\\">"+e.message+"</div>"})' +
        '.prosesAksi(' + JSON.stringify(aksi) + ',' + JSON.stringify(String(id)) + ',' + JSON.stringify(sig) + ',document.querySelector("input[name=j]:checked").value,document.getElementById("alasan").value)}</script>';
    }
  } catch (e) {
    isi = '<h2>Tidak dapat dibuka</h2><div class="msg err">' + escH_(e.message) + '</div>';
  }
  const html = '<!DOCTYPE html><html><head><base target="_top"><meta charset="utf-8"><style>' +
    'body{margin:0;font-family:Arial,sans-serif;font-size:14px;background:#F5F5F6;color:#222}' +
    'header{background:#fff;border-bottom:4px solid #BE353F;padding:10px 16px}header img{height:42px;max-width:100%}' +
    'main{max-width:620px;margin:16px auto;padding:0 16px}.card{background:#fff;border:1px solid #E2E4E8;border-radius:10px;padding:18px}' +
    'h2{margin:0 0 14px;font-size:18px;color:#3A3A3A}.sub{color:#6B6F76;font-size:12px}' +
    '.pilih{display:flex;gap:10px;margin:18px 0 12px;flex-wrap:wrap}.op{flex:1;min-width:160px;border:2px solid #E2E4E8;border-radius:10px;padding:14px;font-weight:bold;cursor:pointer;font-size:15px}' +
    '.op.ya:has(input:checked){border-color:#1E7B4A;background:#E7F5EC}.op.tidak:has(input:checked){border-color:#B42318;background:#FDECEA}' +
    'label{display:block;font-weight:bold;font-size:12px;color:#6B6F76;margin-bottom:4px}textarea{width:100%;box-sizing:border-box;font:inherit;padding:8px;border:1px solid #CBCED4;border-radius:6px}' +
    'button{margin-top:12px;background:#BE353F;color:#fff;border:0;border-radius:8px;padding:12px 26px;font-weight:bold;font-size:15px;cursor:pointer}button:disabled{background:#B9BCC2}' +
    '.msg{padding:12px;border-radius:8px;margin-top:14px}.ok{background:#E7F5EC;color:#1E7B4A}.err{background:#FDECEA;color:#B42318}.info{background:#EFF0F2}' +
    '</style></head><body><header><img src="' + LOGO_DATA_URI + '" alt="PT Citra Banjar Abadi"></header><main><div class="card">' + isi + '</div></main></body></html>';
  return HtmlService.createHtmlOutput(html).setTitle('SPL — PT Citra Banjar Abadi')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Dipanggil dari halaman aksi. aksi: 'k' (konfirmasi karyawan) | 't' (persetujuan atasan); jawaban: 'ya'|'tidak' */
function prosesAksi(aksi, id, sig, jawaban, alasan) {
  if (['k', 't'].indexOf(aksi) < 0 || sig !== tandaTangan_(aksi, id)) throw new Error('Link tidak valid.');
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('Server sibuk, coba lagi.');
  let nomor, pesan;
  try {
    _K = null; _P = null;
    pastikanKolom_(SH.SPL, KOLOM_SPL_TAMBAHAN);
    const S = bacaSPL_(), r = S.rows.find(x => String(x['ID Baris']) === String(id));
    if (!r) throw new Error('Data tidak ditemukan.');
    nomor = r['Nomor SPL'];
    const now = new Date(), alas = String(alasan || '').trim().slice(0, 500);
    const tgl = r['Hari'] + ', ' + tglIndo_(r['Tanggal Lembur']);

    if (aksi === 'k') {
      if (r['Status'] !== ST.KONFIRMASI) return 'Sudah dijawab sebelumnya — status: ' + r['Status'];
      if (jawaban === 'tidak') {
        ubahBaris_(S, r._row, { 'Status': ST.TIDAK, 'Konfirmasi Pada': now, 'Alasan Tidak Bersedia': alas });
        const k = karyawan_()[String(r['Employee ID']).toUpperCase()] || {};
        const isi = '<p><b>' + escH_(r['Full Name']) + '</b> menyatakan <b style="color:#B42318">TIDAK BERSEDIA</b> lembur.</p>' + tabelDetail_(r) +
          (alas ? '<p><b>Alasan:</b> ' + escH_(alas) + '</p>' : '');
        kirimEmail_(r['Email Pembuat'], 'Tidak Bersedia Lembur — ' + r['Full Name'] + ' (' + tgl + ')', isi, k.email || '');
        logAksi_('SPL', nomor, 'TIDAK BERSEDIA', r['Full Name'], alas);
        pesan = 'Terima kasih. Ketidaksediaan Anda sudah dicatat dan diberitahukan ke admin.';
      } else {
        const atasan = cariAtasan_(r['Employee ID']);
        const ke = atasan ? atasan.email : r['Email Pembuat'];
        ubahBaris_(S, r._row, { 'Status': ST.ATASAN, 'Konfirmasi Pada': now,
          'Atasan': atasan ? atasan.nama : '(admin — atasan belum diisi)', 'Email Atasan': ke });
        const isi = (atasan ? '<p>Yth. <b>' + escH_(atasan.nama) + '</b>,</p>' : '<p><i>Atasan langsung karyawan ini belum diisi di Data Karyawan, sehingga permintaan persetujuan dikirim ke admin.</i></p>') +
          '<p><b>' + escH_(r['Full Name']) + '</b> telah menyatakan <b style="color:#1E7B4A">BERSEDIA</b> lembur. Mohon persetujuan Anda:</p>' + tabelDetail_(r) +
          '<p style="margin-top:18px">' + tombol_(linkAksi_('t', id, 'ya'), '✅ Setujui', '#1E7B4A') + tombol_(linkAksi_('t', id, 'tidak'), '❌ Tolak', '#B42318') + '</p>';
        const err = kirimEmail_(ke, 'Persetujuan Lembur — ' + r['Full Name'] + ' (' + tgl + ')', isi);
        ubahBaris_(S, r._row, { 'Status Email': err ? 'GAGAL kirim ke atasan: ' + err : 'Persetujuan terkirim ke ' + ke });
        logAksi_('SPL', nomor, 'BERSEDIA', r['Full Name'], 'Diteruskan ke ' + ke);
        pesan = 'Terima kasih. Kesediaan Anda sudah dicatat dan diteruskan ke ' + (atasan ? 'atasan (' + escH_(atasan.nama) + ')' : 'admin') + ' untuk persetujuan.';
      }
    } else {
      if (r['Status'] !== ST.ATASAN) return 'Sudah diproses sebelumnya — status: ' + r['Status'];
      const setuju = jawaban !== 'tidak';
      ubahBaris_(S, r._row, { 'Status': setuju ? ST.SETUJU : ST.TOLAK, 'Disetujui Oleh': r['Atasan'], 'Disetujui Pada': now, 'Catatan Atasan': alas });
      logAksi_('SPL', nomor, setuju ? 'DISETUJUI ATASAN' : 'DITOLAK ATASAN', r['Email Atasan'], r['Full Name'] + (alas ? ' — ' + alas : ''));
      pesan = 'Terima kasih. Lembur ' + escH_(r['Full Name']) + ' telah <b>' + (setuju ? 'DISETUJUI' : 'DITOLAK') + '</b>.';
    }
    SpreadsheetApp.flush();
  } finally { lock.releaseLock(); }
  try { cekSelesaiSPL_(nomor); } catch (e) { console.error('PDF gagal: ' + e); }
  return pesan;
}

/** Atasan langsung dari kolom "Atasan" (Employee ID). */
function cariAtasan_(employeeId) {
  const K = karyawan_(), k = K[String(employeeId).toUpperCase()];
  if (!k || !k.atasan) return null;
  const a = K[k.atasan];
  return a && a.email && a.aktif ? a : null;
}

/* ======================= 4. PDF SPL ke admin ======================= */
/** Bila semua baris SPL sudah final & PDF belum dibuat → buat PDF & kirim ke admin. */
function cekSelesaiSPL_(nomor, paksa) {
  const S = bacaSPL_(), rows = S.rows.filter(r => r['Nomor SPL'] === nomor);
  if (!rows.length) return 'Nomor SPL tidak ditemukan.';
  const belum = rows.filter(r => ST_FINAL.indexOf(r['Status']) < 0);
  if (!paksa && (belum.length || rows.some(r => r['Link PDF']))) return '';
  const pdf = buatPdfSPL_(rows);
  const folder = folderPdf_();
  const file = folder.createFile(pdf);
  rows.forEach(r => ubahBaris_(S, r._row, { 'Link PDF': file.getUrl() }));
  const disetujui = rows.filter(r => r['Status'] === ST.SETUJU);
  const isi = '<p>Rekap <b>Surat Perintah Lembur ' + escH_(nomor) + '</b> — ' + escH_(rows[0]['Hari']) + ', ' + tglIndo_(rows[0]['Tanggal Lembur']) + '.</p>' +
    '<p>' + disetujui.length + ' dari ' + rows.length + ' karyawan <b>disetujui</b> lembur, total ' +
    String(disetujui.reduce((s, r) => s + (Number(r['Jam Efektif']) || 0), 0)).replace('.', ',') + ' jam.' +
    (belum.length ? ' <i>(' + belum.length + ' karyawan belum selesai diproses saat PDF ini dibuat)</i>' : '') + '</p>' +
    '<p>PDF terlampir dan tersimpan di Google Drive: <a href="' + file.getUrl() + '">' + escH_(file.getName()) + '</a></p>';
  const admin = String(param_().EMAIL_ADMIN || '').trim().toLowerCase();
  const pembuat = String(rows[0]['Email Pembuat']).toLowerCase();
  kirimEmail_(pembuat, 'PDF SPL ' + nomor + ' — ' + tglIndo_(rows[0]['Tanggal Lembur']), isi, admin && admin !== pembuat ? admin : '', pdf);
  logAksi_('SPL', nomor, 'PDF DIKIRIM', pembuat, file.getUrl());
  return 'PDF dibuat & dikirim ke ' + pembuat;
}

function folderPdf_() {
  const parent = DriveApp.getFileById(ss_().getId()).getParents();
  const induk = parent.hasNext() ? parent.next() : DriveApp.getRootFolder();
  const it = induk.getFoldersByName('PDF SPL');
  return it.hasNext() ? it.next() : induk.createFolder('PDF SPL');
}

function buatPdfSPL_(rows) {
  const r0 = rows[0], P = param_();
  const warna = { 'DISETUJUI': '#1E7B4A', 'DITOLAK': '#B42318', 'TIDAK BERSEDIA': '#B42318', 'DIBATALKAN': '#888' };
  const th = 'style="background:#3A3A3A;color:#fff;padding:5px 4px;font-size:9px;border:1px solid #3A3A3A"';
  const td = (s, x) => '<td style="padding:4px;border:1px solid #BBB;font-size:9px;vertical-align:top;' + (x || '') + '">' + s + '</td>';
  const setuju = rows.filter(r => r['Status'] === ST.SETUJU);
  const totJam = setuju.reduce((s, r) => s + (Number(r['Jam Efektif']) || 0), 0);
  const isiTabel = rows.map((r, i) => '<tr>' +
    td(i + 1, 'text-align:center') + td(escH_(r['Employee ID'])) + td('<b>' + escH_(r['Full Name']) + '</b>') +
    td(escH_(r['Jam Mulai']), 'text-align:center') + td(escH_(r['Jam Selesai']), 'text-align:center') + td(escH_(r['Istirahat']), 'text-align:center') +
    td(String(r['Jam Efektif']).replace('.', ','), 'text-align:center') + td(escH_(r['Uraian Pekerjaan'])) +
    td('<b style="color:' + (warna[r['Status']] || '#9A6200') + '">' + escH_(r['Status']) + '</b>' +
      (r['Alasan Tidak Bersedia'] ? '<br>' + escH_(r['Alasan Tidak Bersedia']) : '') + (r['Catatan Atasan'] ? '<br>' + escH_(r['Catatan Atasan']) : '')) +
    td(r['Status'] === ST.SETUJU || r['Status'] === ST.TOLAK ? escH_(r['Disetujui Oleh']) + '<br><span style="color:#777">' + waktuIndo_(r['Disetujui Pada']) + '</span>' : '-') +
    '</tr>').join('');
  const html = '<html><head><meta charset="utf-8"><style>@page{size:A4 landscape;margin:12mm}body{font-family:Arial,sans-serif;font-size:10px;color:#222}</style></head><body>' +
    '<div style="border-bottom:3px solid #BE353F;padding-bottom:6px"><img src="' + LOGO_DATA_URI + '" style="height:46px"></div>' +
    '<div style="text-align:center;margin:12px 0 2px;font-size:15px;font-weight:bold;letter-spacing:.5px">SURAT PERINTAH LEMBUR (SPL)</div>' +
    '<div style="text-align:center;font-size:11px;margin-bottom:10px">Nomor : ' + escH_(r0['Nomor SPL']) + '</div>' +
    '<table style="font-size:10px;margin-bottom:8px"><tr><td style="width:110px">Tanggal Lembur</td><td>: <b>' + escH_(r0['Hari']) + ', ' + tglIndo_(r0['Tanggal Lembur']) + '</b></td></tr>' +
    '<tr><td>Status Hari</td><td>: ' + escH_(r0['Status Hari']) + '</td></tr><tr><td>Dibuat oleh</td><td>: ' + escH_(r0['Dibuat Oleh']) + ' — ' + waktuIndo_(r0['Dibuat Pada']) + '</td></tr></table>' +
    '<table style="border-collapse:collapse;width:100%"><tr><th ' + th + '>No</th><th ' + th + '>Employee ID</th><th ' + th + '>Full Name</th><th ' + th + '>Mulai</th><th ' + th + '>Selesai</th>' +
    '<th ' + th + '>Istirahat</th><th ' + th + '>Jam Efektif</th><th ' + th + '>Uraian Pekerjaan</th><th ' + th + '>Status</th><th ' + th + '>Persetujuan Atasan</th></tr>' + isiTabel +
    '<tr><td colspan="6" style="padding:5px;border:1px solid #BBB;text-align:right;font-weight:bold;font-size:9px;background:#EFF0F2">TOTAL JAM LEMBUR DISETUJUI (' + setuju.length + ' dari ' + rows.length + ' karyawan)</td>' +
    '<td style="padding:5px;border:1px solid #BBB;text-align:center;font-weight:bold;font-size:9px;background:#EFF0F2">' + String(totJam).replace('.', ',') + '</td><td colspan="3" style="border:1px solid #BBB;background:#EFF0F2"></td></tr></table>' +
    '<p style="font-size:9px;color:#555;margin-top:10px">Konfirmasi karyawan dan persetujuan atasan dilakukan secara elektronik melalui Sistem SPL ' + escH_(P.NAMA_PERUSAHAAN || '') +
    '. Dokumen dibuat otomatis pada ' + waktuIndo_(new Date()) + '.</p>' +
    '<table style="width:100%;margin-top:18px;font-size:10px;text-align:center"><tr><td>Dibuat oleh,</td><td>Diketahui oleh,</td><td>Disetujui oleh,</td></tr>' +
    '<tr><td style="height:55px"></td><td></td><td></td></tr>' +
    '<tr><td>( ' + escH_(r0['Dibuat Oleh']) + ' )</td><td>(.........................)</td><td>(.........................)</td></tr>' +
    '<tr><td>Admin</td><td>Manager Departemen</td><td>HRD</td></tr></table>' +
    '</body></html>';
  return HtmlService.createHtmlOutput(html).getBlob().getAs('application/pdf')
    .setName('SPL ' + String(r0['Nomor SPL']).replace(/\s*\/\s*/g, '-') + ' (' + Utilities.formatDate(r0['Tanggal Lembur'], TZ, 'dd-MM-yyyy') + ').pdf');
}

/* ======================= Menu admin ======================= */
function menuPdfSPL() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('Buat & kirim PDF SPL', 'Masukkan Nomor SPL (mis. 001 / SPL / CBA / X / 2026).\nPDF dibuat walau masih ada karyawan yang belum selesai diproses.', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  ui.alert(cekSelesaiSPL_(r.getResponseText().trim(), true) || 'Selesai.');
}
function menuKirimUlangKonfirmasi() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('Kirim ulang email konfirmasi', 'Masukkan Nomor SPL. Email dikirim ulang ke karyawan yang statusnya masih MENUNGGU KONFIRMASI.', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  kirimKonfirmasiSPL_(r.getResponseText().trim());
  ui.alert('Selesai. Lihat kolom "Status Email" di sheet Data SPL.');
}
/**
 * SISTEM SPL & SPD — File : Logo.gs
 * Logo PT Citra Banjar Abadi (JPG base64) untuk header form & kop surat PDF.
 * Jangan diubah manual. Untuk ganti logo, minta file Logo.gs baru.
 */
const LOGO_DATA_URI = 'data:image/jpeg;base64,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';
