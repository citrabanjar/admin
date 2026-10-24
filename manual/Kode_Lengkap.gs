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
const PERAN = ['KARYAWAN', 'ADMIN', 'OWNER'];   // OWNER = pemilik: atasannya diri sendiri, menerima rekap lembur semua admin, tidak bisa diberi surat; ADMIN = boleh login form & membuat surat
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
const CACHE_KARYAWAN = 'KARYAWAN_V5';
function hapusCacheKaryawan_() { try { CacheService.getScriptCache().remove(CACHE_KARYAWAN); } catch (e) {} _K = null; }
function karyawan_() {
  if (_K) return _K;
  try { const c = CacheService.getScriptCache().get(CACHE_KARYAWAN); if (c) return (_K = JSON.parse(c)); } catch (e) {}
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
      // akses form: kotak centang "Akses Lembur" / "Akses Dinas" (bila kolomnya ada), atau kolom teks lama "Akses Form" (kosong = semua)
      aksesCek: true,   // kolom "Akses Form": kartu yang dipilih saja (kosong = tidak ada akses)
      akses: aksesDari_(r['Akses Form']),
      aktif: aktifKol ? aktifKol !== 'TIDAK' : !/resign|keluar|berhenti|non ?aktif|tidak aktif|phk/i.test(statusEmp)
    };
  });
  // disimpan sementara 5 menit agar tiap klik tidak membaca ulang sheet (dihapus otomatis saat sheet karyawan diedit)
  try { CacheService.getScriptCache().put(CACHE_KARYAWAN, JSON.stringify(_K), 300); } catch (e) {}
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

/** "ANANG EFENDI" → "Sdr. Anang Efendi"; teks di dalam kurung tetap apa adanya: "HARTONO (HL)" → "Sdr. Hartono (HL)" */
function sdr_(nama) {
  return 'Sdr. ' + String(nama || '').trim().split(/(\([^)]*\))/).map(b => b.charAt(0) === '(' ? b :
    b.toLowerCase().replace(/(^|[\s.'-])([a-z])/g, (m, a, c) => a + c.toUpperCase())).join('');
}

/* ---------- Hak akses form per admin (kolom "Akses Form" di Data Karyawan) ----------
 *  Kosong / SEMUA = semua form.  Isi: LEMBUR, DINAS (boleh lebih dari satu, pisahkan koma). */
/** Daftar kartu di Beranda form. Tambah kartu baru = tambah 1 baris di sini → Setup awal membuat kolom centangnya. */
const KARTU = [
  { kode: 'LEMBUR', pilihan: 'Lembur', nama: 'Surat Perintah Lembur', alias: [] },
  { kode: 'DINAS', pilihan: 'Dinas', nama: 'Surat Perintah Dinas', alias: [] },
  { kode: 'PTJ', pilihan: 'Res Dinas', nama: 'Pertanggungjawaban Dinas', alias: ['PERTANGGUNGJAWABAN DINAS', 'PERTANGGUNGJAWABAN', 'PTJ'] }
];
/** "LEMBUR, PERTANGGUNGJAWABAN DINAS" → ['LEMBUR','PTJ'] */
function aksesDari_(teks) {
  const t = String(teks || '').toUpperCase().split(/[,;\n]+/).map(x => x.trim()).filter(Boolean);
  return KARTU.filter(x => [x.pilihan.toUpperCase(), x.kode].concat(x.alias || []).some(a => t.indexOf(a) >= 0)).map(x => x.kode);
}
const DAFTAR_AKSES = KARTU.map(x => x.kode);
function punyaAkses_(k, jenis) {
  const a = (k && k.akses) || [];
  if (k && k.aksesCek) return a.indexOf(jenis) >= 0;   // mode kotak centang: harus dicentang
  return !a.length || a.indexOf('SEMUA') >= 0 || a.indexOf(jenis) >= 0;
}
function cekAkses_(k, jenis) {
  if (!punyaAkses_(k, jenis)) throw new Error('Anda tidak memiliki akses ke form ' + ((KARTU.find(x => x.kode === jenis) || {}).nama || jenis) + '.');
}
/**
 * =====================================================================
 *  SISTEM SPL & SPD — File : Menu.gs   (TAHAP 1)
 *  Menu "⚙️ SPL / SPD" di spreadsheet + setup awal + cek data.
 * =====================================================================
 */

/** Simple trigger: edit di sheet karyawan → hapus cache data karyawan */
function onEdit(e) {
  try {
    const n = e && e.range ? e.range.getSheet().getName() : '';
    if ([SH.KARYAWAN].concat(SH_ALTERNATIF[SH.KARYAWAN] || []).indexOf(n) >= 0) CacheService.getScriptCache().remove(CACHE_KARYAWAN);
  } catch (err) {}
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('⚙️ SPL / SPD')
    .addItem('1. Setup awal (jalankan sekali)', 'setupAwal')
    .addItem('2. Cek data karyawan', 'cekDataKaryawan')
    .addItem('3. Tes hitung lembur & SPD', 'tesHitung')
    .addSeparator()
    .addItem('🔗 Link form', 'bukaForm')
    .addSeparator()
    .addItem('✉ Kirim ulang email konfirmasi / persetujuan', 'menuKirimUlangKonfirmasi')
    .addItem('📄 Buat & kirim PDF (SPL/SPD) sekarang', 'menuPdfSPL')
    .addItem('Lihat nomor surat berikutnya', 'lihatNomorBerikutnya')
    .addItem('🔑 Kunci API pengecekan KantorKu', 'menuKunciApi')
    .addToUi();
}

/** Setup awal: zona waktu, simpan ID file, kunci sheet sistem, format kolom jam. */
function setupAwal() {
  hapusCacheKaryawan_();
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
  // Kolom "Akses Form": kartu apa saja yang boleh dibuka ADMIN (pilih satu atau beberapa, dipisah koma)
  let colAkses = kolom_(SH.KARYAWAN)['Akses Form'];
  if (!colAkses) {
    colAkses = shK.getLastColumn() + 1;
    shK.getRange(1, colAkses).setValue('Akses Form').setFontWeight('bold');
  }
  shK.getRange(1, colAkses).setNote('Khusus ADMIN — pilih kartu yang boleh dibuka (boleh lebih dari satu): ' + KARTU.map(x => x.pilihan).join(', ') + '. Kosong = tidak ada akses.');
  {
    const n = shK.getLastRow() - 1, c = kolom_(SH.KARYAWAN);
    if (n > 0) {
      const v = shK.getRange(2, 1, n, shK.getLastColumn()).getValues();
      const kolomCentang = { LEMBUR: c['Akses Lembur'], DINAS: c['Akses Dinas'], PTJ: c['Akses Pertanggungjawaban Dinas'] };
      const adaCentang = Object.values(kolomCentang).some(Boolean);
      // admin yang Akses Form-nya masih kosong: isi dari kotak centang lama (bila ada) atau semua kartu
      const isi = v.map(r => {
        const lama = String(r[colAkses - 1] || '').trim();
        if (lama || String(r[(c['Peran'] || 0) - 1] || '').trim().toUpperCase() !== 'ADMIN') return [lama];
        const pilih = KARTU.filter(x => adaCentang ? (kolomCentang[x.kode] && r[kolomCentang[x.kode] - 1] === true) : true).map(x => x.pilihan);
        return [pilih.join(', ')];
      });
      shK.getRange(2, colAkses, n, 1).setValues(isi);
    }
    shK.getRange(2, colAkses, shK.getMaxRows() - 1, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(KARTU.map(x => x.pilihan), true).setAllowInvalid(true).build());
  }
  hapusCacheKaryawan_();
  // Parameter tambahan (dibuat bila belum ada)
  const shP = sheet_(SH.PARAM), adaP = shP.getRange(1, 1, Math.max(shP.getLastRow(), 1), 1).getValues().map(r => String(r[0]).trim());
  [['URL_WEB_APP', 'https://citrabanjar.github.io/admin/', 'Alamat form & link di email. Kosong = link Apps Script (script.google.com).'],
   ['URL_KANTORKU', 'https://hris.kantorku.id', 'Link aplikasi KantorKu yang ditampilkan setelah karyawan klik Bersedia.']]
    .forEach(r => { if (adaP.indexOf(r[0]) < 0) shP.appendRow(r); });
  _P = null;

  // Kolom tambahan Data SPL untuk alur email
  pastikanKolom_(SH.SPL, KOLOM_SPL_TAMBAHAN);
  pastikanKolom_(SH.PERSONEL, KOLOM_SPD_TAMBAHAN);
  rahasia_();

  logAksi_('-', '-', 'SETUP AWAL', Session.getActiveUser().getEmail(), 'Zona waktu ' + TZ + ', proteksi sheet sistem.');
  SpreadsheetApp.getUi().alert('✅ Setup awal selesai.\n\n• Zona waktu: ' + TZ +
    '\n• Sheet sistem dikunci (peringatan bila diedit manual)\n• Kolom Peran & Atasan tersedia di sheet karyawan:\n   – Peran: isi ADMIN untuk user yang boleh membuat surat (kosong = KARYAWAN)\n   – Atasan: isi Employee ID atasan langsung\n\nLangkah berikut: menu ⚙️ SPL / SPD → 2. Cek data karyawan.');
}

/** Periksa kelengkapan master karyawan sebelum notifikasi email diaktifkan. */
function cekDataKaryawan() {
  hapusCacheKaryawan_();
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
  const owner = aktif.filter(isOwner_);
  if (!owner.length) masalah.push({ j: '⚠ Belum ada OWNER', isi: ['Isi Peran = OWNER untuk pemilik (menerima rekap lembur semua admin).'] });
  if (owner.some(o => !o.email)) masalah.push({ j: '❌ OWNER tanpa email (tidak menerima rekap)', isi: owner.filter(o => !o.email).map(k => k.id + ' — ' + k.nama) });
  const tanpaAtasan = aktif.filter(k => !k.atasan && !isOwner_(k));
  if (tanpaAtasan.length) masalah.push({ j: '⚠ Atasan belum diisi (' + tanpaAtasan.length + ' orang) — persetujuan lembur mereka dikirim ke admin', isi: tanpaAtasan.map(k => k.id + ' — ' + k.nama) });
  const atasanSalah = aktif.filter(k => k.atasan && !Kmap[k.atasan] && !(k.atasan === 'BOSS'));
  if (atasanSalah.length) masalah.push({ j: '❌ Employee ID atasan tidak ditemukan', isi: atasanSalah.map(k => k.id + ' — ' + k.nama + ' → ' + k.atasan) });
  const atasanTanpaEmail = aktif.filter(k => k.atasan && Kmap[k.atasan] && !Kmap[k.atasan].email);
  if (atasanTanpaEmail.length) masalah.push({ j: '⚠ Atasan belum punya email — persetujuan dikirim ke admin', isi: atasanTanpaEmail.map(k => k.id + ' — ' + k.nama + ' → ' + k.atasan + ' ' + Kmap[k.atasan].nama) });
  const diriSendiri = aktif.filter(k => k.atasan && k.atasan === k.id && !isOwner_(k));   // owner boleh atasannya diri sendiri
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
const STATUS_BATAL = ['DITOLAK', 'DIBATALKAN', 'TIDAK BERSEDIA'];
const SESI_HARI = 30;   // login form berlaku 30 hari di perangkat yang sama (OTP hanya saat login pertama)

/* ------------------------------ Tampilan ------------------------------ */
function doGet(e) {
  if (e && e.parameter && e.parameter.a) return halamanAksi_(e.parameter);   // link dari email (Notif.gs)
  if (e && e.parameter && e.parameter.api) return apiData_(e.parameter);       // data untuk pengecekan KantorKu (Api.gs)
  const t = HtmlService.createTemplateFromFile('Form');
  // rumus yang SAMA dengan server dipakai juga di browser (satu sumber rumus)
  t.rumusBersama = 'const NAMA_HARI = ' + JSON.stringify(NAMA_HARI) + ';\n' +
    [ymd_, hariKe_, keTanggal_, jamKeAngka_, angkaKeJam_, rentang_, statusHari_, hitungLembur_, hitungSPD_, sdr_]
      .map(f => f.toString()).join('\n');
  t.logo = LOGO_DATA_URI;
  return t.evaluate()
    .setTitle('Form — ' + (param_().NAMA_PERUSAHAAN || ''))
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Menu spreadsheet: tampilkan link form */
function bukaForm() {
  const url = String(param_().URL_WEB_APP || '').trim() || ScriptApp.getService().getUrl();
  const ui = SpreadsheetApp.getUi();
  if (!url) { ui.alert('Form belum di-deploy.\n\nDeploy → New deployment → Web app (lihat panduan).'); return; }
  ui.showModalDialog(HtmlService.createHtmlOutput(
    '<div style="font-family:Arial;font-size:13px">Link form (bagikan ke user ADMIN):<br><br>' +
    '<input style="width:100%;padding:6px" value="' + url + '" onclick="this.select()" readonly><br><br>' +
    '<a href="' + url + '" target="_blank">Buka form ↗</a></div>').setWidth(520).setHeight(170), 'Link Form');
}

/* ------------------------------ Login ------------------------------ */
function pembuatByEmail_(email) {
  const k = cariKaryawanByEmail_(email);
  return (k && k.aktif && PERAN_PEMBUAT.indexOf(k.peran) >= 0) ? k : null;
}
function infoUser_(k) { return { akses: DAFTAR_AKSES.filter(j => punyaAkses_(k, j)), id: k.id, nama: k.nama, jabatan: k.jabatan, departemen: k.departemen, peran: k.peran, email: k.email }; }
/** Sesi login disimpan di Script Properties (bertahan 30 hari, tidak hilang seperti cache). */
function sesiBuat_(email) {
  const tok = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  const P = PropertiesService.getScriptProperties();
  P.setProperty('S_' + tok, JSON.stringify({ e: email, x: Date.now() + SESI_HARI * 864e5 }));
  // bersihkan sesi kedaluwarsa
  const semua = P.getProperties(), kini = Date.now();
  Object.keys(semua).forEach(k => { if (k.indexOf('S_') === 0) { try { if (JSON.parse(semua[k]).x < kini) P.deleteProperty(k); } catch (e) { P.deleteProperty(k); } } });
  return tok;
}
function sesiCek_(tok) {
  let s = null;
  try { s = tok ? JSON.parse(PropertiesService.getScriptProperties().getProperty('S_' + tok) || 'null') : null; } catch (e) {}
  if (!s || s.x < Date.now()) throw new Error('SESI_HABIS');
  const k = pembuatByEmail_(s.e);
  if (!k) throw new Error('Akun ini tidak berhak membuat SPL / SPD.');
  return k;
}
/** Keluar: hapus sesi di server */
function keluarSesi(tok) { if (tok) PropertiesService.getScriptProperties().deleteProperty('S_' + tok); return true; }

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
    subject: 'Kode masuk form Anda: ' + kode,
    htmlBody: '<div style="font-family:Arial;font-size:14px">Kode masuk form Anda:<br>' +
      '<div style="font-size:28px;font-weight:bold;letter-spacing:6px;margin:12px 0;color:#BE353F">' + kode + '</div>' +
      'Admin PT. Citra Banjar Abadi</div>'
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
  logAksi_('-', '-', 'LOGIN', email, 'Login form');
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
  const kar = Object.values(karyawan_()).filter(k => k.aktif && !isOwner_(k))   // owner tidak muncul di pencarian karyawan
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
    if (spdAktif[r['Nomor SPD']] && STATUS_BATAL.indexOf(String(r['Status'] || '')) < 0) out.spd[String(r['Employee ID']).toUpperCase()] = r['Nomor SPD'];
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
  cekAkses_(u, 'LEMBUR');
  pastikanKolom_(SH.SPL, KOLOM_SPL_TAMBAHAN);
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('Server sibuk, coba simpan lagi beberapa detik lagi.');
  let hasilSimpan;
  try {
    _K = null; _L = null; _P = null;
    const P = param_(), K = karyawan_(), L = liburEfektif_(data && data.tanggal, data && data.liburNasional);
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
      if (isOwner_(k)) { err.push(sdr_(k.nama) + ' adalah OWNER — tidak dapat diberi Surat Perintah Lembur.'); return null; }
      if (id === String(u.id).toUpperCase()) err.push(no + 'Anda tidak bisa memasukkan diri sendiri — Surat Perintah Lembur untuk Anda harus diinput oleh admin lain.');
      if (sudah[id]) err.push(sdr_(k.nama) + ' tercantum lebih dari sekali.');
      sudah[id] = true;
      if (!String(b.uraian || '').trim()) err.push('Uraian pekerjaan ' + sdr_(k.nama) + ' wajib diisi.');
      if (aktif.spl[id]) err.push(sdr_(k.nama) + ' sudah terdaftar di SPL ' + String(aktif.spl[id]).replace(/\s*\(\d{4}-\d{2}-\d{2}\)$/, '') + '.');
      if (aktif.spd[id]) err.push(sdr_(k.nama) + ' sedang dinas luar (' + aktif.spd[id] + ') — lembur tidak dapat diajukan.');
      const h = hitungLembur_({ tanggal: tglStr, mulai: b.mulai, selesai: b.selesai, gaji: k.gaji }, P, L);
      h.error.forEach(e => err.push(sdr_(k.nama) + ' — ' + e));
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
      'Catatan': [].concat(data.liburNasional && !liburSet_()[tglStr] ? ['Ditandai libur nasional oleh admin'] : [], x.h.peringatan).join(' | '),
      'Status Email': 'Mengirim…',
      'ID Baris': Utilities.getUuid().slice(0, 8)
    })), ['Jam Mulai', 'Jam Selesai']);
    SpreadsheetApp.flush();

    if (data.liburNasional) catatLibur_(tglStr, u.nama);
    const totalJam = hasil.reduce((s, x) => s + x.h.jamEfektif, 0);
    logAksi_('SPL', nomor, 'DIBUAT', u.email, hasil.length + ' karyawan, ' + totalJam + ' jam, tgl ' + tglStr);
    hasilSimpan = { ok: true, nomor: nomor, jumlah: hasil.length, totalJam: totalJam, status: ST.KONFIRMASI };
  } finally {
    lock.releaseLock();
  }
  // Email konfirmasi ke tiap karyawan (di luar kunci agar pengguna lain tidak menunggu)
  try { antri_({ t: 'konfirmasi', jenis: 'SPL', nomor: hasilSimpan.nomor }); } catch (e) { logAksi_('SPL', hasilSimpan.nomor, 'EMAIL GAGAL', u.email, String(e)); }
  return hasilSimpan;
}

/** Daftar libur + tanggal yang dicentang "Libur Nasional" di form */
function liburEfektif_(tglStr, centang) {
  const L = Object.assign({}, liburSet_());
  if (centang && /^\d{4}-\d{2}-\d{2}$/.test(String(tglStr || '')) && !L[tglStr]) L[tglStr] = 'Libur nasional';
  return L;
}
/** Tanggal yang dicentang libur nasional → dicatat ke sheet Libur (bila belum ada) */
function catatLibur_(tglStr, oleh) {
  if (liburSet_()[tglStr]) return;
  sheet_(SH.LIBUR).appendRow([Utilities.parseDate(tglStr + ' 00:00', TZ, 'yyyy-MM-dd HH:mm'), '', 'Libur nasional (ditandai ' + oleh + ' via form SPL)']);
  _L = null;
}

/** Riwayat SPL buatan user (30 terakhir), dikelompokkan per nomor, dengan status per karyawan. */
function riwayatSPL(token) {
  const u = sesiCek_(token);
  if (!punyaAkses_(u, 'LEMBUR')) return [];
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
      jam: Number(r['Jam Efektif']) || 0, uraian: r['Uraian Pekerjaan'], status: r['Status'],
      batalOleh: r['Dibatalkan Oleh'] || '', alasanBatal: r['Alasan Pembatalan'] || '' });
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
 *          transportasi, sumberBiaya, uangKas, personel:[{id, lama, keterangan}], biaya:[{jenis, uraian, jumlah, bukti}]}} d
 */
function simpanSPD(token, d) {
  const u = sesiCek_(token);
  cekAkses_(u, 'DINAS');
  pastikanKolom_(SH.PERSONEL, KOLOM_SPD_TAMBAHAN);
  pastikanKolom_(SH.SPD, ['Uang Kas Perjalanan']);
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('Server sibuk, coba simpan lagi beberapa detik lagi.');
  let hasilSimpan;
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
    const uangKas = Number(d.uangKas);
    if (d.uangKas === '' || d.uangKas == null || !(uangKas >= 0)) err.push('Uang kas dalam perjalanan wajib diisi (isi 0 bila tidak ada).');

    const personel = (d.personel || []).filter(p => p && teks(p.id)).map(p => Object.assign({}, p, { id: teks(p.id).toUpperCase() }));
    const biaya = (d.biaya || []).filter(b => b && (teks(b.jenis) || Number(b.jumlah)));
    personel.forEach((p, i) => { if (!K[p.id] || !K[p.id].aktif || isOwner_(K[p.id])) err.push('Personel ke-' + (i + 1) + ': Employee ID ' + p.id + ' tidak ditemukan / tidak aktif.'); });
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
        if (aktif.spd[p.id]) err.push(sdr_(k.nama) + ' sudah ditugaskan di ' + aktif.spd[p.id] + ' pada tanggal yang beririsan.');
        if (aktif.spl[p.id]) warn.push(sdr_(k.nama) + ' punya SPL ' + aktif.spl[p.id] + ' di rentang dinas — SPL tersebut perlu ditinjau.');
      });
    }
    if (err.length) return { ok: false, error: err };

    const alur = { status: 'PROSES 0/' + personel.length };
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
      'Uang Kas Perjalanan': uangKas || 0, 'Total Uang Saku': h.totalUangSaku, 'Total Biaya Lain': h.totalBiayaLain, 'Total Biaya': h.totalBiaya,
      'Status Dinas': 'Aktif', 'Departemen': u.departemen,
      'Dibuat Oleh': u.nama, 'Email Pembuat': u.email, 'Dibuat Pada': now,
      'Status': alur.status,
      'Diketahui Oleh': alur.diketahuiOleh || '', 'Diketahui Pada': alur.diketahuiOleh ? now : '',
      'Catatan': [].concat(warn, alur.catatan ? [alur.catatan] : []).join(' | '),
      'Status Email': 'Konfirmasi dikirim ke personel'
    }]);
    tulisBaris_(SH.PERSONEL, h.personel.map((p, i) => ({
      'Nomor SPD': nomor, 'No': i + 1, 'Employee ID': p.id, 'Full Name': K[p.id].nama, 'Job Position': K[p.id].jabatan,
      'Lama (hari)': Number(p.lama), 'Uang Saku': p.uangSaku, 'Keterangan': teks(p.keterangan),
      'Status': ST.KONFIRMASI, 'Status Email': 'Mengirim…', 'ID Baris': Utilities.getUuid().slice(0, 8)
    })));
    tulisBaris_(SH.BIAYA, biaya.map((b, i) => ({
      'Nomor SPD': nomor, 'No': i + 1, 'Jenis Biaya': teks(b.jenis), 'Uraian': teks(b.uraian),
      'Jumlah (Rp)': Number(b.jumlah) || 0, 'Bukti': teks(b.bukti)
    })));
    SpreadsheetApp.flush();

    logAksi_('SPD', nomor, 'DIBUAT', u.email, personel.length + ' personel, ' + h.lama + ' hari ke ' + teks(d.tempat) + ', total Rp' + h.totalBiaya + ' → ' + alur.status);
    hasilSimpan = { ok: true, nomor: nomor, jumlah: personel.length, lama: h.lama, totalBiaya: h.totalBiaya, status: ST.KONFIRMASI, peringatan: warn };
  } finally {
    lock.releaseLock();
  }
  try { antri_({ t: 'konfirmasi', jenis: 'SPD', nomor: hasilSimpan.nomor }); } catch (e) { logAksi_('SPD', hasilSimpan.nomor, 'EMAIL GAGAL', u.email, String(e)); }
  return hasilSimpan;
}

/** Riwayat SPD buatan user (30 terakhir). */
function riwayatSPD(token) {
  const u = sesiCek_(token);
  if (!punyaAkses_(u, 'DINAS') && !punyaAkses_(u, 'PTJ')) return [];
  const per = {};
  bacaTabel_(SH.PERSONEL, 'Nomor SPD').forEach(r => {
    (per[r['Nomor SPD']] = per[r['Nomor SPD']] || []).push({ nama: r['Full Name'], lama: Number(r['Lama (hari)']) || 0, keterangan: r['Keterangan'], status: r['Status'] || '' });
  });
  return bacaTabel_(SH.SPD, 'Nomor SPD')
    .filter(r => String(r['Email Pembuat']).toLowerCase() === u.email)
    .map(r => ({
      jenis: 'SPD', nomor: r['Nomor SPD'], status: r['Status'],
      tanggal: r['Tanggal Surat'] instanceof Date ? ymd_(r['Tanggal Surat']) : '',
      berangkat: r['Tgl Berangkat'] instanceof Date ? ymd_(r['Tgl Berangkat']) : '',
      kembali: r['Tgl Kembali'] instanceof Date ? ymd_(r['Tgl Kembali']) : '',
      lama: Number(r['Lama (hari)']) || 0, tempat: r['Tempat Tujuan'], maksud: r['Maksud / Tujuan'],
      totalBiaya: Number(r['Total Biaya']) || 0, uangKas: Number(r['Uang Kas Perjalanan']) || 0,
      dibuat: r['Dibuat Pada'] instanceof Date ? r['Dibuat Pada'].getTime() : 0,
      orang: per[r['Nomor SPD']] || [], pdf: r['Link PDF'] || '', pdfLaporan: r['Link PDF Laporan'] || '',
      statusDinas: String(r['Status Dinas'] || 'Aktif') === 'Selesai' ? 'SELESAI' : 'OPEN'
    }))
    .sort((a, b) => b.dibuat - a.dibuat).slice(0, 30);
}
/**
 * =====================================================================
 *  SISTEM SPL & SPD — File : Notif.gs
 *  Alur email (sama untuk SPL & SPD):
 *   1. Admin buat surat → email ke tiap karyawan/personel: [Bersedia] / [Tidak Bersedia]
 *      (tanpa email → email konfirmasi dikirim ke admin pembuat)
 *   2. Tidak Bersedia → email ke admin pembuat + karyawan ybs
 *   3. Bersedia → email ke ATASAN LANGSUNG (kolom "Atasan" di Data Karyawan):
 *      [Setujui] / [Tolak]   (atasan kosong → dikirim ke admin pembuat)
 *   4. Semua orang di surat sudah final → PDF (kop surat) dikirim ke admin
 *      & disimpan di folder "PDF SPL" / "PDF SPD" (satu folder dengan spreadsheet)
 *  Status per orang: MENUNGGU KONFIRMASI → TIDAK BERSEDIA | MENUNGGU ATASAN → DISETUJUI | DITOLAK
 * =====================================================================
 */

const ST = {
  KONFIRMASI: 'MENUNGGU KONFIRMASI', TIDAK: 'TIDAK BERSEDIA', ATASAN: 'MENUNGGU ATASAN',
  SETUJU: 'DISETUJUI', TOLAK: 'DITOLAK', BATAL: 'DIBATALKAN'
};
const ST_FINAL = [ST.TIDAK, ST.SETUJU, ST.TOLAK, ST.BATAL];
const KOLOM_SPL_TAMBAHAN = ['Atasan', 'Email Atasan', 'Konfirmasi Pada', 'Alasan Tidak Bersedia', 'Catatan Atasan', 'Link PDF', 'Link PDF Owner', 'Link PDF Atasan', 'KantorKu', 'Dibatalkan Oleh', 'Dibatalkan Pada', 'Alasan Pembatalan'];
const KOLOM_SPD_TAMBAHAN = ['Status', 'Konfirmasi Pada', 'Alasan Tidak Bersedia', 'Atasan', 'Email Atasan',
  'Disetujui Oleh', 'Disetujui Pada', 'Catatan Atasan', 'Status Email', 'ID Baris'];

/** Konfigurasi per jenis surat. Kode aksi di link email: k/t (SPL), dk/dt (SPD). */
const JENIS = {
  SPL: { sheet: 'Data SPL', kunci: 'Nomor SPL', tambahan: KOLOM_SPL_TAMBAHAN, k: 'k', t: 't', kata: 'Lembur', surat: 'Surat Perintah Lembur', folder: 'PDF SPL' },
  SPD: { sheet: 'Personel SPD', kunci: 'Nomor SPD', tambahan: KOLOM_SPD_TAMBAHAN, k: 'dk', t: 'dt', kata: 'Dinas Luar', surat: 'Surat Perintah Dinas', folder: 'PDF SPD' }
};
const AKSI = { k: ['SPL', 'k'], t: ['SPL', 't'], dk: ['SPD', 'k'], dt: ['SPD', 't'] };

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
function urlWeb_() { return String(param_().URL_WEB_APP || '').trim() || ScriptApp.getService().getUrl(); }
function linkAksi_(aksi, id, pilih) {
  return urlWeb_() + '?a=' + aksi + '&id=' + encodeURIComponent(id) + '&s=' + tandaTangan_(aksi, id) + (pilih ? '&p=' + pilih : '');
}
function logoBlob_() { return Utilities.newBlob(Utilities.base64Decode(LOGO_DATA_URI.split(',')[1]), 'image/jpeg', 'logo.jpg'); }
const escH_ = s => String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const rpH_ = n => 'Rp ' + Math.round(Number(n) || 0).toString().replace(/\B(?=(\d{3})+(?!\d))/g, '.');
function tglIndo_(d) {
  if (!(d instanceof Date)) return String(d || '');
  const bln = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
  const [y, m, dd] = Utilities.formatDate(d, TZ, 'yyyy-M-d').split('-').map(Number);
  return dd + ' ' + bln[m - 1] + ' ' + y;
}
function hariIndo_(d) { return d instanceof Date ? NAMA_HARI[hariKe_(d) - 1] : ''; }
function waktuIndo_(d) { return d instanceof Date ? tglIndo_(d) + ' ' + Utilities.formatDate(d, TZ, 'HH:mm') : ''; }

/** Baca baris orang per jenis → {sh, c, rows:[{_row, ...kolom}]} */
function bacaBaris_(jenis) {
  const J = JENIS[jenis], sh = sheet_(J.sheet), c = kolom_(J.sheet);
  const n = sh.getLastRow() - 1;
  const v = n > 0 ? sh.getRange(2, 1, n, sh.getLastColumn()).getValues() : [];
  const head = Object.keys(c), rows = [];
  v.forEach((r, i) => {
    if (!String(r[c[J.kunci] - 1]).trim()) return;
    const o = { _row: i + 2 }; head.forEach(h => o[h] = r[c[h] - 1]); rows.push(o);
  });
  return { sh: sh, c: c, rows: rows };
}
function bacaSPL_() { return bacaBaris_('SPL'); }
function ubahBaris_(S, row, data) {
  Object.keys(data).forEach(j => { if (S.c[j]) S.sh.getRange(row, S.c[j]).setValue(data[j]); });
}
/** Data induk surat: SPL = baris itu sendiri; SPD = baris di sheet Data SPD (+ _row) */
function induk_(jenis, nomor, r) {
  if (jenis === 'SPL') return r;
  const S = bacaBaris_SPD_induk_();
  return S.rows.find(x => x['Nomor SPD'] === nomor) || {};
}
function bacaBaris_SPD_induk_() {
  const sh = sheet_(SH.SPD), c = kolom_(SH.SPD), n = sh.getLastRow() - 1;
  const v = n > 0 ? sh.getRange(2, 1, n, sh.getLastColumn()).getValues() : [];
  const rows = [];
  v.forEach((r, i) => { if (String(r[c['Nomor SPD'] - 1]).trim()) { const o = { _row: i + 2 }; Object.keys(c).forEach(h => o[h] = r[c[h] - 1]); rows.push(o); } });
  return { sh: sh, c: c, rows: rows };
}
function ringkasTanggal_(jenis, r, I) {
  return jenis === 'SPL' ? r['Hari'] + ', ' + tglIndo_(r['Tanggal Lembur'])
    : I['Tempat Tujuan'] + ', ' + tglIndo_(I['Tgl Berangkat']) + ' s/d ' + tglIndo_(I['Tgl Kembali']);
}

/** Kirim email (MailApp) dengan logo inline. '' bila sukses, pesan error bila gagal. */
function kirimEmail_(ke, subjek, isiHtml, cc, lampiran) {
  try {
    if (!ke) return 'tanpa email';
    if (MailApp.getRemainingDailyQuota() < 1) return 'kuota email harian habis';
    const opt = { to: ke, subject: subjek,
      htmlBody: bungkusEmail_(isiHtml), inlineImages: { logo: logoBlob_() } };
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
    '<div style="font-size:11px;color:#888;border-top:1px solid #eee;padding-top:8px">Admin PT. Citra Banjar Abadi — email otomatis, mohon tidak membalas email ini.</div></div>';
}
function tombol_(href, teks, warna) {
  return '<a href="' + href + '" style="display:inline-block;background:' + warna + ';color:#fff;text-decoration:none;font-weight:bold;padding:12px 22px;border-radius:8px;margin:4px 8px 4px 0">' + teks + '</a>';
}
function tabelDetail_(jenis, r, I, denganKompensasi, tanpaPersonel) {
  const baris = jenis === 'SPL' ? [
    ['Nomor SPL', r['Nomor SPL']],
    ['Tanggal Lembur', r['Hari'] + ', ' + tglIndo_(r['Tanggal Lembur']) + ' (' + r['Status Hari'] + ')'],
    ['Karyawan', sdr_(r['Full Name']) + ' (' + r['Employee ID'] + ')'],
    ['Jam Lembur', r['Jam Mulai'] + ' – ' + r['Jam Selesai'] + (r['Istirahat'] && r['Istirahat'] !== '-' ? ' (istirahat ' + r['Istirahat'] + ')' : '')],
    ['Jam Efektif', String(r['Jam Efektif']).replace('.', ',') + ' jam'],
    ['Uraian Pekerjaan', r['Uraian Pekerjaan']]
  ].concat(denganKompensasi ? [['Estimasi Kompensasi', rpH_(r['Estimasi Kompensasi'])]] : [], [
    ['Diperintahkan oleh', sdr_(namaAtasan_(r['Employee ID']) || r['Dibuat Oleh'])]
  ]) : [
    ['Nomor SPD', r['Nomor SPD']],
    ['Maksud / Tujuan', I['Maksud / Tujuan']],
    ['Tempat Tujuan', I['Tempat Tujuan']],
    ['Tanggal Dinas', hariIndo_(I['Tgl Berangkat']) + ', ' + tglIndo_(I['Tgl Berangkat']) + ' s/d ' + hariIndo_(I['Tgl Kembali']) + ', ' + tglIndo_(I['Tgl Kembali']) + ' (' + I['Lama (hari)'] + ' hari)'],
    ['Personel', sdr_(r['Full Name']) + ' (' + r['Employee ID'] + ')' + (r['Job Position'] ? ' — ' + r['Job Position'] : '')],
    ['Lama Tugas', r['Lama (hari)'] + ' hari' + (r['Keterangan'] ? ' — ' + r['Keterangan'] : '')],
    ['Uang Saku', rpH_(r['Uang Saku']) + ' (dibayarkan melalui aplikasi KantorKu)'],
  ].concat(Number(I['Uang Kas Perjalanan']) ? [['Uang Kas Perjalanan', rpH_(I['Uang Kas Perjalanan']) + ' (untuk seluruh tim, dipertanggungjawabkan setelah kembali)']] : [], [
    ['Transportasi', I['Transportasi']],
    ['Customer / Dasar', [I['Nama Customer'], I['Dasar Penugasan']].filter(Boolean).join(' — ') || '-'],
    ['Diperintahkan oleh', sdr_((owner_()[0] || {}).nama || I['Dibuat Oleh'])]   // Surat Perintah Dinas diperintahkan langsung oleh owner
  ]).filter(b => !tanpaPersonel || ['Personel', 'Lama Tugas', 'Uang Saku'].indexOf(b[0]) < 0);
  return '<table style="border-collapse:collapse;width:100%;font-size:13px">' + baris.map(b =>
    '<tr><td style="padding:6px 8px;background:#F4F5F7;width:150px;border:1px solid #E2E4E8;font-weight:bold">' + escH_(b[0]) +
    '</td><td style="padding:6px 8px;border:1px solid #E2E4E8">' + escH_(b[1]) + '</td></tr>').join('') + '</table>';
}

/* ======================= Instruksi KantorKu (setelah karyawan Bersedia) ======================= */
function htmlInstruksiKantorku_() {
  const url = String(param_().URL_KANTORKU || 'https://hris.kantorku.id').trim();
  return '<div class="instruksi">Silakan lanjutkan pengajuan lembur melalui aplikasi KantorKu:' +
    '<a class="kk" href="' + escH_(url) + '" target="_blank" rel="noopener">' +
    '<img src="' + LOGO_KANTORKU_URI + '" alt="KantorKu"><span>Buka aplikasi ›</span></a></div>';
}

/* ======================= 1. Email konfirmasi ======================= */
function kirimKonfirmasi_(jenis, nomor) {
  const J = JENIS[jenis], S = bacaBaris_(jenis), K = karyawan_();
  const rows = S.rows.filter(r => r[J.kunci] === nomor && r['Status'] === ST.KONFIRMASI);
  if (!rows.length) return;
  const I = induk_(jenis, nomor, rows[0]);
  rows.forEach(r => {
    const Ir = jenis === 'SPL' ? r : I;
    const k = K[String(r['Employee ID']).toUpperCase()] || {};
    const id = String(r['ID Baris']), keAdmin = !k.email;
    const isi = (keAdmin
      ? '<p><b>' + escH_(sdr_(r['Full Name'])) + '</b> belum memiliki email di sistem. Mohon konfirmasi kesediaan ' + J.kata.toLowerCase() + ' <b>atas nama karyawan tersebut</b>:</p>'
      : '<p>Yth. <b>' + escH_(sdr_(r['Full Name'])) + '</b>,</p><p>Anda mendapat <b>' + J.surat + '</b> dengan rincian berikut. Mohon konfirmasi kesediaan Anda:</p>') +
      tabelDetail_(jenis, r, Ir, true) + '<p style="margin-top:18px">' +
      tombol_(linkAksi_(J.k, id, 'ya'), '✅ Bersedia', '#1E7B4A') + tombol_(linkAksi_(J.k, id, 'tidak'), '❌ Tidak Bersedia', '#B42318') + '</p>';
    const err = kirimEmail_(keAdmin ? Ir['Email Pembuat'] : k.email,
      'Konfirmasi ' + J.kata + ' ' + ringkasTanggal_(jenis, r, Ir) + (keAdmin ? ' — atas nama ' + sdr_(r['Full Name']) : ''), isi);
    ubahBaris_(S, r._row, { 'Status Email': err ? 'GAGAL: ' + err : 'Konfirmasi terkirim ke ' + (keAdmin ? 'admin ' : '') + (keAdmin ? Ir['Email Pembuat'] : k.email) + ' ' + Utilities.formatDate(new Date(), TZ, 'dd/MM HH:mm') });
  });
}
function kirimKonfirmasiSPL_(nomor) { kirimKonfirmasi_('SPL', nomor); }

/* ======================= Halaman aksi (link dari email) ======================= */
function halamanAksi_(p) {
  if (p.a === 'b') return halamanBatal_(p);   // pembatalan lembur oleh atasan / owner (Batal.gs)
  const aksi = p.a, id = p.id, sig = p.s;
  let isi;
  try {
    if (!AKSI[aksi] || !id || sig !== tandaTangan_(aksi, id)) throw new Error('Link tidak valid atau sudah diubah.');
    const [jenis, tipe] = AKSI[aksi], J = JENIS[jenis];
    const S = bacaBaris_(jenis), r = S.rows.find(x => String(x['ID Baris']) === String(id));
    if (!r) throw new Error('Data tidak ditemukan (mungkin sudah dihapus).');
    const I = induk_(jenis, r[J.kunci], r);
    const sudah = tipe === 'k' ? r['Status'] !== ST.KONFIRMASI : r['Status'] !== ST.ATASAN;
    isi = '<h2>' + (tipe === 'k' ? 'Konfirmasi Kesediaan ' : 'Persetujuan ') + J.kata + '</h2>' + tabelDetail_(jenis, r, I, tipe === 'k').replace(/width:150px/g, 'width:130px');   // estimasi kompensasi hanya untuk halaman karyawan
    if (tipe === 't') isi += '<p class="sub">Karyawan telah menyatakan <b>BERSEDIA</b>' + (r['Konfirmasi Pada'] instanceof Date ? ' pada ' + waktuIndo_(r['Konfirmasi Pada']) : '') + '.</p>';
    if (sudah) {
      isi += '<div class="msg info">Sudah diproses — status saat ini: <b>' + escH_(r['Status']) + '</b>.</div>';
      if (jenis === 'SPL' && tipe === 'k' && [ST.ATASAN, ST.SETUJU].indexOf(r['Status']) >= 0) isi += htmlInstruksiKantorku_();
    } else {
      const ya = tipe === 'k' ? 'Bersedia' : 'Setujui', tidak = tipe === 'k' ? 'Tidak Bersedia' : 'Tolak';
      const pil = p.p === 'tidak' ? 'tidak' : 'ya';
      isi += '<div class="pilih">' +
        '<label class="op ya"><input type="radio" name="j" value="ya"' + (pil === 'ya' ? ' checked' : '') + ' onchange="ubah()"> ✅ ' + ya + '</label>' +
        '<label class="op tidak"><input type="radio" name="j" value="tidak"' + (pil === 'tidak' ? ' checked' : '') + ' onchange="ubah()"> ❌ ' + tidak + '</label></div>' +
        '<div id="al"><label>' + (tipe === 'k' ? 'Alasan tidak bersedia' : 'Catatan') + ' <span class="sub">(opsional)</span></label><textarea id="alasan" rows="3"></textarea></div>' +
        '<button id="b" onclick="kirim()">Kirim</button><div id="o"></div>' +
        '<script>function ubah(){var t=document.querySelector("input[name=j]:checked").value;document.getElementById("al").style.display=(t==="tidak"||' + (tipe === 't') + ')?"block":"none"}ubah();' +
        'function kirim(){var b=document.getElementById("b"),o=document.getElementById("o");b.disabled=true;b.textContent="Mengirim…";' +
        'google.script.run.withSuccessHandler(function(m){document.querySelector(".pilih").remove();document.getElementById("al").remove();b.remove();o.innerHTML="<div class=\\"msg ok\\">"+m+"</div>"})' +
        '.withFailureHandler(function(e){b.disabled=false;b.textContent="Kirim";o.innerHTML="<div class=\\"msg err\\">"+e.message+"</div>"})' +
        '.prosesAksi(' + JSON.stringify(aksi) + ',' + JSON.stringify(String(id)) + ',' + JSON.stringify(sig) + ',document.querySelector("input[name=j]:checked").value,document.getElementById("alasan").value)}</script>';
    }
  } catch (e) {
    isi = '<h2>Tidak dapat dibuka</h2><div class="msg err">' + escH_(e.message) + '</div>';
  }
  return halamanHtml_(isi);
}
function halamanHtml_(isi) {
  const html = '<!DOCTYPE html><html><head><base target="_top"><meta charset="utf-8"><style>' +
    'body{margin:0;font-family:Arial,sans-serif;font-size:14px;background:#F5F5F6;color:#222}' +
    'header{background:#fff;border-bottom:4px solid #BE353F;padding:10px 16px}header img{height:42px;max-width:100%}' +
    'main{max-width:620px;margin:16px auto;padding:0 16px}.card{background:#fff;border:1px solid #E2E4E8;border-radius:10px;padding:18px}' +
    'h2{margin:0 0 14px;font-size:18px;color:#3A3A3A}.sub{color:#6B6F76;font-size:12px}' +
    '.pilih{display:flex;gap:10px;margin:18px 0 12px;flex-wrap:wrap}.op{flex:1;min-width:160px;border:2px solid #E2E4E8;border-radius:10px;padding:14px;font-weight:bold;cursor:pointer;font-size:15px}' +
    '.op.ya:has(input:checked){border-color:#1E7B4A;background:#E7F5EC}.op.tidak:has(input:checked){border-color:#B42318;background:#FDECEA}' +
    'label{display:block;font-weight:bold;font-size:12px;color:#6B6F76;margin-bottom:4px}textarea{width:100%;box-sizing:border-box;font:inherit;padding:8px;border:1px solid #CBCED4;border-radius:6px}' +
    'button{margin-top:12px;background:#BE353F;color:#fff;border:0;border-radius:8px;padding:12px 26px;font-weight:bold;font-size:15px;cursor:pointer}button:disabled{background:#B9BCC2}' +
    '.instruksi{margin-top:14px;padding:14px 16px;border:1px solid #F3C9CC;background:#FFF7F7;border-radius:10px;font-size:14px;line-height:1.5}' +
    '.instruksi a.kk{display:flex;align-items:center;justify-content:space-between;gap:12px;margin-top:12px;padding:10px 14px;background:#FAF8FE;border:1px solid #E3DCF5;border-radius:12px;text-decoration:none}' +
    '.instruksi a.kk:hover{box-shadow:0 4px 14px rgba(108,58,214,.15)}.instruksi a.kk img{height:40px;display:block}' +
    '.instruksi a.kk span{color:#6C3AD6;font-weight:bold;font-size:14px;white-space:nowrap}' +
    '.msg{padding:12px;border-radius:8px;margin-top:14px}.ok{background:#E7F5EC;color:#1E7B4A}.err{background:#FDECEA;color:#B42318}.info{background:#EFF0F2}' +
    '</style></head><body><header><img src="' + LOGO_DATA_URI + '" alt="PT Citra Banjar Abadi"></header><main><div class="card">' + isi + '</div></main></body></html>';
  return HtmlService.createHtmlOutput(html).setTitle('Form — PT Citra Banjar Abadi')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Dipanggil dari halaman aksi. jawaban: 'ya' | 'tidak' */
function prosesAksi(aksi, id, sig, jawaban, alasan) {
  if (!AKSI[aksi] || sig !== tandaTangan_(aksi, id)) throw new Error('Link tidak valid.');
  const [jenis, tipe] = AKSI[aksi], J = JENIS[jenis];
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('Server sibuk, coba lagi.');
  let nomor, pesan;
  try {
    _K = null; _P = null;
    pastikanKolom_(J.sheet, J.tambahan);
    const S = bacaBaris_(jenis), r = S.rows.find(x => String(x['ID Baris']) === String(id));
    if (!r) throw new Error('Data tidak ditemukan.');
    nomor = r[J.kunci];
    const I = induk_(jenis, nomor, r);
    const now = new Date(), alas = String(alasan || '').trim().slice(0, 500);
    const tgl = ringkasTanggal_(jenis, r, I);

    if (tipe === 'k') {
      if (r['Status'] !== ST.KONFIRMASI) return 'Sudah dijawab sebelumnya — status: ' + r['Status'];
      if (jawaban === 'tidak') {
        ubahBaris_(S, r._row, { 'Status': ST.TIDAK, 'Konfirmasi Pada': now, 'Alasan Tidak Bersedia': alas });
        const k = karyawan_()[String(r['Employee ID']).toUpperCase()] || {};
        const isi = '<p><b>' + escH_(sdr_(r['Full Name'])) + '</b> menyatakan <b style="color:#B42318">TIDAK BERSEDIA</b> ' + J.kata.toLowerCase() + '.</p>' +
          tabelDetail_(jenis, r, I) + (alas ? '<p><b>Alasan:</b> ' + escH_(alas) + '</p>' : '');
        kirimEmail_(I['Email Pembuat'], 'Tidak Bersedia ' + J.kata + ' — ' + sdr_(r['Full Name']) + ' (' + tgl + ')', isi, k.email || '');
        logAksi_(jenis, nomor, 'TIDAK BERSEDIA', r['Full Name'], alas);
        pesan = 'Terima kasih. Ketidaksediaan Anda sudah dicatat dan diberitahukan ke admin.';
      } else {
        const atasan = jenis === 'SPD' ? penyetujuDinas_(r['Employee ID']) : cariAtasan_(r['Employee ID']);   // dinas: 1 level di bawah owner
        const ke = atasan ? atasan.email : I['Email Pembuat'];
        ubahBaris_(S, r._row, { 'Status': ST.ATASAN, 'Konfirmasi Pada': now,
          'Atasan': atasan ? atasan.nama : '(admin — atasan belum diisi)', 'Email Atasan': ke });
        if (jenis === 'SPD') ubahBaris_(S, r._row, { 'Status Email': 'Persetujuan menunggu personel lain (dikirim gabungan)' });
        else kirimPersetujuan_(jenis, S, r, I, atasan, ke);
        logAksi_(jenis, nomor, 'BERSEDIA', r['Full Name'], 'Diteruskan ke ' + ke);
        pesan = 'Terima kasih. Kesediaan Anda sudah dicatat dan diteruskan ke ' + (atasan ? 'atasan (' + escH_(sdr_(atasan.nama)) + ')' : 'admin') + ' untuk persetujuan.' +
          (jenis === 'SPL' ? '</div>' + htmlInstruksiKantorku_() + '<div>' : '');
      }
    } else {
      if (r['Status'] !== ST.ATASAN) return 'Sudah diproses sebelumnya — status: ' + r['Status'];
      const setuju = jawaban !== 'tidak';
      ubahBaris_(S, r._row, { 'Status': setuju ? ST.SETUJU : ST.TOLAK, 'Disetujui Oleh': r['Atasan'], 'Disetujui Pada': now, 'Catatan Atasan': alas });
      logAksi_(jenis, nomor, setuju ? 'DISETUJUI ATASAN' : 'DITOLAK ATASAN', r['Email Atasan'], r['Full Name'] + (alas ? ' — ' + alas : ''));
      pesan = 'Terima kasih. ' + J.kata + ' ' + escH_(sdr_(r['Full Name'])) + ' telah <b>' + (setuju ? 'DISETUJUI' : 'DITOLAK') + '</b>.';
    }
    SpreadsheetApp.flush();
  } finally { lock.releaseLock(); }
  if (jenis === 'SPD' && tipe === 'k') {
    try { antri_({ t: 'persetujuanDinas', nomor: nomor }); } catch (e) { logAksi_(jenis, nomor, 'EMAIL PERSETUJUAN GAGAL', '', String(e)); }
  }
  try { antri_({ t: 'selesai', jenis: jenis, nomor: nomor }); } catch (e) { logAksi_(jenis, nomor, 'PDF GAGAL', '', String(e)); }
  return pesan;
}

/** Nama atasan langsung (tanpa jabatan) — dipakai untuk "Diperintahkan oleh". Kosong bila atasan belum diisi. */
function namaAtasan_(employeeId) {
  const K = karyawan_(), k = K[String(employeeId).toUpperCase()];
  if (k && isOwner_(k)) return k.nama;   // owner: atasannya diri sendiri
  const a = k && k.atasan ? (K[k.atasan] || (k.atasan === 'BOSS' ? owner_()[0] : null)) : null;
  return a ? a.nama : '';
}
/** Rantai atasan seorang karyawan (atasan langsung, atasannya, dst.) TANPA owner.
 *  Setiap orang di rantai ini menerima rekap & tombol pembatalan untuk karyawan tsb. */
function rantaiAtasan_(employeeId) {
  const out = [], lihat = {};
  let a = cariAtasan_(employeeId);
  while (a && !isOwner_(a) && !lihat[a.id]) { lihat[a.id] = 1; out.push(a); a = cariAtasan_(a.id); }
  return out;
}
/** Penyetuju Surat Perintah Dinas: atasan tertinggi di rantai karyawan, tepat 1 level di bawah owner.
 *  Bila atasan langsung karyawan adalah owner (atau karyawan sendiri 1 level di bawah owner) → owner yang menyetujui. */
function penyetujuDinas_(employeeId) {
  const rantai = rantaiAtasan_(employeeId);
  if (rantai.length) return rantai[rantai.length - 1];
  const a = cariAtasan_(employeeId);
  return a || null;
}
/** Owner (BOSS): karyawan aktif ber-email yang kolom Atasan-nya = BOSS (atau Peran / Employee ID = BOSS) */
function owner_() {
  return Object.values(karyawan_()).filter(k => k.aktif && k.email && isOwner_(k));
}
/** Peran OWNER (juga dikenali dari format lama: Atasan = BOSS) */
function isOwner_(k) {
  return !!k && (k.peran === 'OWNER' || k.peran === 'BOSS' || k.atasan === 'BOSS' || k.id === 'BOSS');
}
/** Atasan langsung dari kolom "Atasan" (Employee ID). */
function cariAtasan_(employeeId) {
  const K = karyawan_(), k = K[String(employeeId).toUpperCase()];
  if (k && isOwner_(k)) return k.email && k.aktif ? k : null;   // owner: atasannya diri sendiri
  if (!k || !k.atasan) return null;
  const a = K[k.atasan] || (k.atasan === 'BOSS' ? owner_()[0] : null);
  return a && a.email && a.aktif ? a : null;
}

/* ======================= 4. PDF ke admin ======================= */
/** Bila semua orang di surat sudah final & PDF belum dibuat → buat PDF & kirim ke admin. */
function cekSelesai_(jenis, nomor, paksa) {
  if (jenis === 'SPL') return rekapSPL_(nomor, paksa);
  const J = JENIS[jenis], S = bacaBaris_(jenis), rows = S.rows.filter(r => r[J.kunci] === nomor);
  if (!rows.length) return 'Nomor ' + nomor + ' tidak ditemukan.';
  const belum = rows.filter(r => ST_FINAL.indexOf(r['Status']) < 0);
  const Ind = jenis === 'SPD' ? bacaBaris_SPD_induk_() : null;
  const I = jenis === 'SPD' ? (Ind.rows.find(x => x['Nomor SPD'] === nomor) || {}) : rows[0];
  const sudahPdf = jenis === 'SPL' ? rows.some(r => r['Link PDF']) : !!I['Link PDF'];
  if (!paksa && (belum.length || sudahPdf)) {
    if (jenis === 'SPD' && I._row) ubahBaris_(Ind, I._row, { 'Status': belum.length ? 'PROSES ' + (rows.length - belum.length) + '/' + rows.length : 'SELESAI' });
    return '';
  }
  const pdf = jenis === 'SPL' ? buatPdfSPL_(rows) : buatPdfSPD_(I, rows);
  const file = folderPdf_(J.folder).createFile(pdf);
  if (jenis === 'SPL') rows.forEach(r => ubahBaris_(S, r._row, { 'Link PDF': file.getUrl() }));
  else ubahBaris_(Ind, I._row, { 'Link PDF': file.getUrl(), 'Status': belum.length ? 'PROSES ' + (rows.length - belum.length) + '/' + rows.length : 'SELESAI' });

  const ok = rows.filter(r => r['Status'] === ST.SETUJU);
  const ringkas = jenis === 'SPL'
    ? ok.length + ' dari ' + rows.length + ' karyawan <b>disetujui</b> lembur, total ' + String(ok.reduce((s, r) => s + (Number(r['Jam Efektif']) || 0), 0)).replace('.', ',') + ' jam.'
    : ok.length + ' dari ' + rows.length + ' personel <b>disetujui</b> dinas, uang saku ' + rpH_(ok.reduce((s, r) => s + (Number(r['Uang Saku']) || 0), 0)) + '.';
  const isi = '<p>Rekap <b>' + J.surat + ' ' + escH_(nomor) + '</b> — ' + escH_(ringkasTanggal_(jenis, rows[0], I)) + '.</p><p>' + ringkas +
    (belum.length ? ' <i>(' + belum.length + ' orang belum selesai diproses saat PDF ini dibuat)</i>' : '') + '</p>' +
    '<p>PDF terlampir dan tersimpan di Google Drive: <a href="' + file.getUrl() + '">' + escH_(file.getName()) + '</a></p>';
  const admin = String(param_().EMAIL_ADMIN || '').trim().toLowerCase();
  const pembuat = String(I['Email Pembuat']).toLowerCase();
  kirimEmail_(pembuat, 'PDF ' + jenis + ' ' + nomor + ' — ' + ringkasTanggal_(jenis, rows[0], I), isi, admin && admin !== pembuat ? admin : '', pdf);
  logAksi_(jenis, nomor, 'PDF DIKIRIM', pembuat, file.getUrl());
  return 'PDF dibuat & dikirim ke ' + pembuat;
}
function cekSelesaiSPL_(nomor, paksa) { return cekSelesai_('SPL', nomor, paksa); }

/** Pindahkan file PDF lama (dari link) ke sampah Drive */
function hapusPdfLama_(url) {
  const m = String(url || '').match(/\/d\/([\w-]{20,})/);
  if (!m) return false;
  try { DriveApp.getFileById(m[1]).setTrashed(true); return true; } catch (e) { return false; }
}

/**
 * Rekap PDF Surat Perintah Lembur per TANGGAL LEMBUR.
 *  - Admin : rekap semua surat yang DIA input pada tanggal itu (tanpa estimasi kompensasi).
 *  - Owner : rekap semua surat dari seluruh admin pada tanggal itu (dengan estimasi kompensasi).
 *  PDF dibuat setelah semua karyawan di rekap itu selesai diproses. Bila ada surat susulan di
 *  tanggal yang sama setelah PDF terbentuk, PDF lama dihapus & PDF baru dibuat + dikirim ulang.
 */
function rekapSPL_(nomor, paksa, sudah) {
  sudah = sudah || {};
  const S = bacaBaris_('SPL');
  const r0 = S.rows.find(r => r['Nomor SPL'] === nomor);
  if (!r0) return 'Nomor ' + nomor + ' tidak ditemukan.';
  const tgl = r0['Tanggal Lembur'] instanceof Date ? ymd_(r0['Tanggal Lembur']) : '';
  const hari = S.rows.filter(r => r['Tanggal Lembur'] instanceof Date && ymd_(r['Tanggal Lembur']) === tgl);
  const fin = rows => rows.every(r => ST_FINAL.indexOf(r['Status']) >= 0);
  const hasil = [];

  // 1. rekap admin pembuat surat ini
  const em = String(r0['Email Pembuat']).toLowerCase();
  const milik = hari.filter(r => String(r['Email Pembuat']).toLowerCase() === em);
  if (!sudah['adm:' + em] && (paksa || (fin(milik) && milik.some(r => !r['Link PDF'])))) {
    sudah['adm:' + em] = 1;
    const lama = [...new Set(milik.map(r => r['Link PDF']).filter(Boolean))];
    const pdf = buatPdfSPL_(milik, { estimasi: false, nama: 'Rekap Lembur ' + tglFile_(r0) + ' - ' + r0['Dibuat Oleh'] });
    const file = folderPdf_(JENIS.SPL.folder).createFile(pdf);
    lama.forEach(hapusPdfLama_);
    milik.forEach(r => ubahBaris_(S, r._row, { 'Link PDF': file.getUrl() }));
    kirimRekap_(em, milik, file, pdf, lama.length > 0, '', false);
    logAksi_('SPL', nomor, lama.length ? 'PDF REKAP ADMIN DIPERBARUI' : 'PDF REKAP ADMIN', em, file.getUrl());
    hasil.push('PDF rekap dikirim ke ' + em);
  }

  // 2. rekap owner (semua admin) — setelah seluruh surat di tanggal itu selesai
  const own = owner_();
  if (own.length && !sudah.owner && (paksa || (fin(hari) && hari.some(r => !r['Link PDF Owner'])))) {
    sudah.owner = 1;
    const lama = [...new Set(hari.map(r => r['Link PDF Owner']).filter(Boolean))];
    const pdf = buatPdfSPL_(hari, { estimasi: true, owner: true, nama: 'Rekap Lembur ' + tglFile_(r0) + ' - Semua Admin' });
    const file = folderPdf_(JENIS.SPL.folder).createFile(pdf);
    lama.forEach(hapusPdfLama_);
    hari.forEach(r => ubahBaris_(S, r._row, { 'Link PDF Owner': file.getUrl() }));
    kirimRekap_(own.map(o => o.email).join(','), hari, file, pdf, lama.length > 0, 'Semua Admin', true, 'O.' + own[0].id + '.' + tgl, Object.assign({}, own[0], { semua: true }));
    logAksi_('SPL', nomor, lama.length ? 'PDF REKAP OWNER DIPERBARUI' : 'PDF REKAP OWNER', own.map(o => o.email).join(','), file.getUrl());
    hasil.push('PDF rekap owner dikirim ke ' + own.map(o => o.nama).join(', '));
  }

  // 3. rekap atasan di atas atasan langsung (selain owner) — karyawan di bawah kewenangannya, dengan estimasi & tombol batal
  const rantai = r => rantaiAtasan_(r['Employee ID']).slice(1);   // tanpa atasan langsung (sudah menyetujui per orang)
  const daftarAtasan = {};
  S.rows.filter(r => r['Nomor SPL'] === nomor).forEach(r => rantai(r).forEach(a => daftarAtasan[a.id] = a));
  Object.values(daftarAtasan).forEach(a => {
    const grp = hari.filter(r => rantai(r).some(x => x.id === a.id));
    if (sudah['ats:' + a.id] || !grp.length || !(paksa || (fin(grp) && grp.some(r => !r['Link PDF Atasan'])))) return;
    sudah['ats:' + a.id] = 1;
    const lama = [...new Set(grp.map(r => r['Link PDF Atasan']).filter(Boolean))];
    const pdf = buatPdfSPL_(grp, { estimasi: true, owner: true, nama: 'Rekap Lembur ' + tglFile_(r0) + ' - Atasan ' + a.nama });
    const file = folderPdf_(JENIS.SPL.folder).createFile(pdf);
    lama.forEach(hapusPdfLama_);
    grp.forEach(r => ubahBaris_(S, r._row, { 'Link PDF Atasan': file.getUrl() }));
    kirimRekap_(a.email, grp, file, pdf, lama.length > 0, 'Karyawan di bawah ' + sdr_(a.nama), true, 'A.' + a.id + '.' + tgl, a);
    logAksi_('SPL', nomor, lama.length ? 'PDF REKAP ATASAN DIPERBARUI' : 'PDF REKAP ATASAN', a.email, file.getUrl());
    hasil.push('PDF rekap atasan dikirim ke ' + a.nama);
  });
  return hasil.join('\n');
}
/** "DISETUJUI 2 orang · DITOLAK 1 orang · ..." (hanya status yang ada) */
function ringkasStatus_(rows) {
  const n = st => rows.filter(r => r['Status'] === st).length, bagian = [['DISETUJUI', n(ST.SETUJU)], ['DITOLAK', n(ST.TOLAK)], ['DIBATALKAN', n(ST.BATAL)],
    ['BELUM SELESAI', rows.filter(r => ST_FINAL.indexOf(r['Status']) < 0).length]];
  return bagian.filter((b, i) => i === 0 || b[1]).map(b => b[0] + ' ' + b[1] + ' orang').join(' · ');
}
function tglFile_(r) { return Utilities.formatDate(r['Tanggal Lembur'], TZ, 'dd-MM-yyyy'); }

function kirimRekap_(ke, rows, file, pdf, revisi, label, owner, batalId, atasan) {
  const r0 = rows[0], ok = rows.filter(r => r['Status'] === ST.SETUJU);
  const tgl = r0['Hari'] + ', ' + tglIndo_(r0['Tanggal Lembur']);
  const nomor = [...new Set(rows.map(r => r['Nomor SPL']))];
  const belum = rows.filter(r => ST_FINAL.indexOf(r['Status']) < 0).length;
  const nAda = st => rows.filter(r => r['Status'] === st).length;
  const isi = atasan ? (   // rekap untuk atasan dari atasan langsung & owner (atasan.semua)
    '<p>Yth. <b>' + escH_(sdr_(atasan.nama)) + '</b>,</p>' +
    (revisi ? '<p style="background:#FFF4D6;padding:8px 10px;border-radius:6px"><b>Laporan diperbarui</b> — laporan sebelumnya sudah dihapus, gunakan PDF terbaru ini.</p>' : '') +
    '<p>Terlampir kami sampaikan rekap Surat Perintah Lembur tanggal <b>' + escH_(tgl) + '</b> ' + (atasan.semua ? 'untuk seluruh karyawan' : 'di bawah kelolaan ' + escH_(sdr_(atasan.nama))) + ' dengan rincian:</p>' +
    '<table style="font-size:14px;border-collapse:collapse">' +
    [['Total karyawan yang akan lembur', ok.length + ' orang'],
     ['Estimasi biaya lembur', rpH_(ok.reduce((t, r) => t + (Number(r['Estimasi Kompensasi']) || 0), 0))]].concat(
      nAda(ST.TOLAK) ? [['Ditolak atasan langsung', nAda(ST.TOLAK) + ' orang']] : [],
      nAda(ST.TIDAK) ? [['Tidak bersedia', nAda(ST.TIDAK) + ' orang']] : [],
      nAda(ST.BATAL) ? [['Dibatalkan', nAda(ST.BATAL) + ' orang']] : [],
      belum ? [['Belum selesai diproses', belum + ' orang']] : [])
      .map(x => '<tr><td style="padding:3px 12px 3px 0">' + x[0] + '</td><td style="padding:3px 0">: <b>' + x[1] + '</b></td></tr>').join('') + '</table>' +
    (batalId && rows.some(r => BISA_BATAL_(r)) ? '<p style="margin-top:18px;border-top:1px solid #eee;padding-top:12px">Bila lembur di atas tidak diperlukan, Anda dapat membatalkan seluruhnya:<br>' +
      tombol_(linkAksi_('b', batalId), '🚫 Batalkan Lembur', '#B42318') + '</p>' : '')
  ) : '<p>Rekap <b>Surat Perintah Lembur ' + escH_(tgl) + '</b>' + (label ? ' — ' + escH_(label.charAt(0).toLowerCase() + label.slice(1)) : '') + '.</p>' +
    (revisi ? '<p style="background:#FFF4D6;padding:8px 10px;border-radius:6px"><b>Laporan diperbarui</b> karena ada Surat Perintah Lembur susulan di tanggal yang sama. ' +
      'Laporan sebelumnya sudah dihapus — gunakan PDF terbaru ini.</p>' : '') +
    '<p>' + nomor.length + ' surat (' + nomor.map(escH_).join(', ') + ')<br>' +
    escH_(ringkasStatus_(rows.filter(r => r['Status'] !== ST.TIDAK))) + (rows.some(r => r['Status'] === ST.TIDAK) ? ' · TIDAK BERSEDIA ' + rows.filter(r => r['Status'] === ST.TIDAK).length + ' orang' : '') +
    (owner ? ', estimasi kompensasi ' + rpH_(ok.reduce((s, r) => s + (Number(r['Estimasi Kompensasi']) || 0), 0)) : '') + '.' +
    (belum ? ' <i>(' + belum + ' orang belum selesai diproses saat PDF ini dibuat)</i>' : '') + '</p>' +
    '<p>PDF terlampir dan tersimpan di Google Drive: <a href="' + file.getUrl() + '">' + escH_(file.getName()) + '</a></p>' +
    (batalId && rows.some(r => BISA_BATAL_(r)) ? '<p style="margin-top:18px;border-top:1px solid #eee;padding-top:12px">Bila lembur di atas tidak diperlukan, Anda dapat membatalkan seluruhnya:<br>' +
      tombol_(linkAksi_('b', batalId), '🚫 Batalkan Lembur', '#B42318') + '</p>' : '');
  const err = kirimEmail_(ke, (revisi ? '[DIPERBARUI] ' : '') + 'Rekap Surat Perintah Lembur — ' + tgl + (label ? ' (' + label + ')' : ''), isi, '', pdf);
  logAksi_('SPL', nomor.join(', '), err ? 'EMAIL REKAP GAGAL' : 'EMAIL REKAP TERKIRIM', ke, err || (label || 'Admin'));   // bukti email benar-benar dikirim
  return err;
}

function folderPdf_(nama) {
  const parent = DriveApp.getFileById(ss_().getId()).getParents();
  const induk = parent.hasNext() ? parent.next() : DriveApp.getRootFolder();
  const it = induk.getFoldersByName(nama);
  return it.hasNext() ? it.next() : induk.createFolder(nama);
}

const WARNA_ST_ = { 'DISETUJUI': '#1E7B4A', 'DITOLAK': '#B42318', 'TIDAK BERSEDIA': '#B42318', 'DIBATALKAN': '#888' };
const tdP_ = (s, x) => '<td style="padding:4px;border:1px solid #BBB;font-size:9px;vertical-align:top;' + (x || '') + '">' + s + '</td>';
const thP_ = s => '<th style="background:#3A3A3A;color:#fff;padding:5px 4px;font-size:9px;border:1px solid #3A3A3A">' + s + '</th>';
function selStatus_(r) {
  return '<b style="color:' + (WARNA_ST_[r['Status']] || '#9A6200') + '">' + escH_(r['Status']) + '</b>' +
    (r['Alasan Tidak Bersedia'] ? '<br>' + escH_(r['Alasan Tidak Bersedia']) : '') + (r['Catatan Atasan'] ? '<br>' + escH_(r['Catatan Atasan']) : '');
}
function selAtasan_(r) {
  return r['Status'] === ST.SETUJU || r['Status'] === ST.TOLAK ? escH_(r['Disetujui Oleh']) + '<br><span style="color:#777">' + waktuIndo_(r['Disetujui Pada']) + '</span>' : '-';
}
function kopPdf_(judul, nomor) {
  return '<div style="border-bottom:3px solid #BE353F;padding-bottom:6px"><img src="' + LOGO_DATA_URI + '" style="height:46px"></div>' +
    '<div style="text-align:center;margin:12px 0 2px;font-size:15px;font-weight:bold;letter-spacing:.5px">' + judul + '</div>' +
    (nomor ? '<div style="text-align:center;font-size:11px;margin-bottom:10px">Nomor : ' + escH_(nomor) + '</div>' : '<div style="margin-bottom:10px"></div>');
}
function pdfDari_(html, nama) {
  return HtmlService.createHtmlOutput(html).getBlob().getAs('application/pdf').setName(nama);
}

function buatPdfSPL_(rows, opt) {
  opt = opt || { estimasi: false };
  rows = rows.slice().sort((x, y) => String(x['Nomor SPL']).localeCompare(String(y['Nomor SPL'])) || x._row - y._row);
  const r0 = rows[0], P = param_();
  const nomor = [...new Set(rows.map(r => r['Nomor SPL']))];
  const admin = [...new Set(rows.map(r => r['Dibuat Oleh']))];
  const banyak = nomor.length > 1, kolAdmin = !!opt.owner && admin.length > 1;
  const noSurat = r => String(r['Nomor SPL']).split('/')[0].trim();
  const tidak = rows.filter(r => r['Status'] === ST.TIDAK);
  const bersedia = rows.filter(r => r['Status'] !== ST.TIDAK);
  const setuju = bersedia.filter(r => r['Status'] === ST.SETUJU);
  const totKomp = setuju.reduce((s, r) => s + (Number(r['Estimasi Kompensasi']) || 0), 0);
  const judulTabel = t => '<div style="font-weight:bold;font-size:11px;margin:12px 0 4px;color:#3A3A3A">' + t + '</div>';
  const tot = (lbl, span, isi, sisa) => '<tr><td colspan="' + span + '" style="padding:5px;border:1px solid #BBB;text-align:right;font-weight:bold;font-size:9px;background:#EFF0F2">' + lbl + '</td>' +
    isi.map(x => '<td style="padding:5px;border:1px solid #BBB;text-align:' + x[1] + ';font-weight:bold;font-size:9px;background:#EFF0F2;white-space:nowrap">' + x[0] + '</td>').join('') +
    (sisa ? '<td colspan="' + sisa + '" style="border:1px solid #BBB;background:#EFF0F2"></td>' : '') + '</tr>';
  const persetujuan = r => {
    if (r['Status'] === ST.BATAL)
      return '<b style="color:#888">DIBATALKAN</b>' + (r['Dibatalkan Oleh'] ? '<br>oleh ' + escH_(r['Dibatalkan Oleh']) + '<br><span style="color:#777">' + waktuIndo_(r['Dibatalkan Pada']) + '</span>' : '') +
        (r['Alasan Pembatalan'] ? '<br><i>' + escH_(r['Alasan Pembatalan']) + '</i>' : '');
    if (r['Status'] === ST.SETUJU || r['Status'] === ST.TOLAK)
      return '<b style="color:' + (WARNA_ST_[r['Status']] || '#222') + '">' + escH_(r['Status']) + '</b><br>' + escH_(r['Disetujui Oleh']) +
        '<br><span style="color:#777">' + waktuIndo_(r['Disetujui Pada']) + '</span>' + (r['Catatan Atasan'] ? '<br><i>' + escH_(r['Catatan Atasan']) + '</i>' : '');
    return '<b style="color:#9A6200">' + escH_(r['Status']) + '</b>';
  };
  // kolom awal yang sama untuk tabel A & B
  const headAwal = ['No'].concat(banyak ? ['No. Surat'] : [], kolAdmin ? ['Diinput oleh'] : [], ['Employee ID', 'Full Name']);
  const selAwal = (r, i) => tdP_(i + 1, 'text-align:center') + (banyak ? tdP_(escH_(noSurat(r)), 'text-align:center') : '') +
    (kolAdmin ? tdP_(escH_(r['Dibuat Oleh'])) : '') + tdP_(escH_(r['Employee ID'])) + tdP_('<b>' + escH_(r['Full Name']) + '</b>');

  // A. Karyawan yang bersedia lembur
  let html = judulTabel('A. KARYAWAN YANG BERSEDIA LEMBUR (' + bersedia.length + ' orang)');
  if (bersedia.length) {
    html += '<table style="border-collapse:collapse;width:100%"><tr>' +
      headAwal.concat(['Mulai', 'Selesai', 'Istirahat', 'Jam Efektif'], opt.estimasi ? ['Estimasi Kompensasi'] : [], ['Uraian Pekerjaan', 'Persetujuan Atasan']).map(thP_).join('') + '</tr>' +
      bersedia.map((r, i) => '<tr>' + selAwal(r, i) +
        tdP_(escH_(r['Jam Mulai']), 'text-align:center') + tdP_(escH_(r['Jam Selesai']), 'text-align:center') + tdP_(escH_(r['Istirahat']), 'text-align:center') +
        tdP_(String(r['Jam Efektif']).replace('.', ','), 'text-align:center') +
        (opt.estimasi ? tdP_(rpH_([ST.TOLAK, ST.BATAL].indexOf(r['Status']) >= 0 ? 0 : r['Estimasi Kompensasi']), 'text-align:right;white-space:nowrap') : '') +   // ditolak / dibatalkan = Rp 0
        tdP_(escH_(r['Uraian Pekerjaan'])) + tdP_(persetujuan(r)) + '</tr>').join('') +
      (opt.estimasi ? tot(ringkasStatus_(bersedia), headAwal.length + 4, [[rpH_(totKomp), 'right']], 2)
        : tot(ringkasStatus_(bersedia), headAwal.length + 6, [], 0)) + '</table>';
  } else {
    html += '<div style="font-size:9px;color:#777">Tidak ada karyawan yang bersedia lembur.</div>';
  }

  // B. Karyawan yang tidak bersedia lembur (hanya bila ada)
  if (tidak.length) {
    html += judulTabel('B. KARYAWAN YANG TIDAK BERSEDIA LEMBUR (' + tidak.length + ' orang)') +
      '<table style="border-collapse:collapse;width:100%"><tr>' +
      headAwal.concat(['Rencana Jam', 'Uraian Pekerjaan', 'Alasan Tidak Bersedia', 'Waktu Konfirmasi']).map(thP_).join('') + '</tr>' +
      tidak.map((r, i) => '<tr>' + selAwal(r, i) +
        tdP_(escH_(r['Jam Mulai']) + ' – ' + escH_(r['Jam Selesai']), 'text-align:center;white-space:nowrap') + tdP_(escH_(r['Uraian Pekerjaan'])) +
        tdP_(escH_(r['Alasan Tidak Bersedia'] || '-')) + tdP_(waktuIndo_(r['Konfirmasi Pada']) || '-') + '</tr>').join('') + '</table>';
  }

  const daftarNomor = banyak ? nomor.map(n => escH_(noSurat({ 'Nomor SPL': n })) + ' = ' + escH_(n)).join('<br>') : escH_(nomor[0]);
  const doc = '<html><head><meta charset="utf-8"><style>@page{size:A4 landscape;margin:12mm}body{font-family:Arial,sans-serif;font-size:10px;color:#222}</style></head><body>' +
    kopPdf_(banyak || opt.owner ? 'REKAP SURAT PERINTAH LEMBUR' : 'SURAT PERINTAH LEMBUR', banyak ? '' : nomor[0]) +
    '<table style="font-size:10px;margin-bottom:4px"><tr><td style="width:110px">Tanggal Lembur</td><td>: <b>' + escH_(r0['Hari']) + ', ' + tglIndo_(r0['Tanggal Lembur']) + '</b></td></tr>' +
    '<tr><td>Status Hari</td><td>: ' + escH_(r0['Status Hari']) + '</td></tr>' +
    (banyak && !opt.owner ? '<tr><td style="vertical-align:top">Nomor Surat</td><td>: ' + daftarNomor + '</td></tr>' : '') +   // owner/atasan: nomor ada di kolom tabel
    (kolAdmin ? '' : '<tr><td>Diinput oleh</td><td>: ' + admin.map(escH_).join(', ') + '</td></tr>') + '</table>' +
    html +
    '<p style="font-size:9px;color:#555;margin-top:10px">Konfirmasi karyawan dan persetujuan atasan dilakukan secara elektronik' +
    '. Dokumen dibuat otomatis pada ' + waktuIndo_(new Date()) + '.</p>' + (opt.owner ? ttdPdf_(null) : ttdPdf_(admin.join(', '), 'Admin')) + '</body></html>';
  return pdfDari_(doc, (opt.nama || 'SPL ' + String(r0['Nomor SPL']).replace(/\s*\/\s*/g, '-') + ' (' + tglFile_(r0) + ')') + '.pdf');
}
function ttdPdf_(nama, jabatan) {   // nama null → tanpa kolom "Dibuat oleh" (rekap owner / atasan)
  const d = nama != null;
  return '<table style="width:100%;margin-top:18px;font-size:10px;text-align:center"><tr>' + (d ? '<td>Dibuat oleh,</td>' : '') + '<td>Diketahui oleh,</td><td>Disetujui oleh,</td></tr>' +
    '<tr><td style="height:55px"></td><td></td>' + (d ? '<td></td>' : '') + '</tr>' +
    '<tr>' + (d ? '<td>( ' + escH_(nama) + ' )</td>' : '') + '<td>(.........................)</td><td>(.........................)</td></tr>' +
    '<tr>' + (d ? '<td>' + escH_(jabatan) + '</td>' : '') + '<td>Manager Departemen</td><td>HRD</td></tr></table>';
}

/** PDF SPD mengikuti formulir Excel: A. Referensi, B. Detail, C. Personel, D. Biaya, E. Ketentuan, F. Pengesahan. */
function buatPdfSPD_(I, rows) {
  const P = param_();
  const biaya = bacaTabel_(SH.BIAYA, 'Nomor SPD').filter(b => b['Nomor SPD'] === I['Nomor SPD']);
  const setuju = rows.filter(r => r['Status'] === ST.SETUJU);
  const totSaku = setuju.reduce((s, r) => s + (Number(r['Uang Saku']) || 0), 0);
  const totLain = biaya.reduce((s, b) => s + (Number(b['Jumlah (Rp)']) || 0), 0);
  const sek = t => '<div style="background:#3A3A3A;color:#fff;font-weight:bold;font-size:10px;padding:4px 6px;margin:10px 0 4px">' + t + '</div>';
  const kv = arr => '<table style="width:100%;font-size:10px;border-collapse:collapse">' + arr.map(r => '<tr>' + r.map((x, i) =>
    '<td style="padding:3px 4px;' + (i % 2 ? '' : 'width:110px;color:#555') + '">' + (i % 2 ? ': <b>' + escH_(x) + '</b>' : escH_(x)) + '</td>').join('') + '</tr>').join('') + '</table>';
  const tglB = I['Tgl Berangkat'], tglK = I['Tgl Kembali'];
  // kop & nomor surat diulang di setiap halaman (thead tabel pembungkus)
  const html = '<html><head><meta charset="utf-8"><style>@page{size:A4;margin:12mm}body{font-family:Arial,sans-serif;font-size:10px;color:#222}table.hal{width:100%;border-collapse:collapse}table.hal>thead{display:table-header-group}table.hal>thead>tr>td,table.hal>tbody>tr>td{padding:0}tr{page-break-inside:avoid}</style></head><body>' +
    '<table class="hal"><thead><tr><td>' + kopPdf_('SURAT PERINTAH DINAS', I['Nomor SPD']) + '</td></tr></thead><tbody><tr><td>' +
    sek('A. REFERENSI SURAT') +
    kv([['Nomor Surat', I['No Surat Referensi'] || '-', 'Tanggal Surat', tglIndo_(I['Tanggal Surat'])], ['Nama Customer', I['Nama Customer'] || '-', 'Dasar Penugasan', I['Dasar Penugasan'] || '-']]) +
    sek('B. DETAIL PERJALANAN DINAS') +
    kv([['Maksud / Tujuan', I['Maksud / Tujuan'], 'Tempat Tujuan', I['Tempat Tujuan']],
        ['Tgl. Berangkat', hariIndo_(tglB) + ', ' + tglIndo_(tglB), 'Tgl. Kembali', hariIndo_(tglK) + ', ' + tglIndo_(tglK)],
        ['Lama Dinas', I['Lama (hari)'] + ' hari', 'Transportasi', I['Transportasi']],
        ['Sumber Biaya', I['Sumber Biaya'], 'Uang Saku / Hari', rpH_(P.UANG_SAKU_HARI)],
        ['Uang Kas Perjalanan', rpH_(Number(I['Uang Kas Perjalanan']) || 0)]]) +
    sek('C. PERSONEL YANG DITUGASKAN') +
    '<table style="border-collapse:collapse;width:100%"><tr>' + ['No', 'Employee ID', 'Full Name', 'Job Position', 'Lama (hari)', 'Uang Saku', 'Keterangan', 'Status', 'Persetujuan Atasan'].map(thP_).join('') + '</tr>' +
    rows.map((r, i) => '<tr>' + tdP_(i + 1, 'text-align:center') + tdP_(escH_(r['Employee ID'])) + tdP_('<b>' + escH_(r['Full Name']) + '</b>') + tdP_(escH_(r['Job Position'])) +
      tdP_(r['Lama (hari)'], 'text-align:center') + tdP_(rpH_(r['Uang Saku']), 'text-align:right') + tdP_(escH_(r['Keterangan'])) + tdP_(selStatus_(r)) + tdP_(selAtasan_(r)) + '</tr>').join('') +
    '<tr><td colspan="5" style="padding:5px;border:1px solid #BBB;text-align:right;font-weight:bold;font-size:9px;background:#EFF0F2">TOTAL UANG SAKU PERSONEL DISETUJUI (' + setuju.length + ' dari ' + rows.length + ' orang)</td>' +
    '<td style="padding:5px;border:1px solid #BBB;text-align:right;font-weight:bold;font-size:9px;background:#EFF0F2">' + rpH_(totSaku) + '</td><td colspan="3" style="border:1px solid #BBB;background:#EFF0F2"></td></tr></table>' +
    (biaya.length ? (   // biaya lainnya kini diisi di Pertanggungjawaban Dinas; tampil hanya untuk surat lama yang punya rencana biaya
    sek('D. BIAYA PERJALANAN DINAS LAINNYA') +
    (biaya.length ? '<table style="border-collapse:collapse;width:100%"><tr>' + ['No', 'Jenis Biaya', 'Uraian', 'Jumlah (Rp)', 'Bukti'].map(thP_).join('') + '</tr>' +
      biaya.map((b, i) => '<tr>' + tdP_(i + 1, 'text-align:center') + tdP_(escH_(b['Jenis Biaya'])) + tdP_(escH_(b['Uraian'])) + tdP_(rpH_(b['Jumlah (Rp)']), 'text-align:right') + tdP_(escH_(b['Bukti']), 'text-align:center') + '</tr>').join('') +
      '<tr><td colspan="3" style="padding:5px;border:1px solid #BBB;text-align:right;font-weight:bold;font-size:9px;background:#EFF0F2">TOTAL BIAYA LAINNYA</td><td style="padding:5px;border:1px solid #BBB;text-align:right;font-weight:bold;font-size:9px;background:#EFF0F2">' + rpH_(totLain) + '</td><td style="border:1px solid #BBB;background:#EFF0F2"></td></tr></table>'
      : '<div style="font-size:9px;color:#777">Tidak ada biaya lainnya.</div>') +
    '<table style="width:100%;margin-top:6px;border-collapse:collapse"><tr><td style="padding:6px;border:2px solid #BE353F;font-weight:bold;font-size:11px">TOTAL BIAYA PERJALANAN DINAS (Uang Saku disetujui + Biaya Lainnya)</td>' +
    '<td style="padding:6px;border:2px solid #BE353F;font-weight:bold;font-size:11px;text-align:right;width:130px;color:#BE353F">' + rpH_(totSaku + totLain) + '</td></tr></table>'
    ) : '') +
    sek((biaya.length ? 'E' : 'D') + '. KETENTUAN') +
    '<ol style="font-size:9px;margin:4px 0 0 14px;padding:0;line-height:1.5">' +
    '<li>Karyawan yang tercantum di atas ditugaskan melaksanakan perjalanan dinas sesuai maksud, tempat, dan tanggal yang tercantum dalam surat ini.</li>' +
    '<li>Selama tanggal dinas luar tidak dapat diajukan lembur; pengajuan lembur pada tanggal tersebut akan ditolak.</li>' +
    '<li>Uang saku diberikan per orang per hari dinas (' + rpH_(P.UANG_SAKU_HARI) + '/hari) dan dibayarkan melalui aplikasi KantorKu.</li>' +
    '<li>Uang kas dalam perjalanan dipertanggungjawabkan setelah kembali: uang kas dikurangi biaya lainnya (sesuai bukti) = uang kas yang harus dikembalikan.</li>' +
    '<li>Mintakan tanda tangan & stempel pada kolom Pengesahan di tempat tujuan sebagai bukti kehadiran.</li>' +
    '<li>Laporan perjalanan dinas beserta bukti pengeluaran diserahkan ke Admin paling lambat 3 hari kerja setelah kembali.</li></ol>' +
    sek((biaya.length ? 'F' : 'E') + '. PENGESAHAN DI TEMPAT TUJUAN') +
    '<table style="width:100%;border-collapse:collapse;font-size:9px"><tr>' + ['TIBA DI TEMPAT TUJUAN', 'BERANGKAT DARI TEMPAT TUJUAN'].map(t =>
      '<td style="width:50%;border:1px solid #BBB;padding:6px;vertical-align:top"><b>' + t + '</b><br><br>Tempat : ............................<br><br>Tanggal : ............................<br><br>Jam : ............................' +
      '<div style="height:50px"></div><div style="text-align:center;color:#555">( Nama, Tanda Tangan & Stempel Instansi )</div></td>').join('') + '</tr></table>' +
    '<p style="text-align:right;font-size:10px;margin-top:12px">Dikeluarkan di : ' + (P.KOTA_SURAT ? escH_(P.KOTA_SURAT) + ', ' : '') + tglIndo_(I['Tanggal Surat']) + '</p>' +
    ttdPdf_(I['Dibuat Oleh'], 'Admin').replace('Manager Departemen', 'Manager Departemen') +
    '<p style="font-size:8px;color:#777;margin-top:8px">Konfirmasi personel dan persetujuan atasan dilakukan secara elektronik. Dokumen dibuat otomatis pada ' + waktuIndo_(new Date()) + '.</p>' +
    '</td></tr></tbody></table></body></html>';
  return pdfDari_(html, 'SPD ' + String(I['Nomor SPD']).replace(/\s*\/\s*/g, '-') + ' (' + Utilities.formatDate(tglB, TZ, 'dd-MM-yyyy') + ').pdf');
}

/* ======================= Menu admin ======================= */
function jenisDariNomor_(n) { return /\/\s*SPD\s*\//i.test(n) ? 'SPD' : 'SPL'; }
function menuPdfSPL() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('Buat & kirim PDF', 'Masukkan Nomor SPL / SPD (mis. 001 / SPL / CBA / X / 2026).\nPDF dibuat walau masih ada yang belum selesai diproses.', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const n = r.getResponseText().trim();
  ui.alert(cekSelesai_(jenisDariNomor_(n), n, true) || 'Selesai.');
}
/** Email permintaan persetujuan ke atasan (dipakai saat karyawan Bersedia & saat kirim ulang). */
function kirimPersetujuan_(jenis, S, r, I, atasan, ke) {
  const J = JENIS[jenis], id = String(r['ID Baris']), tgl = ringkasTanggal_(jenis, r, I);
        const isi = (atasan ? '<p>Yth. <b>' + escH_(sdr_(atasan.nama)) + '</b>,</p>' : '<p><i>Atasan langsung karyawan ini belum diisi di Data Karyawan, sehingga permintaan persetujuan dikirim ke admin.</i></p>') +
          '<p><b>' + escH_(sdr_(r['Full Name'])) + '</b> telah menyatakan <b style="color:#1E7B4A">BERSEDIA</b> ' + J.kata.toLowerCase() + '. Mohon persetujuan Anda:</p>' + tabelDetail_(jenis, r, I) +
          '<p style="margin-top:18px">' + tombol_(linkAksi_(J.t, id, 'ya'), '✅ Setujui', '#1E7B4A') + tombol_(linkAksi_(J.t, id, 'tidak'), '❌ Tolak', '#B42318') + '</p>';
        const err = kirimEmail_(ke, 'Persetujuan ' + J.kata + ' — ' + sdr_(r['Full Name']) + ' (' + tgl + ')', isi);
        ubahBaris_(S, r._row, { 'Status Email': err ? 'GAGAL kirim ke atasan: ' + err : 'Persetujuan terkirim ke ' + ke });
  return err;
}

/**
 * Email persetujuan Surat Perintah Dinas DIGABUNG per penyetuju.
 * Personel dengan penyetuju yang sama ditunggu sampai semuanya menjawab konfirmasi
 * (Bersedia / Tidak Bersedia), lalu dikirim SATU email berisi semua personel yang Bersedia —
 * masing-masing dengan tombol Setujui / Tolak sendiri. Penyetuju berbeda → email terpisah.
 * paksa = true (menu kirim ulang): kirim semua yang MENUNGGU ATASAN tanpa menunggu.
 * Mengembalikan jumlah email yang dikirim.
 */
function kirimPersetujuanDinas_(nomor, paksa) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('Server sibuk, coba lagi.');
  let n = 0;
  try {
    _K = null;
    const J = JENIS.SPD, S = bacaBaris_('SPD');
    const semua = S.rows.filter(x => x[J.kunci] === nomor);
    if (!semua.length) return 0;
    const grup = {};
    semua.forEach(x => {
      const a = penyetujuDinas_(x['Employee ID']), kunci = a ? a.id : '-';
      (grup[kunci] = grup[kunci] || { a: a, rows: [] }).rows.push(x);
    });
    Object.values(grup).forEach(g => {
      if (!paksa && g.rows.some(x => x['Status'] === ST.KONFIRMASI)) return;   // tunggu personel lain menjawab
      const kirim = g.rows.filter(x => x['Status'] === ST.ATASAN && (paksa || !/^Persetujuan terkirim/.test(String(x['Status Email'] || ''))));
      if (!kirim.length) return;
      const I = induk_('SPD', nomor, kirim[0]), a = g.a, ke = a ? a.email : I['Email Pembuat'];
      const th = 'padding:6px 8px;background:#F4F5F7;border:1px solid #E2E4E8;text-align:left', td = 'padding:6px 8px;border:1px solid #E2E4E8;vertical-align:top';
      const isi = (a ? '<p>Yth. <b>' + escH_(sdr_(a.nama)) + '</b>,</p>' : '<p><i>Atasan karyawan belum diisi di Data Karyawan, sehingga permintaan persetujuan dikirim ke admin.</i></p>') +
        '<p>Personel berikut telah menyatakan <b style="color:#1E7B4A">BERSEDIA</b> melaksanakan perjalanan dinas. Mohon persetujuan Anda untuk masing-masing personel:</p>' +
        tabelDetail_('SPD', kirim[0], I, false, true) +
        '<table style="border-collapse:collapse;width:100%;font-size:13px;margin-top:12px"><tr><th style="' + th + '">Personel</th><th style="' + th + '">Uang Saku</th><th style="' + th + '">Persetujuan</th></tr>' +
        kirim.map(x => '<tr><td style="' + td + '"><b>' + escH_(sdr_(x['Full Name'])) + '</b><div style="font-size:11px;color:#6B6F76">' + escH_(x['Employee ID']) +
          (x['Job Position'] ? ' — ' + escH_(x['Job Position']) : '') + ' · ' + escH_(x['Lama (hari)']) + ' hari' + (x['Keterangan'] ? ' — ' + escH_(x['Keterangan']) : '') + '</div></td>' +
          '<td style="' + td + ';white-space:nowrap">' + rpH_(x['Uang Saku']) + '</td><td style="' + td + ';white-space:nowrap">' +
          tombol_(linkAksi_(J.t, String(x['ID Baris']), 'ya'), '✅ Setujui', '#1E7B4A') + tombol_(linkAksi_(J.t, String(x['ID Baris']), 'tidak'), '❌ Tolak', '#B42318') + '</td></tr>').join('') +
        '</table><p style="font-size:12px;color:#6B6F76">Uang saku dibayarkan melalui aplikasi KantorKu.</p>';
      const err = kirimEmail_(ke, 'Persetujuan Dinas Luar — ' + (kirim.length > 1 ? kirim.length + ' personel' : sdr_(kirim[0]['Full Name'])) + ' (' + ringkasTanggal_('SPD', kirim[0], I) + ')', isi);
      kirim.forEach(x => ubahBaris_(S, x._row, { 'Status Email': err ? 'GAGAL kirim ke atasan: ' + err : 'Persetujuan terkirim ke ' + ke + ' ' + Utilities.formatDate(new Date(), TZ, 'dd/MM HH:mm') + (kirim.length > 1 ? ' (gabungan ' + kirim.length + ' personel)' : '') }));
      logAksi_('SPD', nomor, err ? 'EMAIL PERSETUJUAN GAGAL' : 'EMAIL PERSETUJUAN TERKIRIM', ke, kirim.map(x => x['Full Name']).join(', ') + (err ? ' — ' + err : ''));
      if (!err) n++;
    });
    SpreadsheetApp.flush();
  } finally { lock.releaseLock(); }
  return n;
}

function menuKirimUlangKonfirmasi() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('Kirim ulang email', 'Masukkan Nomor SPL / SPD. Email dikirim ulang ke karyawan yang MENUNGGU KONFIRMASI dan ke atasan untuk yang MENUNGGU ATASAN.', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const n = r.getResponseText().trim();
  const jenis = jenisDariNomor_(n);
  kirimKonfirmasi_(jenis, n);
  // kirim ulang juga permintaan persetujuan yang masih MENUNGGU ATASAN
  const J = JENIS[jenis], S = bacaBaris_(jenis);
  let nAtasan = 0;
  if (jenis === 'SPD') nAtasan = kirimPersetujuanDinas_(n, true);
  else S.rows.filter(x => x[J.kunci] === n && x['Status'] === ST.ATASAN).forEach(x => {
    const I = induk_(jenis, n, x), atasan = jenis === 'SPD' ? penyetujuDinas_(x['Employee ID']) : cariAtasan_(x['Employee ID']);
    kirimPersetujuan_(jenis, S, x, I, atasan, atasan ? atasan.email : I['Email Pembuat']); nAtasan++;
  });
  ui.alert('Selesai. Email konfirmasi (yang masih MENUNGGU KONFIRMASI) dan ' + nAtasan + ' email persetujuan (MENUNGGU ATASAN) dikirim ulang.\nLihat kolom "Status Email" di sheet ' + (jenis === 'SPD' ? 'Personel SPD' : 'Data SPL') + '.');
}
/**
 * =====================================================================
 *  SISTEM SPL & SPD — File : Batal.gs
 *  Pembatalan lembur dari tombol di email rekap atasan / owner.
 *  id link: 'A.<idAtasan>.<yyyy-mm-dd>' (bawahan atasan itu) atau
 *           'O.<idOwner>.<yyyy-mm-dd>'  (semua lembur di tanggal itu).
 *  Setelah dibatalkan:
 *   1. email ke tiap karyawan terdampak: dibatalkan oleh siapa + alasan
 *   2. email ke admin pembuat surat (boleh input ulang di tanggal yang sama)
 *      + tembusan ke atasan langsung karyawan (bila dibatalkan oleh owner)
 *   3. rekap PDF admin / atasan / owner dibuat ulang (riwayat & alasan tetap tercatat)
 * =====================================================================
 */
const BISA_BATAL_ = r => [ST.TIDAK, ST.TOLAK, ST.BATAL].indexOf(r['Status']) < 0;

function cakupanBatal_(id) {
  const m = String(id || '').match(/^([AO])\.([^.]+)\.(\d{4}-\d{2}-\d{2})$/);
  if (!m) throw new Error('Link tidak valid.');
  const K = karyawan_(), p = K[m[2].toUpperCase()];
  if (!p || !p.aktif) throw new Error('Data atasan / owner tidak ditemukan.');
  if (m[1] === 'O' && !owner_().some(o => o.id === p.id)) throw new Error('Link pembatalan owner tidak berlaku lagi.');
  const S = bacaBaris_('SPL');
  const rows = S.rows.filter(r => r['Tanggal Lembur'] instanceof Date && ymd_(r['Tanggal Lembur']) === m[3] &&
    (m[1] === 'O' || rantaiAtasan_(r['Employee ID']).slice(1).some(a => a.id === p.id)));
  return { S: S, oleh: p, owner: m[1] === 'O', tgl: m[3], rows: rows, aktif: rows.filter(BISA_BATAL_) };
}

function halamanBatal_(p) {
  let isi;
  try {
    if (!p.id || p.s !== tandaTangan_('b', p.id)) throw new Error('Link tidak valid atau sudah diubah.');
    const C = cakupanBatal_(p.id), r0 = C.rows[0];
    if (!r0) throw new Error('Data lembur tidak ditemukan.');
    isi = '<h2>Pembatalan Lembur</h2><p>' + escH_(r0['Hari']) + ', <b>' + tglIndo_(r0['Tanggal Lembur']) + '</b> — ' +
      (C.owner ? 'seluruh karyawan' : 'karyawan di bawah <b>' + escH_(sdr_(C.oleh.nama)) + '</b>') + '</p>';
    if (!C.aktif.length) {
      isi += '<div class="msg info">Tidak ada lembur yang bisa dibatalkan (semua sudah dibatalkan / ditolak / tidak bersedia).</div>';
    } else {
      const th = 'padding:6px;background:#F4F5F7', td = 'padding:6px;border-top:1px solid #eee';
      isi += '<table style="border-collapse:collapse;width:100%;font-size:13px">' +
        '<tr><th style="text-align:left;' + th + '">Karyawan</th><th style="' + th + '">Jam</th><th style="' + th + '">Status</th></tr>' +
        C.aktif.map(r => '<tr><td style="' + td + '"><b>' + escH_(sdr_(r['Full Name'])) + '</b><div class="sub">' + escH_(r['Uraian Pekerjaan']) +
          ' · ' + escH_(r['Nomor SPL']) + '</div></td><td style="' + td + ';text-align:center;white-space:nowrap">' + escH_(r['Jam Mulai']) + '–' + escH_(r['Jam Selesai']) +
          '</td><td style="' + td + ';font-size:11px;text-align:center">' + escH_(r['Status']) + '</td></tr>').join('') + '</table>' +
        '<p style="margin-top:14px"><b>' + C.aktif.length + ' lembur</b> di atas akan <b style="color:#B42318">DIBATALKAN</b> atas nama <b>' + escH_(sdr_(C.oleh.nama)) + '</b>.</p>' +
        '<div id="al"><label>Alasan pembatalan *</label><textarea id="alasan" rows="3" placeholder="Wajib diisi"></textarea></div>' +
        '<button id="b" onclick="kirim()">Batalkan Lembur</button><div id="o"></div>' +
        '<script>function kirim(){var b=document.getElementById("b"),o=document.getElementById("o"),a=document.getElementById("alasan").value.trim();' +
        'if(!a){o.innerHTML="<div class=\\"msg err\\">Alasan pembatalan wajib diisi.</div>";return}' +
        'b.disabled=true;b.textContent="Memproses…";' +
        'google.script.run.withSuccessHandler(function(m){document.getElementById("al").remove();b.remove();o.innerHTML="<div class=\\"msg ok\\">"+m+"</div>"})' +
        '.withFailureHandler(function(e){b.disabled=false;b.textContent="Batalkan Lembur";o.innerHTML="<div class=\\"msg err\\">"+e.message+"</div>"})' +
        '.prosesBatal(' + JSON.stringify(String(p.id)) + ',' + JSON.stringify(p.s) + ',a)}</script>';
    }
  } catch (e) {
    isi = '<h2>Tidak dapat dibuka</h2><div class="msg err">' + escH_(e.message) + '</div>';
  }
  return halamanHtml_(isi);
}

/** Dipanggil dari halaman pembatalan. */
function prosesBatal(id, sig, alasan) {
  if (!id || sig !== tandaTangan_('b', id)) throw new Error('Link tidak valid.');
  alasan = String(alasan || '').trim().slice(0, 500);
  if (!alasan) throw new Error('Alasan pembatalan wajib diisi.');
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('Server sibuk, coba lagi.');
  let C, kena;
  try {
    _K = null; _P = null;
    pastikanKolom_(SH.SPL, KOLOM_SPL_TAMBAHAN);
    C = cakupanBatal_(id); kena = C.aktif;
    if (!kena.length) return 'Tidak ada lembur yang perlu dibatalkan — semua sudah dibatalkan sebelumnya.';
    const now = new Date();
    kena.forEach(r => {
      ubahBaris_(C.S, r._row, { 'Status': ST.BATAL, 'Dibatalkan Oleh': C.oleh.nama, 'Dibatalkan Pada': now, 'Alasan Pembatalan': alasan });
      r['Status'] = ST.BATAL; r['Dibatalkan Oleh'] = C.oleh.nama; r['Dibatalkan Pada'] = now; r['Alasan Pembatalan'] = alasan;
    });
    SpreadsheetApp.flush();
    logAksi_('SPL', [...new Set(kena.map(r => r['Nomor SPL']))].join(', '), 'DIBATALKAN', C.oleh.email, kena.length + ' karyawan, tgl ' + C.tgl + ' — ' + alasan);
  } finally { lock.releaseLock(); }
  const tglTeks = kena[0]['Hari'] + ', ' + tglIndo_(kena[0]['Tanggal Lembur']);
  try { antri_({ t: 'batal', id: id, alasan: alasan, ids: kena.map(r => String(r['ID Baris'])) }); }
  catch (e) { logAksi_('SPL', id, 'EMAIL BATAL GAGAL', '', String(e)); }
  return 'Lembur ' + kena.length + ' karyawan pada ' + escH_(tglTeks) + ' telah <b>DIBATALKAN</b>. Karyawan dan admin akan diberi tahu melalui email.';
}

/** Lanjutan pembatalan (email + rekap baru) — dijalankan dari antrian. */
function batalLanjutan_(t) {
  const alasan = t.alasan, m = String(t.id).match(/^([AO])\.([^.]+)\./);
  const C = { oleh: karyawan_()[m[2].toUpperCase()] || { id: m[2], nama: m[2] } };
  const set = {}; t.ids.forEach(i => set[i] = 1);
  const kena = bacaBaris_('SPL').rows.filter(r => set[String(r['ID Baris'])]);
  if (!kena.length) return;
  const K = karyawan_(), tgl = kena[0]['Hari'] + ', ' + tglIndo_(kena[0]['Tanggal Lembur']);
  // 1. email ke karyawan terdampak
  kena.forEach(r => {
    const k = K[String(r['Employee ID']).toUpperCase()];
    if (!k || !k.email) return;
    kirimEmail_(k.email, 'Lembur Dibatalkan — ' + tgl,
      '<p>Yth. <b>' + escH_(sdr_(r['Full Name'])) + '</b>,</p><p>Surat Perintah Lembur Anda pada <b>' + escH_(tgl) + '</b> ' +
      '<b style="color:#B42318">DIBATALKAN</b> oleh <b>' + escH_(sdr_(C.oleh.nama)) + '</b>.</p>' + tabelDetail_('SPL', r) +
      '<p><b>Alasan pembatalan:</b> ' + escH_(alasan) + '</p><p>Mohon tidak melaksanakan lembur tersebut. Bila sudah mengajukan lembur di aplikasi KantorKu, mohon dibatalkan.</p>');
  });
  // 2. email ke admin pembuat surat
  const perAdmin = {};
  kena.forEach(r => { const em = String(r['Email Pembuat']).toLowerCase(); (perAdmin[em] = perAdmin[em] || []).push(r); });
  Object.keys(perAdmin).forEach(em => kirimEmail_(em, 'Lembur Dibatalkan oleh ' + sdr_(C.oleh.nama) + ' — ' + tgl,
    '<p>Lembur berikut pada <b>' + escH_(tgl) + '</b> <b style="color:#B42318">DIBATALKAN</b> oleh <b>' + escH_(sdr_(C.oleh.nama)) + '</b>:</p><ul>' +
    perAdmin[em].map(r => '<li><b>' + escH_(sdr_(r['Full Name'])) + '</b> ' + escH_(r['Jam Mulai']) + '–' + escH_(r['Jam Selesai']) + ' (' + escH_(r['Nomor SPL']) + ')</li>').join('') +
    '</ul><p><b>Alasan:</b> ' + escH_(alasan) + '</p><p>Bila diperlukan, karyawan tersebut dapat diinput ulang melalui form pada tanggal yang sama. Riwayat pembatalan tetap tersimpan.</p>'));
  // 2b. tembusan ke atasan langsung karyawan (bila yang membatalkan bukan atasan itu sendiri)
  const perAtasan = {};
  kena.forEach(r => rantaiAtasan_(r['Employee ID']).forEach(a => {   // semua atasan di rantai (tanpa owner)
    if (a && a.id !== C.oleh.id && !isOwner_(a)) (perAtasan[a.id] = perAtasan[a.id] || { a: a, rows: [] }).rows.push(r); }));
  Object.values(perAtasan).forEach(x => kirimEmail_(x.a.email, 'Tembusan: Lembur Dibatalkan oleh ' + sdr_(C.oleh.nama) + ' — ' + tgl,
    '<p>Yth. <b>' + escH_(sdr_(x.a.nama)) + '</b>,</p><p>Lembur karyawan di bawah Anda pada <b>' + escH_(tgl) + '</b> <b style="color:#B42318">DIBATALKAN</b> oleh <b>' + escH_(sdr_(C.oleh.nama)) + '</b>:</p><ul>' +
    x.rows.map(r => '<li><b>' + escH_(sdr_(r['Full Name'])) + '</b> ' + escH_(r['Jam Mulai']) + '–' + escH_(r['Jam Selesai']) + ' — ' + escH_(r['Uraian Pekerjaan']) + '</li>').join('') +
    '</ul><p><b>Alasan:</b> ' + escH_(alasan) + '</p>'));
  // 3. rekap PDF dibuat ulang
  const sudah = {};
  [...new Set(kena.map(r => r['Nomor SPL']))].forEach(n => { try { rekapSPL_(n, true, sudah); } catch (e) { logAksi_('SPL', n, 'PDF GAGAL', '', String(e)); } });
}
/**
 * =====================================================================
 *  SISTEM SPL & SPD — File : Antrian.gs
 *  Pekerjaan berat (kirim email konfirmasi, buat PDF rekap, email pembatalan)
 *  dikerjakan di LATAR BELAKANG agar tombol di form / email langsung selesai.
 *
 *  Butuh 1 pemicu (trigger) waktu: fungsi prosesAntrian, setiap 1 menit
 *  (dibuat sekali dari editor Apps Script → Pemicu). Selama pemicu belum
 *  aktif, pekerjaan tetap dijalankan langsung seperti sebelumnya.
 * =====================================================================
 */
const ANTRIAN_KEY = 'Q', ANTRIAN_DETAK = 'Q_DETAK';

/** Pemicu aktif bila prosesAntrian berjalan dalam 3 menit terakhir */
function antrianAktif_() {
  return Date.now() - Number(PropertiesService.getScriptProperties().getProperty(ANTRIAN_DETAK) || 0) < 3 * 60 * 1000;
}

/** Masukkan tugas ke antrian (atau jalankan langsung bila pemicu belum aktif). */
function antri_(tugas) {
  if (!antrianAktif_()) return jalankanTugas_(tugas);
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return jalankanTugas_(tugas);
  try {
    const P = PropertiesService.getScriptProperties();
    const q = JSON.parse(P.getProperty(ANTRIAN_KEY) || '[]');
    const kunci = JSON.stringify(tugas);
    if (!q.some(t => JSON.stringify(t) === kunci)) q.push(tugas);
    P.setProperty(ANTRIAN_KEY, JSON.stringify(q));
  } finally { lock.releaseLock(); }
}

/** Dipanggil pemicu waktu setiap 1 menit. */
function prosesAntrian() {
  const P = PropertiesService.getScriptProperties();
  P.setProperty(ANTRIAN_DETAK, String(Date.now()));
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return;
  let q;
  try {
    q = JSON.parse(P.getProperty(ANTRIAN_KEY) || '[]');
    if (!q.length) return;
    P.setProperty(ANTRIAN_KEY, '[]');
  } finally { lock.releaseLock(); }
  q.forEach(t => {
    try { _K = null; _P = null; jalankanTugas_(t); }
    catch (e) { logAksi_(t.jenis || 'SPL', t.nomor || t.id || '-', 'ANTRIAN GAGAL', '', t.t + ': ' + String(e)); }
  });
}

function jalankanTugas_(t) {
  if (t.t === 'konfirmasi') return kirimKonfirmasi_(t.jenis, t.nomor);
  if (t.t === 'selesai') return cekSelesai_(t.jenis, t.nomor);
  if (t.t === 'batal') return batalLanjutan_(t);
  if (t.t === 'persetujuanDinas') return kirimPersetujuanDinas_(t.nomor);
  throw new Error('Tugas tidak dikenal: ' + t.t);
}
/**
 * =====================================================================
 *  SISTEM SPL & SPD — File : Api.gs
 *  "Pintu data" HANYA-BACA untuk pengecekan pengajuan lembur di KantorKu.
 *  api=spl (SPL disetujui & belum dieksekusi) · api=laporan (simpan hasil cocok) · api=setuju (baris dicentang)
 *  api=tandai (tandai sudah dieksekusi) — semua butuh &key=<KUNCI>
 *  - Tanpa data gaji / estimasi kompensasi.
 *  - Kunci dilihat lewat menu ⚙️ SPL / SPD → 🔑 Kunci API pengecekan KantorKu.
 * =====================================================================
 */

function kunciApi_() {
  const p = PropertiesService.getScriptProperties();
  let k = p.getProperty('API_KEY');
  if (!k) { k = Utilities.getUuid().replace(/-/g, ''); p.setProperty('API_KEY', k); }
  return k;
}

const SH_CEK = 'Cek KantorKu';
const KOLOM_CEK = ['Dicek Pada', 'ID Baris', 'Nomor SPL', 'Tanggal Lembur', 'Employee ID', 'Full Name', 'Jam SPL', 'Jam Efektif SPL',
  'Jam KantorKu', 'Durasi KantorKu', 'Check Out', 'Catatan', 'Setuju', 'Status Eksekusi', 'Dieksekusi Pada'];

/** Sheet laporan pencocokan (dibuat otomatis) */
function sheetCek_() {
  const ss = ss_();
  let sh = ss.getSheetByName(SH_CEK);
  if (!sh) {
    sh = ss.insertSheet(SH_CEK);
    sh.getRange(1, 1, 1, KOLOM_CEK.length).setValues([KOLOM_CEK]).setFontWeight('bold').setBackground('#3A3A3A').setFontColor('#FFFFFF');
    sh.setFrozenRows(1);
    sh.getRange('M2:M').insertCheckboxes();
    sh.getRange('M1').setNote('Centang baris yang disetujui, lalu ketik "eksekusi" ke Claude.');
    sh.setColumnWidth(6, 200); sh.setColumnWidth(12, 260);
  }
  return sh;
}

function apiData_(q) {
  const keluar = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
  if (!q.key || q.key !== kunciApi_()) return keluar({ ok: false, error: 'kunci salah' });
  try {
    if (q.api === 'spl') return keluar(apiSpl_(q));
    if (q.api === 'laporan') return keluar(apiLaporan_(q));
    if (q.api === 'setuju') return keluar(apiSetuju_());
    if (q.api === 'tandai') return keluar(apiTandai_(q));
    if (q.api === 'diag') {   // diagnosa: log terakhir + baris lembur terbaru (tanpa gaji)
      const t = v => v instanceof Date ? Utilities.formatDate(v, TZ, 'dd/MM HH:mm:ss') : v;
      const log = sheet_(SH.LOG).getDataRange().getValues().slice(-40).map(r => r.map(t).join(' | '));
      const spl = bacaTabel_(SH.SPL, 'Nomor SPL').slice(-15).map(r => [r['Nomor SPL'], t(r['Tanggal Lembur']), r['Employee ID'], r['Full Name'], r['Status'], r['Email Pembuat'],
        'pdf:' + (r['Link PDF'] ? 1 : 0) + ' owner:' + (r['Link PDF Owner'] ? 1 : 0) + ' atasan:' + (r['Link PDF Atasan'] ? 1 : 0), r['Status Email'],
        'rantai:' + rantaiAtasan_(r['Employee ID']).map(a => a.id).join('>'), 'owner:' + owner_().map(o => o.id).join(',')].join(' | '));
      const spd = bacaTabel_(SH.PERSONEL, 'Nomor SPD').slice(-10).map(r => [r['Nomor SPD'], r['Employee ID'], r['Full Name'], r['Status'], r['Status Email'], t(r['Konfirmasi Pada']), r['Atasan'], r['Email Atasan'],
        'email:' + String((karyawan_()[String(r['Employee ID']).toUpperCase()] || {}).email || '-').replace(/^(.{3}).*(@.*)$/, '$1•••$2'),
        'penyetuju:' + ((penyetujuDinas_(r['Employee ID']) || {}).id || '-'), 'rantai:' + rantaiAtasan_(r['Employee ID']).map(a => a.id).join('>')].join(' | '));
      return keluar({ ok: true, log: log, spl: spl, spd: spd, kuota: MailApp.getRemainingDailyQuota() });
    }
    if (q.api === 'antrian') { const P = PropertiesService.getScriptProperties(), d = Number(P.getProperty(ANTRIAN_DETAK) || 0);
      return keluar({ ok: true, aktif: antrianAktif_(), detakTerakhir: d ? Utilities.formatDate(new Date(d), TZ, 'dd/MM/yyyy HH:mm:ss') : '-', antrian: JSON.parse(P.getProperty(ANTRIAN_KEY) || '[]').length }); }
    if (q.api === 'struktur') {   // struktur kolom Data Karyawan (tanpa gaji / email / data pribadi)
      const sh = sheet_(SH.KARYAWAN), v = sh.getDataRange().getValues(), head = v.shift().map(String);
      const pakai = head.map((h, i) => /atasan|peran|level|jabatan|position|posisi|grade|organi|depart|divisi|employee id|full name|status employee|nama/i.test(h) && !/salary|gaji|email|phone|telp|hp|nik|ktp|npwp|rek|bank|alamat|lahir|birth/i.test(h) ? i : -1).filter(i => i >= 0);
      return keluar({ ok: true, kolom: head, baris: v.filter(r => String(r[0]).trim()).map(r => Object.fromEntries(pakai.map(i => [head[i], r[i] instanceof Date ? ymd_(r[i]) : r[i]]))) });
    }
    if (q.api === 'peran') return keluar({ ok: true, karyawan: Object.values(karyawan_()).filter(k => k.peran !== 'KARYAWAN' || isOwner_(k))
      .map(k => ({ id: k.id, nama: k.nama, peran: k.peran, atasan: k.atasan, akses: k.akses, adaEmail: !!k.email, aktif: k.aktif })) });   // cek ADMIN/BOSS (tanpa gaji & email)
    return keluar({ ok: false, error: 'api tidak dikenal' });
  } catch (e) { return keluar({ ok: false, error: String(e.message || e) }); }
}

/** SPL yang DISETUJUI atasan & BELUM dieksekusi di KantorKu (default 30 hari terakhir). */
function apiSpl_(q) {
  const tglOk = s => /^\d{4}-\d{2}-\d{2}$/.test(String(s || ''));
  const kini = new Date(), hariIni = ymd_(kini);
  const dari = tglOk(q.dari) ? q.dari : ymd_(new Date(kini.getTime() - 30 * 864e5)), sampai = tglOk(q.sampai) ? q.sampai : hariIni;
  const spl = bacaTabel_(SH.SPL, 'Nomor SPL').filter(r => {
    const t = r['Tanggal Lembur'] instanceof Date ? ymd_(r['Tanggal Lembur']) : '';
    return t && dari <= t && t <= sampai && r['Status'] === 'DISETUJUI' && !String(r['KantorKu'] || '').trim();
  }).map(r => ({
    idBaris: String(r['ID Baris']), nomor: r['Nomor SPL'], tanggal: ymd_(r['Tanggal Lembur']), statusHari: r['Status Hari'],
    id: String(r['Employee ID']).toUpperCase(), nama: r['Full Name'],
    mulai: String(r['Jam Mulai']), selesai: String(r['Jam Selesai']), istirahat: r['Istirahat'], jamEfektif: Number(r['Jam Efektif']) || 0
  }));
  return { ok: true, dari: dari, sampai: sampai, spl: spl };
}

/** Simpan hasil cocok ke sheet "Cek KantorKu". q.data = JSON array
 *  [{idBaris, jamKantorku, durasiKantorku, checkOut, catatan}] (data SPL dilengkapi otomatis). */
function apiLaporan_(q) {
  const data = JSON.parse(q.data || '[]');
  const spl = {}; bacaTabel_(SH.SPL, 'Nomor SPL').forEach(r => spl[String(r['ID Baris'])] = r);
  const sh = sheetCek_(), n = sh.getLastRow() - 1;
  const ada = {};
  (n > 0 ? sh.getRange(2, 2, n, 1).getValues() : []).forEach((r, i) => ada[String(r[0])] = i + 2);
  let baru = 0, ubah = 0;
  data.forEach(d => {
    const r = spl[String(d.idBaris)]; if (!r) return;
    const baris = [new Date(), String(d.idBaris), r['Nomor SPL'], r['Tanggal Lembur'], r['Employee ID'], r['Full Name'],
      r['Jam Mulai'] + ' - ' + r['Jam Selesai'], Number(r['Jam Efektif']) || 0,
      d.jamKantorku || '', d.durasiKantorku || '', d.checkOut || '', d.catatan || ''];
    const row = ada[String(d.idBaris)];
    if (row) {
      if (String(sh.getRange(row, 14).getValue()).trim()) return;   // sudah dieksekusi → jangan diubah
      sh.getRange(row, 1, 1, 12).setValues([baris]); ubah++;
    } else {
      const r0 = sh.getLastRow() + 1;
      sh.getRange(r0, 1, 1, 12).setValues([baris]);
      sh.getRange(r0, 13).insertCheckboxes(); baru++;
    }
  });
  logAksi_('SPL', '-', 'CEK KANTORKU', 'Claude', baru + ' baris baru, ' + ubah + ' diperbarui');
  return { ok: true, baru: baru, diperbarui: ubah };
}

/** Baris yang sudah dicentang "Setuju" dan belum dieksekusi */
function apiSetuju_() {
  const sh = sheetCek_(), n = sh.getLastRow() - 1;
  if (n < 1) return { ok: true, baris: [] };
  const v = sh.getRange(2, 1, n, KOLOM_CEK.length).getValues();
  const st = {}; bacaTabel_(SH.SPL, 'Nomor SPL').forEach(r => st[String(r['ID Baris'])] = r['Status']);
  // lembur yang sudah dibatalkan / tidak lagi DISETUJUI tidak boleh di-approve di KantorKu
  v.forEach((r, i) => { if (r[12] === true && !String(r[13]).trim() && st[String(r[1])] && st[String(r[1])] !== ST.SETUJU) {
    sh.getRange(i + 2, 14).setValue('TIDAK DIEKSEKUSI: lembur ' + st[String(r[1])]); r[13] = 'x'; } });
  return { ok: true, baris: v.filter(r => r[12] === true && !String(r[13]).trim()).map(r => ({
    idBaris: String(r[1]), nomor: r[2], tanggal: r[3] instanceof Date ? ymd_(r[3]) : String(r[3]),
    id: String(r[4]).toUpperCase(), nama: r[5], jamSpl: r[6], jamEfektif: r[7], jamKantorku: r[8], durasiKantorku: r[9]
  })) };
}

/** Tandai sudah dieksekusi. q.ids = "idBaris1,idBaris2", q.status = APPROVED | GAGAL: alasan */
function apiTandai_(q) {
  const ids = String(q.ids || '').split(',').map(s => s.trim()).filter(Boolean);
  const status = String(q.status || 'APPROVED').slice(0, 200), now = new Date();
  const label = status + ' ' + Utilities.formatDate(now, TZ, 'dd/MM/yyyy HH:mm');
  const sh = sheetCek_(), n = sh.getLastRow() - 1;
  (n > 0 ? sh.getRange(2, 2, n, 1).getValues() : []).forEach((r, i) => {
    if (ids.indexOf(String(r[0])) >= 0) sh.getRange(i + 2, 14, 1, 2).setValues([[status, now]]);
  });
  if (status === 'APPROVED') {
    pastikanKolom_(SH.SPL, KOLOM_SPL_TAMBAHAN);
    const S = bacaBaris_('SPL');
    S.rows.filter(r => ids.indexOf(String(r['ID Baris'])) >= 0).forEach(r => ubahBaris_(S, r._row, { 'KantorKu': label }));
  }
  logAksi_('SPL', '-', 'EKSEKUSI KANTORKU', 'Claude', status + ': ' + ids.join(', '));
  return { ok: true, ditandai: ids.length };
}

/** Menu: tampilkan kunci & contoh URL */
function menuKunciApi() {
  const url = ScriptApp.getService().getUrl();
  SpreadsheetApp.getUi().alert('Kunci API pengecekan KantorKu:\n\n' + kunciApi_() +
    '\n\nContoh URL (data SPL hari ini):\n' + url + '?api=spl&key=' + kunciApi_() +
    '\n\nKunci ini hanya membuka data SPL (tanpa gaji). Jangan dibagikan ke orang lain.');
}
/**
 * =====================================================================
 *  SISTEM SPL & SPD — File : Ptj.gs
 *  Pertanggungjawaban Surat Perintah Dinas.
 *  - Surat Perintah Dinas tetap OPEN (Status Dinas = Aktif) sampai laporan
 *    pertanggungjawaban diisi admin → SELESAI.
 *  - Realisasi: tanggal berangkat/kembali aktual, lama per personel (uang saku),
 *    biaya aktual + bukti, ringkasan kegiatan, pengesahan di tempat tujuan.
 *  - PDF laporan dikirim ke admin & disimpan di folder "PDF SPD".
 * =====================================================================
 */

const SH_PTJ = { INDUK: 'Ptj SPD', PERSONEL: 'Ptj Personel SPD', BIAYA: 'Ptj Biaya SPD' };
const KOLOM_PTJ = {
  'Ptj SPD': ['Nomor SPD', 'Tgl Berangkat Aktual', 'Tgl Kembali Aktual', 'Lama Aktual (hari)', 'Ringkasan Kegiatan', 'Pengesahan Diterima',
    'Uang Saku Rencana', 'Uang Saku Realisasi', 'Biaya Lain Rencana', 'Biaya Lain Realisasi', 'Total Rencana', 'Total Realisasi', 'Selisih',
    'Uang Kas Perjalanan', 'Uang Kas Dikembalikan',
    'Dilaporkan Oleh', 'Dilaporkan Pada', 'Link PDF'],
  'Ptj Personel SPD': ['Nomor SPD', 'Employee ID', 'Full Name', 'Lama Rencana', 'Lama Aktual', 'Uang Saku Rencana', 'Uang Saku Realisasi', 'Keterangan'],
  'Ptj Biaya SPD': ['Nomor SPD', 'No', 'Jenis Biaya', 'Uraian', 'Jumlah (Rp)', 'Bukti']
};

function sheetPtj_(nama) {
  const ss = ss_();
  let sh = ss.getSheetByName(nama);
  if (!sh) {
    sh = ss.insertSheet(nama);
    sh.getRange(1, 1, 1, KOLOM_PTJ[nama].length).setValues([KOLOM_PTJ[nama]]).setFontWeight('bold').setBackground('#3A3A3A').setFontColor('#FFFFFF');
    sh.setFrozenRows(1);
  } else pastikanKolom_(nama, KOLOM_PTJ[nama]);
  return sh;
}

/** Data lengkap 1 Surat Perintah Dinas (induk + personel + biaya) */
function dataSpd_(nomor) {
  const I = bacaTabel_(SH.SPD, 'Nomor SPD').find(r => r['Nomor SPD'] === nomor);
  if (!I) return null;
  const tgl = d => d instanceof Date ? ymd_(d) : '';
  const per = bacaTabel_(SH.PERSONEL, 'Nomor SPD').filter(r => r['Nomor SPD'] === nomor).map(r => ({
    id: String(r['Employee ID']).toUpperCase(), nama: r['Full Name'], jabatan: r['Job Position'],
    lama: Number(r['Lama (hari)']) || 0, uangSaku: Number(r['Uang Saku']) || 0, keterangan: r['Keterangan'] || '', status: r['Status'] || ''
  }));
  const biaya = bacaTabel_(SH.BIAYA, 'Nomor SPD').filter(r => r['Nomor SPD'] === nomor).map(r => ({
    jenis: r['Jenis Biaya'], uraian: r['Uraian'], jumlah: Number(r['Jumlah (Rp)']) || 0, bukti: r['Bukti']
  }));
  const selesaiApproval = per.length > 0 && per.every(p => ST_FINAL.indexOf(p.status) >= 0);
  const disetujui = per.filter(p => p.status === ST.SETUJU);
  return {
    nomor: nomor, tanggalSurat: tgl(I['Tanggal Surat']), noReferensi: I['No Surat Referensi'] || '', customer: I['Nama Customer'] || '',
    dasar: I['Dasar Penugasan'] || '', maksud: I['Maksud / Tujuan'], tempat: I['Tempat Tujuan'],
    berangkat: tgl(I['Tgl Berangkat']), kembali: tgl(I['Tgl Kembali']), lama: Number(I['Lama (hari)']) || 0,
    transportasi: I['Transportasi'], sumberBiaya: I['Sumber Biaya'], statusDinas: I['Status Dinas'] || 'Aktif',
    statusApproval: I['Status'] || '', dibuatOleh: I['Dibuat Oleh'], emailPembuat: I['Email Pembuat'],
    personel: per, biaya: biaya, selesaiApproval: selesaiApproval,
    uangKas: Number(I['Uang Kas Perjalanan']) || 0,
    uangSakuRencana: disetujui.reduce((s, p) => s + p.uangSaku, 0), biayaRencana: biaya.reduce((s, b) => s + b.jumlah, 0)
  };
}

/** Daftar Surat Perintah Dinas yang masih OPEN (belum ada pertanggungjawaban). Semua admin bisa melihat. */
function spdOpen(token) {
  const u0 = sesiCek_(token);
  if (!punyaAkses_(u0, 'PTJ')) return [];
  return bacaTabel_(SH.SPD, 'Nomor SPD')
    .filter(r => String(r['Status Dinas'] || 'Aktif') !== 'Selesai')
    .map(r => dataSpd_(r['Nomor SPD']))
    .filter(Boolean)
    // semua personel ditolak / tidak bersedia → tidak perlu pertanggungjawaban, tetap ditampilkan agar bisa ditutup
    .sort((a, b) => (a.berangkat < b.berangkat ? -1 : 1));
}

/**
 * Simpan pertanggungjawaban.
 * d = {nomor, berangkat, kembali, ringkasan, pengesahan:boolean,
 *      personel:[{id, lama, keterangan}], biaya:[{jenis, uraian, jumlah, bukti}]}
 */
function simpanPtjSPD(token, d) {
  const u = sesiCek_(token);
  cekAkses_(u, 'PTJ');
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('Server sibuk, coba lagi beberapa detik lagi.');
  let hasil;
  try {
    _K = null; _P = null;
    const P = param_(), S = dataSpd_(String(d && d.nomor || ''));
    const err = [];
    if (!S) throw new Error('Surat Perintah Dinas tidak ditemukan.');
    if (S.statusDinas === 'Selesai') throw new Error('Pertanggungjawaban surat ini sudah diselesaikan.');
    if (!S.selesaiApproval) err.push('Proses konfirmasi / persetujuan personel belum selesai.');
    const tglOk = s => /^\d{4}-\d{2}-\d{2}$/.test(String(s || ''));
    if (!tglOk(d.berangkat) || !tglOk(d.kembali)) err.push('Tanggal berangkat & kembali aktual wajib diisi.');
    else if (d.kembali < d.berangkat) err.push('Tanggal kembali aktual lebih awal dari tanggal berangkat.');
    const lamaAktual = tglOk(d.berangkat) && tglOk(d.kembali) ? Math.max(Math.round((keTanggal_(d.kembali) - keTanggal_(d.berangkat)) / 864e5) + 1, 0) : 0;
    if (!String(d.ringkasan || '').trim()) err.push('Ringkasan hasil kegiatan wajib diisi.');

    const disetujui = {}; S.personel.filter(p => p.status === ST.SETUJU).forEach(p => disetujui[p.id] = p);
    const per = (d.personel || []).map(p => {
      const id = String(p.id).toUpperCase(), ref = disetujui[id], lama = Number(p.lama);
      if (!ref) { err.push('Personel ' + id + ' tidak termasuk personel yang disetujui.'); return null; }
      if (!(lama >= 0 && lama === Math.floor(lama))) err.push(sdr_(ref.nama) + ': lama aktual tidak valid.');
      else if (lama > lamaAktual) err.push(sdr_(ref.nama) + ': lama aktual melebihi lama dinas (' + lamaAktual + ' hari).');
      return { ref: ref, lama: lama, keterangan: String(p.keterangan || '').trim(), uangSaku: (lama || 0) * Number(P.UANG_SAKU_HARI) };
    }).filter(Boolean);
    Object.keys(disetujui).forEach(id => { if (!per.some(p => p.ref.id === id)) err.push(sdr_(disetujui[id].nama) + ' belum diisi lama aktualnya.'); });
    const biaya = (d.biaya || []).filter(b => b && (String(b.jenis || '').trim() || Number(b.jumlah)));
    biaya.forEach((b, i) => {
      if (!String(b.jenis || '').trim()) err.push('Biaya ke-' + (i + 1) + ': jenis biaya wajib dipilih.');
      if (!(Number(b.jumlah) >= 0)) err.push('Biaya ke-' + (i + 1) + ': jumlah tidak valid.');
    });
    if (err.length) return { ok: false, error: err };

    const sakuReal = per.reduce((s, p) => s + p.uangSaku, 0);
    const lainReal = biaya.reduce((s, b) => s + (Number(b.jumlah) || 0), 0);
    const totRencana = S.uangSakuRencana + S.biayaRencana, totReal = sakuReal + lainReal;
    const kasKembali = S.uangKas - lainReal;   // + = dikembalikan ke perusahaan, - = kekurangan (dibayar perusahaan)
    const now = new Date(), tp = s => Utilities.parseDate(s + ' 00:00', TZ, 'yyyy-MM-dd HH:mm');

    const ptj = { berangkat: d.berangkat, kembali: d.kembali, lamaAktual: lamaAktual, ringkasan: String(d.ringkasan).trim(),
      pengesahan: !!d.pengesahan, personel: per, biaya: biaya, sakuReal: sakuReal, lainReal: lainReal, totRencana: totRencana, totReal: totReal,
      uangKas: S.uangKas, kasKembali: kasKembali,
      oleh: u.nama, pada: now };
    const pdf = buatPdfPtj_(S, ptj);
    const file = folderPdf_(JENIS.SPD.folder).createFile(pdf);

    Object.keys(KOLOM_PTJ).forEach(sheetPtj_);
    tulisBaris_(SH_PTJ.INDUK, [{ 'Nomor SPD': S.nomor, 'Tgl Berangkat Aktual': tp(d.berangkat), 'Tgl Kembali Aktual': tp(d.kembali), 'Lama Aktual (hari)': lamaAktual,
      'Ringkasan Kegiatan': ptj.ringkasan, 'Pengesahan Diterima': ptj.pengesahan ? 'Ya' : 'Tidak',
      'Uang Saku Rencana': S.uangSakuRencana, 'Uang Saku Realisasi': sakuReal, 'Biaya Lain Rencana': S.biayaRencana, 'Biaya Lain Realisasi': lainReal,
      'Total Rencana': totRencana, 'Total Realisasi': totReal, 'Selisih': totReal - totRencana,
      'Uang Kas Perjalanan': S.uangKas, 'Uang Kas Dikembalikan': kasKembali,
      'Dilaporkan Oleh': u.nama, 'Dilaporkan Pada': now, 'Link PDF': file.getUrl() }]);
    tulisBaris_(SH_PTJ.PERSONEL, per.map(p => ({ 'Nomor SPD': S.nomor, 'Employee ID': p.ref.id, 'Full Name': p.ref.nama, 'Lama Rencana': p.ref.lama,
      'Lama Aktual': p.lama, 'Uang Saku Rencana': p.ref.uangSaku, 'Uang Saku Realisasi': p.uangSaku, 'Keterangan': p.keterangan })));
    tulisBaris_(SH_PTJ.BIAYA, biaya.map((b, i) => ({ 'Nomor SPD': S.nomor, 'No': i + 1, 'Jenis Biaya': String(b.jenis).trim(),
      'Uraian': String(b.uraian || '').trim(), 'Jumlah (Rp)': Number(b.jumlah) || 0, 'Bukti': String(b.bukti || '') })));

    // tutup Surat Perintah Dinas
    pastikanKolom_(SH.SPD, ['Tgl Laporan', 'Total Realisasi', 'Selisih', 'Uang Kas Dikembalikan', 'Link PDF Laporan']);
    const Ind = bacaBaris_SPD_induk_(), I = Ind.rows.find(x => x['Nomor SPD'] === S.nomor);
    if (I) ubahBaris_(Ind, I._row, { 'Status Dinas': 'Selesai', 'Tgl Laporan': now, 'Total Realisasi': totReal, 'Selisih': totReal - totRencana, 'Uang Kas Dikembalikan': kasKembali, 'Link PDF Laporan': file.getUrl() });
    SpreadsheetApp.flush();

    const isi = '<p>Pertanggungjawaban <b>Surat Perintah Dinas ' + escH_(S.nomor) + '</b> (' + escH_(S.tempat) + ') telah diselesaikan oleh ' + escH_(sdr_(u.nama)) + '.</p>' +
      '<p>Uang kas dalam perjalanan ' + rpH_(S.uangKas) + ' − biaya lainnya ' + rpH_(lainReal) + ' = <b>' + kasTeks_(kasKembali) + '</b>.<br>' +
      'Uang saku ' + rpH_(sakuReal) + ' dibayarkan melalui aplikasi KantorKu.</p>' +
      '<p>PDF laporan terlampir: <a href="' + file.getUrl() + '">' + escH_(file.getName()) + '</a></p>';
    const admin = String(P.EMAIL_ADMIN || '').trim().toLowerCase(), pembuat = String(S.emailPembuat || u.email).toLowerCase();
    kirimEmail_(pembuat, 'Laporan Pertanggungjawaban Dinas ' + S.nomor, isi, admin && admin !== pembuat ? admin : '', pdf);
    logAksi_('SPD', S.nomor, 'PERTANGGUNGJAWABAN', u.email, 'Kas ' + rpH_(S.uangKas) + ' − biaya ' + rpH_(lainReal) + ' → ' + kasTeks_(kasKembali));
    hasil = { ok: true, nomor: S.nomor, uangKas: S.uangKas, biayaLain: lainReal, kasKembali: kasKembali, uangSaku: sakuReal, pdf: file.getUrl() };
  } finally { lock.releaseLock(); }
  return hasil;
}

/** Teks hasil uang kas: positif = dikembalikan, negatif = kekurangan */
function kasTeks_(n) {
  return n > 0 ? 'uang kas dikembalikan ' + rpH_(n) : n < 0 ? 'kekurangan kas ' + rpH_(-n) + ' (dibayar perusahaan)' : 'uang kas habis terpakai (Rp 0)';
}

function buatPdfPtj_(S, R) {
  const sek = t => '<div style="background:#3A3A3A;color:#fff;font-weight:bold;font-size:10px;padding:4px 6px;margin:10px 0 4px">' + t + '</div>';
  const kv = arr => '<table style="width:100%;font-size:10px;border-collapse:collapse">' + arr.map(r => '<tr>' + r.map((x, i) =>
    '<td style="padding:3px 4px;' + (i % 2 ? '' : 'width:120px;color:#555') + '">' + (i % 2 ? ': <b>' + escH_(x) + '</b>' : escH_(x)) + '</td>').join('') + '</tr>').join('') + '</table>';
  const tg = s => { const t = keTanggal_(s); return t ? hariIndo_(t) + ', ' + tglIndo_(t) : '-'; };
  const totRow = (lbl, span, nilai, sisa) => '<tr><td colspan="' + span + '" style="padding:5px;border:1px solid #BBB;text-align:right;font-weight:bold;font-size:9px;background:#EFF0F2">' + lbl +
    '</td><td style="padding:5px;border:1px solid #BBB;text-align:right;font-weight:bold;font-size:9px;background:#EFF0F2;white-space:nowrap">' + nilai + '</td>' +
    (sisa ? '<td colspan="' + sisa + '" style="border:1px solid #BBB;background:#EFF0F2"></td>' : '') + '</tr>';
  const html = '<html><head><meta charset="utf-8"><style>@page{size:A4;margin:12mm}body{font-family:Arial,sans-serif;font-size:10px;color:#222}table.hal{width:100%;border-collapse:collapse}table.hal>thead{display:table-header-group}table.hal>thead>tr>td,table.hal>tbody>tr>td{padding:0}tr{page-break-inside:avoid}</style></head><body>' +
    '<table class="hal"><thead><tr><td>' + kopPdf_('LAPORAN PERTANGGUNGJAWABAN PERJALANAN DINAS', S.nomor) + '</td></tr></thead><tbody><tr><td>' +
    sek('A. DATA SURAT PERINTAH DINAS') +
    kv([['Maksud / Tujuan', S.maksud, 'Tempat Tujuan', S.tempat],
        ['Rencana Berangkat', tg(S.berangkat), 'Rencana Kembali', tg(S.kembali)],
        ['Nama Customer', S.customer || '-', 'Transportasi', S.transportasi || '-']]) +
    sek('B. REALISASI PERJALANAN') +
    kv([['Berangkat Aktual', tg(R.berangkat), 'Kembali Aktual', tg(R.kembali)],
        ['Lama Aktual', R.lamaAktual + ' hari', 'Pengesahan Tujuan', R.pengesahan ? 'Diterima' : 'Belum diterima']]) +
    sek('C. PERSONEL & UANG SAKU (DIBAYARKAN MELALUI APLIKASI KANTORKU)') +
    '<table style="border-collapse:collapse;width:100%"><tr>' + ['No', 'Employee ID', 'Full Name', 'Lama Rencana', 'Lama Aktual', 'Uang Saku Rencana', 'Uang Saku Realisasi', 'Keterangan'].map(thP_).join('') + '</tr>' +
    R.personel.map((p, i) => '<tr>' + tdP_(i + 1, 'text-align:center') + tdP_(escH_(p.ref.id)) + tdP_('<b>' + escH_(p.ref.nama) + '</b>') +
      tdP_(p.ref.lama + ' hari', 'text-align:center') + tdP_(p.lama + ' hari', 'text-align:center') + tdP_(rpH_(p.ref.uangSaku), 'text-align:right') +
      tdP_(rpH_(p.uangSaku), 'text-align:right') + tdP_(escH_(p.keterangan)) + '</tr>').join('') +
    totRow('TOTAL UANG SAKU REALISASI', 6, rpH_(R.sakuReal), 1) + '</table>' +
    sek('D. BIAYA PERJALANAN DINAS LAINNYA (REALISASI)') +
    (R.biaya.length ? '<table style="border-collapse:collapse;width:100%"><tr>' + ['No', 'Jenis Biaya', 'Uraian', 'Jumlah (Rp)', 'Bukti'].map(thP_).join('') + '</tr>' +
      R.biaya.map((b, i) => '<tr>' + tdP_(i + 1, 'text-align:center') + tdP_(escH_(b.jenis)) + tdP_(escH_(b.uraian)) + tdP_(rpH_(b.jumlah), 'text-align:right') + tdP_(escH_(b.bukti), 'text-align:center') + '</tr>').join('') +
      totRow('TOTAL BIAYA LAINNYA', 3, rpH_(R.lainReal), 1) + '</table>' : '<div style="font-size:9px;color:#777">Tidak ada biaya lainnya.</div>') +
    sek('E. PERHITUNGAN UANG KAS DALAM PERJALANAN') +
    '<table style="width:100%;border-collapse:collapse;font-size:10px">' +
    [['Uang kas dalam perjalanan (diterima)', rpH_(R.uangKas)], ['Dikurangi: total biaya lainnya (realisasi, bagian D)', '( ' + rpH_(R.lainReal) + ' )'],
     [R.kasKembali >= 0 ? 'UANG KAS YANG HARUS DIKEMBALIKAN' : 'KEKURANGAN UANG KAS (dibayarkan perusahaan)', rpH_(Math.abs(R.kasKembali))]]
      .map((r, i) => '<tr><td style="padding:5px;border:' + (i === 2 ? '2px solid #BE353F' : '1px solid #BBB') + ';font-weight:bold">' + r[0] + '</td><td style="padding:5px;border:' +
        (i === 2 ? '2px solid #BE353F;color:#BE353F' : '1px solid #BBB') + ';text-align:right;font-weight:bold;width:140px">' + r[1] + '</td></tr>').join('') + '</table>' +
    '<div style="font-size:9px;color:#555;margin-top:4px">Uang saku personel ' + rpH_(R.sakuReal) + ' dibayarkan melalui aplikasi KantorKu dan tidak termasuk perhitungan uang kas.</div>' +
    sek('F. RINGKASAN HASIL KEGIATAN') +
    '<div style="border:1px solid #BBB;padding:8px;font-size:10px;white-space:pre-wrap;min-height:40px">' + escH_(R.ringkasan) + '</div>' +
    '<p style="text-align:right;font-size:10px;margin-top:12px">' + (param_().KOTA_SURAT ? escH_(param_().KOTA_SURAT) + ', ' : '') + tglIndo_(R.pada) + '</p>' +
    '<table style="width:100%;margin-top:6px;font-size:10px;text-align:center"><tr><td>Dilaporkan oleh,</td><td>Diperiksa oleh,</td><td>Disetujui oleh,</td></tr>' +
    '<tr><td style="height:55px"></td><td></td><td></td></tr>' +
    '<tr><td>( ' + escH_(R.oleh) + ' )</td><td>(.........................)</td><td>(.........................)</td></tr>' +
    '<tr><td>Admin</td><td>Finance</td><td>HRD</td></tr></table>' +
    '<p style="font-size:8px;color:#777;margin-top:8px">Dokumen dibuat otomatis pada ' + waktuIndo_(new Date()) + '.</p></td></tr></tbody></table></body></html>';
  return pdfDari_(html, 'Laporan Dinas ' + String(S.nomor).replace(/\s*\/\s*/g, '-') + '.pdf');
}
/**
 * SISTEM SPL & SPD — File : Logo.gs
 * Logo PT Citra Banjar Abadi (JPG base64) untuk header form & kop surat PDF.
 * Jangan diubah manual. Untuk ganti logo, minta file Logo.gs baru.
 */
const LOGO_DATA_URI = 'data:image/jpeg;base64,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';

/** Logo aplikasi KantorKu (tombol menuju hris.kantorku.id di halaman konfirmasi) */
const LOGO_KANTORKU_URI = 'data:image/png;base64,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';
