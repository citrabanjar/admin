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
 *    Karyawan HL (Employee ID mengandung "HL"): faktor 1 (flat), hari kerja maupun hari libur
 * @param {{tanggal, mulai, selesai, gaji, id}} x
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

  const faktor = /HL/i.test(String(x.id || '')) ? 1 : Number(libur_ ? P.FAKTOR_HARI_LIBUR : P.FAKTOR_HARI_KERJA);
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
/** Teks isian surat dinas disimpan HURUF BESAR */
function besar_(s) { return String(s == null ? '' : s).trim().toUpperCase(); }
function besarSemua_(d, kunci, daftar) {
  if (!d) return d;
  kunci.forEach(k => { if (d[k] != null) d[k] = besar_(d[k]); });
  Object.keys(daftar || {}).forEach(k => (d[k] || []).forEach(x => daftar[k].forEach(f => { if (x && x[f] != null) x[f] = besar_(x[f]); })));
  return d;
}
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
