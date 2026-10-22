/**
 * =====================================================================
 *  SISTEM SPL & SPD — File : Spl.gs   (TAHAP 2–3 : FORM WEB SPL)
 *  - Login pembuat SPL dengan kode OTP ke email (aman untuk Gmail biasa
 *    maupun Google Workspace; karyawan TIDAK perlu akses ke spreadsheet)
 *  - Data untuk form (tanpa gaji)
 *  - Simpan SPL: validasi ulang di server, nomor otomatis, anti-bentrok
 *  - Riwayat "SPL Saya"
 *  Butuh: Inti.gs, FormSPL.html
 * =====================================================================
 */

const PERAN_PEMBUAT = ['PEMBUAT', 'MANAGER', 'HRD', 'ADMIN'];
const STATUS_BATAL = ['DITOLAK', 'DIBATALKAN'];
const SESI_DETIK = 6 * 3600;   // login berlaku 6 jam

/* ------------------------------ Web App ------------------------------ */
function doGet() {
  const t = HtmlService.createTemplateFromFile('FormSPL');
  // fungsi hitung yang SAMA dengan server dipakai juga di browser (satu sumber rumus)
  t.rumusBersama = 'const NAMA_HARI = ' + JSON.stringify(NAMA_HARI) + ';\n' +
    [ymd_, hariKe_, keTanggal_, jamKeAngka_, angkaKeJam_, rentang_, statusHari_, hitungLembur_].map(f => f.toString()).join('\n');
  t.namaPerusahaan = param_().NAMA_PERUSAHAAN || '';
  return t.evaluate()
    .setTitle('Surat Perintah Lembur')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Menu spreadsheet: tampilkan link form */
function bukaFormSPL() {
  const url = ScriptApp.getService().getUrl();
  const ui = SpreadsheetApp.getUi();
  if (!url) { ui.alert('Form belum di-deploy.\n\nDeploy → New deployment → Web app (lihat panduan).'); return; }
  ui.showModalDialog(HtmlService.createHtmlOutput(
    '<div style="font-family:Arial;font-size:13px">Link form SPL (bagikan ke PEMBUAT / MANAGER):<br><br>' +
    '<input style="width:100%;padding:6px" value="' + url + '" onclick="this.select()" readonly><br><br>' +
    '<a href="' + url + '" target="_blank">Buka form ↗</a></div>').setWidth(520).setHeight(170), 'Form SPL');
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
  if (!k) throw new Error('Akun ini tidak berhak membuat SPL.');
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
  if (!k) throw new Error('Email tidak terdaftar sebagai pembuat SPL. Hubungi Admin/HRD.');
  const c = CacheService.getScriptCache();
  const n = Number(c.get('N_' + email) || 0);
  if (n >= 5) throw new Error('Terlalu banyak permintaan kode. Coba lagi 1 jam lagi.');
  c.put('N_' + email, String(n + 1), 3600);
  const kode = String(Math.floor(100000 + Math.random() * 900000));
  c.put('OTP_' + email, kode, 600);
  c.remove('G_' + email);
  MailApp.sendEmail({
    to: email,
    subject: 'Kode masuk Form SPL: ' + kode,
    htmlBody: '<div style="font-family:Arial;font-size:14px">Halo ' + k.nama + ',<br><br>Kode masuk Form SPL Anda:<br>' +
      '<div style="font-size:28px;font-weight:bold;letter-spacing:6px;margin:12px 0;color:#1F3864">' + kode + '</div>' +
      'Berlaku 10 menit. Abaikan email ini jika Anda tidak meminta kode.<br><br>— Sistem SPL ' + (param_().NAMA_PERUSAHAAN || '') + '</div>'
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
  logAksi_('SPL', '-', 'LOGIN', email, 'Login form SPL');
  return { token: sesiBuat_(email), user: infoUser_(k) };
}

/* ------------------------------ Data form ------------------------------ */
function dataForm(token) {
  const u = sesiCek_(token);
  const P = param_();
  const kar = Object.values(karyawan_()).filter(k => k.aktif)
    .map(k => ({ id: k.id, nama: k.nama, jabatan: k.jabatan, departemen: k.departemen }))   // TANPA gaji
    .sort((a, b) => (a.departemen + a.nama).localeCompare(b.departemen + b.nama));
  const jam = []; for (let h = 0; h <= 24; h++) jam.push(angkaKeJam_(h));
  const pk = ['JAM_MULAI_MIN_SENJUM', 'JAM_MULAI_MIN_SABTU', 'JAM_MULAI_MIN_LIBUR', 'JAM_SELESAI_MAKS',
    'ISTIRAHAT_HARI_KERJA', 'ISTIRAHAT_HARI_LIBUR', 'BATAS_SELESAI_LIBUR', 'MAKS_KARYAWAN_PER_SPL'];
  return {
    user: infoUser_(u),
    karyawan: kar,
    libur: liburSet_(),
    P: Object.fromEntries(pk.map(k => [k, P[k]])),
    jam: jam
  };
}

/* ------------------------------ Simpan SPL ------------------------------ */
/**
 * @param {string} token
 * @param {{tanggal:'yyyy-mm-dd', baris:[{id, mulai, selesai, uraian}]}} data
 */
function simpanSPL(token, data) {
  const u = sesiCek_(token);
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('Server sibuk, coba simpan lagi beberapa detik lagi.');
  try {
    _K = null; _L = null; _P = null;
    const P = param_(), K = karyawan_(), L = liburSet_();
    const err = [];
    const tglStr = String(data && data.tanggal || '');
    if (!/^\d{4}-\d{2}-\d{2}$/.test(tglStr)) throw new Error('Tanggal lembur wajib diisi.');
    const baris = (data.baris || []).filter(b => b && String(b.id || '').trim());
    if (!baris.length) throw new Error('Isi minimal 1 karyawan.');
    if (baris.length > Number(P.MAKS_KARYAWAN_PER_SPL)) throw new Error('Maksimal ' + P.MAKS_KARYAWAN_PER_SPL + ' karyawan per SPL.');

    const bentrok = cekBentrok_(tglStr);
    const sudah = {};
    const hasil = baris.map((b, i) => {
      const id = String(b.id).trim().toUpperCase(), no = 'Baris ' + (i + 1) + ': ';
      const k = K[id];
      if (!k || !k.aktif) { err.push(no + 'Employee ID ' + id + ' tidak ditemukan / tidak aktif.'); return null; }
      if (sudah[id]) err.push(no + k.nama + ' tercantum lebih dari sekali.');
      sudah[id] = true;
      if (!String(b.uraian || '').trim()) err.push(no + 'Uraian pekerjaan ' + k.nama + ' wajib diisi.');
      if (bentrok.spl[id]) err.push(no + k.nama + ' sudah ada di ' + bentrok.spl[id] + ' pada tanggal yang sama.');
      if (bentrok.spd[id]) err.push(no + k.nama + ' sedang dinas luar (' + bentrok.spd[id] + ') — lembur tidak dapat diajukan.');
      const h = hitungLembur_({ tanggal: tglStr, mulai: b.mulai, selesai: b.selesai, gaji: k.gaji }, P, L);
      h.error.forEach(e => err.push(no + k.nama + ' — ' + e));
      return { k: k, b: b, h: h };
    });
    if (err.length) return { ok: false, error: err };

    // Tentukan alur approval
    const alur = alurApprovalSPL_(u);
    const tgl = Utilities.parseDate(tglStr + ' 00:00', TZ, 'yyyy-MM-dd HH:mm');
    const nomor = nomorBerikutnya_('SPL', tgl);
    const now = new Date();

    const sh = sheet_(SH.SPL), c = kolom_(SH.SPL), lebar = sh.getLastColumn();
    const rows = hasil.map(x => {
      const r = new Array(lebar).fill('');
      const set = (judul, v) => { if (c[judul]) r[c[judul] - 1] = v; };
      set('Nomor SPL', nomor); set('Tanggal Lembur', tgl); set('Hari', x.h.hari); set('Status Hari', x.h.statusHari);
      set('Employee ID', x.k.id); set('Full Name', x.k.nama); set('Departemen', x.k.departemen);
      set('Jam Mulai', angkaKeJam_(jamKeAngka_(x.b.mulai))); set('Jam Selesai', angkaKeJam_(jamKeAngka_(x.b.selesai)));
      set('Istirahat', x.h.istirahat); set('Jam Efektif', x.h.jamEfektif); set('Estimasi Kompensasi', x.h.kompensasi);
      set('Uraian Pekerjaan', String(x.b.uraian).trim());
      set('Dibuat Oleh', u.nama); set('Email Pembuat', u.email); set('Dibuat Pada', now);
      set('Status', alur.status);
      if (alur.diketahuiOleh) { set('Diketahui Oleh', alur.diketahuiOleh); set('Diketahui Pada', now); }
      set('Catatan', [].concat(x.h.peringatan, alur.catatan ? [alur.catatan] : []).join(' | '));
      set('Status Email', 'Belum dikirim');
      set('ID Baris', Utilities.getUuid().slice(0, 8));
      return r;
    });
    const r0 = sh.getLastRow() + 1;
    sh.getRange(r0, c['Jam Mulai'], rows.length, 2).setNumberFormat('@');
    sh.getRange(r0, 1, rows.length, lebar).setValues(rows);
    SpreadsheetApp.flush();

    const totalJam = hasil.reduce((s, x) => s + x.h.jamEfektif, 0);
    logAksi_('SPL', nomor, 'DIBUAT', u.email, rows.length + ' karyawan, ' + totalJam + ' jam, tgl ' + tglStr + ' → ' + alur.status);
    return { ok: true, nomor: nomor, jumlah: rows.length, totalJam: totalJam, status: alur.status };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Alur approval: Dibuat (pembuat) → Diketahui (Manager departemen pembuat) → Disetujui (HRD).
 * - Pembuat adalah MANAGER  → otomatis "Diketahui" oleh dirinya, langsung ke HRD
 * - Departemen tanpa MANAGER → Parameter TANPA_MANAGER: "LANGSUNG HRD" atau email pengganti
 */
function alurApprovalSPL_(pembuat) {
  if (pembuat.peran === 'MANAGER')
    return { status: 'MENUNGGU HRD', diketahuiOleh: pembuat.nama, approver: daftarHRD_() };
  const mgr = managerDepartemen_(pembuat.departemen);
  if (mgr.length) return { status: 'MENUNGGU MANAGER', approver: mgr };
  const t = String(param_().TANPA_MANAGER || 'LANGSUNG HRD').trim();
  if (t.indexOf('@') > 0) {
    const pengganti = cariKaryawanByEmail_(t);
    return { status: 'MENUNGGU MANAGER', approver: [pengganti || { nama: t, email: t.toLowerCase() }] };
  }
  return { status: 'MENUNGGU HRD', approver: daftarHRD_(), catatan: 'Departemen tanpa Manager — langsung ke HRD' };
}

/** Karyawan yang sudah punya SPL aktif / sedang SPD pada tanggal tsb. → {spl:{id:nomor}, spd:{id:nomor}} */
function cekBentrok_(tglStr) {
  const out = { spl: {}, spd: {} };
  bacaTabel_(SH.SPL, 'Nomor SPL').forEach(r => {
    if (r['Tanggal Lembur'] instanceof Date && ymd_(r['Tanggal Lembur']) === tglStr && STATUS_BATAL.indexOf(String(r['Status'])) < 0)
      out.spl[String(r['Employee ID']).toUpperCase()] = r['Nomor SPL'];
  });
  const spdAktif = {};
  bacaTabel_(SH.SPD, 'Nomor SPD').forEach(r => {
    const b = r['Tgl Berangkat'], k = r['Tgl Kembali'];
    if (b instanceof Date && k instanceof Date && STATUS_BATAL.indexOf(String(r['Status'])) < 0 &&
        ymd_(b) <= tglStr && tglStr <= ymd_(k)) spdAktif[r['Nomor SPD']] = true;
  });
  bacaTabel_(SH.PERSONEL, 'Nomor SPD').forEach(r => {
    if (spdAktif[r['Nomor SPD']]) out.spd[String(r['Employee ID']).toUpperCase()] = r['Nomor SPD'];
  });
  return out;
}

/** Riwayat SPL buatan user (30 terakhir), dikelompokkan per nomor. */
function riwayatSPL(token) {
  const u = sesiCek_(token);
  const g = {};
  bacaTabel_(SH.SPL, 'Nomor SPL').forEach(r => {
    if (String(r['Email Pembuat']).toLowerCase() !== u.email) return;
    const n = r['Nomor SPL'];
    if (!g[n]) g[n] = {
      nomor: n, tanggal: r['Tanggal Lembur'] instanceof Date ? ymd_(r['Tanggal Lembur']) : '',
      hari: r['Hari'], statusHari: r['Status Hari'], status: r['Status'],
      dibuat: r['Dibuat Pada'] instanceof Date ? r['Dibuat Pada'].getTime() : 0, orang: [], totalJam: 0
    };
    g[n].orang.push({ nama: r['Full Name'], mulai: String(r['Jam Mulai']), selesai: String(r['Jam Selesai']), jam: Number(r['Jam Efektif']) || 0, uraian: r['Uraian Pekerjaan'] });
    g[n].totalJam += Number(r['Jam Efektif']) || 0;
  });
  return Object.values(g).sort((a, b) => b.dibuat - a.dibuat).slice(0, 30);
}
