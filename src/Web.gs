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
  const url = String(param_().URL_WEB_APP || '').trim() || ScriptApp.getService().getUrl();
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
    if (spdAktif[r['Nomor SPD']] && STATUS_BATAL.indexOf(String(r['Status'] || '')) < 0) out.spd[String(r['Employee ID']).toUpperCase()] = r['Nomor SPD'];
  });
  return out;
}
