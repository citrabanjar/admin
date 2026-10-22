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
    .addItem('✉ Kirim ulang email konfirmasi (SPL/SPD)', 'menuKirimUlangKonfirmasi')
    .addItem('📄 Buat & kirim PDF (SPL/SPD) sekarang', 'menuPdfSPL')
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
  // Parameter tambahan (dibuat bila belum ada)
  const shP = sheet_(SH.PARAM), adaP = shP.getRange(1, 1, Math.max(shP.getLastRow(), 1), 1).getValues().map(r => String(r[0]).trim());
  [['URL_WEB_APP', 'https://citrabanjar.github.io/admin/', 'Alamat form & link di email. Kosong = link Apps Script (script.google.com).']]
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
