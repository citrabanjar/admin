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
    .addItem('🧪 Tes pengiriman email (diagnosa bounce)', 'menuTesEmail')
    .addItem('🧪 Tes email isi asli (diagnosa bounce 2)', 'menuTesEmail2')
    .addItem('🧪 Tes email per bagian isi (diagnosa bounce 3)', 'menuTesEmail3')
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
  const appr = aktif.filter(k => k.peran === 'APPROVAL');
  if (appr.some(o => !o.email)) masalah.push({ j: '❌ APPROVAL tanpa email (tidak menerima rekap lembur)', isi: appr.filter(o => !o.email).map(k => k.id + ' — ' + k.nama) });
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
 * Diagnosa penolakan Gmail: kirim 5 variasi email ke 1 alamat (jeda 20 detik),
 * lalu lihat di Email Log Search variasi mana yang Bounced.
 */
function menuTesEmail() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('Tes pengiriman email', 'Alamat email tujuan tes (mis. budimul3006@gmail.com).\nAkan dikirim 5 email berbeda, jeda 20 detik (± 1,5 menit).', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const ke = r.getResponseText().trim();
  if (!/^\S+@\S+\.\S+$/.test(ke)) return ui.alert('Alamat email tidak valid.');
  const url = urlWeb_(), tombol = tombol_(url + '?a=dk&id=TES123&s=TES', '✅ Bersedia', '#1E7B4A') + tombol_(url + '?a=dk&id=TES123&s=TES&p=tidak', '❌ Tidak Bersedia', '#B42318');
  const v = [
    ['TES 1 — teks polos', { body: 'Tes 1: email teks polos tanpa logo dan tanpa link.' }],
    ['TES 2 — HTML + logo', { htmlBody: bungkusEmail_('<p>Tes 2: HTML dengan logo, tanpa link.</p>'), inlineImages: { logo: logoBlob_() } }],
    ['TES 3 — link alamat form', { htmlBody: '<p>Tes 3: HTML tanpa logo, dengan link biasa ke <a href="' + url + '">' + url + '</a>.</p>' }],
    ['TES 4 — tombol link aksi', { htmlBody: '<p>Tes 4: HTML tanpa logo, dengan tombol link aksi:</p><p>' + tombol + '</p>' }],
    ['TES 5 — lengkap (logo + tombol)', { htmlBody: bungkusEmail_('<p>Tes 5: seperti email konfirmasi asli (logo + tombol):</p><p>' + tombol + '</p>'), inlineImages: { logo: logoBlob_() } }]
  ];
  const hasil = v.map((x, i) => {
    if (i) Utilities.sleep(20000);
    try { MailApp.sendEmail(Object.assign({ to: ke, subject: x[0] + ' (' + Utilities.formatDate(new Date(), TZ, 'HH:mm:ss') + ')', body: x[1].body || 'Tes' }, x[1])); return x[0] + ': terkirim'; }
    catch (e) { return x[0] + ': GAGAL ' + e.message; }
  });
  logAksi_('SPL', '-', 'TES EMAIL', ke, hasil.join(' | '));
  ui.alert('Selesai.\n\n' + hasil.join('\n') + '\n\nCek Admin Console → Email Log Search untuk alamat ' + ke + ' (± 5 menit): mana yang Delivered dan mana yang Bounced.\nLink alamat yang dipakai: ' + url);
}

/** Diagnosa 2: kirim isi email konfirmasi ASLI dari 1 surat dinas.
 *  A = 1 email sendiri; jeda 30 detik; B & C = 2 email isi sama dikirim beruntun (seperti ke 2 personel sekaligus). */
function menuTesEmail2() {
  const ui = SpreadsheetApp.getUi();
  const r1 = ui.prompt('Tes email isi asli', 'Alamat email tujuan tes:', ui.ButtonSet.OK_CANCEL);
  if (r1.getSelectedButton() !== ui.Button.OK) return;
  const r2 = ui.prompt('Tes email isi asli', 'Nomor Surat Perintah Dinas yang dipakai isinya (mis. 002 / SPD / CBA / X / 2026):', ui.ButtonSet.OK_CANCEL);
  if (r2.getSelectedButton() !== ui.Button.OK) return;
  const ke = r1.getResponseText().trim(), nomor = r2.getResponseText().trim();
  const S = bacaBaris_('SPD'), r = S.rows.find(x => x['Nomor SPD'] === nomor);
  if (!r) return ui.alert('Nomor surat tidak ditemukan.');
  const I = induk_('SPD', nomor, r), J = JENIS.SPD, id = String(r['ID Baris']);
  const isi = '<p>Yth. <b>' + escH_(sdr_(r['Full Name'])) + '</b>,</p><p>Anda mendapat <b>' + J.surat + '</b> dengan rincian berikut. Mohon konfirmasi kesediaan Anda:</p>' +
    tabelDetail_('SPD', r, I, true) + '<p style="margin-top:18px">' + tombol_(linkAksi_(J.k, id, 'ya'), '✅ Bersedia', '#1E7B4A') + tombol_(linkAksi_(J.k, id, 'tidak'), '❌ Tidak Bersedia', '#B42318') + '</p>';
  const subj = 'Konfirmasi ' + J.kata + ' ' + ringkasTanggal_('SPD', r, I);
  const hasil = [];
  hasil.push('A (sendiri): ' + (kirimEmail_(ke, '[TES A] ' + subj, isi) || 'terkirim'));
  Utilities.sleep(30000);
  hasil.push('B (beruntun): ' + (kirimEmail_(ke, '[TES B] ' + subj, isi) || 'terkirim'));
  hasil.push('C (beruntun): ' + (kirimEmail_(ke, '[TES C] ' + subj, isi) || 'terkirim'));
  logAksi_('SPD', nomor, 'TES EMAIL 2', ke, hasil.join(' | '));
  ui.alert('Selesai.\n\n' + hasil.join('\n') + '\n\nCek Email Log Search untuk ' + ke + ': status TES A, TES B, TES C.\n(Tombol di email tes ini adalah link asli — jangan diklik.)');
}

/** Diagnosa 3: isi email konfirmasi dipecah per bagian (jeda 30 detik) untuk mencari bagian yang ditolak Gmail. */
function menuTesEmail3() {
  const ui = SpreadsheetApp.getUi();
  const r1 = ui.prompt('Tes email per bagian', 'Alamat email tujuan tes:', ui.ButtonSet.OK_CANCEL);
  if (r1.getSelectedButton() !== ui.Button.OK) return;
  const r2 = ui.prompt('Tes email per bagian', 'Nomor Surat Perintah Dinas yang dipakai isinya:', ui.ButtonSet.OK_CANCEL);
  if (r2.getSelectedButton() !== ui.Button.OK) return;
  const ke = r1.getResponseText().trim(), nomor = r2.getResponseText().trim();
  const r = bacaBaris_('SPD').rows.find(x => x['Nomor SPD'] === nomor);
  if (!r) return ui.alert('Nomor surat tidak ditemukan.');
  const I = induk_('SPD', nomor, r), id = String(r['ID Baris']);
  const tabel = rows => '<table style="border-collapse:collapse;width:100%;font-size:13px">' + rows.map(b => '<tr><td style="padding:6px 8px;background:#F4F5F7;width:150px;border:1px solid #E2E4E8;font-weight:bold">' +
    escH_(b[0]) + '</td><td style="padding:6px 8px;border:1px solid #E2E4E8">' + escH_(b[1]) + '</td></tr>').join('') + '</table>';
  const umum = [['Nomor SPD', nomor], ['Maksud / Tujuan', I['Maksud / Tujuan']], ['Tempat Tujuan', I['Tempat Tujuan']],
    ['Tanggal Dinas', tglIndo_(I['Tgl Berangkat']) + ' s/d ' + tglIndo_(I['Tgl Kembali'])], ['Personel', sdr_(r['Full Name']) + ' (' + r['Employee ID'] + ')'], ['Transportasi', I['Transportasi']]];
  const uang = [['Uang Saku', rpH_(r['Uang Saku'])], ['Uang Kas Perjalanan', rpH_(kasTotal_(I)) + ' (untuk seluruh tim, dipertanggungjawabkan setelah kembali)']];
  const tombol = '<p style="margin-top:18px">' + tombol_(linkAksi_('dk', id, 'ya'), '✅ Bersedia', '#1E7B4A') + tombol_(linkAksi_('dk', id, 'tidak'), '❌ Tidak Bersedia', '#B42318') + '</p>';
  const v = [
    ['TES D — tanpa nominal uang + tombol', tabel(umum) + tombol],
    ['TES E — tabel lengkap tanpa tombol', tabel(umum.concat(uang))],
    ['TES F — hanya nominal uang + tombol', tabel(uang) + tombol],
    ['TES G — tabel lengkap + tombol, judul netral', tabel(umum.concat(uang)) + tombol]
  ];
  const hasil = v.map((x, i) => { if (i) Utilities.sleep(30000);
    return x[0] + ': ' + (kirimEmail_(ke, x[0] === v[3][0] ? '[TES G] Surat Perintah ' + nomor : '[' + x[0].slice(0, 5) + '] Konfirmasi Dinas Luar ' + I['Tempat Tujuan'], '<p>Yth. <b>' + escH_(sdr_(r['Full Name'])) + '</b>,</p>' + x[1]) || 'terkirim'); });
  logAksi_('SPD', nomor, 'TES EMAIL 3', ke, hasil.join(' | '));
  ui.alert('Selesai.\n\n' + hasil.join('\n') + '\n\nCek Email Log Search: status TES D, E, F, G. (Jangan klik tombol di email tes.)');
}
