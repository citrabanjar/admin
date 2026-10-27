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
    if (lampiran) opt.attachments = Array.isArray(lampiran) ? lampiran : [lampiran];
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
/** Uang kas dinas = kas awal + tambahan kas */
function kasTotal_(I) { return (Number(I['Uang Kas Perjalanan']) || 0) + (Number(I['Tambahan Kas']) || 0); }
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
  ].concat(kasTotal_(I) ? [['Uang Kas Dinas', rpH_(kasTotal_(I)) + ' (untuk seluruh tim, dipertanggungjawabkan setelah kembali)']] : [], [
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
    // Surat Perintah Dinas: format ringkas yang terbukti tidak ditolak Gmail (uji 26/10: judul netral, tanpa kalimat "Anda mendapat… mohon konfirmasi")
    const dinas = jenis === 'SPD' && !keAdmin;
    const err = kirimEmail_(keAdmin ? Ir['Email Pembuat'] : k.email,
      dinas ? 'Surat Perintah Dinas ' + nomor : 'Konfirmasi ' + J.kata + ' ' + ringkasTanggal_(jenis, r, Ir) + (keAdmin ? ' — atas nama ' + sdr_(r['Full Name']) : ''),
      dinas ? '<p>Yth. <b>' + escH_(sdr_(r['Full Name'])) + '</b>,</p>' + tabelRingkasDinas_(r, Ir) + '<p style="margin-top:18px">' +
        tombol_(linkAksi_(J.k, id, 'ya'), '✅ Bersedia', '#1E7B4A') + tombol_(linkAksi_(J.k, id, 'tidak'), '❌ Tidak Bersedia', '#B42318') + '</p>' : isi);
    ubahBaris_(S, r._row, { 'Status Email': err ? 'GAGAL: ' + err : 'Konfirmasi terkirim ke ' + (keAdmin ? 'admin ' : '') + (keAdmin ? Ir['Email Pembuat'] : k.email) + ' ' + Utilities.formatDate(new Date(), TZ, 'dd/MM HH:mm') });
  });
}
/** Tabel ringkas Surat Perintah Dinas untuk email ke karyawan */
function tabelRingkasDinas_(r, I) {
  const baris = [['Nomor SPD', I['Nomor SPD']], ['Maksud / Tujuan', I['Maksud / Tujuan']], ['Tempat Tujuan', I['Tempat Tujuan']],
    ['Tanggal Dinas', tglIndo_(I['Tgl Berangkat']) + ' s/d ' + tglIndo_(I['Tgl Kembali'])], ['Personel', sdr_(r['Full Name']) + ' (' + r['Employee ID'] + ')'],
    ['Transportasi', I['Transportasi']]].concat(kasTotal_(I) ? [['Uang Kas Dinas', rpH_(kasTotal_(I)) + ' (untuk seluruh tim, dipertanggungjawabkan setelah kembali)']] : [],
    [['Diperintahkan oleh', sdr_((owner_()[0] || {}).nama || I['Dibuat Oleh'])]]);
  return '<table style="border-collapse:collapse;width:100%;font-size:13px">' + baris.map(b => '<tr><td style="padding:6px 8px;background:#F4F5F7;width:150px;border:1px solid #E2E4E8;font-weight:bold">' +
    escH_(b[0]) + '</td><td style="padding:6px 8px;border:1px solid #E2E4E8">' + escH_(b[1]) + '</td></tr>').join('') + '</table>';
}
function kirimKonfirmasiSPL_(nomor) { kirimKonfirmasi_('SPL', nomor); }

/* ======================= Halaman aksi (link dari email) ======================= */
function halamanAksi_(p) {
  if (p.a === 'b') return halamanBatal_(p);
  if (p.a === 'dg') return halamanGrupDinas_(p);
  if (p.a === 'pp') return halamanPerpanjangan_(p);   // persetujuan perpanjangan dinas (Ubah.gs)
  if (p.a === 'kt') return halamanTerimaKas_(p);      // konfirmasi terima tambahan kas (Ubah.gs)   // pembatalan lembur oleh atasan / owner (Batal.gs)
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
/** Persetujuan gabungan Surat Perintah Dinas: 1 tombol untuk seluruh karyawan di email. id = 'IDBaris.IDBaris...' */
function barisGrup_(id) {
  const ids = String(id || '').split('.').filter(Boolean), S = bacaBaris_('SPD');
  return { S: S, rows: ids.map(i => S.rows.find(x => String(x['ID Baris']) === i)).filter(Boolean) };
}
function halamanGrupDinas_(p) {
  let isi;
  try {
    if (!p.id || p.s !== tandaTangan_('dg', p.id)) throw new Error('Link tidak valid atau sudah diubah.');
    const G = barisGrup_(p.id), r0 = G.rows[0];
    if (!r0) throw new Error('Data tidak ditemukan (mungkin sudah dihapus).');
    const I = induk_('SPD', r0['Nomor SPD'], r0), td = 'padding:6px;border-top:1px solid #eee';
    const aktif = G.rows.filter(x => x['Status'] === ST.ATASAN);
    isi = '<h2>Persetujuan Dinas Luar</h2>' + tabelDetail_('SPD', r0, I, false, true).replace(/width:150px/g, 'width:130px') +
      '<table style="border-collapse:collapse;width:100%;font-size:13px;margin-top:12px"><tr><th style="text-align:left;padding:6px;background:#F4F5F7">Nama Karyawan</th><th style="padding:6px;background:#F4F5F7">Status</th></tr>' +
      G.rows.map(x => '<tr><td style="' + td + '"><b>' + escH_(sdr_(x['Full Name'])) + '</b><div class="sub">' + escH_(x['Employee ID']) + (x['Job Position'] ? ' — ' + escH_(x['Job Position']) : '') +
        '</div></td><td style="' + td + ';font-size:11px;text-align:center">' + escH_(x['Status']) + '</td></tr>').join('') + '</table>';
    if (!aktif.length) {
      isi += '<div class="msg info">Sudah diproses seluruhnya.</div>';
    } else {
      const pil = p.p === 'tidak' ? 'tidak' : 'ya';
      isi += '<div class="pilih">' +
        '<label class="op ya"><input type="radio" name="j" value="ya"' + (pil === 'ya' ? ' checked' : '') + '> ✅ Setujui</label>' +
        '<label class="op tidak"><input type="radio" name="j" value="tidak"' + (pil === 'tidak' ? ' checked' : '') + '> ❌ Tolak</label></div>' +
        '<div id="al"><label>Catatan <span class="sub">(opsional)</span></label><textarea id="alasan" rows="3"></textarea></div>' +
        '<button id="b" onclick="kirim()">Kirim</button><div id="o"></div>' +
        '<script>function kirim(){var b=document.getElementById("b"),o=document.getElementById("o");b.disabled=true;b.textContent="Mengirim…";' +
        'google.script.run.withSuccessHandler(function(m){document.querySelector(".pilih").remove();document.getElementById("al").remove();b.remove();o.innerHTML="<div class=\\"msg ok\\">"+m+"</div>"})' +
        '.withFailureHandler(function(e){b.disabled=false;b.textContent="Kirim";o.innerHTML="<div class=\\"msg err\\">"+e.message+"</div>"})' +
        '.prosesGrupDinas(' + JSON.stringify(String(p.id)) + ',' + JSON.stringify(p.s) + ',document.querySelector("input[name=j]:checked").value,document.getElementById("alasan").value)}</script>';
    }
  } catch (e) {
    isi = '<h2>Tidak dapat dibuka</h2><div class="msg err">' + escH_(e.message) + '</div>';
  }
  return halamanHtml_(isi);
}
/** Dipanggil dari halaman persetujuan gabungan dinas. */
function prosesGrupDinas(id, sig, jawaban, alasan) {
  if (!id || sig !== tandaTangan_('dg', id)) throw new Error('Link tidak valid.');
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('Server sibuk, coba lagi.');
  let nomor, kena;
  const setuju = jawaban !== 'tidak', alas = String(alasan || '').trim().slice(0, 500), now = new Date();
  try {
    _K = null; _P = null;
    pastikanKolom_(JENIS.SPD.sheet, JENIS.SPD.tambahan);
    const G = barisGrup_(id);
    if (!G.rows.length) throw new Error('Data tidak ditemukan.');
    nomor = G.rows[0]['Nomor SPD'];
    kena = G.rows.filter(x => x['Status'] === ST.ATASAN);
    if (!kena.length) return 'Sudah diproses sebelumnya.';
    kena.forEach(r => ubahBaris_(G.S, r._row, { 'Status': setuju ? ST.SETUJU : ST.TOLAK, 'Disetujui Oleh': r['Atasan'], 'Disetujui Pada': now, 'Catatan Atasan': alas }));
    logAksi_('SPD', nomor, setuju ? 'DISETUJUI ATASAN' : 'DITOLAK ATASAN', kena[0]['Email Atasan'], kena.map(r => r['Full Name']).join(', ') + (alas ? ' — ' + alas : ''));
    SpreadsheetApp.flush();
  } finally { lock.releaseLock(); }
  try { antri_({ t: 'selesai', jenis: 'SPD', nomor: nomor }); } catch (e) { logAksi_('SPD', nomor, 'PDF GAGAL', '', String(e)); }
  return 'Terima kasih. Dinas luar ' + kena.length + ' karyawan (' + kena.map(r => escH_(sdr_(r['Full Name']))).join(', ') + ') telah <b>' + (setuju ? 'DISETUJUI' : 'DITOLAK') + '</b>.';
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
    : ok.length + ' dari ' + rows.length + ' karyawan <b>disetujui</b> dinas.';
  const isi = '<p>Rekap <b>' + J.surat + ' ' + escH_(nomor) + '</b> — ' + escH_(ringkasTanggal_(jenis, rows[0], I)) + '.</p><p>' + ringkas +
    (belum.length ? ' <i>(' + belum.length + ' orang belum selesai diproses saat PDF ini dibuat)</i>' : '') + '</p>' +
    '<p>PDF terlampir.</p>';
  const admin = String(param_().EMAIL_ADMIN || '').trim().toLowerCase();
  const pembuat = String(I['Email Pembuat']).toLowerCase();
  kirimEmail_(pembuat, 'PDF ' + jenis + ' ' + nomor + ' — ' + ringkasTanggal_(jenis, rows[0], I), isi, admin && admin !== pembuat ? admin : '', pdf);
  logAksi_(jenis, nomor, 'PDF DIKIRIM', pembuat, file.getUrl());
  if (jenis === 'SPD' && !belum.length && ok.length) {
    // surat final (semua atasan sudah memproses) → PDF terlampir hanya ke 1 orang (personel disetujui pertama yang ber-email);
    // personel lain yang disetujui menerima konfirmasi persetujuan tanpa lampiran
    const K = karyawan_(), em = x => (K[String(x['Employee ID']).toUpperCase()] || {}).email || '';
    const pemegang = ok.find(em);
    const tim = '<p>Personel yang ditugaskan:</p><ol>' + ok.map(x => '<li><b>' + escH_(sdr_(x['Full Name'])) + '</b>' +
      (x['Job Position'] ? ' — ' + escH_(x['Job Position']) : '') + '</li>').join('') + '</ol>';
    ok.forEach(r => {
      const ke = em(r);
      if (!ke) return;
      const utama = r === pemegang;
      const err = kirimEmail_(ke, 'Surat Perintah Dinas ' + nomor + ' Disetujui — ' + ringkasTanggal_(jenis, r, I),
        '<p>Yth. <b>' + escH_(sdr_(r['Full Name'])) + '</b>,</p><p>Surat Perintah Dinas berikut telah <b style="color:#1E7B4A">DISETUJUI</b>.' +
        (utama ? ' Terlampir Surat Perintah Dinas.' : ' Surat Perintah Dinas dikirimkan kepada <b>' + escH_(sdr_(pemegang['Full Name'])) + '</b>.') + '</p>' +
        tabelDetail_('SPD', r, I, false, ok.length > 1) + (ok.length > 1 ? tim : ''), '', utama ? pdf : null);
      ubahBaris_(S, r._row, { 'Status Email': err ? 'GAGAL kirim: ' + err : (utama ? 'PDF surat terkirim ke ' : 'Konfirmasi disetujui terkirim ke ') + ke + ' ' + Utilities.formatDate(new Date(), TZ, 'dd/MM HH:mm') });
      if (err) logAksi_(jenis, nomor, 'EMAIL KARYAWAN GAGAL', ke, err);
    });
    if (pemegang) logAksi_(jenis, nomor, 'PDF KE KARYAWAN', em(pemegang), pemegang['Full Name']);
  }
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
    kirimRekapApproval_(hari, file, pdf, lama.length > 0, nomor, hasil);
  }
  if (!own.length && !sudah.owner && (paksa || (fin(hari) && hari.some(r => !r['Link PDF Owner'])))) {   // tanpa owner: approval tetap menerima rekap
    sudah.owner = 1;
    const lama = [...new Set(hari.map(r => r['Link PDF Owner']).filter(Boolean))];
    const pdf = buatPdfSPL_(hari, { estimasi: true, owner: true, nama: 'Rekap Lembur ' + tglFile_(r0) + ' - Semua Admin' });
    const file = folderPdf_(JENIS.SPL.folder).createFile(pdf);
    lama.forEach(hapusPdfLama_);
    hari.forEach(r => ubahBaris_(S, r._row, { 'Link PDF Owner': file.getUrl() }));
    kirimRekapApproval_(hari, file, pdf, lama.length > 0, nomor, hasil);
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
/** Peran APPROVAL: penyetuju lembur di aplikasi KantorKu — menerima rekap lembur seluruh karyawan (sama dengan owner, tanpa tombol pembatalan) */
function approval_() { return Object.values(karyawan_()).filter(k => k.aktif && k.email && k.peran === 'APPROVAL'); }
function kirimRekapApproval_(hari, file, pdf, revisi, nomor, hasil) {
  approval_().forEach(a => {
    kirimRekap_(a.email, hari, file, pdf, revisi, 'Semua Admin', true, null, Object.assign({}, a, { semua: true, approval: true }));
    logAksi_('SPL', nomor, revisi ? 'PDF REKAP APPROVAL DIPERBARUI' : 'PDF REKAP APPROVAL', a.email, file.getUrl());
    hasil.push('PDF rekap approval dikirim ke ' + a.nama);
  });
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
    (atasan.approval ? '<p style="margin-top:14px">Mohon lakukan <b>approval pengajuan lembur di aplikasi KantorKu</b> satu per satu untuk karyawan berstatus <b>DISETUJUI</b> pada rekap terlampir ' +
      '(nama, tanggal & jam lembur sesuai rekap). Pengajuan karyawan yang tidak tercantum / tidak berstatus DISETUJUI tidak perlu disetujui.</p>' : '') +
    (batalId && rows.some(r => BISA_BATAL_(r)) ? '<p style="margin-top:18px;border-top:1px solid #eee;padding-top:12px">Bila lembur di atas tidak diperlukan, Anda dapat membatalkan seluruhnya:<br>' +
      tombol_(linkAksi_('b', batalId), '🚫 Batalkan Lembur', '#B42318') + '</p>' : '')
  ) : '<p>Rekap <b>Surat Perintah Lembur ' + escH_(tgl) + '</b>' + (label ? ' — ' + escH_(label.charAt(0).toLowerCase() + label.slice(1)) : '') + '.</p>' +
    (revisi ? '<p style="background:#FFF4D6;padding:8px 10px;border-radius:6px"><b>Laporan diperbarui</b> karena ada Surat Perintah Lembur susulan di tanggal yang sama. ' +
      'Laporan sebelumnya sudah dihapus — gunakan PDF terbaru ini.</p>' : '') +
    '<p>' + nomor.length + ' surat (' + nomor.map(escH_).join(', ') + ')<br>' +
    escH_(ringkasStatus_(rows.filter(r => r['Status'] !== ST.TIDAK))) + (rows.some(r => r['Status'] === ST.TIDAK) ? ' · TIDAK BERSEDIA ' + rows.filter(r => r['Status'] === ST.TIDAK).length + ' orang' : '') +
    (owner ? ', estimasi kompensasi ' + rpH_(ok.reduce((s, r) => s + (Number(r['Estimasi Kompensasi']) || 0), 0)) : '') + '.' +
    (belum ? ' <i>(' + belum + ' orang belum selesai diproses saat PDF ini dibuat)</i>' : '') + '</p>' +
    '<p>PDF terlampir.</p>' +
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
  const setuju = rows.filter(r => [ST.TIDAK, ST.TOLAK, ST.BATAL].indexOf(r['Status']) < 0);   // personel yang tidak bersedia / ditolak tidak dicantumkan
  const ubah = perubahanSpd_(I['Nomor SPD']).filter(x => x.jenis === 'TAMBAH KAS' || x.jenis === 'SERAH KAS' || x.status === UBAH_ST.SETUJU);   // riwayat perubahan (revisi)
  const totSaku = setuju.reduce((s, r) => s + (Number(r['Uang Saku']) || 0), 0);
  const totLain = biaya.reduce((s, b) => s + (Number(b['Jumlah (Rp)']) || 0), 0);
  const sek = t => '<div style="background:#3A3A3A;color:#fff;font-weight:bold;font-size:10px;padding:4px 6px;margin:10px 0 4px">' + t + '</div>';
  const kv = arr => '<table style="width:100%;font-size:10px;border-collapse:collapse">' + arr.map(r => '<tr>' + r.map((x, i) =>
    '<td style="padding:3px 4px;' + (i % 2 ? '' : 'width:110px;color:#555') + '">' + (i % 2 ? ': <b>' + escH_(x) + '</b>' : escH_(x)) + '</td>').join('') + '</tr>').join('') + '</table>';
  const tglB = I['Tgl Berangkat'], tglK = I['Tgl Kembali'];
  // kop & nomor surat diulang di setiap halaman (thead tabel pembungkus)
  const html = '<html><head><meta charset="utf-8"><style>@page{size:A4;margin:12mm}body{font-family:Arial,sans-serif;font-size:10px;color:#222}table.hal{width:100%;border-collapse:collapse}table.hal>thead{display:table-header-group}table.hal>thead>tr>td,table.hal>tbody>tr>td{padding:0}tr{page-break-inside:avoid}</style></head><body>' +
    '<table class="hal"><thead><tr><td>' + kopPdf_('SURAT PERINTAH DINAS', I['Nomor SPD'] + (ubah.length ? ' — Revisi ' + ubah.length : '')) + '</td></tr></thead><tbody><tr><td>' +
    sek('A. REFERENSI SURAT') +
    kv([['Nomor Surat', I['No Surat Referensi'] || '-', 'Tanggal Surat', tglIndo_(I['Tanggal Surat'])], ['Nama Customer', I['Nama Customer'] || '-', 'Dasar Penugasan', I['Dasar Penugasan'] || '-']]) +
    sek('B. DETAIL PERJALANAN DINAS') +
    kv([['Maksud / Tujuan', I['Maksud / Tujuan'], 'Tempat Tujuan', I['Tempat Tujuan']],
        ['Tgl. Berangkat', hariIndo_(tglB) + ', ' + tglIndo_(tglB), 'Tgl. Kembali', hariIndo_(tglK) + ', ' + tglIndo_(tglK)],
        ['Lama Dinas', I['Lama (hari)'] + ' hari', 'Transportasi', I['Transportasi']],
        ['Sumber Biaya', I['Sumber Biaya'], 'Uang Saku / Hari', rpH_(P.UANG_SAKU_HARI)],
        ['Uang Kas Dinas', rpH_(kasTotal_(I)) + (Number(I['Tambahan Kas']) ? ' (termasuk tambahan ' + rpH_(I['Tambahan Kas']) + ')' : '')]]) +
    sek('C. KARYAWAN YANG DITUGASKAN') +
    '<table style="border-collapse:collapse;width:100%"><tr>' + ['No', 'Employee ID', 'Full Name', 'Job Position', 'Lama (hari)', 'Keterangan'].map(thP_).join('') + '</tr>' +
    setuju.map((r, i) => '<tr>' + tdP_(i + 1, 'text-align:center') + tdP_(escH_(r['Employee ID'])) + tdP_('<b>' + escH_(r['Full Name']) + '</b>') + tdP_(escH_(r['Job Position'])) +
      tdP_(r['Lama (hari)'], 'text-align:center') + tdP_(escH_(r['Keterangan'])) + '</tr>').join('') +
    '<tr><td colspan="6" style="padding:5px;border:1px solid #BBB;text-align:right;font-weight:bold;font-size:9px;background:#EFF0F2">TOTAL KARYAWAN : ' + setuju.length + ' orang</td></tr></table>' +
    (biaya.length ? (   // biaya lainnya kini diisi di Pertanggungjawaban Dinas; tampil hanya untuk surat lama yang punya rencana biaya
    sek('D. BIAYA PERJALANAN DINAS LAINNYA') +
    (biaya.length ? '<table style="border-collapse:collapse;width:100%"><tr>' + ['No', 'Jenis Biaya', 'Uraian', 'Jumlah (Rp)', 'Bukti'].map(thP_).join('') + '</tr>' +
      biaya.map((b, i) => '<tr>' + tdP_(i + 1, 'text-align:center') + tdP_(escH_(b['Jenis Biaya'])) + tdP_(escH_(b['Uraian'])) + tdP_(rpH_(b['Jumlah (Rp)']), 'text-align:right') + tdP_(escH_(b['Bukti']), 'text-align:center') + '</tr>').join('') +
      '<tr><td colspan="3" style="padding:5px;border:1px solid #BBB;text-align:right;font-weight:bold;font-size:9px;background:#EFF0F2">TOTAL BIAYA LAINNYA</td><td style="padding:5px;border:1px solid #BBB;text-align:right;font-weight:bold;font-size:9px;background:#EFF0F2">' + rpH_(totLain) + '</td><td style="border:1px solid #BBB;background:#EFF0F2"></td></tr></table>'
      : '<div style="font-size:9px;color:#777">Tidak ada biaya lainnya.</div>') +
    '<table style="width:100%;margin-top:6px;border-collapse:collapse"><tr><td style="padding:6px;border:2px solid #BE353F;font-weight:bold;font-size:11px">TOTAL BIAYA PERJALANAN DINAS (Uang Saku disetujui + Biaya Lainnya)</td>' +
    '<td style="padding:6px;border:2px solid #BE353F;font-weight:bold;font-size:11px;text-align:right;width:130px;color:#BE353F">' + rpH_(totSaku + totLain) + '</td></tr></table>'
    ) : '') +
    (ubah.length ? sek((biaya.length ? 'E' : 'D') + '. RIWAYAT PERUBAHAN SURAT') +
      '<table style="border-collapse:collapse;width:100%"><tr>' + ['No', 'Tanggal', 'Perubahan', 'Keterangan', 'Status'].map(thP_).join('') + '</tr>' +
      ubah.map((x, i) => '<tr>' + tdP_(i + 1, 'text-align:center') + tdP_(x.dibuatPada ? tglIndo_(x.dibuatPada) : '-') +
        tdP_(x.jenis === 'PERPANJANGAN' ? 'Perpanjangan s/d <b>' + tglIndo_(keTanggal_(x.kembaliBaru)) + '</b> (semula ' + tglIndo_(keTanggal_(x.kembaliLama)) + ')<br>' +
          x.personel.filter(p => p.status === UBAH_ST.SETUJU).map(p => escH_(p.nama)).join(', ') : (x.jenis === 'SERAH KAS' ? 'Serah terima uang kas <b>' : 'Tambahan uang kas <b>') + rpH_(x.tambahKas) + '</b>') +
        tdP_(escH_(x.alasan)) + tdP_(escH_(x.jenis !== 'PERPANJANGAN' ? (x.status === UBAH_ST.DITERIMA ? 'Diterima ketua tim' : 'Menunggu konfirmasi') : x.status)) + '</tr>').join('') + '</table>' : '') +
    sek(String.fromCharCode(68 + (biaya.length ? 1 : 0) + (ubah.length ? 1 : 0)) + '. KETENTUAN') +
    '<ol style="font-size:9px;margin:4px 0 0 14px;padding:0;line-height:1.5">' +
    '<li>Selama dinas karyawan tidak dapat mengajukan lembur.</li>' +
    '<li>Mintakan tanda tangan & stempel pada kolom Pengesahan di tempat tujuan sebagai bukti kehadiran.</li>' +
    '<li>Laporan perjalanan dinas dan pertanggungjawaban kas wajib dilaporkan ke admin paling lambat 3 hari kerja setelah kembali.</li></ol>' +
    sek(String.fromCharCode(69 + (biaya.length ? 1 : 0) + (ubah.length ? 1 : 0)) + '. PENGESAHAN DI TEMPAT TUJUAN') +
    '<table style="width:100%;border-collapse:collapse;font-size:9px"><tr>' + ['TIBA DI TEMPAT TUJUAN', 'BERANGKAT DARI TEMPAT TUJUAN'].map(t =>
      '<td style="width:50%;border:1px solid #BBB;padding:6px;vertical-align:top"><b>' + t + '</b><br><br>Tempat : ............................<br><br>Tanggal : ............................<br><br>Jam : ............................' +
      '<div style="height:50px"></div><div style="text-align:center;color:#555">( Nama, Tanda Tangan & Stempel Instansi )</div></td>').join('') + '</tr></table>' +
    '<p style="text-align:right;font-size:10px;margin-top:12px">Dikeluarkan di : ' + (P.KOTA_SURAT ? escH_(P.KOTA_SURAT) + ', ' : '') + tglIndo_(new Date()) + '</p>' +
    ttdPdf_(I['Dibuat Oleh'], 'Admin').replace('<td>(.........................)</td>', '<td>( ' + escH_((owner_()[0] || {}).nama || '.........................') + ' )</td>').replace('Manager Departemen', 'General Manager') +
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
        '<p>Personel berikut telah menyatakan <b style="color:#1E7B4A">BERSEDIA</b> melaksanakan perjalanan dinas. Mohon persetujuan Anda:</p>' +
        tabelDetail_('SPD', kirim[0], I, false, true) +
        '<table style="border-collapse:collapse;width:100%;font-size:13px;margin-top:12px"><tr><th style="' + th + '">No</th><th style="' + th + '">Nama Karyawan</th></tr>' +
        kirim.map((x, i) => '<tr><td style="' + td + ';text-align:center">' + (i + 1) + '</td><td style="' + td + '"><b>' + escH_(sdr_(x['Full Name'])) + '</b><div style="font-size:11px;color:#6B6F76">' + escH_(x['Employee ID']) +
          (x['Job Position'] ? ' — ' + escH_(x['Job Position']) : '') + ' · ' + escH_(x['Lama (hari)']) + ' hari' + (x['Keterangan'] ? ' — ' + escH_(x['Keterangan']) : '') + '</div></td></tr>').join('') + '</table>' +
        '<p style="margin-top:18px">' + tombol_(linkAksi_('dg', kirim.map(x => x['ID Baris']).join('.'), 'ya'), '✅ Setujui', '#1E7B4A') +
        tombol_(linkAksi_('dg', kirim.map(x => x['ID Baris']).join('.'), 'tidak'), '❌ Tolak', '#B42318') + '</p>';
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
