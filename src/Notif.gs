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
const KOLOM_SPL_TAMBAHAN = ['Atasan', 'Email Atasan', 'Konfirmasi Pada', 'Alasan Tidak Bersedia', 'Catatan Atasan', 'Link PDF'];
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
function tabelDetail_(jenis, r, I, denganKompensasi) {
  const baris = jenis === 'SPL' ? [
    ['Nomor SPL', r['Nomor SPL']],
    ['Tanggal Lembur', r['Hari'] + ', ' + tglIndo_(r['Tanggal Lembur']) + ' (' + r['Status Hari'] + ')'],
    ['Karyawan', r['Full Name'] + ' (' + r['Employee ID'] + ')'],
    ['Jam Lembur', r['Jam Mulai'] + ' – ' + r['Jam Selesai'] + (r['Istirahat'] && r['Istirahat'] !== '-' ? ' (istirahat ' + r['Istirahat'] + ')' : '')],
    ['Jam Efektif', String(r['Jam Efektif']).replace('.', ',') + ' jam'],
    ['Uraian Pekerjaan', r['Uraian Pekerjaan']]
  ].concat(denganKompensasi ? [['Estimasi Kompensasi', rpH_(r['Estimasi Kompensasi'])]] : [], [
    ['Diperintahkan oleh', namaAtasan_(r['Employee ID']) || r['Dibuat Oleh']]
  ]) : [
    ['Nomor SPD', r['Nomor SPD']],
    ['Maksud / Tujuan', I['Maksud / Tujuan']],
    ['Tempat Tujuan', I['Tempat Tujuan']],
    ['Tanggal Dinas', hariIndo_(I['Tgl Berangkat']) + ', ' + tglIndo_(I['Tgl Berangkat']) + ' s/d ' + hariIndo_(I['Tgl Kembali']) + ', ' + tglIndo_(I['Tgl Kembali']) + ' (' + I['Lama (hari)'] + ' hari)'],
    ['Personel', r['Full Name'] + ' (' + r['Employee ID'] + ')' + (r['Job Position'] ? ' — ' + r['Job Position'] : '')],
    ['Lama Tugas', r['Lama (hari)'] + ' hari' + (r['Keterangan'] ? ' — ' + r['Keterangan'] : '')],
    ['Uang Saku', rpH_(r['Uang Saku'])],
    ['Transportasi', I['Transportasi']],
    ['Customer / Dasar', [I['Nama Customer'], I['Dasar Penugasan']].filter(Boolean).join(' — ') || '-'],
    ['Diperintahkan oleh', namaAtasan_(r['Employee ID']) || I['Dibuat Oleh']]
  ];
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
      ? '<p><b>' + escH_(r['Full Name']) + '</b> belum memiliki email di sistem. Mohon konfirmasi kesediaan ' + J.kata.toLowerCase() + ' <b>atas nama karyawan tersebut</b>:</p>'
      : '<p>Yth. <b>' + escH_(r['Full Name']) + '</b>,</p><p>Anda mendapat <b>' + J.surat + '</b> dengan rincian berikut. Mohon konfirmasi kesediaan Anda:</p>') +
      tabelDetail_(jenis, r, Ir, true) + '<p style="margin-top:18px">' +
      tombol_(linkAksi_(J.k, id, 'ya'), '✅ Bersedia', '#1E7B4A') + tombol_(linkAksi_(J.k, id, 'tidak'), '❌ Tidak Bersedia', '#B42318') + '</p>';
    const err = kirimEmail_(keAdmin ? Ir['Email Pembuat'] : k.email,
      'Konfirmasi ' + J.kata + ' ' + ringkasTanggal_(jenis, r, Ir) + (keAdmin ? ' — atas nama ' + r['Full Name'] : ''), isi);
    ubahBaris_(S, r._row, { 'Status Email': err ? 'GAGAL: ' + err : 'Konfirmasi terkirim ' + (keAdmin ? 'ke admin ' : '') + Utilities.formatDate(new Date(), TZ, 'dd/MM HH:mm') });
  });
}
function kirimKonfirmasiSPL_(nomor) { kirimKonfirmasi_('SPL', nomor); }

/* ======================= Halaman aksi (link dari email) ======================= */
function halamanAksi_(p) {
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
  return HtmlService.createHtmlOutput(html).setTitle('SPL & SPD — PT Citra Banjar Abadi')
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
        const isi = '<p><b>' + escH_(r['Full Name']) + '</b> menyatakan <b style="color:#B42318">TIDAK BERSEDIA</b> ' + J.kata.toLowerCase() + '.</p>' +
          tabelDetail_(jenis, r, I) + (alas ? '<p><b>Alasan:</b> ' + escH_(alas) + '</p>' : '');
        kirimEmail_(I['Email Pembuat'], 'Tidak Bersedia ' + J.kata + ' — ' + r['Full Name'] + ' (' + tgl + ')', isi, k.email || '');
        logAksi_(jenis, nomor, 'TIDAK BERSEDIA', r['Full Name'], alas);
        pesan = 'Terima kasih. Ketidaksediaan Anda sudah dicatat dan diberitahukan ke admin.';
      } else {
        const atasan = cariAtasan_(r['Employee ID']);
        const ke = atasan ? atasan.email : I['Email Pembuat'];
        ubahBaris_(S, r._row, { 'Status': ST.ATASAN, 'Konfirmasi Pada': now,
          'Atasan': atasan ? atasan.nama : '(admin — atasan belum diisi)', 'Email Atasan': ke });
        const isi = (atasan ? '<p>Yth. <b>' + escH_(atasan.nama) + '</b>,</p>' : '<p><i>Atasan langsung karyawan ini belum diisi di Data Karyawan, sehingga permintaan persetujuan dikirim ke admin.</i></p>') +
          '<p><b>' + escH_(r['Full Name']) + '</b> telah menyatakan <b style="color:#1E7B4A">BERSEDIA</b> ' + J.kata.toLowerCase() + '. Mohon persetujuan Anda:</p>' + tabelDetail_(jenis, r, I) +
          '<p style="margin-top:18px">' + tombol_(linkAksi_(J.t, id, 'ya'), '✅ Setujui', '#1E7B4A') + tombol_(linkAksi_(J.t, id, 'tidak'), '❌ Tolak', '#B42318') + '</p>';
        const err = kirimEmail_(ke, 'Persetujuan ' + J.kata + ' — ' + r['Full Name'] + ' (' + tgl + ')', isi);
        ubahBaris_(S, r._row, { 'Status Email': err ? 'GAGAL kirim ke atasan: ' + err : 'Persetujuan terkirim ke ' + ke });
        logAksi_(jenis, nomor, 'BERSEDIA', r['Full Name'], 'Diteruskan ke ' + ke);
        pesan = 'Terima kasih. Kesediaan Anda sudah dicatat dan diteruskan ke ' + (atasan ? 'atasan (' + escH_(atasan.nama) + ')' : 'admin') + ' untuk persetujuan.' +
          (jenis === 'SPL' ? '</div>' + htmlInstruksiKantorku_() + '<div>' : '');
      }
    } else {
      if (r['Status'] !== ST.ATASAN) return 'Sudah diproses sebelumnya — status: ' + r['Status'];
      const setuju = jawaban !== 'tidak';
      ubahBaris_(S, r._row, { 'Status': setuju ? ST.SETUJU : ST.TOLAK, 'Disetujui Oleh': r['Atasan'], 'Disetujui Pada': now, 'Catatan Atasan': alas });
      logAksi_(jenis, nomor, setuju ? 'DISETUJUI ATASAN' : 'DITOLAK ATASAN', r['Email Atasan'], r['Full Name'] + (alas ? ' — ' + alas : ''));
      pesan = 'Terima kasih. ' + J.kata + ' ' + escH_(r['Full Name']) + ' telah <b>' + (setuju ? 'DISETUJUI' : 'DITOLAK') + '</b>.';
    }
    SpreadsheetApp.flush();
  } finally { lock.releaseLock(); }
  try { cekSelesai_(jenis, nomor); } catch (e) { logAksi_(jenis, nomor, 'PDF GAGAL', '', String(e)); }
  return pesan;
}

/** Nama atasan langsung (tanpa jabatan) — dipakai untuk "Diperintahkan oleh". Kosong bila atasan belum diisi. */
function namaAtasan_(employeeId) {
  const K = karyawan_(), k = K[String(employeeId).toUpperCase()];
  return k && k.atasan && K[k.atasan] ? K[k.atasan].nama : '';
}
/** Atasan langsung dari kolom "Atasan" (Employee ID). */
function cariAtasan_(employeeId) {
  const K = karyawan_(), k = K[String(employeeId).toUpperCase()];
  if (!k || !k.atasan) return null;
  const a = K[k.atasan];
  return a && a.email && a.aktif ? a : null;
}

/* ======================= 4. PDF ke admin ======================= */
/** Bila semua orang di surat sudah final & PDF belum dibuat → buat PDF & kirim ke admin. */
function cekSelesai_(jenis, nomor, paksa) {
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
    '<div style="text-align:center;font-size:11px;margin-bottom:10px">Nomor : ' + escH_(nomor) + '</div>';
}
function pdfDari_(html, nama) {
  return HtmlService.createHtmlOutput(html).getBlob().getAs('application/pdf').setName(nama);
}

function buatPdfSPL_(rows) {
  const r0 = rows[0], P = param_();
  const setuju = rows.filter(r => r['Status'] === ST.SETUJU);
  const totJam = setuju.reduce((s, r) => s + (Number(r['Jam Efektif']) || 0), 0);
  const isiTabel = rows.map((r, i) => '<tr>' +
    tdP_(i + 1, 'text-align:center') + tdP_(escH_(r['Employee ID'])) + tdP_('<b>' + escH_(r['Full Name']) + '</b>') +
    tdP_(escH_(r['Jam Mulai']), 'text-align:center') + tdP_(escH_(r['Jam Selesai']), 'text-align:center') + tdP_(escH_(r['Istirahat']), 'text-align:center') +
    tdP_(String(r['Jam Efektif']).replace('.', ','), 'text-align:center') + tdP_(rpH_(r['Estimasi Kompensasi']), 'text-align:right;white-space:nowrap') +
    tdP_(escH_(r['Uraian Pekerjaan'])) + tdP_(selStatus_(r)) + tdP_(selAtasan_(r)) + '</tr>').join('');
  const totKomp = setuju.reduce((s, r) => s + (Number(r['Estimasi Kompensasi']) || 0), 0);
  const html = '<html><head><meta charset="utf-8"><style>@page{size:A4 landscape;margin:12mm}body{font-family:Arial,sans-serif;font-size:10px;color:#222}</style></head><body>' +
    kopPdf_('SURAT PERINTAH LEMBUR (SPL)', r0['Nomor SPL']) +
    '<table style="font-size:10px;margin-bottom:8px"><tr><td style="width:110px">Tanggal Lembur</td><td>: <b>' + escH_(r0['Hari']) + ', ' + tglIndo_(r0['Tanggal Lembur']) + '</b></td></tr>' +
    '<tr><td>Status Hari</td><td>: ' + escH_(r0['Status Hari']) + '</td></tr><tr><td>Dibuat oleh</td><td>: ' + escH_(r0['Dibuat Oleh']) + ' — ' + waktuIndo_(r0['Dibuat Pada']) + '</td></tr></table>' +
    '<table style="border-collapse:collapse;width:100%"><tr>' + ['No', 'Employee ID', 'Full Name', 'Mulai', 'Selesai', 'Istirahat', 'Jam Efektif', 'Estimasi Kompensasi', 'Uraian Pekerjaan', 'Status', 'Persetujuan Atasan'].map(thP_).join('') + '</tr>' + isiTabel +
    '<tr><td colspan="6" style="padding:5px;border:1px solid #BBB;text-align:right;font-weight:bold;font-size:9px;background:#EFF0F2">TOTAL JAM LEMBUR DISETUJUI (' + setuju.length + ' dari ' + rows.length + ' karyawan)</td>' +
    '<td style="padding:5px;border:1px solid #BBB;text-align:center;font-weight:bold;font-size:9px;background:#EFF0F2">' + String(totJam).replace('.', ',') + '</td><td style="padding:5px;border:1px solid #BBB;text-align:right;font-weight:bold;font-size:9px;background:#EFF0F2;white-space:nowrap">' + rpH_(totKomp) + '</td><td colspan="3" style="border:1px solid #BBB;background:#EFF0F2"></td></tr></table>' +
    '<p style="font-size:9px;color:#555;margin-top:10px">Konfirmasi karyawan dan persetujuan atasan dilakukan secara elektronik melalui Sistem SPL ' + escH_(P.NAMA_PERUSAHAAN || '') +
    '. Dokumen dibuat otomatis pada ' + waktuIndo_(new Date()) + '.</p>' + ttdPdf_(r0['Dibuat Oleh'], 'Admin') + '</body></html>';
  return pdfDari_(html, 'SPL ' + String(r0['Nomor SPL']).replace(/\s*\/\s*/g, '-') + ' (' + Utilities.formatDate(r0['Tanggal Lembur'], TZ, 'dd-MM-yyyy') + ').pdf');
}
function ttdPdf_(nama, jabatan) {
  return '<table style="width:100%;margin-top:18px;font-size:10px;text-align:center"><tr><td>Dibuat oleh,</td><td>Diketahui oleh,</td><td>Disetujui oleh,</td></tr>' +
    '<tr><td style="height:55px"></td><td></td><td></td></tr>' +
    '<tr><td>( ' + escH_(nama) + ' )</td><td>(.........................)</td><td>(.........................)</td></tr>' +
    '<tr><td>' + escH_(jabatan) + '</td><td>Manager Departemen</td><td>HRD</td></tr></table>';
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
  const html = '<html><head><meta charset="utf-8"><style>@page{size:A4;margin:12mm}body{font-family:Arial,sans-serif;font-size:10px;color:#222}</style></head><body>' +
    kopPdf_('SURAT PERINTAH DINAS (SPD)', I['Nomor SPD']) +
    sek('A. REFERENSI SURAT') +
    kv([['Nomor Surat', I['No Surat Referensi'] || '-', 'Tanggal Surat', tglIndo_(I['Tanggal Surat'])], ['Nama Customer', I['Nama Customer'] || '-', 'Dasar Penugasan', I['Dasar Penugasan'] || '-']]) +
    sek('B. DETAIL PERJALANAN DINAS') +
    kv([['Maksud / Tujuan', I['Maksud / Tujuan'], 'Tempat Tujuan', I['Tempat Tujuan']],
        ['Tgl. Berangkat', hariIndo_(tglB) + ', ' + tglIndo_(tglB), 'Tgl. Kembali', hariIndo_(tglK) + ', ' + tglIndo_(tglK)],
        ['Lama Dinas', I['Lama (hari)'] + ' hari', 'Transportasi', I['Transportasi']],
        ['Sumber Biaya', I['Sumber Biaya'], 'Uang Saku / Hari', rpH_(P.UANG_SAKU_HARI)]]) +
    sek('C. PERSONEL YANG DITUGASKAN') +
    '<table style="border-collapse:collapse;width:100%"><tr>' + ['No', 'Employee ID', 'Full Name', 'Job Position', 'Lama (hari)', 'Uang Saku', 'Keterangan', 'Status', 'Persetujuan Atasan'].map(thP_).join('') + '</tr>' +
    rows.map((r, i) => '<tr>' + tdP_(i + 1, 'text-align:center') + tdP_(escH_(r['Employee ID'])) + tdP_('<b>' + escH_(r['Full Name']) + '</b>') + tdP_(escH_(r['Job Position'])) +
      tdP_(r['Lama (hari)'], 'text-align:center') + tdP_(rpH_(r['Uang Saku']), 'text-align:right') + tdP_(escH_(r['Keterangan'])) + tdP_(selStatus_(r)) + tdP_(selAtasan_(r)) + '</tr>').join('') +
    '<tr><td colspan="5" style="padding:5px;border:1px solid #BBB;text-align:right;font-weight:bold;font-size:9px;background:#EFF0F2">TOTAL UANG SAKU PERSONEL DISETUJUI (' + setuju.length + ' dari ' + rows.length + ' orang)</td>' +
    '<td style="padding:5px;border:1px solid #BBB;text-align:right;font-weight:bold;font-size:9px;background:#EFF0F2">' + rpH_(totSaku) + '</td><td colspan="3" style="border:1px solid #BBB;background:#EFF0F2"></td></tr></table>' +
    sek('D. BIAYA PERJALANAN DINAS LAINNYA') +
    (biaya.length ? '<table style="border-collapse:collapse;width:100%"><tr>' + ['No', 'Jenis Biaya', 'Uraian', 'Jumlah (Rp)', 'Bukti'].map(thP_).join('') + '</tr>' +
      biaya.map((b, i) => '<tr>' + tdP_(i + 1, 'text-align:center') + tdP_(escH_(b['Jenis Biaya'])) + tdP_(escH_(b['Uraian'])) + tdP_(rpH_(b['Jumlah (Rp)']), 'text-align:right') + tdP_(escH_(b['Bukti']), 'text-align:center') + '</tr>').join('') +
      '<tr><td colspan="3" style="padding:5px;border:1px solid #BBB;text-align:right;font-weight:bold;font-size:9px;background:#EFF0F2">TOTAL BIAYA LAINNYA</td><td style="padding:5px;border:1px solid #BBB;text-align:right;font-weight:bold;font-size:9px;background:#EFF0F2">' + rpH_(totLain) + '</td><td style="border:1px solid #BBB;background:#EFF0F2"></td></tr></table>'
      : '<div style="font-size:9px;color:#777">Tidak ada biaya lainnya.</div>') +
    '<table style="width:100%;margin-top:6px;border-collapse:collapse"><tr><td style="padding:6px;border:2px solid #BE353F;font-weight:bold;font-size:11px">TOTAL BIAYA PERJALANAN DINAS (Uang Saku disetujui + Biaya Lainnya)</td>' +
    '<td style="padding:6px;border:2px solid #BE353F;font-weight:bold;font-size:11px;text-align:right;width:130px;color:#BE353F">' + rpH_(totSaku + totLain) + '</td></tr></table>' +
    sek('E. KETENTUAN') +
    '<ol style="font-size:9px;margin:4px 0 0 14px;padding:0;line-height:1.5">' +
    '<li>Karyawan yang tercantum di atas ditugaskan melaksanakan perjalanan dinas sesuai maksud, tempat, dan tanggal yang tercantum dalam surat ini.</li>' +
    '<li>Selama tanggal dinas luar tidak dapat diajukan lembur; pengajuan lembur pada tanggal tersebut akan ditolak.</li>' +
    '<li>Uang saku diberikan per orang per hari dinas (' + rpH_(P.UANG_SAKU_HARI) + '/hari).</li>' +
    '<li>Mintakan tanda tangan & stempel pada kolom Pengesahan di tempat tujuan sebagai bukti kehadiran.</li>' +
    '<li>Laporan perjalanan dinas beserta bukti pengeluaran diserahkan ke Admin paling lambat 3 hari kerja setelah kembali.</li></ol>' +
    sek('F. PENGESAHAN DI TEMPAT TUJUAN') +
    '<table style="width:100%;border-collapse:collapse;font-size:9px"><tr>' + ['TIBA DI TEMPAT TUJUAN', 'BERANGKAT DARI TEMPAT TUJUAN'].map(t =>
      '<td style="width:50%;border:1px solid #BBB;padding:6px;vertical-align:top"><b>' + t + '</b><br><br>Tempat : ............................<br><br>Tanggal : ............................<br><br>Jam : ............................' +
      '<div style="height:50px"></div><div style="text-align:center;color:#555">( Nama, Tanda Tangan & Stempel Instansi )</div></td>').join('') + '</tr></table>' +
    '<p style="text-align:right;font-size:10px;margin-top:12px">Dikeluarkan di : ' + (P.KOTA_SURAT ? escH_(P.KOTA_SURAT) + ', ' : '') + tglIndo_(I['Tanggal Surat']) + '</p>' +
    ttdPdf_(I['Dibuat Oleh'], 'Admin').replace('Manager Departemen', 'Manager Departemen') +
    '<p style="font-size:8px;color:#777;margin-top:8px">Konfirmasi personel dan persetujuan atasan dilakukan secara elektronik melalui Sistem SPL & SPD. Dokumen dibuat otomatis pada ' + waktuIndo_(new Date()) + '.</p>' +
    '</body></html>';
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
function menuKirimUlangKonfirmasi() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('Kirim ulang email konfirmasi', 'Masukkan Nomor SPL / SPD. Email dikirim ulang ke yang statusnya masih MENUNGGU KONFIRMASI.', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const n = r.getResponseText().trim();
  kirimKonfirmasi_(jenisDariNomor_(n), n);
  ui.alert('Selesai. Lihat kolom "Status Email" di sheet ' + (jenisDariNomor_(n) === 'SPD' ? 'Personel SPD' : 'Data SPL') + '.');
}
