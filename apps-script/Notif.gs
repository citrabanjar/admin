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
