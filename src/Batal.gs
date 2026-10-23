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
  return 'Lembur ' + kena.length + ' karyawan pada ' + escH_(tgl) + ' telah <b>DIBATALKAN</b>. Karyawan dan admin sudah diberi tahu melalui email.';
}
