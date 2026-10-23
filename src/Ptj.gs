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
  sesiCek_(token);
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
      if (!(lama >= 0 && lama === Math.floor(lama))) err.push(ref.nama + ': lama aktual tidak valid.');
      else if (lama > lamaAktual) err.push(ref.nama + ': lama aktual melebihi lama dinas (' + lamaAktual + ' hari).');
      return { ref: ref, lama: lama, keterangan: String(p.keterangan || '').trim(), uangSaku: (lama || 0) * Number(P.UANG_SAKU_HARI) };
    }).filter(Boolean);
    Object.keys(disetujui).forEach(id => { if (!per.some(p => p.ref.id === id)) err.push(disetujui[id].nama + ' belum diisi lama aktualnya.'); });
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

    const isi = '<p>Pertanggungjawaban <b>Surat Perintah Dinas ' + escH_(S.nomor) + '</b> (' + escH_(S.tempat) + ') telah diselesaikan oleh ' + escH_(u.nama) + '.</p>' +
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
  const html = '<html><head><meta charset="utf-8"><style>@page{size:A4;margin:12mm}body{font-family:Arial,sans-serif;font-size:10px;color:#222}</style></head><body>' +
    kopPdf_('LAPORAN PERTANGGUNGJAWABAN PERJALANAN DINAS', S.nomor) +
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
    '<p style="font-size:8px;color:#777;margin-top:8px">Dokumen dibuat otomatis pada ' + waktuIndo_(new Date()) + '.</p></body></html>';
  return pdfDari_(html, 'Laporan Dinas ' + String(S.nomor).replace(/\s*\/\s*/g, '-') + '.pdf');
}
