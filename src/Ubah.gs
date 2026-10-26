/**
 * =====================================================================
 *  SISTEM SPL & SPD — File : Ubah.gs
 *  Perubahan Surat Perintah Dinas yang masih berjalan:
 *   1. PERPANJANGAN — admin mengajukan tanggal kembali baru untuk personel tertentu.
 *      Persetujuan oleh atasan penyetuju dinas (1 email + 1 tombol per atasan),
 *      owner mendapat pemberitahuan (mengetahui). Setelah semua atasan memproses →
 *      lama & uang saku personel diperbarui, PDF surat direvisi.
 *   2. TAMBAH KAS — admin mencatat tambahan uang kas dalam perjalanan. Ketua tim
 *      (personel disetujui No. 1) menerima email + PDF revisi dan mengonfirmasi
 *      sudah menerima uang kas tambahan.
 *  Link email: a=pp (persetujuan perpanjangan), a=kt (konfirmasi terima kas).
 * =====================================================================
 */
const SH_UBAH = { UBAH: 'Perubahan SPD', PP: 'Perpanjangan Personel' };
const KOLOM_UBAH = {
  'Perubahan SPD': ['ID', 'Nomor SPD', 'Jenis', 'Tgl Kembali Lama', 'Tgl Kembali Baru', 'Tambah Kas', 'Alasan', 'Status',
    'Dibuat Oleh', 'Email Pembuat', 'Dibuat Pada', 'Dikonfirmasi Oleh', 'Dikonfirmasi Pada', 'Status Email'],
  'Perpanjangan Personel': ['ID Perubahan', 'Nomor SPD', 'Employee ID', 'Full Name', 'Tambahan Hari', 'Atasan', 'Email Atasan',
    'Status', 'Disetujui Oleh', 'Disetujui Pada', 'Catatan Atasan', 'ID Baris', 'Status Email']
};
const UBAH_ST = { MENUNGGU: 'MENUNGGU ATASAN', SETUJU: 'DISETUJUI', TOLAK: 'DITOLAK', KONFIRMASI: 'MENUNGGU KONFIRMASI', DITERIMA: 'DITERIMA' };

function sheetUbah_(nama) {
  const ss = ss_();
  let sh = cariSheet_(ss, nama);
  if (!sh) {
    sh = ss.insertSheet(nama);
    sh.getRange(1, 1, 1, KOLOM_UBAH[nama].length).setValues([KOLOM_UBAH[nama]]).setFontWeight('bold').setBackground('#3A3A3A').setFontColor('#FFFFFF');
    sh.setFrozenRows(1);
  } else pastikanKolom_(nama, KOLOM_UBAH[nama]);
  return sh;
}
/** Baca sheet → {sh, c, rows:[{_row, ...}]} (sheet dibuat bila belum ada) */
function bacaUbah_(nama) {
  const sh = sheetUbah_(nama), c = kolom_(nama), n = sh.getLastRow() - 1;
  const v = n > 0 ? sh.getRange(2, 1, n, sh.getLastColumn()).getValues() : [];
  const kunci = KOLOM_UBAH[nama][0], rows = [];
  v.forEach((r, i) => { if (String(r[c[kunci] - 1]).trim()) { const o = { _row: i + 2 }; Object.keys(c).forEach(h => o[h] = r[c[h] - 1]); rows.push(o); } });
  return { sh: sh, c: c, rows: rows };
}
const tglYmd_ = d => d instanceof Date ? ymd_(d) : String(d || '').slice(0, 10);
const tambahHari_ = (s, n) => ymd_(new Date(keTanggal_(s).getTime() + n * 864e5));
const selisihHari_ = (a, b) => Math.round((keTanggal_(tglYmd_(b)) - keTanggal_(tglYmd_(a))) / 864e5);

/** Riwayat perubahan 1 surat (dibaca tanpa membuat sheet baru). */
function perubahanSpd_(nomor) {
  const ss = ss_();
  if (!cariSheet_(ss, SH_UBAH.UBAH)) return [];
  const pp = cariSheet_(ss, SH_UBAH.PP) ? bacaTabel_(SH_UBAH.PP, 'ID Perubahan').filter(r => r['Nomor SPD'] === nomor) : [];
  return bacaTabel_(SH_UBAH.UBAH, 'ID').filter(r => r['Nomor SPD'] === nomor).map(r => ({
    id: String(r['ID']), jenis: r['Jenis'], status: r['Status'], alasan: r['Alasan'] || '',
    kembaliLama: tglYmd_(r['Tgl Kembali Lama']), kembaliBaru: tglYmd_(r['Tgl Kembali Baru']), tambahKas: Number(r['Tambah Kas']) || 0,
    dibuatOleh: r['Dibuat Oleh'], dibuatPada: r['Dibuat Pada'] instanceof Date ? r['Dibuat Pada'] : null,
    dikonfirmasiPada: r['Dikonfirmasi Pada'] instanceof Date ? r['Dikonfirmasi Pada'] : null,
    personel: pp.filter(p => String(p['ID Perubahan']) === String(r['ID'])).map(p => ({ id: p['Employee ID'], nama: p['Full Name'], status: p['Status'], tambahan: Number(p['Tambahan Hari']) || 0 }))
  }));
}
/** Total tambahan kas (semua catatan TAMBAH KAS) */
function tambahanKas_(nomor) { return perubahanSpd_(nomor).filter(x => x.jenis === 'TAMBAH KAS').reduce((s, x) => s + x.tambahKas, 0); }

/** Ketua tim = personel DISETUJUI pertama (urutan No.) */
function ketuaTim_(nomor) {
  const K = karyawan_(), ditunjuk = String(induk_('SPD', nomor)['Ketua Tim'] || '').toUpperCase();   // berganti bila ada serah terima kas
  const ok = bacaTabel_(SH.PERSONEL, 'Nomor SPD').filter(x => x['Nomor SPD'] === nomor && x['Status'] === ST.SETUJU);
  const r = ok.find(x => String(x['Employee ID']).toUpperCase() === ditunjuk) || ok[0];
  if (!r) return null;
  const k = K[String(r['Employee ID']).toUpperCase()] || {};
  return { id: String(r['Employee ID']).toUpperCase(), nama: r['Full Name'], email: k.email || '' };
}

function cekUbahBoleh_(u) {
  if (!punyaAkses_(u, 'DINAS') && !punyaAkses_(u, 'PTJ')) throw new Error('Anda tidak memiliki akses untuk mengubah Surat Perintah Dinas.');
}
function cekSuratBerjalan_(S) {
  if (!S) throw new Error('Surat Perintah Dinas tidak ditemukan.');
  if (S.statusDinas !== 'Aktif') throw new Error('Surat ini sudah tidak berjalan (status ' + S.statusDinas + ').');
  if (!S.selesaiApproval) throw new Error('Konfirmasi / persetujuan personel surat ini belum selesai.');
}

/* ============================== PERPANJANGAN ============================== */
/** d = {nomor, kembali:'yyyy-MM-dd', alasan, ids:[Employee ID]} */
function ajukanPerpanjangan(token, d) {
  const u = sesiCek_(token);
  cekUbahBoleh_(u);
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('Server sibuk, coba lagi beberapa detik lagi.');
  let id;
  const warn = [];
  try {
    _K = null; _P = null;
    const S = dataSpd_(String(d && d.nomor || ''));
    cekSuratBerjalan_(S);
    const err = [];
    if (S.perpanjanganPending) err.push('Masih ada perpanjangan surat ini yang menunggu persetujuan atasan.');
    if (!/^\d{4}-\d{2}-\d{2}$/.test(String(d.kembali || ''))) err.push('Tanggal kembali baru wajib diisi.');
    else if (d.kembali <= S.kembali) err.push('Tanggal kembali baru harus setelah ' + tglIndo_(keTanggal_(S.kembali)) + '.');
    const alasan = besar_(d.alasan).slice(0, 500);
    if (!alasan) err.push('Alasan perpanjangan wajib diisi.');
    const setuju = {}; S.personel.filter(p => p.status === ST.SETUJU).forEach(p => setuju[p.id] = p);
    const ids = [...new Set((d.ids || []).map(x => String(x).toUpperCase()))].filter(x => setuju[x]);
    ids.filter(x => !setuju[x].masihTugas).forEach(x => err.push(sdr_(setuju[x].nama) + ' sudah selesai bertugas (' + tglIndo_(keTanggal_(setuju[x].sampai)) + ') — tidak dapat diperpanjang.'));
    if (!ids.length) err.push('Pilih minimal 1 personel yang diperpanjang.');
    if (!err.length) {
      const aktif = suratAktif_(tambahHari_(S.kembali, 1), d.kembali);
      ids.forEach(x => {
        if (aktif.spd[x] && aktif.spd[x] !== S.nomor) err.push(sdr_(setuju[x].nama) + ' sudah ditugaskan di ' + aktif.spd[x] + ' pada tanggal perpanjangan.');
        if (aktif.spl[x]) warn.push(sdr_(setuju[x].nama) + ' punya SPL ' + aktif.spl[x] + ' di rentang perpanjangan — SPL tersebut perlu ditinjau.');
      });
    }
    if (err.length) return { ok: false, error: err };

    const tambah = selisihHari_(S.kembali, d.kembali), now = new Date(), tp = s => keTanggal_(s);
    id = 'P' + Utilities.getUuid().slice(0, 8);
    sheetUbah_(SH_UBAH.UBAH); sheetUbah_(SH_UBAH.PP);
    tulisBaris_(SH_UBAH.UBAH, [{ 'ID': id, 'Nomor SPD': S.nomor, 'Jenis': 'PERPANJANGAN', 'Tgl Kembali Lama': tp(S.kembali), 'Tgl Kembali Baru': tp(d.kembali),
      'Alasan': alasan, 'Status': UBAH_ST.MENUNGGU, 'Dibuat Oleh': u.nama, 'Email Pembuat': u.email, 'Dibuat Pada': now }]);
    tulisBaris_(SH_UBAH.PP, ids.map(x => {
      const a = penyetujuDinas_(x);
      return { 'ID Perubahan': id, 'Nomor SPD': S.nomor, 'Employee ID': x, 'Full Name': setuju[x].nama, 'Tambahan Hari': tambah,
        'Atasan': a ? a.nama : '(admin — atasan belum diisi)', 'Email Atasan': a ? a.email : u.email, 'Status': UBAH_ST.MENUNGGU, 'ID Baris': Utilities.getUuid().slice(0, 8) };
    }));
    SpreadsheetApp.flush();
    logAksi_('SPD', S.nomor, 'PERPANJANGAN DIAJUKAN', u.email, ids.length + ' personel s/d ' + d.kembali + ' (+' + tambah + ' hari) — ' + alasan);
  } finally { lock.releaseLock(); }
  try { antri_({ t: 'perpanjangan', id: id }); } catch (e) { logAksi_('SPD', d.nomor, 'EMAIL PERPANJANGAN GAGAL', '', String(e)); }
  return { ok: true, peringatan: warn };
}

/** Email persetujuan perpanjangan — 1 email per atasan, 1 tombol untuk semua personel. */
function kirimPersetujuanPerpanjangan_(idUbah) {
  const U = bacaUbah_(SH_UBAH.UBAH), r0 = U.rows.find(x => String(x['ID']) === String(idUbah));
  if (!r0) return;
  const P = bacaUbah_(SH_UBAH.PP), rows = P.rows.filter(x => String(x['ID Perubahan']) === String(idUbah) && x['Status'] === UBAH_ST.MENUNGGU);
  const I = induk_('SPD', r0['Nomor SPD']), grup = {};
  rows.forEach(x => { const k = String(x['Email Atasan']).toLowerCase(); (grup[k] = grup[k] || []).push(x); });
  Object.keys(grup).forEach(ke => {
    const g = grup[ke], idg = g.map(x => x['ID Baris']).join('.');
    const isi = '<p>Yth. <b>' + escH_(sdr_(g[0]['Atasan'])) + '</b>,</p><p>Admin mengajukan <b>perpanjangan perjalanan dinas</b> untuk karyawan di bawah kelolaan Anda. Mohon persetujuan Anda:</p>' +
      htmlPerpanjangan_(I, r0, g) +
      '<p style="margin-top:18px">' + tombol_(linkAksi_('pp', idg, 'ya'), '✅ Setujui', '#1E7B4A') + tombol_(linkAksi_('pp', idg, 'tidak'), '❌ Tolak', '#B42318') + '</p>';
    const err = kirimEmail_(ke, 'Persetujuan Perpanjangan Dinas — ' + (g.length > 1 ? g.length + ' personel' : sdr_(g[0]['Full Name'])) + ' (' + I['Tempat Tujuan'] + ')', isi);
    g.forEach(x => ubahBaris_(P, x._row, { 'Status Email': err ? 'GAGAL: ' + err : 'Persetujuan terkirim ke ' + ke + ' ' + Utilities.formatDate(new Date(), TZ, 'dd/MM HH:mm') }));
    logAksi_('SPD', r0['Nomor SPD'], err ? 'EMAIL PERPANJANGAN GAGAL' : 'EMAIL PERPANJANGAN TERKIRIM', ke, g.map(x => x['Full Name']).join(', ') + (err ? ' — ' + err : ''));
  });
}
function htmlPerpanjangan_(I, r0, g) {
  const th = 'padding:6px 8px;background:#F4F5F7;border:1px solid #E2E4E8;text-align:left', td = 'padding:6px 8px;border:1px solid #E2E4E8;vertical-align:top';
  const baris = [['Nomor SPD', I['Nomor SPD']], ['Maksud / Tujuan', I['Maksud / Tujuan']], ['Tempat Tujuan', I['Tempat Tujuan']],
    ['Tgl. Berangkat', hariIndo_(I['Tgl Berangkat']) + ', ' + tglIndo_(I['Tgl Berangkat'])],
    ['Tgl. Kembali Semula', hariIndo_(r0['Tgl Kembali Lama']) + ', ' + tglIndo_(r0['Tgl Kembali Lama'])],
    ['Diperpanjang s/d', hariIndo_(r0['Tgl Kembali Baru']) + ', ' + tglIndo_(r0['Tgl Kembali Baru']) + ' (+' + selisihHari_(r0['Tgl Kembali Lama'], r0['Tgl Kembali Baru']) + ' hari)'],
    ['Alasan', r0['Alasan']], ['Diajukan oleh', sdr_(r0['Dibuat Oleh'])]];
  return '<table style="border-collapse:collapse;width:100%;font-size:13px">' + baris.map(b => '<tr><td style="' + th + ';width:150px;font-weight:bold">' + escH_(b[0]) +
      '</td><td style="' + td + '">' + escH_(b[1]) + '</td></tr>').join('') + '</table>' +
    '<table style="border-collapse:collapse;width:100%;font-size:13px;margin-top:12px"><tr><th style="' + th + '">No</th><th style="' + th + '">Nama Karyawan</th><th style="' + th + '">Tambahan</th></tr>' +
    g.map((x, i) => '<tr><td style="' + td + ';text-align:center">' + (i + 1) + '</td><td style="' + td + '"><b>' + escH_(sdr_(x['Full Name'])) + '</b><div style="font-size:11px;color:#6B6F76">' + escH_(x['Employee ID']) + '</div></td>' +
      '<td style="' + td + ';white-space:nowrap">+' + escH_(x['Tambahan Hari']) + ' hari</td></tr>').join('') + '</table>';
}
function barisPerpanjangan_(id) {
  const ids = String(id || '').split('.').filter(Boolean), P = bacaUbah_(SH_UBAH.PP);
  return { P: P, rows: ids.map(i => P.rows.find(x => String(x['ID Baris']) === i)).filter(Boolean) };
}
function halamanPerpanjangan_(p) {
  let isi;
  try {
    if (!p.id || p.s !== tandaTangan_('pp', p.id)) throw new Error('Link tidak valid atau sudah diubah.');
    const G = barisPerpanjangan_(p.id), x0 = G.rows[0];
    if (!x0) throw new Error('Data tidak ditemukan (mungkin sudah dihapus).');
    const r0 = bacaUbah_(SH_UBAH.UBAH).rows.find(x => String(x['ID']) === String(x0['ID Perubahan'])) || {};
    const I = induk_('SPD', x0['Nomor SPD']), aktif = G.rows.filter(x => x['Status'] === UBAH_ST.MENUNGGU);
    isi = '<h2>Persetujuan Perpanjangan Dinas</h2>' + htmlPerpanjangan_(I, r0, G.rows).replace(/width:150px/g, 'width:130px') +
      (aktif.length ? '' : '<div class="msg info">Sudah diproses — status: <b>' + escH_(G.rows.map(x => sdr_(x['Full Name']) + ' ' + x['Status']).join(', ')) + '</b>.</div>');
    if (aktif.length) isi += formPilih_(p, 'Setujui', 'Tolak', 'prosesPerpanjangan');
  } catch (e) { isi = '<h2>Tidak dapat dibuka</h2><div class="msg err">' + escH_(e.message) + '</div>'; }
  return halamanHtml_(isi);
}
/** Potongan HTML pilihan ya/tidak + catatan + tombol Kirim → memanggil fungsi server fn(id, sig, jawaban, catatan) */
function formPilih_(p, ya, tidak, fn) {
  const pil = p.p === 'tidak' ? 'tidak' : 'ya';
  return '<div class="pilih">' +
    '<label class="op ya"><input type="radio" name="j" value="ya"' + (pil === 'ya' ? ' checked' : '') + '> ✅ ' + ya + '</label>' +
    '<label class="op tidak"><input type="radio" name="j" value="tidak"' + (pil === 'tidak' ? ' checked' : '') + '> ❌ ' + tidak + '</label></div>' +
    '<div id="al"><label>Catatan <span class="sub">(opsional)</span></label><textarea id="alasan" rows="3"></textarea></div>' +
    '<button id="b" onclick="kirim()">Kirim</button><div id="o"></div>' +
    '<script>function kirim(){var b=document.getElementById("b"),o=document.getElementById("o");b.disabled=true;b.textContent="Mengirim…";' +
    'google.script.run.withSuccessHandler(function(m){document.querySelector(".pilih").remove();document.getElementById("al").remove();b.remove();o.innerHTML="<div class=\\"msg ok\\">"+m+"</div>"})' +
    '.withFailureHandler(function(e){b.disabled=false;b.textContent="Kirim";o.innerHTML="<div class=\\"msg err\\">"+e.message+"</div>"})' +
    '.' + fn + '(' + JSON.stringify(String(p.id)) + ',' + JSON.stringify(p.s) + ',document.querySelector("input[name=j]:checked").value,document.getElementById("alasan").value)}</script>';
}
function prosesPerpanjangan(id, sig, jawaban, catatan) {
  if (!id || sig !== tandaTangan_('pp', id)) throw new Error('Link tidak valid.');
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('Server sibuk, coba lagi.');
  const setuju = jawaban !== 'tidak', alas = String(catatan || '').trim().slice(0, 500), now = new Date();
  let kena, idUbah;
  try {
    _K = null; _P = null;
    const G = barisPerpanjangan_(id);
    if (!G.rows.length) throw new Error('Data tidak ditemukan.');
    idUbah = G.rows[0]['ID Perubahan'];
    kena = G.rows.filter(x => x['Status'] === UBAH_ST.MENUNGGU);
    if (!kena.length) return 'Sudah diproses sebelumnya.';
    kena.forEach(x => ubahBaris_(G.P, x._row, { 'Status': setuju ? UBAH_ST.SETUJU : UBAH_ST.TOLAK, 'Disetujui Oleh': x['Atasan'], 'Disetujui Pada': now, 'Catatan Atasan': alas }));
    SpreadsheetApp.flush();
    logAksi_('SPD', kena[0]['Nomor SPD'], setuju ? 'PERPANJANGAN DISETUJUI' : 'PERPANJANGAN DITOLAK', kena[0]['Email Atasan'], kena.map(x => x['Full Name']).join(', ') + (alas ? ' — ' + alas : ''));
  } finally { lock.releaseLock(); }
  try { antri_({ t: 'perpanjanganCek', id: idUbah }); } catch (e) { logAksi_('SPD', '-', 'PERPANJANGAN GAGAL', '', String(e)); }
  return 'Terima kasih. Perpanjangan dinas ' + kena.map(x => escH_(sdr_(x['Full Name']))).join(', ') + ' telah <b>' + (setuju ? 'DISETUJUI' : 'DITOLAK') + '</b>.';
}

/** Dijalankan dari antrian setelah atasan memproses: bila semua selesai → terapkan & kirim PDF revisi. */
function selesaiPerpanjangan_(idUbah) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('Server sibuk.');
  let r0, ok, rows;
  try {
    _K = null; _P = null;
    const U = bacaUbah_(SH_UBAH.UBAH);
    r0 = U.rows.find(x => String(x['ID']) === String(idUbah));
    if (!r0 || [UBAH_ST.SETUJU, UBAH_ST.TOLAK].indexOf(r0['Status']) >= 0) return;
    rows = bacaUbah_(SH_UBAH.PP).rows.filter(x => String(x['ID Perubahan']) === String(idUbah));
    const belum = rows.filter(x => x['Status'] === UBAH_ST.MENUNGGU);
    if (belum.length) { ubahBaris_(U, r0._row, { 'Status': 'PROSES ' + (rows.length - belum.length) + '/' + rows.length }); return; }
    ok = rows.filter(x => x['Status'] === UBAH_ST.SETUJU);
    ubahBaris_(U, r0._row, { 'Status': ok.length ? UBAH_ST.SETUJU : UBAH_ST.TOLAK });
    if (ok.length) {
      const nomor = r0['Nomor SPD'], P = param_(), S = bacaBaris_('SPD'), set = {};
      ok.forEach(x => set[String(x['Employee ID']).toUpperCase()] = Number(x['Tambahan Hari']) || 0);
      S.rows.filter(x => x['Nomor SPD'] === nomor && set[String(x['Employee ID']).toUpperCase()] != null).forEach(x => {
        const lama = (Number(x['Lama (hari)']) || 0) + set[String(x['Employee ID']).toUpperCase()];
        ubahBaris_(S, x._row, { 'Lama (hari)': lama, 'Uang Saku': lama * Number(P.UANG_SAKU_HARI) });
        x['Uang Saku'] = lama * Number(P.UANG_SAKU_HARI);
      });
      const Ind = bacaBaris_SPD_induk_(), I = Ind.rows.find(x => x['Nomor SPD'] === nomor);
      if (I) {
        const baru = r0['Tgl Kembali Baru'] > I['Tgl Kembali'] ? r0['Tgl Kembali Baru'] : I['Tgl Kembali'];
        const totSaku = S.rows.filter(x => x['Nomor SPD'] === nomor && [ST.TIDAK, ST.TOLAK, ST.BATAL].indexOf(x['Status']) < 0).reduce((s, x) => s + (Number(x['Uang Saku']) || 0), 0);
        ubahBaris_(Ind, I._row, { 'Tgl Kembali': baru, 'Lama (hari)': selisihHari_(I['Tgl Berangkat'], baru) + 1, 'Total Uang Saku': totSaku });
      }
    }
    SpreadsheetApp.flush();
  } finally { lock.releaseLock(); }
  const nomor = r0['Nomor SPD'], tolak = rows.filter(x => x['Status'] === UBAH_ST.TOLAK);
  const ringkas = '<p>Perpanjangan perjalanan dinas s/d <b>' + escH_(hariIndo_(r0['Tgl Kembali Baru']) + ', ' + tglIndo_(r0['Tgl Kembali Baru'])) + '</b> — alasan: ' + escH_(r0['Alasan']) + '.</p>' +
    (ok.length ? '<p><b style="color:#1E7B4A">Disetujui</b>: ' + ok.map(x => escH_(sdr_(x['Full Name']))).join(', ') + '</p>' : '') +
    (tolak.length ? '<p><b style="color:#B42318">Ditolak</b>: ' + tolak.map(x => escH_(sdr_(x['Full Name'])) + (x['Catatan Atasan'] ? ' (' + escH_(x['Catatan Atasan']) + ')' : '')).join(', ') + '</p>' : '');
  if (!ok.length) {
    kirimEmail_(r0['Email Pembuat'], 'Perpanjangan Dinas Ditolak — ' + nomor, ringkas);
    tolakKePersonel_(nomor, r0, tolak);
    logAksi_('SPD', nomor, 'PERPANJANGAN SELESAI', '', 'ditolak seluruhnya');
    return;
  }
  kirimRevisiSpd_(nomor, 'Perpanjangan Dinas Disetujui', ringkas, ok.map(x => String(x['Employee ID']).toUpperCase()), true, r0['Email Pembuat']);
  tolakKePersonel_(nomor, r0, tolak);
}
/** Personel yang perpanjangannya ditolak: kembali sesuai jadwal semula. */
function tolakKePersonel_(nomor, r0, tolak) {
  const K = karyawan_();
  tolak.forEach(x => {
    const k = K[String(x['Employee ID']).toUpperCase()]; if (!k || !k.email) return;
    kirimEmail_(k.email, 'Perpanjangan Dinas Tidak Disetujui — ' + nomor, '<p>Yth. <b>' + escH_(sdr_(k.nama)) + '</b>,</p><p>Perpanjangan perjalanan dinas Anda s/d ' +
      escH_(tglIndo_(r0['Tgl Kembali Baru'])) + ' <b style="color:#B42318">TIDAK DISETUJUI</b>' + (x['Catatan Atasan'] ? ' — catatan: ' + escH_(x['Catatan Atasan']) : '') +
      '.</p><p>Mohon kembali sesuai jadwal semula, <b>' + escH_(hariIndo_(r0['Tgl Kembali Lama']) + ', ' + tglIndo_(r0['Tgl Kembali Lama'])) + '</b>.</p>');
  });
}

/** Buat PDF Surat Perintah Dinas revisi, simpan, dan kirim:
 *  ketua tim (PDF terlampir) · personel terkait lainnya (tanpa lampiran) · admin pembuat (PDF) · owner (bila keOwner, mengetahui).
 *  Mengembalikan {file, pdf, ketua}. */
function kirimRevisiSpd_(nomor, judul, ringkasHtml, idsTerkait, keOwner, emailPembuat, tanpaKetua) {
  const Ind = bacaBaris_SPD_induk_(), I = Ind.rows.find(x => x['Nomor SPD'] === nomor);
  if (!I) return null;
  const rows = bacaBaris_('SPD').rows.filter(x => x['Nomor SPD'] === nomor);
  const pdf = buatPdfSPD_(I, rows), file = folderPdf_(JENIS.SPD.folder).createFile(pdf);
  hapusPdfLama_(I['Link PDF']);
  ubahBaris_(Ind, I._row, { 'Link PDF': file.getUrl() });
  const K = karyawan_(), ketua = ketuaTim_(nomor), tgl = ringkasTanggal_('SPD', rows[0] || {}, I);
  const detail = tabelDetail_('SPD', rows[0] || {}, I, false, true);
  if (ketua && ketua.email && !tanpaKetua) {
    kirimEmail_(ketua.email, judul + ' — ' + nomor + ' (' + tgl + ')', '<p>Yth. <b>' + escH_(sdr_(ketua.nama)) + '</b>,</p>' + ringkasHtml +
      '<p>Terlampir Surat Perintah Dinas yang telah direvisi.</p>' + detail, '', pdf);
  }
  (idsTerkait || []).filter(x => !ketua || x !== ketua.id).forEach(x => {
    const k = K[x]; if (!k || !k.email) return;
    kirimEmail_(k.email, judul + ' — ' + nomor + ' (' + tgl + ')', '<p>Yth. <b>' + escH_(sdr_(k.nama)) + '</b>,</p>' + ringkasHtml +
      (ketua ? '<p>Surat Perintah Dinas revisi dikirimkan kepada <b>' + escH_(sdr_(ketua.nama)) + '</b>.</p>' : '') + detail);
  });
  const admin = String(param_().EMAIL_ADMIN || '').trim().toLowerCase(), pembuat = String(emailPembuat || I['Email Pembuat']).toLowerCase();
  kirimEmail_(pembuat, judul + ' — ' + nomor + ' (' + tgl + ')', ringkasHtml + '<p>PDF Surat Perintah Dinas revisi terlampir: <a href="' + file.getUrl() + '">' + escH_(file.getName()) + '</a></p>',
    admin && admin !== pembuat ? admin : '', pdf);
  if (keOwner) owner_().forEach(o => kirimEmail_(o.email, 'Pemberitahuan: ' + judul + ' — ' + nomor,
    '<p>Yth. <b>' + escH_(sdr_(o.nama)) + '</b>,</p><p>Sebagai pemberitahuan (mengetahui):</p>' + ringkasHtml + detail, '', pdf));
  logAksi_('SPD', nomor, 'PDF REVISI', ketua ? ketua.email : '', judul + ' — ' + file.getUrl());
  return { file: file, pdf: pdf, ketua: ketua };
}

/* ============================== TAMBAH KAS ============================== */
/** d = {nomor, nominal, alasan} */
function tambahKasSPD(token, d) {
  const u = sesiCek_(token);
  cekUbahBoleh_(u);
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('Server sibuk, coba lagi beberapa detik lagi.');
  let id, S;
  try {
    _K = null; _P = null;
    S = dataSpd_(String(d && d.nomor || ''));
    cekSuratBerjalan_(S);
    const nominal = Math.round(Number(d.nominal)), alasan = besar_(d.alasan).slice(0, 500), err = [];
    if (!(nominal > 0)) err.push('Nominal tambahan kas wajib diisi.');
    if (!alasan) err.push('Alasan / keperluan tambahan kas wajib diisi.');
    if (!S.ketua) err.push('Belum ada personel yang disetujui pada surat ini.');
    if (err.length) return { ok: false, error: err };
    id = 'K' + Utilities.getUuid().slice(0, 8);
    sheetUbah_(SH_UBAH.UBAH);
    tulisBaris_(SH_UBAH.UBAH, [{ 'ID': id, 'Nomor SPD': S.nomor, 'Jenis': 'TAMBAH KAS', 'Tambah Kas': nominal, 'Alasan': alasan,
      'Status': UBAH_ST.KONFIRMASI, 'Dibuat Oleh': u.nama, 'Email Pembuat': u.email, 'Dibuat Pada': new Date() }]);
    pastikanKolom_(SH.SPD, ['Tambahan Kas']);
    const Ind = bacaBaris_SPD_induk_(), I = Ind.rows.find(x => x['Nomor SPD'] === S.nomor);
    if (I) ubahBaris_(Ind, I._row, { 'Tambahan Kas': (Number(I['Tambahan Kas']) || 0) + nominal });
    SpreadsheetApp.flush();
    logAksi_('SPD', S.nomor, 'TAMBAH KAS', u.email, rpH_(nominal) + ' — ' + alasan);
  } finally { lock.releaseLock(); }
  try { antri_({ t: 'tambahKas', id: id }); } catch (e) { logAksi_('SPD', S.nomor, 'EMAIL TAMBAH KAS GAGAL', '', String(e)); }
  return { ok: true, ketua: S.ketua ? sdr_(S.ketua.nama) : '' };
}
/** Email ke ketua tim: PDF revisi + tombol konfirmasi sudah menerima. */
function kirimTambahKas_(id) {
  const U = bacaUbah_(SH_UBAH.UBAH), r = U.rows.find(x => String(x['ID']) === String(id));
  if (!r) return;
  const nomor = r['Nomor SPD'], total = (Number(induk_('SPD', nomor)['Uang Kas Perjalanan']) || 0) + tambahanKas_(nomor);
  const ringkas = '<p>Uang kas dalam perjalanan ditambah sebesar <b>' + rpH_(r['Tambah Kas']) + '</b> — keperluan: ' + escH_(r['Alasan']) + '.<br>' +
    'Total uang kas dalam perjalanan menjadi <b>' + rpH_(total) + '</b>.</p>';
  const R = kirimRevisiSpd_(nomor, 'Tambahan Uang Kas Dinas', ringkas, [], false, r['Email Pembuat'], true);
  const ketua = R && R.ketua;
  if (!ketua || !ketua.email) { ubahBaris_(U, r._row, { 'Status Email': 'GAGAL: ketua tim tanpa email' }); return; }
  const I = induk_('SPD', nomor);
  const err = kirimEmail_(ketua.email, 'Konfirmasi Tambahan Uang Kas Dinas — ' + nomor + ' (' + I['Tempat Tujuan'] + ')',
    '<p>Yth. <b>' + escH_(sdr_(ketua.nama)) + '</b>,</p>' + ringkas + '<p>Mohon konfirmasi bahwa Anda sudah menerima uang kas tambahan tersebut:</p>' +
    '<p style="margin-top:14px">' + tombol_(linkAksi_('kt', String(id)), '✅ Sudah Diterima', '#1E7B4A') + '</p>' +
    '<p>Terlampir Surat Perintah Dinas yang telah direvisi.</p>' + tabelDetail_('SPD', {}, I, false, true), '', R.pdf);
  ubahBaris_(U, r._row, { 'Status Email': err ? 'GAGAL: ' + err : 'Konfirmasi terkirim ke ' + ketua.email + ' ' + Utilities.formatDate(new Date(), TZ, 'dd/MM HH:mm') });
}
function halamanTerimaKas_(p) {
  let isi;
  try {
    if (!p.id || p.s !== tandaTangan_('kt', p.id)) throw new Error('Link tidak valid atau sudah diubah.');
    const r = bacaUbah_(SH_UBAH.UBAH).rows.find(x => String(x['ID']) === String(p.id));
    if (!r) throw new Error('Data tidak ditemukan.');
    const I = induk_('SPD', r['Nomor SPD']);
    const serah = r['Jenis'] === 'SERAH KAS';
    isi = '<h2>Konfirmasi ' + (serah ? 'Serah Terima' : 'Tambahan') + ' Uang Kas</h2><p><b>' + escH_(r['Nomor SPD']) + '</b><br>' + escH_(I['Tempat Tujuan']) + '</p>' +
      '<p>' + (serah ? 'Uang kas yang diterima' : 'Tambahan uang kas dalam perjalanan') + ': <b>' + rpH_(r['Tambah Kas']) + '</b><br>' + (serah ? '' : 'Keperluan: ') + escH_(r['Alasan']) + '</p>';
    if (r['Status'] === UBAH_ST.DITERIMA) isi += '<div class="msg info">Sudah dikonfirmasi diterima' + (r['Dikonfirmasi Pada'] instanceof Date ? ' pada ' + waktuIndo_(r['Dikonfirmasi Pada']) : '') + '.</div>';
    else isi += '<button id="b" onclick="kirim()">✅ Saya sudah menerima uang kas' + (serah ? '' : ' tambahan') + '</button><div id="o"></div>' +
      '<script>function kirim(){var b=document.getElementById("b"),o=document.getElementById("o");b.disabled=true;b.textContent="Mengirim…";' +
      'google.script.run.withSuccessHandler(function(m){b.remove();o.innerHTML="<div class=\\"msg ok\\">"+m+"</div>"})' +
      '.withFailureHandler(function(e){b.disabled=false;b.textContent="✅ Saya sudah menerima uang kas tambahan";o.innerHTML="<div class=\\"msg err\\">"+e.message+"</div>"})' +
      '.prosesTerimaKas(' + JSON.stringify(String(p.id)) + ',' + JSON.stringify(p.s) + ')}</script>';
  } catch (e) { isi = '<h2>Tidak dapat dibuka</h2><div class="msg err">' + escH_(e.message) + '</div>'; }
  return halamanHtml_(isi);
}
function prosesTerimaKas(id, sig) {
  if (!id || sig !== tandaTangan_('kt', id)) throw new Error('Link tidak valid.');
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('Server sibuk, coba lagi.');
  let r, ketua;
  try {
    _K = null;
    const U = bacaUbah_(SH_UBAH.UBAH);
    r = U.rows.find(x => String(x['ID']) === String(id));
    if (!r) throw new Error('Data tidak ditemukan.');
    if (r['Status'] === UBAH_ST.DITERIMA) return 'Sudah dikonfirmasi sebelumnya.';
    ketua = ketuaTim_(r['Nomor SPD']);
    ubahBaris_(U, r._row, { 'Status': UBAH_ST.DITERIMA, 'Dikonfirmasi Oleh': ketua ? ketua.nama : '', 'Dikonfirmasi Pada': new Date() });
    SpreadsheetApp.flush();
    logAksi_('SPD', r['Nomor SPD'], 'TAMBAH KAS DITERIMA', ketua ? ketua.email : '', rpH_(r['Tambah Kas']));
  } finally { lock.releaseLock(); }
  kirimEmail_(r['Email Pembuat'], (r['Jenis'] === 'SERAH KAS' ? 'Serah Terima' : 'Tambahan') + ' Uang Kas Diterima — ' + r['Nomor SPD'],
    '<p><b>' + escH_(sdr_(ketua ? ketua.nama : '-')) + '</b> telah mengonfirmasi <b style="color:#1E7B4A">SUDAH MENERIMA</b> ' + (r['Jenis'] === 'SERAH KAS' ? 'serah terima' : 'tambahan') + ' uang kas dalam perjalanan sebesar <b>' +
    rpH_(r['Tambah Kas']) + '</b> untuk Surat Perintah Dinas ' + escH_(r['Nomor SPD']) + '.</p><p>' + escH_(r['Alasan']) + '</p>');
  return 'Terima kasih. Penerimaan uang kas ' + rpH_(r['Tambah Kas']) + ' telah dikonfirmasi.';
}

/* ============================== KEKURANGAN KAS ============================== */
/** Setelah kekurangan kas dipotong dari gaji ketua tim (manual), admin mencatat potongannya → surat SELESAI.
 *  d = {nomor, nominal, keterangan} */
function selesaikanKurangKas(token, d) {
  const u = sesiCek_(token);
  cekAkses_(u, 'PTJ');
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('Server sibuk, coba lagi beberapa detik lagi.');
  let I;
  const nominal = Math.round(Number(d && d.nominal)), ket = besar_(d && d.keterangan).slice(0, 300);
  try {
    pastikanKolom_(SH.SPD, ['Potongan Gaji Ketua', 'Keterangan Potongan', 'Potongan Dicatat Oleh', 'Potongan Dicatat Pada']);
    const Ind = bacaBaris_SPD_induk_();
    I = Ind.rows.find(x => x['Nomor SPD'] === String(d && d.nomor || ''));
    if (!I) throw new Error('Surat Perintah Dinas tidak ditemukan.');
    if (I['Status Dinas'] !== 'Kurang Kas') throw new Error('Surat ini tidak berstatus Kurang Kas.');
    if (!(nominal > 0)) return { ok: false, error: ['Nominal potongan gaji wajib diisi.'] };
    ubahBaris_(Ind, I._row, { 'Status Dinas': 'Selesai', 'Potongan Gaji Ketua': nominal, 'Keterangan Potongan': ket, 'Potongan Dicatat Oleh': u.nama, 'Potongan Dicatat Pada': new Date() });
    SpreadsheetApp.flush();
    logAksi_('SPD', I['Nomor SPD'], 'POTONGAN GAJI KETUA', u.email, rpH_(nominal) + (ket ? ' — ' + ket : ''));
  } finally { lock.releaseLock(); }
  const ketua = ketuaTim_(I['Nomor SPD']);
  const isi = '<p>Kekurangan uang kas <b>Surat Perintah Dinas ' + escH_(I['Nomor SPD']) + '</b> (' + escH_(I['Tempat Tujuan']) + ') sebesar <b>' + rpH_(nominal) +
    '</b> telah dicatat dipotong dari gaji ketua tim <b>' + escH_(sdr_(ketua ? ketua.nama : '-')) + '</b>' + (ket ? ' — ' + escH_(ket) : '') + '.</p><p>Status surat: <b>SELESAI</b>. Dicatat oleh ' + escH_(sdr_(u.nama)) + '.</p>';
  const admin = String(param_().EMAIL_ADMIN || '').trim().toLowerCase(), pembuat = String(I['Email Pembuat'] || u.email).toLowerCase();
  kirimEmail_(pembuat, 'Kekurangan Kas Dinas Diselesaikan — ' + I['Nomor SPD'], isi, admin && admin !== pembuat ? admin : '');
  return { ok: true };
}

/* ============================== LAPOR PULANG (parsial) ============================== */
/** Personel pulang lebih dulu: uang saku dibayar kasir di kantor, langsung bisa diajukan lembur lagi.
 *  Ketua tim memegang uang kas → tidak lewat sini, melainkan lewat pertanggungjawaban akhir.
 *  d = {nomor, id, pulang:'yyyy-MM-dd', keterangan} */
function laporPulang(token, d) {
  const u = sesiCek_(token);
  cekUbahBoleh_(u);
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('Server sibuk, coba lagi beberapa detik lagi.');
  let hasil;
  try {
    _K = null; _P = null;
    const S = dataSpd_(String(d && d.nomor || ''));
    cekSuratBerjalan_(S);
    const id = String(d.id || '').toUpperCase(), p = S.personel.find(x => x.id === id && x.status === ST.SETUJU), err = [];
    if (!p) throw new Error('Personel tidak ditemukan di surat ini.');
    if (p.pulang) throw new Error(sdr_(p.nama) + ' sudah lapor pulang (' + tglIndo_(keTanggal_(p.pulang)) + ').');
    const isKetua = S.ketua && S.ketua.id === id;
    let baru = null;
    if (isKetua && !S.personel.some(x => x.status === ST.SETUJU && !x.pulang && x.id !== id))
      return { ok: false, error: [sdr_(p.nama) + ' adalah personel terakhir yang masih bertugas — laporkan kepulangannya melalui Selesaikan (pertanggungjawaban akhir).'] };
    if (isKetua) {   // ketua pulang lebih dulu → uang kas diserahkan ke personel yang masih bertugas (jadi ketua tim baru)
      baru = S.personel.find(x => x.id === String(d.ketuaBaru || '').toUpperCase() && x.status === ST.SETUJU && !x.pulang && x.id !== id);
      if (!baru) err.push('Ketua tim memegang uang kas — pilih personel yang masih bertugas sebagai penerima uang kas (ketua tim baru).');
      if (!(Number(d.kasSerah) >= 0) || d.kasSerah === '' || d.kasSerah == null) err.push('Isi sisa uang kas yang diserahkan (isi 0 bila habis).');
    }
    if (S.perpanjanganPending) err.push('Masih ada perpanjangan yang menunggu persetujuan atasan.');
    const pulang = String(d.pulang || '');
    if (!/^\d{4}-\d{2}-\d{2}$/.test(pulang)) err.push('Tanggal pulang wajib diisi.');
    else if (pulang < S.berangkat) err.push('Tanggal pulang lebih awal dari tanggal berangkat.');
    else if (pulang > p.sampai) err.push('Tanggal pulang melewati tanggal selesai tugas ' + sdr_(p.nama) + ' (' + tglIndo_(keTanggal_(p.sampai)) + '). Ajukan perpanjangan bila memang diperpanjang.');
    else if (pulang > ymd_(new Date())) err.push('Tanggal pulang belum terjadi.');
    if (err.length) return { ok: false, error: err };
    const lama = selisihHari_(S.berangkat, pulang) + 1, saku = lama * Number(param_().UANG_SAKU_HARI), now = new Date();
    pastikanKolom_(SH.PERSONEL, ['Tgl Pulang Aktual', 'Lama Aktual', 'Uang Saku Dibayar Kasir', 'Keterangan Pulang', 'Lapor Pulang Oleh', 'Lapor Pulang Pada']);
    const B = bacaBaris_('SPD'), r = B.rows.find(x => x['Nomor SPD'] === S.nomor && String(x['Employee ID']).toUpperCase() === id);
    ubahBaris_(B, r._row, { 'Tgl Pulang Aktual': keTanggal_(pulang), 'Lama Aktual': lama, 'Uang Saku Dibayar Kasir': saku,
      'Keterangan Pulang': besar_(d.keterangan).slice(0, 300), 'Lapor Pulang Oleh': u.nama, 'Lapor Pulang Pada': now });
    SpreadsheetApp.flush();
    logAksi_('SPD', S.nomor, 'LAPOR PULANG', u.email, p.nama + ' pulang ' + pulang + ' (' + lama + ' hari) — uang saku ' + rpH_(saku) + ' dibayar kasir');
    hasil = { ok: true, nama: sdr_(p.nama), lama: lama, uangSaku: saku, pulang: pulang, nomor: S.nomor, tempat: S.tempat, email: S.emailPembuat };
    if (baru) {
      const idSerah = 'S' + Utilities.getUuid().slice(0, 8), nominal = Math.round(Number(d.kasSerah)) || 0;
      pastikanKolom_(SH.SPD, ['Ketua Tim']);
      const Ind = bacaBaris_SPD_induk_(), I = Ind.rows.find(x => x['Nomor SPD'] === S.nomor);
      if (I) ubahBaris_(Ind, I._row, { 'Ketua Tim': baru.id });
      sheetUbah_(SH_UBAH.UBAH);
      tulisBaris_(SH_UBAH.UBAH, [{ 'ID': idSerah, 'Nomor SPD': S.nomor, 'Jenis': 'SERAH KAS', 'Tambah Kas': nominal,
        'Alasan': 'SERAH TERIMA UANG KAS DARI ' + besar_(p.nama) + ' KEPADA ' + besar_(baru.nama) + ' (KETUA TIM PULANG ' + pulang + ')',
        'Status': UBAH_ST.KONFIRMASI, 'Dibuat Oleh': u.nama, 'Email Pembuat': u.email, 'Dibuat Pada': now }]);
      SpreadsheetApp.flush();
      logAksi_('SPD', S.nomor, 'SERAH KAS', u.email, p.nama + ' → ' + baru.nama + ' ' + rpH_(nominal));
      hasil.serah = { id: idSerah, nominal: nominal, dari: sdr_(p.nama), ke: sdr_(baru.nama), keId: baru.id };
    }
  } finally { lock.releaseLock(); }
  const admin = String(param_().EMAIL_ADMIN || '').trim().toLowerCase(), pembuat = String(hasil.email || u.email).toLowerCase();
  kirimEmail_(pembuat, 'Lapor Pulang Dinas — ' + hasil.nama + ' (' + hasil.nomor + ')',
    '<p><b>' + escH_(hasil.nama) + '</b> telah pulang dari perjalanan dinas <b>' + escH_(hasil.nomor) + '</b> (' + escH_(hasil.tempat) + ') pada <b>' +
    escH_(hariIndo_(keTanggal_(hasil.pulang)) + ', ' + tglIndo_(keTanggal_(hasil.pulang))) + '</b> — lama dinas ' + hasil.lama + ' hari.</p>' +
    '<p>Uang saku <b>' + rpH_(hasil.uangSaku) + '</b> dibayarkan oleh kasir di kantor (tidak termasuk pertanggungjawaban ketua tim). ' +
    'Karyawan dapat diajukan lembur kembali mulai hari berikutnya.</p>' +
    (hasil.serah ? '<p>Uang kas dalam perjalanan sebesar <b>' + rpH_(hasil.serah.nominal) + '</b> diserahkan kepada <b>' + escH_(hasil.serah.ke) + '</b> sebagai ketua tim baru ' +
      '(menunggu konfirmasi penerimaan). Kekurangan kas di pertanggungjawaban akhir menjadi tanggungan ketua tim baru.</p>' : '') +
    '<p>Dicatat oleh ' + escH_(sdr_(u.nama)) + '.</p>', admin && admin !== pembuat ? admin : '');
  if (hasil.serah) {
    const k = karyawan_()[hasil.serah.keId], I = induk_('SPD', hasil.nomor);
    if (k && k.email) {
      const err = kirimEmail_(k.email, 'Serah Terima Uang Kas Dinas ' + hasil.nomor,
        '<p>Yth. <b>' + escH_(sdr_(k.nama)) + '</b>,</p><p><b>' + escH_(hasil.serah.dari) + '</b> pulang lebih dulu. Anda ditunjuk sebagai <b>ketua tim</b> dan menerima uang kas dalam perjalanan sebesar <b>' +
        rpH_(hasil.serah.nominal) + '</b>.</p>' + tabelRingkasDinas_({ 'Full Name': k.nama, 'Employee ID': k.id, 'Uang Saku': 0 }, I).replace(/<tr><td[^>]*>Uang Saku<\/td><td[^>]*>[^<]*<\/td><\/tr>/, '').replace('>Personel<', '>Ketua Tim<') +
        '<p>Mohon konfirmasi bahwa Anda sudah menerima uang kas tersebut:</p><p style="margin-top:14px">' + tombol_(linkAksi_('kt', hasil.serah.id), '✅ Sudah Diterima', '#1E7B4A') + '</p>' +
        '<p>Sebagai ketua tim, Anda melaporkan uang kas, biaya &amp; bukti kepada admin setelah dinas selesai.</p>');
      const U = bacaUbah_(SH_UBAH.UBAH), r = U.rows.find(x => String(x['ID']) === hasil.serah.id);
      if (r) ubahBaris_(U, r._row, { 'Status Email': err ? 'GAGAL: ' + err : 'Konfirmasi terkirim ke ' + k.email + ' ' + Utilities.formatDate(new Date(), TZ, 'dd/MM HH:mm') });
    }
  }
  return hasil;
}

/* ============================== PENGINGAT LAPORAN DINAS ============================== */
/** 1x sehari (dipanggil dari prosesAntrian, mulai jam 08.00): surat dinas yang sudah lewat tanggal kembali
 *  tetapi belum dipertanggungjawabkan → email pengingat ke ketua tim. */
function pengingatDinasHarian_() {
  const P = PropertiesService.getScriptProperties(), hari = ymd_(new Date());
  if (P.getProperty('PENGINGAT_DINAS') === hari || Number(Utilities.formatDate(new Date(), TZ, 'H')) < 8) return;
  P.setProperty('PENGINGAT_DINAS', hari);
  let n = 0;
  bacaTabel_(SH.SPD, 'Nomor SPD').filter(r => String(r['Status Dinas'] || 'Aktif') === 'Aktif' && r['Tgl Kembali'] instanceof Date && ymd_(r['Tgl Kembali']) < hari)
    .forEach(r => {
      try {
        const S = dataSpd_(r['Nomor SPD']);
        if (!S || !S.selesaiApproval || S.perpanjanganPending || !S.ketua) return;
        const k = karyawan_()[S.ketua.id];
        if (!k || !k.email) return;
        const lewat = selisihHari_(S.kembali, hari), belum = S.personel.filter(p => p.status === ST.SETUJU && !p.pulang);
        const err = kirimEmail_(k.email, 'Pengingat Laporan Dinas ' + S.nomor,
          '<p>Yth. <b>' + escH_(sdr_(k.nama)) + '</b>,</p><p>Perjalanan dinas berikut seharusnya selesai pada <b>' + escH_(hariIndo_(keTanggal_(S.kembali)) + ', ' + tglIndo_(keTanggal_(S.kembali))) +
          '</b> (' + lewat + ' hari yang lalu), namun laporannya belum diterima.</p>' + tabelRingkasDinas_({ 'Full Name': k.nama, 'Employee ID': k.id, 'Uang Saku': 0 }, induk_('SPD', S.nomor))
            .replace(/<tr><td[^>]*>Uang Saku<\/td><td[^>]*>[^<]*<\/td><\/tr>/, '').replace('>Personel<', '>Ketua Tim<') +
          '<p>Mohon segera melaporkan dinas (tanggal pulang, uang kas, biaya &amp; bukti) kepada admin agar Surat Perintah Dinas dapat diselesaikan. ' +
          'Selama belum dilaporkan, ' + (belum.length > 1 ? 'personel berikut belum dapat diajukan lembur: ' + belum.map(p => escH_(sdr_(p.nama))).join(', ') : 'Anda belum dapat diajukan lembur') + '.</p>');
        logAksi_('SPD', S.nomor, err ? 'PENGINGAT GAGAL' : 'PENGINGAT LAPORAN', k.email, 'lewat ' + lewat + ' hari' + (err ? ' — ' + err : ''));
        n++;
      } catch (e) { logAksi_('SPD', r['Nomor SPD'], 'PENGINGAT GAGAL', '', String(e)); }
    });
  return n;
}
