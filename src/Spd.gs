/**
 * =====================================================================
 *  SISTEM SPL & SPD — File : Spd.gs
 *  Simpan & riwayat Surat Perintah Dinas (dipanggil dari Form.html)
 *  Struktur sama dengan formulir SPD Excel: A. Referensi, B. Detail,
 *  C. Personel (maks 5), D. Biaya lainnya.
 *  Butuh: Inti.gs, Web.gs
 * =====================================================================
 */

/**
 * @param {string} token
 * @param {{tanggalSurat, noReferensi, customer, dasarPenugasan, maksud, tempat, berangkat, kembali,
 *          transportasi, sumberBiaya, uangKas, personel:[{id, lama, keterangan}], biaya:[{jenis, uraian, jumlah, bukti}]}} d
 */
function simpanSPD(token, d) {
  const u = sesiCek_(token);
  pastikanKolom_(SH.PERSONEL, KOLOM_SPD_TAMBAHAN);
  pastikanKolom_(SH.SPD, ['Uang Kas Perjalanan']);
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('Server sibuk, coba simpan lagi beberapa detik lagi.');
  let hasilSimpan;
  try {
    _K = null; _P = null;
    const P = param_(), K = karyawan_(), pil = pilihan_();
    const err = [], warn = [];
    const tglOk = s => /^\d{4}-\d{2}-\d{2}$/.test(String(s || ''));
    const teks = s => String(s || '').trim();

    if (!tglOk(d.tanggalSurat)) err.push('Tanggal surat wajib diisi.');
    if (!teks(d.maksud)) err.push('Maksud / Tujuan wajib diisi.');
    if (!teks(d.tempat)) err.push('Tempat tujuan wajib diisi.');
    if (!tglOk(d.berangkat) || !tglOk(d.kembali)) err.push('Tanggal berangkat & kembali wajib diisi.');
    const cekPilihan = (v, daftar, nama) => { if (teks(v) && (daftar || []).indexOf(teks(v)) < 0) err.push(nama + ' tidak valid.'); };
    cekPilihan(d.dasarPenugasan, pil['Dasar Penugasan'], 'Dasar penugasan');
    cekPilihan(d.transportasi, pil['Transportasi'], 'Transportasi');
    cekPilihan(d.sumberBiaya, pil['Sumber Biaya'], 'Sumber biaya');
    if (!teks(d.transportasi)) err.push('Transportasi wajib dipilih.');
    if (!teks(d.sumberBiaya)) err.push('Sumber biaya wajib dipilih.');
    const uangKas = Number(d.uangKas);
    if (d.uangKas === '' || d.uangKas == null || !(uangKas >= 0)) err.push('Uang kas dalam perjalanan wajib diisi (isi 0 bila tidak ada).');

    const personel = (d.personel || []).filter(p => p && teks(p.id)).map(p => Object.assign({}, p, { id: teks(p.id).toUpperCase() }));
    const biaya = (d.biaya || []).filter(b => b && (teks(b.jenis) || Number(b.jumlah)));
    personel.forEach((p, i) => { if (!K[p.id] || !K[p.id].aktif || isOwner_(K[p.id])) err.push('Personel ke-' + (i + 1) + ': Employee ID ' + p.id + ' tidak ditemukan / tidak aktif.'); });
    biaya.forEach((b, i) => {
      if (!teks(b.jenis)) err.push('Biaya ke-' + (i + 1) + ': jenis biaya wajib dipilih.');
      if (!(Number(b.jumlah) >= 0)) err.push('Biaya ke-' + (i + 1) + ': jumlah tidak valid.');
    });

    const h = hitungSPD_({ berangkat: d.berangkat, kembali: d.kembali, personel: personel, biaya: biaya.map(b => ({ jumlah: Number(b.jumlah) || 0 })) }, P);
    err.push.apply(err, h.error);
    warn.push.apply(warn, h.peringatan);

    if (tglOk(d.berangkat) && tglOk(d.kembali)) {
      const aktif = suratAktif_(d.berangkat, d.kembali);
      personel.forEach(p => {
        const k = K[p.id]; if (!k) return;
        if (aktif.spd[p.id]) err.push(k.nama + ' sudah ditugaskan di ' + aktif.spd[p.id] + ' pada tanggal yang beririsan.');
        if (aktif.spl[p.id]) warn.push(k.nama + ' punya SPL ' + aktif.spl[p.id] + ' di rentang dinas — SPL tersebut perlu ditinjau.');
      });
    }
    if (err.length) return { ok: false, error: err };

    const alur = { status: 'PROSES 0/' + personel.length };
    const tp = s => Utilities.parseDate(s + ' 00:00', TZ, 'yyyy-MM-dd HH:mm');
    const tglSurat = tp(d.tanggalSurat);
    const nomor = nomorBerikutnya_('SPD', tglSurat);
    const now = new Date();

    tulisBaris_(SH.SPD, [{
      'Nomor SPD': nomor, 'Tanggal Surat': tglSurat, 'Tahun': Number(d.tanggalSurat.slice(0, 4)),
      'No Urut': Number(nomor.split('/')[0]), 'No Surat Referensi': teks(d.noReferensi), 'Nama Customer': teks(d.customer),
      'Dasar Penugasan': teks(d.dasarPenugasan), 'Maksud / Tujuan': teks(d.maksud), 'Tempat Tujuan': teks(d.tempat),
      'Tgl Berangkat': tp(d.berangkat), 'Tgl Kembali': tp(d.kembali), 'Lama (hari)': h.lama,
      'Transportasi': teks(d.transportasi), 'Sumber Biaya': teks(d.sumberBiaya),
      'Uang Kas Perjalanan': uangKas || 0, 'Total Uang Saku': h.totalUangSaku, 'Total Biaya Lain': h.totalBiayaLain, 'Total Biaya': h.totalBiaya,
      'Status Dinas': 'Aktif', 'Departemen': u.departemen,
      'Dibuat Oleh': u.nama, 'Email Pembuat': u.email, 'Dibuat Pada': now,
      'Status': alur.status,
      'Diketahui Oleh': alur.diketahuiOleh || '', 'Diketahui Pada': alur.diketahuiOleh ? now : '',
      'Catatan': [].concat(warn, alur.catatan ? [alur.catatan] : []).join(' | '),
      'Status Email': 'Konfirmasi dikirim ke personel'
    }]);
    tulisBaris_(SH.PERSONEL, h.personel.map((p, i) => ({
      'Nomor SPD': nomor, 'No': i + 1, 'Employee ID': p.id, 'Full Name': K[p.id].nama, 'Job Position': K[p.id].jabatan,
      'Lama (hari)': Number(p.lama), 'Uang Saku': p.uangSaku, 'Keterangan': teks(p.keterangan),
      'Status': ST.KONFIRMASI, 'Status Email': 'Mengirim…', 'ID Baris': Utilities.getUuid().slice(0, 8)
    })));
    tulisBaris_(SH.BIAYA, biaya.map((b, i) => ({
      'Nomor SPD': nomor, 'No': i + 1, 'Jenis Biaya': teks(b.jenis), 'Uraian': teks(b.uraian),
      'Jumlah (Rp)': Number(b.jumlah) || 0, 'Bukti': teks(b.bukti)
    })));
    SpreadsheetApp.flush();

    logAksi_('SPD', nomor, 'DIBUAT', u.email, personel.length + ' personel, ' + h.lama + ' hari ke ' + teks(d.tempat) + ', total Rp' + h.totalBiaya + ' → ' + alur.status);
    hasilSimpan = { ok: true, nomor: nomor, jumlah: personel.length, lama: h.lama, totalBiaya: h.totalBiaya, status: ST.KONFIRMASI, peringatan: warn };
  } finally {
    lock.releaseLock();
  }
  try { kirimKonfirmasi_('SPD', hasilSimpan.nomor); } catch (e) { logAksi_('SPD', hasilSimpan.nomor, 'EMAIL GAGAL', u.email, String(e)); }
  return hasilSimpan;
}

/** Riwayat SPD buatan user (30 terakhir). */
function riwayatSPD(token) {
  const u = sesiCek_(token);
  const per = {};
  bacaTabel_(SH.PERSONEL, 'Nomor SPD').forEach(r => {
    (per[r['Nomor SPD']] = per[r['Nomor SPD']] || []).push({ nama: r['Full Name'], lama: Number(r['Lama (hari)']) || 0, keterangan: r['Keterangan'], status: r['Status'] || '' });
  });
  return bacaTabel_(SH.SPD, 'Nomor SPD')
    .filter(r => String(r['Email Pembuat']).toLowerCase() === u.email)
    .map(r => ({
      jenis: 'SPD', nomor: r['Nomor SPD'], status: r['Status'],
      tanggal: r['Tanggal Surat'] instanceof Date ? ymd_(r['Tanggal Surat']) : '',
      berangkat: r['Tgl Berangkat'] instanceof Date ? ymd_(r['Tgl Berangkat']) : '',
      kembali: r['Tgl Kembali'] instanceof Date ? ymd_(r['Tgl Kembali']) : '',
      lama: Number(r['Lama (hari)']) || 0, tempat: r['Tempat Tujuan'], maksud: r['Maksud / Tujuan'],
      totalBiaya: Number(r['Total Biaya']) || 0, uangKas: Number(r['Uang Kas Perjalanan']) || 0,
      dibuat: r['Dibuat Pada'] instanceof Date ? r['Dibuat Pada'].getTime() : 0,
      orang: per[r['Nomor SPD']] || [], pdf: r['Link PDF'] || '', pdfLaporan: r['Link PDF Laporan'] || '',
      statusDinas: String(r['Status Dinas'] || 'Aktif') === 'Selesai' ? 'SELESAI' : 'OPEN'
    }))
    .sort((a, b) => b.dibuat - a.dibuat).slice(0, 30);
}
