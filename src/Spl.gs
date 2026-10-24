/**
 * =====================================================================
 *  SISTEM SPL & SPD — File : Spl.gs
 *  Simpan & riwayat Surat Perintah Lembur (dipanggil dari Form.html)
 *  Butuh: Inti.gs, Web.gs
 * =====================================================================
 */

/**
 * @param {string} token
 * @param {{tanggal:'yyyy-mm-dd', baris:[{id, mulai, selesai, uraian}]}} data
 */
function simpanSPL(token, data) {
  const u = sesiCek_(token);
  cekAkses_(u, 'LEMBUR');
  pastikanKolom_(SH.SPL, KOLOM_SPL_TAMBAHAN);
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('Server sibuk, coba simpan lagi beberapa detik lagi.');
  let hasilSimpan;
  try {
    _K = null; _L = null; _P = null;
    const P = param_(), K = karyawan_(), L = liburEfektif_(data && data.tanggal, data && data.liburNasional);
    const err = [];
    const tglStr = String(data && data.tanggal || '');
    if (!/^\d{4}-\d{2}-\d{2}$/.test(tglStr)) throw new Error('Tanggal lembur wajib diisi.');
    const baris = (data.baris || []).filter(b => b && String(b.id || '').trim());
    if (!baris.length) throw new Error('Isi minimal 1 karyawan.');
    if (baris.length > Number(P.MAKS_KARYAWAN_PER_SPL)) throw new Error('Maksimal ' + P.MAKS_KARYAWAN_PER_SPL + ' karyawan per SPL.');

    const aktif = suratAktif_(tglStr, tglStr);
    const sudah = {};
    const hasil = baris.map((b, i) => {
      const id = String(b.id).trim().toUpperCase(), no = 'Baris ' + (i + 1) + ': ';
      const k = K[id];
      if (!k || !k.aktif) { err.push(no + 'Employee ID ' + id + ' tidak ditemukan / tidak aktif.'); return null; }
      if (isOwner_(k)) { err.push(sdr_(k.nama) + ' adalah OWNER — tidak dapat diberi Surat Perintah Lembur.'); return null; }
      if (id === String(u.id).toUpperCase()) err.push(no + 'Anda tidak bisa memasukkan diri sendiri — Surat Perintah Lembur untuk Anda harus diinput oleh admin lain.');
      if (sudah[id]) err.push(sdr_(k.nama) + ' tercantum lebih dari sekali.');
      sudah[id] = true;
      if (!String(b.uraian || '').trim()) err.push('Uraian pekerjaan ' + sdr_(k.nama) + ' wajib diisi.');
      if (aktif.spl[id]) err.push(sdr_(k.nama) + ' sudah terdaftar di SPL ' + String(aktif.spl[id]).replace(/\s*\(\d{4}-\d{2}-\d{2}\)$/, '') + '.');
      if (aktif.spd[id]) err.push(sdr_(k.nama) + ' sedang dinas luar (' + aktif.spd[id] + ') — lembur tidak dapat diajukan.');
      const h = hitungLembur_({ tanggal: tglStr, mulai: b.mulai, selesai: b.selesai, gaji: k.gaji }, P, L);
      h.error.forEach(e => err.push(sdr_(k.nama) + ' — ' + e));
      return { k: k, b: b, h: h };
    });
    if (err.length) return { ok: false, error: err };

    const tgl = Utilities.parseDate(tglStr + ' 00:00', TZ, 'yyyy-MM-dd HH:mm');
    const nomor = nomorBerikutnya_('SPL', tgl);
    const now = new Date();

    tulisBaris_(SH.SPL, hasil.map(x => ({
      'Nomor SPL': nomor, 'Tanggal Lembur': tgl, 'Hari': x.h.hari, 'Status Hari': x.h.statusHari,
      'Employee ID': x.k.id, 'Full Name': x.k.nama, 'Departemen': x.k.departemen,
      'Jam Mulai': angkaKeJam_(jamKeAngka_(x.b.mulai)), 'Jam Selesai': angkaKeJam_(jamKeAngka_(x.b.selesai)),
      'Istirahat': x.h.istirahat, 'Jam Efektif': x.h.jamEfektif, 'Estimasi Kompensasi': x.h.kompensasi,
      'Uraian Pekerjaan': String(x.b.uraian).trim(),
      'Dibuat Oleh': u.nama, 'Email Pembuat': u.email, 'Dibuat Pada': now,
      'Status': ST.KONFIRMASI,
      'Catatan': [].concat(data.liburNasional && !liburSet_()[tglStr] ? ['Ditandai libur nasional oleh admin'] : [], x.h.peringatan).join(' | '),
      'Status Email': 'Mengirim…',
      'ID Baris': Utilities.getUuid().slice(0, 8)
    })), ['Jam Mulai', 'Jam Selesai']);
    SpreadsheetApp.flush();

    if (data.liburNasional) catatLibur_(tglStr, u.nama);
    const totalJam = hasil.reduce((s, x) => s + x.h.jamEfektif, 0);
    logAksi_('SPL', nomor, 'DIBUAT', u.email, hasil.length + ' karyawan, ' + totalJam + ' jam, tgl ' + tglStr);
    hasilSimpan = { ok: true, nomor: nomor, jumlah: hasil.length, totalJam: totalJam, status: ST.KONFIRMASI };
  } finally {
    lock.releaseLock();
  }
  // Email konfirmasi ke tiap karyawan (di luar kunci agar pengguna lain tidak menunggu)
  try { antri_({ t: 'konfirmasi', jenis: 'SPL', nomor: hasilSimpan.nomor }); } catch (e) { logAksi_('SPL', hasilSimpan.nomor, 'EMAIL GAGAL', u.email, String(e)); }
  return hasilSimpan;
}

/** Daftar libur + tanggal yang dicentang "Libur Nasional" di form */
function liburEfektif_(tglStr, centang) {
  const L = Object.assign({}, liburSet_());
  if (centang && /^\d{4}-\d{2}-\d{2}$/.test(String(tglStr || '')) && !L[tglStr]) L[tglStr] = 'Libur nasional';
  return L;
}
/** Tanggal yang dicentang libur nasional → dicatat ke sheet Libur (bila belum ada) */
function catatLibur_(tglStr, oleh) {
  if (liburSet_()[tglStr]) return;
  sheet_(SH.LIBUR).appendRow([Utilities.parseDate(tglStr + ' 00:00', TZ, 'yyyy-MM-dd HH:mm'), '', 'Libur nasional (ditandai ' + oleh + ' via form SPL)']);
  _L = null;
}

/** Riwayat SPL buatan user (30 terakhir), dikelompokkan per nomor, dengan status per karyawan. */
function riwayatSPL(token) {
  const u = sesiCek_(token);
  if (!punyaAkses_(u, 'LEMBUR')) return [];
  const g = {};
  bacaTabel_(SH.SPL, 'Nomor SPL').forEach(r => {
    if (String(r['Email Pembuat']).toLowerCase() !== u.email) return;
    const n = r['Nomor SPL'];
    if (!g[n]) g[n] = {
      jenis: 'SPL', nomor: n, tanggal: r['Tanggal Lembur'] instanceof Date ? ymd_(r['Tanggal Lembur']) : '',
      hari: r['Hari'], statusHari: r['Status Hari'], pdf: r['Link PDF'] || '',
      dibuat: r['Dibuat Pada'] instanceof Date ? r['Dibuat Pada'].getTime() : 0, orang: [], totalJam: 0
    };
    g[n].orang.push({ nama: r['Full Name'], mulai: String(r['Jam Mulai']), selesai: String(r['Jam Selesai']),
      jam: Number(r['Jam Efektif']) || 0, uraian: r['Uraian Pekerjaan'], status: r['Status'],
      batalOleh: r['Dibatalkan Oleh'] || '', alasanBatal: r['Alasan Pembatalan'] || '' });
    g[n].totalJam += Number(r['Jam Efektif']) || 0;
  });
  return Object.values(g).map(x => {
    const selesai = x.orang.filter(o => ST_FINAL.indexOf(o.status) >= 0).length;
    x.status = selesai === x.orang.length ? 'SELESAI' : 'PROSES ' + selesai + '/' + x.orang.length;
    return x;
  }).sort((a, b) => b.dibuat - a.dibuat).slice(0, 30);
}
