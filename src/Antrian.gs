/**
 * =====================================================================
 *  SISTEM SPL & SPD — File : Antrian.gs
 *  Pekerjaan berat (kirim email konfirmasi, buat PDF rekap, email pembatalan)
 *  dikerjakan di LATAR BELAKANG agar tombol di form / email langsung selesai.
 *
 *  Butuh 1 pemicu (trigger) waktu: fungsi prosesAntrian, setiap 1 menit
 *  (dibuat sekali dari editor Apps Script → Pemicu). Selama pemicu belum
 *  aktif, pekerjaan tetap dijalankan langsung seperti sebelumnya.
 * =====================================================================
 */
const ANTRIAN_KEY = 'Q', ANTRIAN_DETAK = 'Q_DETAK';

/** Pemicu aktif bila prosesAntrian berjalan dalam 3 menit terakhir */
function antrianAktif_() {
  return Date.now() - Number(PropertiesService.getScriptProperties().getProperty(ANTRIAN_DETAK) || 0) < 3 * 60 * 1000;
}

/** Masukkan tugas ke antrian (atau jalankan langsung bila pemicu belum aktif). */
function antri_(tugas) {
  if (!antrianAktif_()) return jalankanTugas_(tugas);
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return jalankanTugas_(tugas);
  try {
    const P = PropertiesService.getScriptProperties();
    const q = JSON.parse(P.getProperty(ANTRIAN_KEY) || '[]');
    const kunci = JSON.stringify(tugas);
    if (!q.some(t => JSON.stringify(t) === kunci)) q.push(tugas);
    P.setProperty(ANTRIAN_KEY, JSON.stringify(q));
  } finally { lock.releaseLock(); }
}

/** Dipanggil pemicu waktu setiap 1 menit. */
function prosesAntrian() {
  const P = PropertiesService.getScriptProperties();
  P.setProperty(ANTRIAN_DETAK, String(Date.now()));
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return;
  let q;
  try {
    q = JSON.parse(P.getProperty(ANTRIAN_KEY) || '[]');
    if (!q.length) return;
    P.setProperty(ANTRIAN_KEY, '[]');
  } finally { lock.releaseLock(); }
  q.forEach(t => {
    try { _K = null; _P = null; jalankanTugas_(t); }
    catch (e) { logAksi_(t.jenis || 'SPL', t.nomor || t.id || '-', 'ANTRIAN GAGAL', '', t.t + ': ' + String(e)); }
  });
}

function jalankanTugas_(t) {
  if (t.t === 'konfirmasi') return kirimKonfirmasi_(t.jenis, t.nomor);
  if (t.t === 'selesai') return cekSelesai_(t.jenis, t.nomor);
  if (t.t === 'batal') return batalLanjutan_(t);
  if (t.t === 'persetujuanDinas') return kirimPersetujuanDinas_(t.nomor);
  if (t.t === 'perpanjangan') return kirimPersetujuanPerpanjangan_(t.id);
  if (t.t === 'perpanjanganCek') return selesaiPerpanjangan_(t.id);
  if (t.t === 'tambahKas') return kirimTambahKas_(t.id);
  throw new Error('Tugas tidak dikenal: ' + t.t);
}
