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

/** Dipanggil pemicu waktu setiap 1 menit.
 *  Gangguan sesaat server Google ("server error… try again") ditangkap & dicatat di Log (bukan email kegagalan);
 *  tugas diambil SATU per SATU dari antrian sehingga bila eksekusi terhenti, tugas lain tetap tersimpan,
 *  dan tugas yang gagal dicoba ulang (maks. 3x) pada menit berikutnya. */
function prosesAntrian() {
  const mulai = Date.now();
  try {
    const P = PropertiesService.getScriptProperties();
    P.setProperty(ANTRIAN_DETAK, String(Date.now()));
    try { pengingatDinasHarian_(); } catch (e) { logAksi_('SPD', '-', 'PENGINGAT GAGAL', '', String(e)); }   // 1x sehari
    try { rekapOwnerHarian_(); } catch (e) { logAksi_('SPD', '-', 'REKAP OWNER GAGAL', '', String(e)); }        // rekap dinas owner 1x sehari setelah 17.00
    while (Date.now() - mulai < 4 * 60 * 1000) {   // sisa antrian dilanjutkan pada menit berikutnya
      const t = ambilTugas_(P);
      if (!t) return;
      try { _K = null; _P = null; jalankanTugas_(t); }
      catch (e) {
        t.coba = (t.coba || 0) + 1;
        logAksi_(t.jenis || 'SPL', t.nomor || t.id || '-', t.coba < 3 ? 'ANTRIAN DIULANG' : 'ANTRIAN GAGAL', '', t.t + ' (percobaan ' + t.coba + '): ' + String(e));
        if (t.coba < 3) kembalikanTugas_(P, t);
        return;   // beri jeda — coba lagi menit berikutnya
      }
    }
  } catch (e) {
    try { console.warn('prosesAntrian: ' + e); } catch (x) {}   // gangguan sesaat server Google — dicoba lagi menit berikutnya
  }
}
function ambilTugas_(P) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return null;
  try {
    const q = JSON.parse(P.getProperty(ANTRIAN_KEY) || '[]');
    if (!q.length) return null;
    const t = q.shift();
    P.setProperty(ANTRIAN_KEY, JSON.stringify(q));
    return t;
  } finally { lock.releaseLock(); }
}
function kembalikanTugas_(P, t) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return;
  try { const q = JSON.parse(P.getProperty(ANTRIAN_KEY) || '[]'); q.push(t); P.setProperty(ANTRIAN_KEY, JSON.stringify(q)); }
  finally { lock.releaseLock(); }
}

function jalankanTugas_(t) {
  if (t.t === 'konfirmasi') return kirimKonfirmasi_(t.jenis, t.nomor);
  if (t.t === 'suratLembur') return kirimSuratLembur_(t.nomor);
  if (t.t === 'selesai') return cekSelesai_(t.jenis, t.nomor);
  if (t.t === 'batal') return batalLanjutan_(t);
  if (t.t === 'persetujuanDinas') return kirimPersetujuanDinas_(t.nomor);
  if (t.t === 'perpanjangan') return kirimPersetujuanPerpanjangan_(t.id);
  if (t.t === 'perpanjanganCek') return selesaiPerpanjangan_(t.id);
  if (t.t === 'tambahKas') return kirimTambahKas_(t.id);
  throw new Error('Tugas tidak dikenal: ' + t.t);
}
