/**
 * =====================================================================
 *  SISTEM SPL & SPD — File : Api.gs
 *  "Pintu data" HANYA-BACA untuk pengecekan pengajuan lembur di KantorKu.
 *  URL : <web app>/exec?api=spl&key=<KUNCI>&dari=yyyy-mm-dd&sampai=yyyy-mm-dd
 *  - Tanpa data gaji / estimasi kompensasi.
 *  - Kunci dilihat lewat menu ⚙️ SPL / SPD → 🔑 Kunci API pengecekan KantorKu.
 * =====================================================================
 */

function kunciApi_() {
  const p = PropertiesService.getScriptProperties();
  let k = p.getProperty('API_KEY');
  if (!k) { k = Utilities.getUuid().replace(/-/g, ''); p.setProperty('API_KEY', k); }
  return k;
}

function apiData_(q) {
  const keluar = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
  if (!q.key || q.key !== kunciApi_()) return keluar({ ok: false, error: 'kunci salah' });
  const tglOk = s => /^\d{4}-\d{2}-\d{2}$/.test(String(s || ''));
  const hariIni = ymd_(new Date());
  const dari = tglOk(q.dari) ? q.dari : hariIni, sampai = tglOk(q.sampai) ? q.sampai : dari;

  const spl = bacaTabel_(SH.SPL, 'Nomor SPL').filter(r => {
    const t = r['Tanggal Lembur'] instanceof Date ? ymd_(r['Tanggal Lembur']) : '';
    return t && dari <= t && t <= sampai;
  }).map(r => ({
    nomor: r['Nomor SPL'], tanggal: ymd_(r['Tanggal Lembur']), hari: r['Hari'], statusHari: r['Status Hari'],
    id: String(r['Employee ID']).toUpperCase(), nama: r['Full Name'], departemen: r['Departemen'],
    mulai: String(r['Jam Mulai']), selesai: String(r['Jam Selesai']), istirahat: r['Istirahat'],
    jamEfektif: Number(r['Jam Efektif']) || 0, uraian: r['Uraian Pekerjaan'], status: r['Status'],
    atasan: r['Atasan'] || '', disetujuiPada: r['Disetujui Pada'] instanceof Date ? r['Disetujui Pada'].toISOString() : ''
  }));

  // Personel yang sedang dinas luar (SPD aktif) di rentang tanggal → lembur tidak boleh di-approve
  const aktif = suratAktif_(dari, sampai);
  return keluar({ ok: true, dari: dari, sampai: sampai, dibuat: new Date().toISOString(), spl: spl, dinasLuar: aktif.spd });
}

/** Menu: tampilkan kunci & contoh URL */
function menuKunciApi() {
  const url = ScriptApp.getService().getUrl();
  SpreadsheetApp.getUi().alert('Kunci API pengecekan KantorKu:\n\n' + kunciApi_() +
    '\n\nContoh URL (data SPL hari ini):\n' + url + '?api=spl&key=' + kunciApi_() +
    '\n\nKunci ini hanya membuka data SPL (tanpa gaji). Jangan dibagikan ke orang lain.');
}
