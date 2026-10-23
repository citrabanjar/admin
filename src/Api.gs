/**
 * =====================================================================
 *  SISTEM SPL & SPD — File : Api.gs
 *  "Pintu data" HANYA-BACA untuk pengecekan pengajuan lembur di KantorKu.
 *  api=spl (SPL disetujui & belum dieksekusi) · api=laporan (simpan hasil cocok) · api=setuju (baris dicentang)
 *  api=tandai (tandai sudah dieksekusi) — semua butuh &key=<KUNCI>
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

const SH_CEK = 'Cek KantorKu';
const KOLOM_CEK = ['Dicek Pada', 'ID Baris', 'Nomor SPL', 'Tanggal Lembur', 'Employee ID', 'Full Name', 'Jam SPL', 'Jam Efektif SPL',
  'Jam KantorKu', 'Durasi KantorKu', 'Check Out', 'Catatan', 'Setuju', 'Status Eksekusi', 'Dieksekusi Pada'];

/** Sheet laporan pencocokan (dibuat otomatis) */
function sheetCek_() {
  const ss = ss_();
  let sh = ss.getSheetByName(SH_CEK);
  if (!sh) {
    sh = ss.insertSheet(SH_CEK);
    sh.getRange(1, 1, 1, KOLOM_CEK.length).setValues([KOLOM_CEK]).setFontWeight('bold').setBackground('#3A3A3A').setFontColor('#FFFFFF');
    sh.setFrozenRows(1);
    sh.getRange('M2:M').insertCheckboxes();
    sh.getRange('M1').setNote('Centang baris yang disetujui, lalu ketik "eksekusi" ke Claude.');
    sh.setColumnWidth(6, 200); sh.setColumnWidth(12, 260);
  }
  return sh;
}

function apiData_(q) {
  const keluar = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
  if (!q.key || q.key !== kunciApi_()) return keluar({ ok: false, error: 'kunci salah' });
  try {
    if (q.api === 'spl') return keluar(apiSpl_(q));
    if (q.api === 'laporan') return keluar(apiLaporan_(q));
    if (q.api === 'setuju') return keluar(apiSetuju_());
    if (q.api === 'tandai') return keluar(apiTandai_(q));
    if (q.api === 'struktur') {   // struktur kolom Data Karyawan (tanpa gaji / email / data pribadi)
      const sh = sheet_(SH.KARYAWAN), v = sh.getDataRange().getValues(), head = v.shift().map(String);
      const pakai = head.map((h, i) => /atasan|peran|level|jabatan|position|posisi|grade|organi|depart|divisi|employee id|full name|status employee|nama/i.test(h) && !/salary|gaji|email|phone|telp|hp|nik|ktp|npwp|rek|bank|alamat|lahir|birth/i.test(h) ? i : -1).filter(i => i >= 0);
      return keluar({ ok: true, kolom: head, baris: v.filter(r => String(r[0]).trim()).map(r => Object.fromEntries(pakai.map(i => [head[i], r[i] instanceof Date ? ymd_(r[i]) : r[i]]))) });
    }
    if (q.api === 'peran') return keluar({ ok: true, karyawan: Object.values(karyawan_()).filter(k => k.peran !== 'KARYAWAN' || isOwner_(k))
      .map(k => ({ id: k.id, nama: k.nama, peran: k.peran, atasan: k.atasan, adaEmail: !!k.email, aktif: k.aktif })) });   // cek ADMIN/BOSS (tanpa gaji & email)
    return keluar({ ok: false, error: 'api tidak dikenal' });
  } catch (e) { return keluar({ ok: false, error: String(e.message || e) }); }
}

/** SPL yang DISETUJUI atasan & BELUM dieksekusi di KantorKu (default 30 hari terakhir). */
function apiSpl_(q) {
  const tglOk = s => /^\d{4}-\d{2}-\d{2}$/.test(String(s || ''));
  const kini = new Date(), hariIni = ymd_(kini);
  const dari = tglOk(q.dari) ? q.dari : ymd_(new Date(kini.getTime() - 30 * 864e5)), sampai = tglOk(q.sampai) ? q.sampai : hariIni;
  const spl = bacaTabel_(SH.SPL, 'Nomor SPL').filter(r => {
    const t = r['Tanggal Lembur'] instanceof Date ? ymd_(r['Tanggal Lembur']) : '';
    return t && dari <= t && t <= sampai && r['Status'] === 'DISETUJUI' && !String(r['KantorKu'] || '').trim();
  }).map(r => ({
    idBaris: String(r['ID Baris']), nomor: r['Nomor SPL'], tanggal: ymd_(r['Tanggal Lembur']), statusHari: r['Status Hari'],
    id: String(r['Employee ID']).toUpperCase(), nama: r['Full Name'],
    mulai: String(r['Jam Mulai']), selesai: String(r['Jam Selesai']), istirahat: r['Istirahat'], jamEfektif: Number(r['Jam Efektif']) || 0
  }));
  return { ok: true, dari: dari, sampai: sampai, spl: spl };
}

/** Simpan hasil cocok ke sheet "Cek KantorKu". q.data = JSON array
 *  [{idBaris, jamKantorku, durasiKantorku, checkOut, catatan}] (data SPL dilengkapi otomatis). */
function apiLaporan_(q) {
  const data = JSON.parse(q.data || '[]');
  const spl = {}; bacaTabel_(SH.SPL, 'Nomor SPL').forEach(r => spl[String(r['ID Baris'])] = r);
  const sh = sheetCek_(), n = sh.getLastRow() - 1;
  const ada = {};
  (n > 0 ? sh.getRange(2, 2, n, 1).getValues() : []).forEach((r, i) => ada[String(r[0])] = i + 2);
  let baru = 0, ubah = 0;
  data.forEach(d => {
    const r = spl[String(d.idBaris)]; if (!r) return;
    const baris = [new Date(), String(d.idBaris), r['Nomor SPL'], r['Tanggal Lembur'], r['Employee ID'], r['Full Name'],
      r['Jam Mulai'] + ' - ' + r['Jam Selesai'], Number(r['Jam Efektif']) || 0,
      d.jamKantorku || '', d.durasiKantorku || '', d.checkOut || '', d.catatan || ''];
    const row = ada[String(d.idBaris)];
    if (row) {
      if (String(sh.getRange(row, 14).getValue()).trim()) return;   // sudah dieksekusi → jangan diubah
      sh.getRange(row, 1, 1, 12).setValues([baris]); ubah++;
    } else {
      const r0 = sh.getLastRow() + 1;
      sh.getRange(r0, 1, 1, 12).setValues([baris]);
      sh.getRange(r0, 13).insertCheckboxes(); baru++;
    }
  });
  logAksi_('SPL', '-', 'CEK KANTORKU', 'Claude', baru + ' baris baru, ' + ubah + ' diperbarui');
  return { ok: true, baru: baru, diperbarui: ubah };
}

/** Baris yang sudah dicentang "Setuju" dan belum dieksekusi */
function apiSetuju_() {
  const sh = sheetCek_(), n = sh.getLastRow() - 1;
  if (n < 1) return { ok: true, baris: [] };
  const v = sh.getRange(2, 1, n, KOLOM_CEK.length).getValues();
  const st = {}; bacaTabel_(SH.SPL, 'Nomor SPL').forEach(r => st[String(r['ID Baris'])] = r['Status']);
  // lembur yang sudah dibatalkan / tidak lagi DISETUJUI tidak boleh di-approve di KantorKu
  v.forEach((r, i) => { if (r[12] === true && !String(r[13]).trim() && st[String(r[1])] && st[String(r[1])] !== ST.SETUJU) {
    sh.getRange(i + 2, 14).setValue('TIDAK DIEKSEKUSI: lembur ' + st[String(r[1])]); r[13] = 'x'; } });
  return { ok: true, baris: v.filter(r => r[12] === true && !String(r[13]).trim()).map(r => ({
    idBaris: String(r[1]), nomor: r[2], tanggal: r[3] instanceof Date ? ymd_(r[3]) : String(r[3]),
    id: String(r[4]).toUpperCase(), nama: r[5], jamSpl: r[6], jamEfektif: r[7], jamKantorku: r[8], durasiKantorku: r[9]
  })) };
}

/** Tandai sudah dieksekusi. q.ids = "idBaris1,idBaris2", q.status = APPROVED | GAGAL: alasan */
function apiTandai_(q) {
  const ids = String(q.ids || '').split(',').map(s => s.trim()).filter(Boolean);
  const status = String(q.status || 'APPROVED').slice(0, 200), now = new Date();
  const label = status + ' ' + Utilities.formatDate(now, TZ, 'dd/MM/yyyy HH:mm');
  const sh = sheetCek_(), n = sh.getLastRow() - 1;
  (n > 0 ? sh.getRange(2, 2, n, 1).getValues() : []).forEach((r, i) => {
    if (ids.indexOf(String(r[0])) >= 0) sh.getRange(i + 2, 14, 1, 2).setValues([[status, now]]);
  });
  if (status === 'APPROVED') {
    pastikanKolom_(SH.SPL, KOLOM_SPL_TAMBAHAN);
    const S = bacaBaris_('SPL');
    S.rows.filter(r => ids.indexOf(String(r['ID Baris'])) >= 0).forEach(r => ubahBaris_(S, r._row, { 'KantorKu': label }));
  }
  logAksi_('SPL', '-', 'EKSEKUSI KANTORKU', 'Claude', status + ': ' + ids.join(', '));
  return { ok: true, ditandai: ids.length };
}

/** Menu: tampilkan kunci & contoh URL */
function menuKunciApi() {
  const url = ScriptApp.getService().getUrl();
  SpreadsheetApp.getUi().alert('Kunci API pengecekan KantorKu:\n\n' + kunciApi_() +
    '\n\nContoh URL (data SPL hari ini):\n' + url + '?api=spl&key=' + kunciApi_() +
    '\n\nKunci ini hanya membuka data SPL (tanpa gaji). Jangan dibagikan ke orang lain.');
}
