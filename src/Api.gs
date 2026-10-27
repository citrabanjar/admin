/**
 * =====================================================================
 *  SISTEM SPL & SPD — File : Api.gs
 *  "Pintu data" HANYA-BACA untuk diagnosa (api=diag, antrian, struktur, peran) — butuh &key=<KUNCI>.
 *  - Tanpa data gaji / estimasi kompensasi.
 *  - Kunci dilihat lewat menu ⚙️ SPL / SPD → 🔑 Kunci API diagnosa.
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
  try {
    if (q.api === 'diag') {   // diagnosa: log terakhir + baris lembur terbaru (tanpa gaji)
      const t = v => v instanceof Date ? Utilities.formatDate(v, TZ, 'dd/MM HH:mm:ss') : v;
      const log = sheet_(SH.LOG).getDataRange().getValues().slice(-40).map(r => r.map(t).join(' | '));
      const spl = bacaTabel_(SH.SPL, 'Nomor SPL').slice(-15).map(r => [r['Nomor SPL'], t(r['Tanggal Lembur']), r['Employee ID'], r['Full Name'], r['Status'], r['Email Pembuat'],
        'pdf:' + (r['Link PDF'] ? 1 : 0) + ' owner:' + (r['Link PDF Owner'] ? 1 : 0) + ' atasan:' + (r['Link PDF Atasan'] ? 1 : 0), r['Status Email'],
        'rantai:' + rantaiAtasan_(r['Employee ID']).map(a => a.id).join('>'), 'owner:' + owner_().map(o => o.id).join(',')].join(' | '));
      const spd = bacaTabel_(SH.PERSONEL, 'Nomor SPD').slice(-10).map(r => [r['Nomor SPD'], r['Employee ID'], r['Full Name'], r['Status'], r['Status Email'], t(r['Konfirmasi Pada']), r['Atasan'], r['Email Atasan'],
        'email:' + String((karyawan_()[String(r['Employee ID']).toUpperCase()] || {}).email || '-').replace(/^(.{3}).*(@.*)$/, '$1•••$2'),
        'penyetuju:' + ((penyetujuDinas_(r['Employee ID']) || {}).id || '-'), 'rantai:' + rantaiAtasan_(r['Employee ID']).map(a => a.id).join('>')].join(' | '));
      return keluar({ ok: true, log: log, spl: spl, spd: spd, kuota: MailApp.getRemainingDailyQuota() });
    }
    if (q.api === 'antrian') { const P = PropertiesService.getScriptProperties(), d = Number(P.getProperty(ANTRIAN_DETAK) || 0);
      return keluar({ ok: true, aktif: antrianAktif_(), detakTerakhir: d ? Utilities.formatDate(new Date(d), TZ, 'dd/MM/yyyy HH:mm:ss') : '-', antrian: JSON.parse(P.getProperty(ANTRIAN_KEY) || '[]').length }); }
    if (q.api === 'struktur') {   // struktur kolom Data Karyawan (tanpa gaji / email / data pribadi)
      const sh = sheet_(SH.KARYAWAN), v = sh.getDataRange().getValues(), head = v.shift().map(String);
      const pakai = head.map((h, i) => /atasan|peran|level|jabatan|position|posisi|grade|organi|depart|divisi|employee id|full name|status employee|nama/i.test(h) && !/salary|gaji|email|phone|telp|hp|nik|ktp|npwp|rek|bank|alamat|lahir|birth/i.test(h) ? i : -1).filter(i => i >= 0);
      return keluar({ ok: true, kolom: head, baris: v.filter(r => String(r[0]).trim()).map(r => Object.fromEntries(pakai.map(i => [head[i], r[i] instanceof Date ? ymd_(r[i]) : r[i]]))) });
    }
    if (q.api === 'peran') return keluar({ ok: true, karyawan: Object.values(karyawan_()).filter(k => k.peran !== 'KARYAWAN' || isOwner_(k))
      .map(k => ({ id: k.id, nama: k.nama, peran: k.peran, atasan: k.atasan, akses: k.akses, adaEmail: !!k.email, aktif: k.aktif })) });   // cek ADMIN/BOSS (tanpa gaji & email)
    return keluar({ ok: false, error: 'api tidak dikenal' });
  } catch (e) { return keluar({ ok: false, error: String(e.message || e) }); }
}

/** Menu: tampilkan kunci API diagnosa */
function menuKunciApi() {
  SpreadsheetApp.getUi().alert('Kunci API diagnosa:\n\n' + kunciApi_() + '\n\nKunci ini hanya membuka data diagnosa (tanpa gaji). Jangan dibagikan ke orang lain.');
}
