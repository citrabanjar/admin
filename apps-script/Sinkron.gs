/**
 * =====================================================================
 *  SISTEM SPL & SPD — File : Sinkron.gs
 *  Menu "4. Sinkron data karyawan": salin data dari sheet karyawan lain
 *  (mis. hasil export HRIS) ke sheet Karyawan, dicocokkan per Employee ID.
 *  - Isi / perbarui: Email, Full Name, Job Position, Organization, Basic Salary
 *  - Karyawan baru → ditambahkan (Peran default dari jabatan, Aktif = YA)
 *  - Peran & Aktif karyawan lama TIDAK diubah
 *  - Karyawan yang tidak ada di sheet sumber hanya DILAPORKAN (tidak dihapus)
 *  Judul kolom sumber dikenali otomatis (Employee ID/NIK, Nama, Email, dst.)
 * =====================================================================
 */

const SINKRON_ALIAS = {
  id:      ['employeeid', 'idkaryawan', 'id', 'nik', 'nip', 'noinduk', 'nomorinduk', 'idemployee', 'empid'],
  nama:    ['fullname', 'namalengkap', 'nama', 'name', 'namakaryawan', 'employeename'],
  jabatan: ['jobposition', 'jabatan', 'posisi', 'position', 'job'],
  org:     ['organization', 'organisasi', 'departemen', 'department', 'divisi', 'bagian', 'unit'],
  gaji:    ['basicsalary', 'gajipokok', 'gaji', 'salary', 'upah'],
  email:   ['email', 'emailaddress', 'alamatemail', 'surel', 'emailkantor', 'workemail']
};
const SINKRON_FIELD = { nama: 'Full Name', jabatan: 'Job Position', org: 'Organization', gaji: 'Basic Salary', email: 'Email' };

/* ===================== FUNGSI MURNI (bisa diuji tanpa spreadsheet) ===================== */

const normJudul_ = s => String(s || '').toLowerCase().replace(/[^a-z0-9]/g, '');

/** Cari baris judul (10 baris pertama) & peta kolom sumber → {baris, kol:{id,nama,...}} */
function kenaliKolom_(values) {
  for (let r = 0; r < Math.min(10, values.length); r++) {
    const head = values[r].map(normJudul_), kol = {};
    Object.keys(SINKRON_ALIAS).forEach(f => {
      for (const a of SINKRON_ALIAS[f]) { const i = head.indexOf(a); if (i >= 0) { kol[f] = i; break; } }
    });
    if (kol.id !== undefined && (kol.email !== undefined || kol.nama !== undefined)) return { baris: r, kol: kol };
  }
  return null;
}

/**
 * Susun rencana sinkron.
 * @param {Object[]} target  baris sheet Karyawan [{id,nama,jabatan,org,gaji,email,peran,aktif}]
 * @param {Object[]} sumber  baris sumber [{id,nama,jabatan,org,gaji,email}] (field tak ada = undefined)
 * @return {{hasil:Object[], ubah:Object[], baru:Object[], hilang:Object[], emailSalah:string[], emailDobel:string[], idDobel:string[]}}
 */
function rencanaSinkron_(target, sumber) {
  const reEmail = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  const peranDefault = j => ({ FOREMAN: 'PEMBUAT', SUPERVISOR: 'PEMBUAT', MANAGER: 'MANAGER' }[String(j || '').toUpperCase()] || 'KARYAWAN');
  const hasil = target.map(t => Object.assign({}, t));
  const idx = {}; hasil.forEach((t, i) => idx[t.id] = i);
  const ubah = [], baru = [], emailSalah = [], idDobel = [], dilihat = {};

  sumber.forEach(s => {
    const id = String(s.id || '').trim().toUpperCase();
    if (!id) return;
    if (dilihat[id]) { idDobel.push(id); return; }
    dilihat[id] = true;
    const v = {};
    Object.keys(SINKRON_FIELD).forEach(f => {
      if (s[f] === undefined || s[f] === null || String(s[f]).trim() === '') return;
      v[f] = f === 'gaji' ? Number(String(s[f]).replace(/[^0-9.-]/g, '')) || 0
           : f === 'email' ? String(s[f]).trim().toLowerCase() : String(s[f]).trim();
    });
    if (v.email && !reEmail.test(v.email)) { emailSalah.push(id + ' — ' + v.email); delete v.email; }

    if (idx[id] !== undefined) {
      const t = hasil[idx[id]], beda = [];
      Object.keys(v).forEach(f => {
        const lama = f === 'gaji' ? Number(t[f]) || 0 : String(t[f] || '').trim();
        if (lama !== v[f]) { beda.push({ field: SINKRON_FIELD[f], dari: lama, ke: v[f] }); t[f] = v[f]; }
      });
      if (beda.length) ubah.push({ id: id, nama: t.nama, beda: beda });
    } else {
      const n = { id: id, nama: v.nama || '', jabatan: v.jabatan || '', org: v.org || '', gaji: v.gaji || 0,
                  email: v.email || '', peran: peranDefault(v.jabatan), aktif: 'YA' };
      idx[id] = hasil.push(n) - 1;
      baru.push(n);
    }
  });

  const hilang = hasil.filter(t => !dilihat[t.id] && String(t.aktif).toUpperCase() !== 'TIDAK');
  const em = hasil.filter(t => t.email && String(t.aktif).toUpperCase() !== 'TIDAK').map(t => t.email);
  const emailDobel = [...new Set(em.filter((x, i) => em.indexOf(x) !== i))];
  return { hasil: hasil, ubah: ubah, baru: baru, hilang: hilang, emailSalah: emailSalah, emailDobel: emailDobel, idDobel: idDobel };
}
/* ===================== akhir fungsi murni ===================== */

/** Menu: buka dialog pilih sheet sumber */
function sinkronKaryawan() {
  const sistem = Object.values(SH).concat(['Panduan']);
  const daftar = SpreadsheetApp.getActiveSpreadsheet().getSheets().map(s => s.getName()).filter(n => sistem.indexOf(n) < 0);
  if (!daftar.length) { SpreadsheetApp.getUi().alert('Tidak ada sheet sumber.\n\nTambahkan dulu sheet berisi data karyawan (dengan kolom Employee ID & Email).'); return; }
  const html = '<div style="font-family:Arial;font-size:13px">' +
    '<label><b>Sheet sumber data karyawan:</b></label><br>' +
    '<select id="s" style="width:100%;padding:6px;margin:6px 0 10px">' + daftar.map(n => '<option>' + n.replace(/</g, '&lt;') + '</option>').join('') + '</select>' +
    '<button id="b1" onclick="cek()" style="padding:8px 14px;background:#BE353F;color:#fff;border:0;border-radius:6px;cursor:pointer">Pratinjau</button> ' +
    '<button id="b2" onclick="terap()" disabled style="padding:8px 14px;background:#3A3A3A;color:#fff;border:0;border-radius:6px;cursor:pointer">Terapkan</button>' +
    '<div id="o" style="margin-top:12px;max-height:380px;overflow:auto"></div></div>' +
    '<script>' +
    'function cek(){o.innerHTML="Membaca…";b2.disabled=true;google.script.run.withSuccessHandler(function(h){o.innerHTML=h;b2.disabled=h.indexOf("TIDAK_BISA")>=0}).withFailureHandler(function(e){o.innerHTML="<p style=color:#B42318>"+e.message+"</p>"}).pratinjauSinkron(s.value)}' +
    'function terap(){if(!confirm("Terapkan perubahan ke sheet Karyawan?"))return;b1.disabled=b2.disabled=true;o.innerHTML="Menyimpan…";google.script.run.withSuccessHandler(function(h){o.innerHTML=h}).withFailureHandler(function(e){o.innerHTML="<p style=color:#B42318>"+e.message+"</p>";b1.disabled=false}).terapkanSinkron(s.value)}' +
    '</script>';
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(640).setHeight(560), 'Sinkron Data Karyawan');
}

function bacaTargetKaryawan_() {
  const sh = sheet_(SH.KARYAWAN), c = kolom_(SH.KARYAWAN);
  const v = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), sh.getLastColumn()).getValues();
  const g = (r, j) => r[c[j] - 1];
  return v.filter(r => String(g(r, 'Employee ID')).trim()).map(r => ({
    id: String(g(r, 'Employee ID')).trim().toUpperCase(), nama: g(r, 'Full Name'), jabatan: g(r, 'Job Position'),
    org: g(r, 'Organization'), gaji: g(r, 'Basic Salary'), email: String(g(r, 'Email') || '').trim().toLowerCase(),
    peran: g(r, 'Peran') || 'KARYAWAN', aktif: g(r, 'Aktif') || 'YA'
  }));
}
function bacaSumber_(namaSheet) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(namaSheet);
  if (!sh) throw new Error('Sheet "' + namaSheet + '" tidak ditemukan.');
  const v = sh.getDataRange().getValues();
  const k = kenaliKolom_(v);
  if (!k) throw new Error('Judul kolom tidak dikenali. Sheet sumber harus punya kolom "Employee ID" (atau NIK/ID Karyawan) dan "Email".');
  const rows = v.slice(k.baris + 1).map(r => {
    const o = {}; Object.keys(k.kol).forEach(f => o[f] = r[k.kol[f]]); return o;
  }).filter(o => String(o.id || '').trim());
  const judul = Object.keys(k.kol).map(f => f + ' ← "' + v[k.baris][k.kol[f]] + '"');
  return { rows: rows, judul: judul };
}

function laporanSinkron_(R, judul, sudah) {
  const daftar = (arr, f) => '<div style="background:#f6f6f6;padding:6px;border-radius:4px;max-height:120px;overflow:auto;font-size:12px">' + arr.map(f).join('<br>') + '</div>';
  const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;');
  let h = '<p style="color:#666;font-size:12px">Kolom terbaca: ' + esc(judul.join(' · ')) + '</p>';
  h += '<p><b>' + (sudah ? '✅ Diterapkan:' : 'Rencana:') + '</b> ' + R.ubah.length + ' karyawan diperbarui · ' + R.baru.length + ' karyawan baru · ' +
       R.hasil.filter(t => t.email).length + ' dari ' + R.hasil.length + ' punya email</p>';
  if (R.ubah.length) h += '<p><b>Diperbarui</b></p>' + daftar(R.ubah, u => '• <b>' + esc(u.id + ' ' + u.nama) + '</b>: ' + u.beda.map(b => esc(b.field) + (b.field === 'Basic Salary' ? ' (berubah)' : ' → ' + esc(b.ke))).join(', '));
  if (R.baru.length) h += '<p><b>Karyawan baru</b> (Peran default dari jabatan — periksa lagi)</p>' + daftar(R.baru, n => '• ' + esc(n.id + ' — ' + n.nama + ' · ' + n.jabatan + ' · ' + n.peran + (n.email ? ' · ' + n.email : ' · (tanpa email)')));
  if (R.hilang.length) h += '<p style="color:#9A6200"><b>⚠ Ada di sheet Karyawan tapi tidak ada di sumber (' + R.hilang.length + ')</b> — tidak dihapus; ubah Aktif = TIDAK bila sudah keluar</p>' + daftar(R.hilang, t => '• ' + esc(t.id + ' — ' + t.nama));
  if (R.emailSalah.length) h += '<p style="color:#B42318"><b>❌ Format email salah (dilewati)</b></p>' + daftar(R.emailSalah, esc);
  if (R.emailDobel.length) h += '<p style="color:#B42318"><b>❌ Email dipakai lebih dari 1 karyawan</b></p>' + daftar(R.emailDobel, esc);
  if (R.idDobel.length) h += '<p style="color:#9A6200"><b>⚠ Employee ID dobel di sumber (baris pertama yang dipakai)</b></p>' + daftar(R.idDobel, esc);
  if (!sudah && !R.ubah.length && !R.baru.length) h += '<p style="color:green">Data sudah sama, tidak ada yang perlu diubah. <span style="display:none">TIDAK_BISA</span></p>';
  if (sudah) h += '<p>Langkah berikut: menu ⚙️ SPL / SPD → <b>2. Cek data karyawan</b>.</p>';
  return h;
}

function pratinjauSinkron(namaSheet) {
  const S = bacaSumber_(namaSheet);
  return laporanSinkron_(rencanaSinkron_(bacaTargetKaryawan_(), S.rows), S.judul, false);
}

function terapkanSinkron(namaSheet) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('Sedang ada proses lain, coba lagi.');
  try {
    const S = bacaSumber_(namaSheet);
    const R = rencanaSinkron_(bacaTargetKaryawan_(), S.rows);
    const sh = sheet_(SH.KARYAWAN), c = kolom_(SH.KARYAWAN);
    const n = R.hasil.length;
    if (sh.getMaxRows() < n + 1) sh.insertRowsAfter(sh.getMaxRows(), n + 1 - sh.getMaxRows());
    // Tulis per kolom (kolom Departemen berisi rumus — tidak disentuh)
    const map = { 'Employee ID': 'id', 'Full Name': 'nama', 'Job Position': 'jabatan', 'Organization': 'org',
                  'Basic Salary': 'gaji', 'Email': 'email', 'Peran': 'peran', 'Aktif': 'aktif' };
    Object.keys(map).forEach(j => { if (c[j]) sh.getRange(2, c[j], n, 1).setValues(R.hasil.map(t => [t[map[j]]])); });
    // pastikan rumus Departemen ada di baris baru
    if (c['Departemen']) {
      const L = sh.getRange(2, c['Organization']).getA1Notation().replace(/\d+/, '');
      const f = [];
      for (let i = 2; i <= n + 1; i++) {
        const org = 'SUBSTITUTE(UPPER(TRIM(' + L + i + ')),"DEPARTMENT","DEPARTEMENT")';
        f.push(['=IF(' + L + i + '="","",IFERROR(LEFT(' + org + ',FIND(" GROUP",' + org + ')-1),' + org + '))']);
      }
      sh.getRange(2, c['Departemen'], n, 1).setFormulas(f);
    }
    SpreadsheetApp.flush();
    _K = null;
    logAksi_('-', '-', 'SINKRON KARYAWAN', Session.getActiveUser().getEmail(),
      'Sumber "' + namaSheet + '": ' + R.ubah.length + ' diperbarui, ' + R.baru.length + ' baru');
    return laporanSinkron_(R, S.judul, true);
  } finally { lock.releaseLock(); }
}

if (typeof module !== 'undefined') module.exports = { kenaliKolom_, rencanaSinkron_ };
