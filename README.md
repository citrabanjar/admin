# Sistem SPL & SPD

Sistem **Surat Perintah Lembur (SPL)** dan **Surat Perintah Dinas (SPD)** berbasis Google Sheets + Google Apps Script.

- **Preview form (demo, data fiktif):** GitHub Pages → folder [`docs/`](docs/) — login: email apa saja, kode `123456`
- **Kode Apps Script:** folder [`apps-script/`](apps-script/)

## Struktur
| Folder / file | Isi |
|---|---|
| `src/` | **Kode Apps Script** — otomatis dikirim ke Google Apps Script oleh GitHub Actions setiap ada perubahan di `main` |
| `src/Inti.gs` | Fungsi inti: baca master & parameter, rumus lembur, rumus SPD, nomor surat, log |
| `src/Menu.gs` | Menu `⚙️ SPL / SPD` di spreadsheet, setup awal, cek data karyawan, tes hitung |
| `src/Web.gs` | Web App: tampilan, login OTP email, data form |
| `src/Spl.gs` / `src/Spd.gs` | Simpan & riwayat SPL / SPD |
| `src/Notif.gs` | Email konfirmasi karyawan, persetujuan atasan, PDF kop surat ke admin |
| `src/Api.gs` | Pintu data hanya-baca (JSON) untuk pengecekan pengajuan lembur KantorKu |
| `src/Logo.gs` | Logo perusahaan (base64) |
| `src/Form.html` | Tampilan form |
| `src/appsscript.json` | Manifest (zona waktu Asia/Jakarta, Web App: execute as owner, akses Anyone) |
| `manual/Kode_Lengkap.gs` | Semua .gs digabung — cadangan bila perlu tempel manual |
| `docs/index.html` | **Pintu masuk form** di https://citrabanjar.github.io/admin/ (memuat Web App Apps Script) |
| `docs/demo/` | Demo tampilan (data fiktif) |
| `.github/workflows/deploy-apps-script.yml` | Otomatis `clasp push` + update deployment |

## Deploy otomatis (GitHub → Apps Script)
Isi 3 secret di **Settings → Secrets and variables → Actions**: `CLASPRC_JSON`, `SCRIPT_ID`, `DEPLOYMENT_ID`.
Setelah itu setiap commit ke `main` yang mengubah `src/` langsung terpasang di Apps Script dan link Web App tetap sama.

> Data karyawan & gaji **tidak** disimpan di repo ini.

## Tahapan
- [x] Tahap 1 — Database & script inti
- [x] Tahap 2–3 — Form SPL & SPD (logo perusahaan), nomor otomatis, simpan
- [x] Tahap 4 — Email SPL: konfirmasi karyawan → persetujuan atasan langsung → PDF ke admin
- [x] Tahap 5 — PDF SPL dengan kop surat (otomatis saat semua karyawan selesai diproses)
- [x] Email & PDF untuk SPD (alur sama dengan SPL, PDF mengikuti formulir Excel A–F)
- [ ] Tahap 6 — Panel admin & rekap
