# Sistem SPL & SPD

Sistem **Surat Perintah Lembur (SPL)** dan **Surat Perintah Dinas (SPD)** berbasis Google Sheets + Google Apps Script.

- **Preview form (demo, data fiktif):** GitHub Pages → folder [`docs/`](docs/) — login: email apa saja, kode `123456`
- **Kode Apps Script:** folder [`apps-script/`](apps-script/)

## Isi `apps-script/`
| File | Fungsi |
|---|---|
| `Inti.gs` | Fungsi inti: baca master & parameter, rumus lembur, rumus SPD, nomor surat, log |
| `Menu.gs` | Menu `⚙️ SPL / SPD` di spreadsheet, setup awal, cek data karyawan, tes hitung |
| `Spl.gs` | Web App form SPL: login OTP email, simpan SPL, riwayat |
| `FormSPL.html` | Tampilan form SPL (dipakai oleh `Spl.gs`) |

## Pasang
1. Database Google Sheet (sheet: Karyawan, Parameter, Pilihan, Libur, Data SPL, Data SPD, Personel SPD, Biaya SPD, Log).
2. Extensions → Apps Script → salin semua file di `apps-script/` (nama file harus sama).
3. Deploy → New deployment → Web app — Execute as **Me**, access **Anyone**.

> Data karyawan & gaji **tidak** disimpan di repo ini.

## Tahapan
- [x] Tahap 1 — Database & script inti
- [x] Tahap 2–3 — Form SPL, nomor otomatis, simpan
- [ ] Tahap 4 — Email notifikasi & approval Manager / HRD
- [ ] Tahap 5 — Cetak PDF surat
- [ ] Tahap 6 — Panel admin & rekap
- [ ] Form SPD
