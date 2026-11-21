# TAR System – PT Harapan Mitra Lestari

Web App Technical Analysis Report (TAR) perawatan alat berat.
Node.js (Express) + PostgreSQL, frontend satu file (`public/index.html`), deploy via Dokploy.

## Struktur
```
server.js            seluruh backend (API, auth, workflow, upload, ekspor, webhook, seed)
public/index.html    seluruh antarmuka (CSS + JS inline) + logo.png
scripts/             manage-users.js, import-master.js (CLI)
seed/                units.csv, components.csv, users.csv (data awal dari database.xlsx)
Dockerfile, .env.example, package.json
```

## Tabel (dibuat otomatis saat start pertama)
`units`, `components` (komponen terdampak), `users`, `reports`, `approval_log` — sesuai permintaan.
Tabel tambahan yang diperlukan: `report_photos` (foto + caption + URL), `notifications` (bell),
`settings` (URL webhook), `tar_counters` (nomor urut TAR atomik; tidak terpakai ulang walau TAR dihapus).

## Deploy di Dokploy
1. **Postgres**: Create Service → Database → PostgreSQL. Catat *Internal Connection URL*.
2. **Aplikasi**: Create Service → Application, sumber Git (repo project ini) atau Docker image. Build type: **Dockerfile**.
3. **Environment** (tab Environment): isi variabel dari `.env.example`:
   `DATABASE_URL`, `JWT_SECRET` (`openssl rand -hex 32`), `APP_URL=https://tar.planthml.my.id`, `NODE_ENV=production`.
4. **Volume foto**: tab Advanced → Volumes → Add Volume Mount, tipe *Volume*, mount path `/data/uploads`.
   (Gunakan volume bernama, bukan bind mount, agar izin folder otomatis sesuai user `node`.)
5. **Domain**: tab Domains → host `tar.planthml.my.id`, port `3000`, HTTPS aktif (Let's Encrypt).
   Pastikan record DNS A domain menunjuk ke IP VPS.
6. Deploy. Saat pertama jalan, tabel dibuat dan data `seed/` dimasukkan (password di-hash bcrypt).
7. Login memakai NRP + password dari spreadsheet, lalu **minta semua user mengganti password** lewat tombol *Password*.

## Lupa password (via webhook n8n/WAHA)
Tombol *Lupa password?* di halaman login meminta username, lalu server membuat **tautan sekali pakai** (berlaku 30 menit, hanya hash-nya yang disimpan)
dan mengirimnya ke URL webhook "Lupa password" (Admin Panel → Webhook Notifikasi). Password tidak pernah dikirim. Payload JSON:
`event:"password_reset", message, link, expires_minutes, recipient{nama, username, wa_number, email}`.
Di n8n: Webhook → WAHA "send text" ke `recipient.wa_number` dengan teks `message` (atau node Email ke `recipient.email`).
Syarat: nomor WA/email user sudah diisi di Admin Panel → User. Admin tetap bisa reset manual lewat Admin Panel.

## Download semua TAR (PDF/ZIP)
Search TAR (Group Leader & Super admin) → **Download PDF**: mengunduh TAR sesuai filter yang tampil (maks 200 per unduhan).
1 TAR = file PDF, 2+ TAR = ZIP berisi PDF per TAR, dengan atau tanpa nilai. PDF dibuat server (pdfkit; foto diolah dengan sharp).

## Foto profil
Klik lingkaran foto di sidebar untuk mengganti/menghapus foto (dipotong persegi 400x400, JPG/PNG/WEBP, maks 3 MB, isi file dicek server).
Foto disimpan di `UPLOAD_DIR/avatars` (ikut volume `/data/uploads`) dan hanya bisa dilihat user yang sudah login.
Admin Panel → Pop Up dapat menampilkan foto + nama seorang user pada pop up pengumuman.

## Tanda tangan elektronik (QR)
Lembar cetak memuat QR untuk Mekanik (saat submit) dan Group Leader (saat approve). QR dibuat server (library `qrcode`, gratis)
dan berisi tautan `https://<domain>/verify/<no TAR>/<mekanik|leader>?c=<kode>`. Kode = HMAC-SHA256 atas no TAR, penandatangan, waktu,
dan hash isi laporan, jadi dokumen yang diubah/dipalsukan otomatis tidak valid. Halaman verifikasi bersifat publik (tanpa login)
dan hanya menampilkan nama, NRP, waktu, serta status — bukan isi laporan.
Isi `SIGN_SECRET` di Environment Dokploy (sekali, jangan diganti) agar QR lama tetap valid.

## Keamanan – wajib dibaca
- `seed/users.csv` berisi password awal **teks biasa** (dari spreadsheet). Pakai repo Git **private**, dan setelah deploy pertama
  kosongkan kolom password / hapus file itu dari repo (seed hanya berjalan bila tabel `users` kosong).
- Login: JWT di cookie `httpOnly`, `SameSite=Strict`, `Secure` (production); rate-limit 15 percobaan/15 menit/IP.
- Foto tidak ada di folder publik; hanya dilayani lewat `/api/photos/:id` setelah cek login dan hak akses laporan.
- Role dicek di server pada setiap endpoint.

## CLI
```
npm run users -- list
npm run users -- add 10199999 "Nama Baru" "Mekanik" [password]
npm run users -- passwd 10199999 [password]
npm run import -- units seed/units.csv
npm run import -- components seed/components.csv
```
Di Dokploy jalankan lewat tab *Terminal* pada container aplikasi (`node scripts/manage-users.js list`).

## Webhook (n8n + WAHA)
Admin Panel → Webhook Notifikasi: isi URL untuk 3 kejadian (`submitted`, `approved`, `rejected`).
Payload JSON: `event, message, tar_number, judul, code_unit, status, owner{}, leader{}, recipient{nama, wa_number, email}, link`.
Di n8n cukup Webhook → WAHA "send text" ke `recipient.wa_number`. Isi nomor WA user (format `62812…`) di Admin Panel → User.

## Pengembangan lokal
```
cp .env.example .env   # isi DATABASE_URL lokal
npm install && npm run dev
```
