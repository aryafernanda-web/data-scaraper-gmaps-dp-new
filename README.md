# Cek Coverage DP — FTTH

Aplikasi untuk mengecek apakah titik-titik data pelanggan berada dalam radius
tertentu (default 350 m) dari titik DP FTTH. Pemrosesan utama (baca Excel,
hitung jarak, geocoding, export ke Excel) tetap berjalan di browser
pengguna, tanpa backend. Satu-satunya bagian yang jalan di server adalah
1 serverless function kecil untuk fitur export ke Notion (lihat bagian
"Fitur baru: Export Hasil ke Notion" di bawah) — dibutuhkan supaya API key
Notion tidak pernah muncul di kode browser.

## Isi folder
- `index.html` — seluruh aplikasi (HTML + CSS + JS jadi satu file)
- `api/notion-export.js` — serverless function (proxy aman ke Notion API)
- `vercel.json` — konfigurasi minimal untuk Vercel

## Cara deploy ke Vercel

### Opsi A — lewat Vercel CLI (paling cepat)
```bash
npm install -g vercel
cd folder-ini
vercel
```
Ikuti prompt-nya (login, pilih scope/project). Setelah selesai, Vercel akan
kasih URL live. Untuk deploy ke production langsung:
```bash
vercel --prod
```

### Opsi B — lewat Vercel Dashboard (drag & drop, tanpa CLI)
1. Buka https://vercel.com/new
2. Pilih **"Deploy a folder"** / drag folder ini ke area upload
   (atau upload lewat Git jika sudah kamu push ke GitHub/GitLab/Bitbucket)
3. Framework preset: pilih **"Other"** (bukan Next.js dkk) — Vercel akan
   otomatis mendeteksi `index.html` sebagai static site
4. Klik **Deploy**

### Opsi C — lewat GitHub
1. Push folder ini ke repo GitHub baru
2. Di Vercel dashboard: **New Project** → **Import Git Repository** → pilih repo itu
3. Framework preset: **Other** → Deploy

## Catatan
- Tidak ada environment variable atau API key yang perlu diatur — semua
  layanan yang dipakai (SheetJS via CDN, Nominatim untuk geocoding alamat)
  publik dan gratis.
- Data DP (359 titik) sudah tertanam langsung di dalam `index.html`. Kalau
  data DP berubah di kemudian hari, file `index.html` perlu diupdate ulang.

## Fitur baru: Export Hasil ke Notion
Selain export ke Excel, sekarang ada tombol **"Export Hasil ke Notion"** yang
mengirim setiap baris hasil langsung sebagai page baru ke database Notion
"Data Scraping".

Karena Notion API tidak bisa dipanggil langsung dari browser (CORS) dan API
key tidak boleh ditaruh di kode client, ditambahkan **1 serverless function**
di `api/notion-export.js`. Ini satu-satunya bagian yang jalan di server;
semua pemrosesan data (baca Excel, hitung jarak, geocoding) tetap 100% di
browser seperti sebelumnya.

### Setup (wajib sebelum tombol "Export ke Notion" bisa jalan)
1. **Share database Notion "Data Scraping" ke integration-nya** (kalau
   belum): buka database di Notion → `...` (kanan atas) → `Connect to` →
   pilih integration kamu.
2. **Isi environment variable di Vercel Dashboard** → project ini →
   **Settings → Environment Variables** → tambahkan:
   - `NOTION_API_KEY` = token integration Notion kamu
   - `NOTION_DATA_SOURCE_ID` = `3c3dcd14-e2c8-8040-8462-000bc13161ec`
     (sudah jadi default di `api/notion-export.js`, isi ini kalau mau
     override ke database lain)
3. **Redeploy** (`vercel --prod` atau lewat dashboard) supaya env var
   terbaca.

Kalau env variable belum diisi, tombol "Export ke Notion" tetap bisa
diklik tapi akan menampilkan pesan gagal untuk setiap baris — export ke
Excel tidak terpengaruh sama sekali.

### Troubleshooting: "Export ke Notion" gagal terus
Cek satu-satu urutan di bawah ini — ini mencakup hampir semua penyebab umum.

1. **Pastikan `notion-export.js` ada di dalam folder `api/`**, persis di path
   `api/notion-export.js` dari root project. Vercel hanya mengenali file di
   dalam folder `api/` sebagai serverless function. Kalau filenya ada di
   root project (sejajar dengan `index.html`), endpoint `/api/notion-export`
   tidak akan pernah ada.
2. **Test endpoint-nya langsung di browser**: buka
   `https://<domain-vercel-kamu>/api/notion-export`.
   - Kalau muncul halaman **404 "This page could not be found" dari Vercel**
     → function belum ke-deploy dengan benar (balik ke poin 1).
   - Kalau muncul JSON seperti `{"ok":true,"info":"...","hasApiKey":false,...}`
     → function-nya sudah jalan, tapi env var `NOTION_API_KEY` belum
     terbaca. Lanjut ke poin 3.
   - Kalau `hasApiKey: true` → function & env var sudah OK, masalahnya ada
     di sisi Notion (poin 4 atau 5).
3. **Env var harus di-redeploy.** Mengisi `NOTION_API_KEY` di Vercel
   Dashboard **tidak otomatis berlaku** untuk deployment yang sudah ada.
   Setelah isi/ubah env var: Deployments → deployment terakhir → menu `...`
   → **Redeploy** (atau `vercel --prod` lagi dari CLI). Pastikan juga env
   var itu dicentang untuk environment **Production** (bukan cuma
   Development/Preview).
4. **Integration Notion harus di-share ke database "Data Scraping".** Buka
   database itu di Notion → `...` (kanan atas) → `Connect to` → pilih
   integration yang API key-nya dipakai. Kalau belum di-share, Notion akan
   menolak request dengan error `object_not_found` walau API key valid.
5. **Lihat pesan error aslinya.** Buka DevTools browser (F12) → tab
   Console, lalu klik lagi "Export Hasil ke Notion" — tiap baris yang gagal
   akan mencetak detail error dari Notion (`console.error('Gagal export
   baris ke Notion:', row, data)`). Bisa juga cek **Vercel Dashboard →
   project ini → Deployments → deployment aktif → tab Functions/Logs**
   untuk log dari sisi server.

## Skema export Excel
Kolom hasil export (`Export Hasil ke Excel`) sudah disamakan dengan skema
database Notion "Data Scraping" supaya bisa langsung di-import:

| Kolom Excel | Isi |
|---|---|
| Name | Nama pelanggan |
| Phone | Nomor telepon |
| Alamat | Alamat pelanggan |
| Maps | Link Google Maps |
| Coverage | `Cover` / `Uncover` (kosong kalau baris invalid) |
| Category | Jarak ke DP terdekat (+ catatan est. alamat/invalid) |
