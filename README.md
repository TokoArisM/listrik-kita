# LISTRIK KITA - PWA POS KASIR (100% Bebas Biaya Server)

Aplikasi Point of Sale (POS) Kasir modern berbasis PWA (Progressive Web App) untuk toko alat listrik **LISTRIK KITA**. Sistem ini tidak membutuhkan biaya server (0 rupiah), berjalan langsung di browser HP/Tablet/Laptop dengan backend Google Apps Script + Google Sheets, serta terhubung langsung ke Printer Thermal Bluetooth 58mm dan Laci Kasir Otomatis (RJ11).

---

## 📁 Struktur File Proyek

```text
d:\APP ARIS LISTRIK KITA\
│
├── Code.gs             # Backend Google Apps Script (doGet & doPost)
├── index.html          # Frontend POS UI (Tailwind CSS, Web Bluetooth, ESC/POS)
├── manifest.json       # Konfigurasi PWA (Add to Home Screen)
├── sw.js               # Service Worker untuk Caching & Offline PWA
├── icons/              # Ikon aplikasi PWA (192x192 & 512x512)
│   ├── icon-192.png
│   └── icon-512.png
└── README.md           # Panduan Lengkap Instalasi & Penggunaan
```

---

## 🛠️ 1. SETUP BACKEND (Google Sheets & Apps Script)

1. Buka [Google Sheets](https://sheets.new) dan buat spreadsheet baru dengan nama **DB POS LISTRIK KITA**.
2. Buat sheet tab pertama bernama **`Produk`** dengan kolom baris 1:
   - `ID` | `Nama` | `Kategori` | `Harga` | `Stok` | `Satuan`
3. Buat sheet tab kedua bernama **`Transaksi`** dengan kolom baris 1:
   - `ID Transaksi` | `Waktu` | `Detail Item` | `Total Belanja` | `Bayar` | `Kembalian` | `Metode`
4. Di Google Sheets, buka menu **Extensions (Ekstensi)** -> **Apps Script**.
5. Salin seluruh isi file [`Code.gs`](file:///d:/APP%20ARIS%20LISTRIK%20KITA/Code.gs) dan tempel ke editor Apps Script.
6. Anda bisa menjalankan fungsi `setupDatabase()` satu kali untuk mengisi data awal produk contoh toko listrik.
7. Klik tombol biru **Deploy** di pojok kanan atas -> pilih **New deployment**:
   - Pilih jenis: **Web app**
   - *Execute as*: **Me (email Anda)**
   - *Who has access*: **Anyone (Siapa saja)** *(Sangat penting agar PWA bisa membaca & menulis)*
8. Klik **Deploy**, beri izin Google Account, lalu salin **Web App URL** (berakhiran `/exec`).

---

## 🌐 2. SETUP FRONTEND & PWA

### Catatan Keamanan Browser:
Fitur **PWA** dan **Web Bluetooth API** (`navigator.bluetooth`) membutuhkan koneksi **HTTPS** atau **localhost**:
- Untuk pengujian di komputer: Buka menggunakan ekstensi VS Code *Live Server* atau jalankan terminal:
  ```powershell
  npx serve .
  ```
- Untuk produksi / dipakai kasir di HP Android: Unggah folder ini ke **GitHub Pages**, **Cloudflare Pages**, atau **Vercel** (semuanya 100% gratis & otomatis HTTPS).

### Menghubungkan ke Google Sheets:
1. Buka aplikasi di browser (misal Chrome di Android atau PC).
2. Klik ikon gear **Pengaturan (Settings)** di navbar kanan atas.
3. Tempelkan **Web App URL** dari Google Apps Script yang sudah Anda salin di langkah sebelumnya.
4. Klik **Simpan & Sinkronkan**. Katalog produk akan otomatis dimuat langsung dari sheet `Produk`.

---

## 🖨️ 3. KONEKSI PRINTER THERMAL & LACI KASIR (RJ11)

1. Pastikan Bluetooth di HP/Laptop dalam keadaan **Aktif** dan Printer Thermal 58mm menyala.
2. Colokkan kabel kabel RJ11 dari laci kasir (Cash Drawer) ke port RJ11 di belakang printer thermal.
3. Di navbar aplikasi kasir, klik tombol **"Hubungkan Printer"**.
4. Browser akan memunculkan dialog pencarian Bluetooth (`navigator.bluetooth.requestDevice`).
5. Pilih printer thermal Anda (contoh: *RPP02N*, *POS-58*, *MPT-II*, dsb) lalu klik **Pair/Hubungkan**.
6. Saat kasir menekan tombol **"Konfirmasi & Cetak Struk"**:
   - Printer akan mencetak struk belanja 32 kolom rapi.
   - Di baris paling akhir, perintah ESC/POS byte array `[27, 112, 0, 25, 250]` otomatis dikirim ke printer untuk men-*trigger* kick pin solenoid RJ11 sehingga **laci kasir terbuka secara otomatis**.
