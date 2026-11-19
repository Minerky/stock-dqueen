# Stok DQueen — Atelier Persediaan

Aplikasi manajemen stok dan mutasi barang lokal DQueen tanpa runtime dependency eksternal (menggunakan bawaan Node.js stdlib `node:sqlite` + `node:zlib`).

## Fitur Utama
- **Struktur Sheet Kategori**: Kardus, Nampan, Standing, Kantor, Pita, Box.
- **Tab Semua Barang**: Katalog lengkap 1.217 barang dengan pencarian cepat.
- **Input Mutasi Taktil**: Tombol `Masuk` dan `Keluar` instan.
- **Ekspor Excel Multi-Sheet**: Format `.xlsx` rapi dengan freeze pane, auto filter, dan warna status.
- **Latar Belakang**: Server berjalan otomatis di latar belakang tanpa jendela CMD.

## Cara Menjalankan
1. Pastikan Node.js (v22+) terpasang di komputer.
2. Klik ganda `start.bat` atau `silent.vbs`.
3. Aplikasi otomatis terbuka di browser pada `http://localhost:3001`.
4. Untuk mematikan server, klik ganda `stop.bat`.
