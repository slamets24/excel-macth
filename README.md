# Excel Match

MVP aplikasi web statis untuk membandingkan dua file Excel langsung di browser. Aplikasi ini tidak memakai database dan tidak mengunggah isi file ke server.

Target repository: https://github.com/slamets24/excel-macth.git

## Fitur MVP

- Upload dua file `.xlsx`, `.xls`, atau `.csv`.
- Pilih sheet dari masing-masing file.
- Baca otomatis kolom mitra dari header `NAMA TOKO` atau `NAMA MITRA`.
- Baca otomatis warna dari header warna dan size dari baris `XS`, `S`, `M`, `L`, `XL`, `XXL`, `XXXL`.
- Tampilkan hanya data mitra + warna + size yang qty-nya berbeda.
- Lihat qty File A, qty File B, dan selisihnya.
- Filter, cari hasil diff, dan export hasil ke Excel.

## Stack

- HTML, CSS, JavaScript browser-native
- SheetJS via CDN
- Vercel static hosting
- Tanpa database

## Deploy ke Vercel

1. Import repository `slamets24/excel-macth` di Vercel.
2. Framework preset: `Other` atau static site.
3. Build command: kosongkan.
4. Output directory: `.` atau root project.
5. Deploy.

## Privacy

File Excel diproses di browser pengguna. Aplikasi ini tidak memiliki API upload, backend, database, localStorage, atau IndexedDB.
