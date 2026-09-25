# Naskah video

Target durasi sekitar tiga menit. Jangan dibaca persis seperti naskah. Pakai kalimat yang terasa wajar saat diucapkan.

## Pembuka, 0:00 sampai 0:25

Halo, saya [Nama Anda]. Saya pernah mengerjakan perbaikan workbook yang punya banyak sheet dan rumus saling terhubung.

File proyek lama saya sudah tidak ada, jadi saya membuat ulang contoh sederhana dengan data fiktif. Saya ingin memperlihatkan cara saya menata workbook seperti yang dijelaskan di brief Anda.

## Sheet START HERE, 0:25 sampai 0:45

Saya membagi workbook ini menjadi area input, master data, perhitungan, dan laporan. Teknisi bekerja di sheet input. Sheet lain dipakai untuk menyimpan referensi, memeriksa rumus, dan membaca hasil.

Sel kuning dipakai untuk input. Sel biru berisi rumus. Pada demo ini proteksinya sengaja saya matikan supaya rumus bisa diperiksa. Nanti proteksi dapat diaktifkan saat file masuk ke tahap UAT.

## Sheet INPUT LOG, 0:45 sampai 1:35

Teknisi mengisi nomor pekerjaan, tanggal, nama, Asset ID, dan tiga hasil pengukuran.

Begitu Asset ID dipilih, workbook mengambil nama alat, lokasi, unit, nilai standar, dan toleransi dari master data. Teknisi tidak perlu mengetik data yang sama lagi.

Saya juga memasang dropdown dan pemeriksaan nilai. Baris yang belum lengkap diberi status INCOMPLETE. Jika rata-rata pengukuran melewati toleransi, statusnya berubah menjadi REVIEW.

## MASTER DATA dan CALCULATION, 1:35 sampai 2:10

Semua referensi alat disimpan di MASTER DATA. Saat nilai standar berubah, pengelola cukup memperbarui tabel ini.

CALCULATION memperlihatkan angka yang dipakai untuk menentukan status. Reviewer bisa mengikuti hasil pengukuran, nilai standar, toleransi, rata-rata, dan deviasinya dalam satu baris.

Rumus di contoh ini hanya untuk demonstrasi. Pada workbook tim Anda, saya akan memakai metode yang sudah disetujui. Jika ada usulan perubahan, saya akan membahasnya lebih dulu dengan penanggung jawab teknis.

## REPORT dan VALIDATION, 2:10 sampai 2:40

REPORT membaca data langsung dari input. Tidak ada pemindahan angka secara manual.

Di VALIDATION, saya menyiapkan lima kasus acuan. Hasil hitung workbook dibandingkan dengan hasil yang diharapkan. Cara yang sama bisa dipakai untuk memeriksa workbook baru terhadap file lama sebelum digunakan oleh tim.

## Penutup, 2:40 sampai 3:05

Jika mendapat akses ke workbook yang sekarang, saya akan mulai dengan memetakan dependensi dan memilih beberapa kasus acuan. Setelah itu baru saya merapikan strukturnya dan menyiapkan versi untuk diuji.

Perkiraan awal saya 8 sampai 12 hari kerja. Angka pastinya bergantung pada jumlah sheet, formula, external link, serta penggunaan VBA atau Power Query.

Terima kasih sudah menonton. Saya siap memulai dari pemeriksaan workbook versi terbaru.

## Caption video

Saya pernah mengerjakan perbaikan workbook sejenis, tetapi file proyek lamanya sudah tidak tersedia. Video ini memakai rekonstruksi dengan data fiktif untuk memperlihatkan cara saya menata input, master data, rumus, laporan, dan pemeriksaan hasil. Untuk workbook tim Anda, saya akan memeriksa file serta menyepakati kasus acuan sebelum mengubah rumus teknis.
