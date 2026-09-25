# Catatan pemeriksaan workbook klien

Gunakan daftar ini saat klien mengirim file. Tidak semua pertanyaan perlu ditanyakan sekaligus. Banyak jawabannya mungkin sudah terlihat dari workbook.

## Pengguna dan cara kerja

1. Siapa yang mengisi workbook setiap hari?
2. Siapa yang memeriksa dan menyetujui hasilnya?
3. Data apa yang sekarang harus diketik atau disalin lebih dari sekali?
4. Apakah tim memakai satu file bersama atau membuat salinan untuk setiap pekerjaan?
5. Versi Excel apa yang dipakai oleh anggota tim?

## Data dan rumus

6. Sheet mana yang dianggap sebagai sumber resmi data standar?
7. Siapa yang boleh memperbarui data tersebut?
8. Rumus atau metode mana yang memerlukan persetujuan sebelum diubah?
9. Apakah workbook memakai VBA, Power Query, external link, add-in, named range, atau formula array?
10. Berapa jumlah baris pada pemakaian normal dan saat paling ramai?
11. Apakah data lama perlu dipindahkan ke format baru?

## Laporan dan pemeriksaan hasil

12. Laporan mana yang dipakai untuk mengambil keputusan atau dikirim ke pihak lain?
13. Adakah beberapa contoh pekerjaan yang hasilnya sudah diperiksa dan dinyatakan benar?
14. Bagian mana yang boleh diubah teknisi?
15. Apakah tim memerlukan catatan perubahan per pengguna atau cukup change log di tingkat workbook?

## Syarat penerimaan

Sebelum file dinyatakan selesai, periksa hal berikut:

- pengguna dapat mengenali semua sel yang perlu diisi;
- data yang sebelumnya diketik berulang sudah diambil dari sumber yang disepakati;
- pemilik dan cara memperbarui master data sudah dicatat;
- formula serta konfigurasi penting tidak bisa terhapus saat teknisi mengisi data;
- hasil kasus acuan sama dengan baseline atau selisihnya sudah disetujui;
- laporan yang masuk ruang lingkup dapat dibuat tanpa copy-paste;
- temuan UAT dengan prioritas tinggi sudah selesai; dan
- panduan penggunaan sudah diberikan kepada tim.

## Cara memeriksa

| Bagian | Percobaan | Hasil yang diharapkan |
|---|---|---|
| Input wajib | Kosongkan salah satu kolom wajib | Workbook menampilkan peringatan atau status belum lengkap. |
| Master data | Pilih beberapa kode aset | Deskripsi dan standar sesuai dengan tabel sumber. |
| Rumus teknis | Jalankan kasus acuan | Hasil sama dengan baseline yang disetujui. |
| Toleransi | Coba nilai di bawah, tepat pada, dan di atas batas | Status mengikuti aturan yang disepakati. |
| Proteksi | Coba mengubah sel formula sebagai teknisi | Excel menolak perubahan. |
| Laporan | Tambahkan transaksi baru | Rekap ikut berubah tanpa pemindahan data manual. |
| Penelusuran | Pilih satu hasil di laporan | Input dan referensi asalnya dapat ditemukan. |

## Catatan perubahan rumus

Untuk setiap rumus atau metode yang berubah, simpan lokasi rumus lama, alasan perubahan, rumus baru, contoh pengaruhnya pada hasil, nama pemberi persetujuan, tanggal persetujuan, dan hasil pemeriksaan setelah perubahan.
