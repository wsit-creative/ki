SISTEM ORDER BARANG SUBSIDI — PANDUAN SINGKAT

File:
- Code.gs: backend Google Apps Script
- index.html: seluruh HTML + CSS + JavaScript frontend dalam satu file
- README.txt: panduan pemasangan

PENTING: Versi ini menyajikan halaman lewat Apps Script HtmlService agar google.script.run bekerja tanpa masalah CORS GitHub Pages. Buka URL Web App Apps Script. Jika harus memakai GitHub Pages, perlu API proxy terpisah (Cloud Run/Functions) dan penyesuaian CORS; jangan kirim NIK lewat URL GET/JSONP.

PEMASANGAN
1. Buat Google Sheets baru untuk pengujian.
2. Spreadsheet > Ekstensi > Apps Script.
3. Tempel Code.gs ke file Code.gs.
4. Tambahkan file HTML bernama index, lalu tempel isi index.html.
5. Jalankan setupSheets() sekali dan izinkan akses.
6. Project Settings > Script Properties, tambahkan dua nilai rahasia yang berbeda:
   - NIK_HMAC_KEY: string acak kuat minimal 32 karakter.
   - ADMIN_AUTH_KEY: string acak kuat minimal 32 karakter.
7. Di Code.gs, ganti nilai password pada oneTimeSetAdminPasswordHmac() dengan password admin unik minimal 14 karakter. Jalankan fungsi itu sekali. Setelah sukses, HAPUS fungsi oneTimeSetAdminPasswordHmac() dari source, lalu simpan. Password tidak boleh tertinggal di source.
8. Deploy > New deployment > Web app. Execute as: Me. Who has access: Anyone jika pelanggan tanpa login Google harus menggunakan halaman. Karena endpoint publik, uji dan pantau dengan hati-hati.
9. Buka URL Web App, login admin, tambahkan barang, lalu buat aturan promo NIK per barang dan interval hari.
10. Uji hanya dengan NIK dummy 16 digit dulu.

CARA KERJA
- NIK dimasukkan oleh admin di halaman promo. Backend langsung menghitung HMAC-SHA-256 menggunakan NIK_HMAC_KEY; NIK asli tidak disimpan di Sheets.
- Setiap aturan mengikat NIK (HMAC) ke satu barang dan interval hari.
- Pelanggan memilih barang dan memasukkan NIK; backend memeriksa aturan aktif dan pengambilan terakhir untuk barang itu.
- Stok dicadangkan ketika pesanan dibuat dan dikembalikan jika pesanan dibatalkan.
- Masa tunggu dimulai saat admin menekan “Selesai diambil”, berdasarkan waktu server, bukan waktu browser.
- LockService mengurangi risiko dua pesanan bersamaan mengubah stok atau mengonfirmasi NIK yang sama.

BATASAN KEAMANAN
- HMAC bukan enkripsi yang bisa dibalik; NIK asli tidak dapat dipulihkan dari HMAC. Ini cocok untuk pencocokan, tetapi bukan jika perlu menampilkan NIK kembali.
- Endpoint publik masih dapat dicoba secara otomatis. Untuk produksi, tambahkan rate limiting kuat melalui API gateway/proxy, pemantauan abuse, dan proteksi tambahan. LockService bukan rate limiter.
- Sesi admin contoh memakai CacheService dan dapat berakhir lebih awal. Untuk penggunaan kritis, gunakan sistem identitas admin yang teruji.
- Spreadsheet dan proyek Apps Script hanya boleh diakses petugas yang berwenang. Editor Apps Script dapat membaca Script Properties/kode dan karena itu berpotensi mengakses kunci.
- Data nama/telepon/catatan pesanan masih tersimpan di sheet Orders. Batasi akses dan tentukan retensi.
- Uji semua alur dengan data dummy, backup, dan pastikan aturan sesuai ketentuan subsidi serta UU Perlindungan Data Pribadi yang berlaku.


FITUR CHECKOUT TAMBAHAN (versi revisi)
- Pembeli dapat menambah beberapa barang ke keranjang dan mengubah jumlah sebelum checkout.
- Checkout meminta nama, nomor WhatsApp, dan alamat. Tombol “Gunakan Alamat Saat Ini” meminta izin GPS dan mencoba mengubah koordinat menjadi alamat; hasil wajib diperiksa pembeli. Jika layanan alamat gagal, koordinat dan tautan peta diisi sebagai cadangan sehingga alamat tetap perlu dilengkapi.
- Kolom voucher/NIK opsional. Namun, bila barang memiliki aturan voucher aktif, backend tetap mewajibkan NIK yang terdaftar untuk barang tersebut.
- Pesanan tersimpan di sheet Orders, termasuk alamat dan ID grup keranjang. Setiap jenis barang disimpan sebagai baris pesanan tersendiri.
- Setelah penyimpanan berhasil, WhatsApp admin 082266897579 dibuka dengan teks berisi ID grup, nama, nomor WhatsApp, alamat, rincian barang/jumlah, dan catatan. Pembeli masih harus menekan tombol Kirim di WhatsApp; website biasa tidak dapat mengirim pesan WhatsApp otomatis tanpa integrasi WhatsApp Business API.
- Jika spreadsheet sebelumnya sudah dibuat, jalankan kembali setupSheets() untuk menambahkan kolom customer_address dan cart_group_id ke sheet Orders. Fungsi ini mempertahankan data lama dan menambahkan header baru di akhir.
