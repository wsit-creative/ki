SISTEM ORDER BARANG SUBSIDI — VERSI GITHUB PAGES

ISI PAKET
- index.html: frontend toko (HTML, CSS, JavaScript dalam satu file) untuk GitHub Pages.
- Code.gs: backend Google Apps Script yang mengakses Google Sheets.
- Bridge.html: jembatan komunikasi iframe antara GitHub Pages dan Apps Script. Tidak memakai JSONP/GET untuk mengirim NIK atau data checkout.
- README.txt: panduan pemasangan.

ARSITEKTUR
GitHub Pages (toko) -> iframe Bridge.html di Apps Script -> google.script.run -> Google Sheets.
Pola bridge ini menghindari pemanggilan google.script.run langsung dari GitHub Pages dan tidak perlu mengirim NIK/pesanan lewat query URL. Jangan menghapus validasi server.

LANGKAH 1 — GOOGLE SHEETS DAN APPS SCRIPT
1. Buat Google Sheets baru khusus pengujian.
2. Dari spreadsheet pilih Ekstensi > Apps Script.
3. Tempel isi Code.gs ke file Code.gs.
4. Klik + > HTML, beri nama persis Bridge (Apps Script akan membuat Bridge.html), lalu tempel isi Bridge.html.
5. Simpan proyek. Jalankan setupSheets() sekali dan izinkan akses yang diminta.

LANGKAH 2 — SCRIPT PROPERTIES
Buka Project Settings > Script Properties, lalu tambahkan:
- NIK_HMAC_KEY: rahasia acak minimal 32 karakter.
- ADMIN_AUTH_KEY: rahasia acak berbeda minimal 32 karakter.
- GITHUB_PAGES_ORIGIN: origin situs GitHub Pages, TANPA path dan tanpa garis miring terakhir. Contoh jika situs https://namakamu.github.io/nama-repo/ maka isi https://namakamu.github.io. Jika memakai custom domain, isi origin domain tersebut, misalnya https://tokoku.example.

LANGKAH 3 — PASSWORD ADMIN
1. Di Code.gs cari oneTimeSetAdminPasswordHmac(). Ganti placeholder GANTI_DENGAN_PASSWORD_ADMIN_KUAT dengan password admin unik minimal 14 karakter.
2. Jalankan fungsi oneTimeSetAdminPasswordHmac() sekali dari editor.
3. Setelah sukses, hapus fungsi oneTimeSetAdminPasswordHmac() dari source code lalu simpan. Jangan meninggalkan password asli di source.

LANGKAH 4 — DEPLOY BACKEND APPS SCRIPT
1. Klik Deploy > New deployment > pilih jenis Web app.
2. Execute as: Me.
3. Who has access: Anyone (agar pembeli tanpa login Google dapat memesan), jika opsi ini tersedia dan sesuai kebijakan akunmu.
4. Deploy dan salin URL Web App yang berakhir /exec. Jangan menambahkan ?bridge=1 ke URL dasar saat menyimpannya; URL itu akan dipakai di langkah berikutnya.
5. Jika mengubah kode Apps Script setelah deployment, buka Deploy > Manage deployments > Edit, pilih New version, lalu Deploy.

LANGKAH 5 — GITHUB PAGES
1. Unggah index.html ke repository GitHub. Agar halaman utama terbuka otomatis, nama file harus index.html dan berada di root repository (atau folder yang dipilih untuk Pages).
2. Buka index.html di GitHub dan cari konstanta GAS_BRIDGE_URL di dekat awal script.
3. Ganti PASTE_URL_WEB_APP_APPS_SCRIPT_DI_SINI dengan URL Web App kamu, sehingga hasilnya seperti:
   const GAS_BRIDGE_URL='https://script.google.com/macros/s/ID_DEPLOYMENT/exec?bridge=1';
   Gunakan URL milikmu sendiri, bukan contoh ini.
4. Commit perubahan, lalu aktifkan GitHub Pages melalui Settings > Pages jika belum aktif.
5. Buka URL GitHub Pages melalui HTTPS. Pastikan origin situs (bagian https://namakamu.github.io) sama persis dengan GITHUB_PAGES_ORIGIN di Script Properties.

LANGKAH 6 — SETUP DAN UJI
1. Buka URL GitHub Pages.
2. Login admin, tambahkan produk, lalu tambahkan aturan voucher jika diperlukan.
3. Uji produk, keranjang, checkout, alamat GPS, penyimpanan sheet, login admin, status pesanan, pembatalan, dan konfirmasi pengambilan dengan data dummy.
4. Jangan gunakan NIK asli untuk uji coba awal. Pastikan stok/aturan voucher benar sebelum dipakai publik.

CARA KERJA FITUR
- NIK diubah di server menjadi HMAC-SHA-256; NIK asli tidak disimpan di sheet.
- Aturan voucher mengikat NIK ke satu barang dan interval hari.
- Jika barang memiliki aturan voucher aktif, pesanan barang tersebut wajib menyertakan NIK yang terdaftar.
- Stok dicadangkan ketika pesanan dibuat dan dikembalikan ketika pesanan pending/dibayar dibatalkan.
- Masa cooldown dimulai saat admin menekan “Selesai diambil”.
- Checkout menyimpan nama, WhatsApp, alamat, catatan, rincian barang dan ID grup ke sheet Orders.
- Setelah pesanan berhasil tersimpan, halaman menyiapkan pesan WhatsApp untuk admin 082266897579. Pembeli tetap perlu menekan Kirim di WhatsApp; pengiriman tanpa tindakan pembeli memerlukan WhatsApp Business API.
- Tombol alamat saat ini menggunakan izin GPS browser dan layanan geocoding eksternal. Pembeli harus memeriksa alamat yang terisi.

PENTING TENTANG KEAMANAN
- Bridge hanya menerima pesan dari origin yang ditentukan di GITHUB_PAGES_ORIGIN dan hanya meneruskan daftar aksi API yang diizinkan.
- Endpoint backend tetap dapat diakses publik melalui deployment; jangan anggap origin check sebagai pengganti autentikasi/rate limiting. Admin dilindungi password/token, tetapi endpoint publik dapat diserang otomatis.
- Untuk penggunaan produksi yang menyimpan NIK/data pribadi, pertimbangkan rate limiting/proxy khusus, monitoring, kebijakan privasi, retensi data, backup, dan uji keamanan independen.
- HMAC tidak dapat dibalik untuk menampilkan NIK asli. Batasi akses spreadsheet dan proyek Apps Script; editor Apps Script dapat mengakses kunci rahasia.
- Harga tidak dikelola oleh sistem versi ini; konfirmasi harga/ketentuan sebelum publikasi jika diperlukan.

TROUBLESHOOTING
- “Koneksi backend belum diatur”: periksa konstanta GAS_BRIDGE_URL di index.html.
- “Koneksi backend belum siap”: pastikan URL /exec?bridge=1 valid, Bridge.html sudah dibuat, backend sudah dideploy sebagai versi terbaru, dan GITHUB_PAGES_ORIGIN sama persis.
- Data produk tidak muncul: jalankan setupSheets(), tambahkan produk lewat admin, dan lihat Apps Script > Executions untuk error server.
- Jika repository berada di https://username.github.io/nama-repo/, GITHUB_PAGES_ORIGIN tetap https://username.github.io (tanpa /nama-repo).
