# SEMAR online (gratis) di Streamlit Community Cloud

Panduan ini memasang SEMAR (Dashboard, Semar, Pemeriksaan Barang) di internet dengan alamat `https://nama-pilihan.streamlit.app`.
HTTPS membuat kamera ponsel bisa dipakai langsung dari halaman Pemeriksaan Barang.

## Yang perlu disiapkan
- Akun GitHub (gratis): github.com
- API key DeepSeek yang saldonya sudah terisi (DEEPSEEK_API_KEY, diawali `sk-`)
- NIP (18 digit) orang yang akan menjadi admin aplikasi

## Langkah

### 1. Buat repositori PRIVAT di GitHub
Di github.com: New repository, beri nama `semar`, pilih **Private**, lalu Create.
Repositori harus privat karena berisi file Excel data Anda.

### 2. Unggah isi folder ini ke repositori
Cara termudah: **GitHub Desktop** (desktop.github.com), pilih Add local repository ke folder ini, lalu Commit dan Publish repository (centang Keep this code private).

Cara lewat web: di repositori klik Add file, Upload files, lalu seret semua berkas. Folder berawalan titik kadang terlewat saat diseret.
Pastikan ada `.streamlit/config.toml` dan `_komponen_sheet/index.html`. Bila hilang, buat lewat Add file, Create new file, ketik nama lengkap dengan foldernya, lalu tempel isinya.

Jangan unggah file `.env` atau `secrets.toml` yang berisi nilai asli.

### 3. Pasang di Streamlit Community Cloud
1. Buka share.streamlit.io lalu masuk dengan akun GitHub (izinkan akses ke repositori privat bila diminta).
2. Klik Create app, pilih repositori `semar`, branch `main`, dan **Main file path: `app_web.py`**.
3. Pilih subdomain, misalnya `semar-kppbc`.
4. Buka **Advanced settings**: Python version `3.12`, dan pada kolom **Secrets** tempel (nilai diganti dengan yang asli):

```toml
DEEPSEEK_API_KEY = "sk-..."
DATABASE_URL = "postgresql://...   (dari Neon, lihat langkah 4)"
SEMAR_ADMIN_NIP = "NIP-admin-18-digit"
```

5. Klik Deploy. Pemasangan pertama memakan beberapa menit.

### 4. Uji
- Buka alamat aplikasi, masukkan kata sandi.
- Dashboard tampil; coba halaman Pemeriksaan Barang dengan satu dokumen.
- Di HP: buka alamat yang sama, di slot foto aktifkan Buka kamera.

### 4. Database untuk foto dokumentasi 7 hari (Neon, gratis)
Foto dan file yang diunggah di halaman Pemeriksaan Barang disimpan 7 hari dan muncul di kolom U **file dokumentasi** pada sheet (tombol Lihat dan Unduh).
Di server, file lokal hilang setiap aplikasi restart, jadi diperlukan database di luar server:

1. Buka neon.com, daftar (boleh dengan akun GitHub atau Google), lalu **Create project**. Pilih region terdekat (mis. Singapore).
2. Di dashboard proyek klik **Connect**, lalu salin **connection string** (diawali `postgresql://`).
3. Tempel sebagai `DATABASE_URL` pada Secrets (di dalam tanda kutip), lalu simpan dan muat ulang aplikasi.
4. Buka halaman Pemeriksaan Barang, buka kotak **Penyimpanan 7 hari**, lalu klik **Periksa penyimpanan**. Harus muncul "Tersambung".

Tabel dibuat otomatis (`semar_baris` dan `semar_berkas`). Baris dan file yang lebih tua dari 7 hari dihapus otomatis setiap ada sesi baru.
Tanpa `DATABASE_URL`, aplikasi memakai file lokal `data_semar/` (cukup untuk komputer sendiri, tidak untuk server).

## Memperbarui
- Ubah kode atau Excel: commit ke GitHub, aplikasi otomatis diperbarui.
- Ganti sandi atau API key: share.streamlit.io, menu titik tiga di aplikasi, Settings, Secrets.

## Login dan akun pengguna
Setiap orang punya akun sendiri. **Masuk memakai NIP (18 digit) dan kata sandi.** Pendaftaran meminta Nama, NIP, Jabatan, Satker, Seksi, dan kata sandi.
- **Pendaftar baru menunggu persetujuan admin.** Admin menyetujui, menolak, menonaktifkan, mereset kata sandi, atau menghapus akun di halaman **Pengguna** (hanya tampil untuk admin).
- **Admin pertama:** orang yang mendaftar dengan NIP yang sama dengan `SEMAR_ADMIN_NIP` otomatis menjadi admin dan langsung aktif. Daftarlah dengan NIP itu terlebih dahulu.
- Di server, tanpa `SEMAR_ADMIN_NIP` pendaftaran ditutup. Tanpa `DATABASE_URL`, akun hilang saat aplikasi restart, jadi isi database dulu (langkah 4).
- Kata sandi disimpan sebagai hash (tidak bisa dibaca siapa pun, termasuk admin). 5 kali salah berturut-turut mengunci akun 15 menit.
- Lupa kata sandi: minta admin mereset (kata sandi sementara), lalu ganti lewat tombol **Ganti kata sandi** di sidebar.
- **Admin lupa kata sandi:** buka Neon, SQL Editor, jalankan `DELETE FROM semar_pengguna WHERE nip = 'NIP-admin';`, lalu daftar ulang dengan NIP itu (otomatis admin lagi).
- Di komputer sendiri, orang yang mendaftar pertama menjadi admin. Untuk melewati login saat pengembangan lokal, atur `SEMAR_TANPA_LOGIN=1` pada file `.env` (diabaikan di server).

## Yang perlu Anda ketahui
- **Tidur otomatis.** Aplikasi tanpa kunjungan 12 jam akan tidur. Buka alamatnya dan klik "Yes, get this app back up!" (sekitar satu menit).
- **Batas sumber daya dan tanpa jaminan.** Layanan gratis ini punya batas CPU/RAM (dokumentasi Streamlit menyebut sekitar 2,7 GB RAM; bisa berubah tanpa pemberitahuan) dan tidak ada SLA.
- **Sheet disimpan 7 hari.** Dengan `DATABASE_URL` terisi, baris sheet dan foto dokumentasi tetap ada setelah refresh atau aplikasi tidur, lalu dihapus otomatis setelah 7 hari. Semua pengguna aktif melihat daftar yang sama. Tetap **unduh .xlsx** untuk arsip, karena data di sheet dan database hilang setelah 7 hari.
- **Kapasitas database gratis terbatas** (cek halaman harga Neon). Foto disimpan maksimal 2400 piksel pada sisi terpanjang dan file di atas 20 MB tidak disimpan.
- **Admin mengendalikan akses.** Akun yang dinonaktifkan atau dihapus admin tertutup paling lambat dalam satu menit, tetapi sheet dan foto dokumentasi tetap dibagi ke semua pengguna aktif.
- **Dokumen keluar dari instansi.** Dokumen dan foto dikirim ke DeepSeek (seperti sekarang) dan diproses di server hosting di luar negeri. Pastikan kebijakan instansi Anda mengizinkan.
- **Website sumber pengetahuan.** Sebagian situs bisa menolak akses dari server. Periksa lewat tombol Lihat isi basis pengetahuan.
- **Suara.** Bila suara pria dari server tidak jalan, tambahkan `SEMAR_MESIN = "peramban"` pada Secrets.
- **Akses lebih ketat (opsional).** Pengaturan Sharing di Streamlit bisa membatasi penonton ke email tertentu. Menurut forum Streamlit, aplikasi privat gratis dibatasi jumlahnya; kebijakan ini bisa berubah.

## Dipakai di komputer sendiri
`jalankan_semar_web.bat` tetap bisa dipakai seperti biasa. Halaman login juga tampil di komputer sendiri; orang yang mendaftar pertama menjadi admin. Untuk melewati login saat dipakai sendiri, tulis `SEMAR_TANPA_LOGIN=1` di file `.env`.
