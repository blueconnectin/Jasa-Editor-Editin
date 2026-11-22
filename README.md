# Pendapatan Editor

Aplikasi Google Apps Script dengan penyimpanan Google Sheets.

Form mencatat tanggal, editor, jasa, harga, metode bayar, dan tarif biaya jasa.
Daftar nama diambil dari kolom A tab Daftar Editor, dengan judul Nama Editor pada A1.
Setup membuat tab tersebut dan mengisinya dari data Excel yang diimpor.

Operasional sama dengan total fee pegawai per bulan.
Sisa hasil usaha sama dengan omzet dikurangi operasional.
Alokasi dari sisa positif menggunakan investasi 3%, dana darurat 3%, iklan 10%,
infaq 2%, Mumu 41%, Hilmi 10%, dan Gina 31%. Gina menyerap selisih pembulatan.

## Hubungkan ke proyek yang sudah digunakan

1. Di GitHub buka Settings > Secrets and variables > Actions.
2. Pada Variables tambahkan APPS_SCRIPT_ID dari Apps Script > Project Settings.
3. Tambahkan APPS_SCRIPT_DEPLOYMENT_ID dari Deploy > Manage deployments.
   Gunakan deployment Web app yang sudah ada agar URL tetap sama.
4. Aktifkan Google Apps Script API di https://script.google.com/home/usersettings.
5. Di komputer pribadi yang memiliki Node.js, jalankan perintah berikut.

   npm install --global @google/clasp@2.4.2
   clasp login

6. Masuk menggunakan akun Google pemilik proyek Apps Script.
   Ambil isi file .clasprc.json pada direktori pengguna komputer.
   Simpan seluruh isi JSON sebagai repository secret CLASPRC_JSON.
   Kredensial ini dimasukkan langsung melalui GitHub Settings.
7. Buka Actions > Deploy Apps Script > Run workflow.

Jika workflow sudah berhasil, perubahan Code.gs, Index.html, atau manifest di main
akan otomatis dikirim ke Apps Script dan deployment diperbarui.
Workflow memerlukan semua konfigurasi sebelum mengirim kode.

## Pemakaian pertama

Jika proyek Apps Script sudah pernah menjalankan setup dan memakai spreadsheet yang
sama, properti BOOK_ID tetap dipakai. Untuk pemasangan baru, buka proyek yang terikat
ke spreadsheet lalu jalankan setup satu kali.

Setup dapat dipakai pada spreadsheet kosong. Bila tab Fee Editor tersedia, data lama
diimpor berdasarkan nomor baris. Pertahankan urutan baris sumber saat mengulang impor.
Tab Daftar Editor dapat ditambah nama mulai A2. Klik Muat ulang di aplikasi.

## Pengembangan

Edit Code.gs untuk server dan Index.html untuk tampilan.
GitHub Actions memeriksa sintaks sebelum sinkronisasi.
clasp push mengganti kode proyek dengan file sumber dalam repository.
Salin perubahan terbaru dari editor Apps Script ke repository sebelum sinkronisasi
pertama jika Anda sudah mengubahnya di luar versi percakapan ini.

Akses web app mengikuti pengaturan deployment Google yang sudah dibuat.
