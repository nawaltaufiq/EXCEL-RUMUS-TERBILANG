EXCEL-RUMUS-TERBILANG

Modul VBA ini dikembangkan untuk membantu mengonversi angka menjadi teks terbilang dalam bahasa Indonesia.

Terbilang VBA Module

Deskripsi

Modul VBA ini dikembangkan untuk membantu mengonversi angka menjadi teks terbilang dalam bahasa Indonesia. Modul ini sangat berguna untuk kebutuhan penulisan faktur atau laporan keuangan, terutama di perusahaan yang masih menggunakan proses manual dalam pembuatan dokumen keuangan.

Latar Belakang

Modul ini terinspirasi dari studi kasus di mana banyak perusahaan yang masih menulis nilai faktur secara manual. Hal ini berisiko tinggi terhadap kesalahan penulisan, terutama untuk angka yang besar atau memiliki desimal. Dengan modul ini, penulisan terbilang menjadi lebih cepat, akurat, dan konsisten.


---

Versi 1: Angka Desimal Tidak Disebutkan

Pada versi pertama, hanya angka sebelum koma yang dikonversi menjadi teks. Angka desimal (di belakang koma) diabaikan sepenuhnya.

Contoh:

Input: 1234.56

Output: *SERIBU DUA RATUS TIGA PULUH EMPAT RUPIAH*


Fitur Versi 1:

1. Angka Tidak Dibulatkan:

Bagian angka sebelum koma dikonversi tanpa pembulatan.



2. Mendukung Bilangan Besar:

Dapat menangani angka hingga jutaan, ribuan, ratusan, puluhan, dan satuan.



3. Format Rupiah:

Otomatis menambahkan kata "Rupiah" di akhir hasil konversi.



4. Teks Kapital dan Dibungkus dengan Asteris (*):

Hasil akhir dikonversi menjadi huruf kapital dan dibungkus dengan asteris di awal dan akhir.



5. Frasa Khusus:

Otomatis mengganti "Satu Ribu" menjadi "Seribu" dan "Satu Ratus" menjadi "Seratus" untuk hasil yang lebih natural.





---

Versi 2: Angka Desimal Disebutkan

Pada versi kedua, angka sebelum dan sesudah koma dikonversi menjadi teks terbilang. Angka desimal disebutkan per digit.

Contoh:

Input: 1234.56

Output: *SERIBU DUA RATUS TIGA PULUH EMPAT LIMA ENAM RUPIAH*


Fitur Versi 2:

1. Angka Desimal Disebutkan:

Bagian angka di belakang koma dikonversi menjadi teks per digit.



2. Mendukung Bilangan Besar:

Dapat menangani angka hingga jutaan, ribuan, ratusan, puluhan, dan satuan.



3. Format Rupiah:

Otomatis menambahkan kata "Rupiah" di akhir hasil konversi.



4. Teks Kapital dan Dibungkus dengan Asteris (*):

Hasil akhir dikonversi menjadi huruf kapital dan dibungkus dengan asteris di awal dan akhir.



5. Frasa Khusus:

Otomatis mengganti "Satu Ribu" menjadi "Seribu" dan "Satu Ratus" menjadi "Seratus" untuk hasil yang lebih natural.





---

Perbedaan Versi 1 dan Versi 2


---

Cara Menggunakan

1. Import Fungsi ke Excel

1. Buka VBA Editor di Excel:

Tekan Alt + F11.



2. Buat modul baru:

Klik kanan pada ThisWorkbook > Insert > Module.



3. Salin kode fungsi ke modul:

Gunakan kode untuk Versi 1 atau Versi 2 sesuai kebutuhan Anda.




2. Gunakan Fungsi di Excel

1. Kembali ke workbook Excel Anda.


2. Ketik formula berikut di salah satu sel:

=Terbilang(A1)


3. Ganti A1 dengan sel yang berisi angka yang ingin dikonversi.


4. Hasilnya akan berupa teks terbilang sesuai dengan versi yang Anda gunakan.



