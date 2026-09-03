# Folder /templates

Dua file template Excel resmi WAJIB ada di sini sebelum deploy, karena kedua
sistem laporan bergantung padanya (mengisi sel & formula di dalam template,
bukan membuat file baru dari nol):

1. `template_laporan_resmi.xlsx`
   Dipakai oleh sistem **Rekap Monev** (tab "Laporan Resmi").
   Ini file yang sebelumnya diletakkan di `assets/template_laporan_resmi.xlsx`
   pada aplikasi Streamlit lama. Salin apa adanya ke sini.

2. `laporan_kolom_isian_template.xlsx`
   Dipakai oleh sistem **Pre & Post-Test** (bagian "Buat Laporan Resmi").
   Sebelumnya ada di `templates/laporan_kolom_isian_template.xlsx` pada
   aplikasi Streamlit lama. Salin apa adanya ke sini juga.

Tanpa kedua file ini, endpoint /api/monev/report dan /api/pbk/report akan
mengembalikan pesan error yang menjelaskan file mana yang belum ada --
dashboard & fitur lain tetap berjalan normal tanpa keduanya.

Vercel otomatis menyertakan folder ini ke dalam bundle setiap fungsi Python
lewat konfigurasi "includeFiles" pada vercel.json.
