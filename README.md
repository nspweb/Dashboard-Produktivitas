# Dasbor Evaluasi &amp; Produktivitas Pelatihan

Gabungan dua sistem lama (Streamlit) dalam satu aplikasi web statis + Vercel
Serverless Functions (Python), siap di-hosting di Vercel tanpa Streamlit:

1. **Rekap Monev** — evaluasi penyelenggaraan pelatihan (dari `app.py` + `report_generator.py`)
2. **Pre &amp; Post-Test** — produktivitas pelatihan (dari `dashboard_pbk.py` + `report_kolom_isian.py`)

## Struktur proyek

```
eval-suite/
├── index.html, css/, js/        -> frontend statis (di-serve langsung oleh Vercel)
├── api/
│   ├── monev/process.py         -> upload & deteksi kolom file evaluasi
│   ├── monev/report.py          -> generate laporan resmi (.xlsx) sistem Monev
│   ├── monev/comment_recap.py   -> generate rekap komentar (.xlsx)
│   ├── pbk/process.py           -> upload & bersihkan data Pre/Post-Test
│   ├── pbk/report.py            -> generate laporan resmi (.xlsx) sistem Pre/Post-Test
│   ├── convert_pdf.py           -> konversi .xlsx -> .pdf lewat CloudConvert
│   └── _lib/                    -> logika inti (diporting langsung dari 4 file Python asli)
├── templates/                   -> WAJIB diisi manual, lihat templates/README.md
├── requirements.txt
└── vercel.json
```

## Yang berubah dari versi Streamlit

- **Tidak ada server yang selalu hidup.** Setiap unggah file diproses oleh
  fungsi serverless yang stateless; hasilnya (JSON) disimpan di memori
  browser (JavaScript), bukan `st.session_state`. Filter & dashboard
  dihitung ulang di browser (cepat, tanpa round-trip ke server).
- **Konversi PDF pindah dari LibreOffice ke CloudConvert**, karena Vercel
  tidak bisa menjalankan binary `soffice`. Anda perlu API key CloudConvert
  (gratis untuk volume kecil): https://cloudconvert.com/dashboard/api/v2/keys
- **Grafik pakai Chart.js** (CDN), bukan Plotly, supaya frontend tetap
  ringan sebagai file statis tanpa build step.
- **Unduhan "data mentah"** kini berupa CSV (bisa dibuka Excel langsung),
  bukan `.xlsx` — menghindari endpoint tambahan hanya untuk konversi format.

## Sebelum deploy — 2 langkah wajib

### 1. Salin template Excel resmi Anda
Lihat `templates/README.md`. Tanpa dua file ini, fitur "Laporan Resmi" pada
kedua sistem akan menampilkan pesan error (tapi dashboard tetap jalan).

### 2. Set environment variable di Vercel
Project Settings → Environment Variables:

| Nama                    | Nilai                                  |
|--------------------------|-----------------------------------------|
| `CLOUDCONVERT_API_KEY`   | API key dari dashboard CloudConvert Anda |

## Deploy

```bash
npm i -g vercel
cd eval-suite
vercel        # deploy preview
vercel --prod # deploy production
```

Atau hubungkan repo GitHub ini ke Vercel lewat dashboard (import project) —
Vercel otomatis mendeteksi `requirements.txt` dan `vercel.json`, tidak perlu
konfigurasi build tambahan.

## Menjalankan lokal

### Cara 1: Menggunakan Python (Direkomendasikan)

Cukup jalankan perintah berikut di terminal:

```bash
python server.py
```

Lalu buka browser di: **http://localhost:5000**

### Cara 2: Menggunakan Vercel CLI

```bash
npm i -g vercel
vercel dev
```
Buka browser di: **http://localhost:3000**

## Catatan batasan platform

- Body request Vercel Functions dibatasi (~4.5 MB pada paket Hobby). File
  evaluasi yang sangat besar (ribuan responden dengan banyak kolom) mungkin
  perlu paket Vercel Pro, atau proses per-batch.
- `maxDuration` fungsi diset 30 detik di `vercel.json` (cukup untuk mengisi
  template + formula; konversi PDF via CloudConvert biasanya di bawah itu,
  tapi bisa dinaikkan jika perlu, tergantung paket Vercel Anda).