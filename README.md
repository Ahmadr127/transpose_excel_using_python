# Sistem Excel Processing - Upload & Download

Sistem web Python yang dapat membaca, mempelajari, dan memproses file Excel dengan kemampuan upload dan download. Sistem ini dirancang untuk mengkonversi data Excel mentah ke format yang terstruktur sesuai kebutuhan.

## Fitur Utama

- **Upload File Excel**: Mendukung format .xlsx dan .xls
- **Preview Data**: Menampilkan preview struktur dan sample data
- **Smart Processing**: Otomatis memetakan kolom input ke output
- **Format Standardisasi**: Mengkonversi ke format output yang diinginkan
- **Download Result**: File Excel hasil pemrosesan siap download
- **Drag & Drop**: Interface modern dengan drag & drop support
- **Responsive Design**: Tampilan yang responsif untuk berbagai device

## Struktur Output

Sistem akan menghasilkan file Excel dengan kolom-kolom berikut:

1. **PROVID** - Provider ID
2. **PROVIDER_NAME** - Nama Provider
3. **SERVICECODE** - Kode Layanan
4. **SERVICECODE DESCRIPTION** - Deskripsi Layanan
5. **KELAS** - Kelas Layanan
6. **RUANG BEDAH (SURGERY)/NON RUANG BEDAH (NON SURGERY)** - Jenis Ruangan
7. **HELPER** - Helper/Asisten
8. **TARIFF** - Tarif Layanan
9. **TARIFF DESCRIPTION** - Deskripsi Tarif
10. **QUANTITY** - Jumlah
11. **TOTAL BILLED** - Total Tagihan
12. **GIVEN DATE** - Tanggal Pemberian
13. **HEAMODIALISA/CHEMOTHERAPY/ODC/PHYSIOTHERAPY/RADIOTHERAPY** - Jenis Terapi
14. **HEAMODIALISA/CHEMOTHERAPY/ODC/PHYSIOTHERAPY/RADIOTHERAPY DESCRIPTION** - Deskripsi Terapi
15. **ICD_X_DIAGNOSIS_PRIMARY** - Diagnosis Utama ICD-X
16. **ICD_X_DESC_PRIMARY** - Deskripsi Diagnosis Utama
17. **ICD_X_DIAGNOSIS_SECONDARY** - Diagnosis Sekunder ICD-X
18. **ICD_X_DESC_SECONDARY** - Deskripsi Diagnosis Sekunder
19. **PHYSICIAN NAME** - Nama Dokter
20. **PHYSICIAN DESCRIPTION** - Deskripsi Dokter (DPJP/IGD/POLICLINIC)
21. **CLIENT NAME** - Nama Klien
22. **CLIENTS DOB** - Tanggal Lahir Klien
23. **CLIENTS SEX** - Jenis Kelamin Klien
24. **CLIENTS ADDRESS** - Alamat Klien
25. **CLIENTS MEMBER ID** - ID Member Klien
26. **CLIENTS MR NUMBER** - Nomor MR Klien
27. **CLIENTS INVOICE NUMBER** - Nomor Invoice Klien
28. **CLIENTSREGISTER NUMBER** - Nomor Registrasi Klien
29. **CLIENTS OTHER NUMBER** - Nomor Lain Klien
30. **admission** - Tanggal Masuk
31. **discharge** - Tanggal Keluar
32. **LoS** - Length of Stay

## Instalasi

### Prerequisites

- Python 3.8 atau lebih tinggi
- pip (Python package manager)

### Langkah Instalasi

1. **Clone atau download project ini**
   ```bash
   git clone <repository-url>
   cd sistem-excel
   ```

2. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```

3. **Buat direktori yang diperlukan**
   ```bash
   mkdir uploads
   mkdir outputs
   mkdir templates
   ```

4. **Jalankan aplikasi**
   ```bash
   python app.py
   ```

5. **Buka browser dan akses**
   ```
   http://localhost:5000
   ```

## Penggunaan

### 1. Upload File
- Drag & drop file Excel (.xlsx/.xls) ke area upload
- Atau klik tombol "Pilih File" untuk memilih file secara manual
- Sistem akan otomatis membaca dan menganalisis struktur data

### 2. Preview Data
- Setelah upload berhasil, sistem akan menampilkan preview data
- Lihat struktur kolom, tipe data, dan sample data
- Verifikasi bahwa data yang dibaca sudah benar

### 3. Proses Data
- Klik tombol "Proses Data" untuk memulai konversi
- Sistem akan memetakan kolom input ke output secara otomatis
- Progress bar akan menunjukkan status pemrosesan

### 4. Download Result
- Setelah proses selesai, file output siap didownload
- Klik tombol "Download File" untuk mengunduh hasil
- File akan otomatis tersimpan dengan nama yang sesuai

## Struktur Project

```
sistem-excel/
├── app.py                 # Aplikasi Flask utama
├── excel_processor.py     # Modul pemrosesan Excel
├── requirements.txt       # Dependencies Python
├── README.md             # Dokumentasi ini
├── templates/            # Template HTML
│   └── index.html       # Halaman utama
├── uploads/             # Direktori file upload (auto-created)
└── outputs/             # Direktori file output (auto-created)
```

## Konfigurasi

### Environment Variables (Opsional)
```bash
export FLASK_ENV=development
export FLASK_DEBUG=1
export MAX_FILE_SIZE=16777216  # 16MB dalam bytes
```

### Customization
Anda dapat memodifikasi file `excel_processor.py` untuk:
- Mengubah mapping kolom
- Menambah logika pemrosesan khusus
- Memodifikasi format output
- Menambah validasi data

## Troubleshooting

### Error Umum

1. **"Module not found"**
   - Pastikan semua dependencies terinstall: `pip install -r requirements.txt`

2. **"Permission denied"**
   - Pastikan direktori `uploads` dan `outputs` memiliki permission write

3. **"File too large"**
   - File Excel tidak boleh lebih dari 16MB
   - Modifikasi `MAX_CONTENT_LENGTH` di `app.py` jika diperlukan

4. **"Invalid file format"**
   - Pastikan file yang diupload berformat .xlsx atau .xls
   - File tidak boleh corrupt atau rusak

### Debug Mode
Untuk development, jalankan dengan debug mode:
```bash
export FLASK_DEBUG=1
python app.py
```

## API Endpoints

- `GET /` - Halaman utama
- `POST /upload` - Upload file Excel
- `POST /process` - Proses data Excel
- `GET /download` - Download file hasil
- `POST /cleanup` - Bersihkan file temporary

## Dependencies

- **Flask** - Web framework
- **pandas** - Data manipulation dan analysis
- **openpyxl** - Excel file handling
- **Werkzeug** - WSGI utilities
- **python-dateutil** - Date utilities

## Lisensi

Project ini dibuat untuk tujuan edukasi dan penggunaan internal. Silakan modifikasi sesuai kebutuhan Anda.

## Support

Jika mengalami masalah atau ada pertanyaan:
1. Periksa log error di console
2. Pastikan semua dependencies terinstall
3. Verifikasi format file Excel yang diupload
4. Cek permission direktori

## Changelog

### v1.0.0
- Initial release
- Basic Excel processing functionality
- Modern web interface
- Drag & drop support
- Automatic column mapping
# transpose_excel_using_python
