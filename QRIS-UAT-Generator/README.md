# QRIS UAT Document Generator

**Bank Sahabat Sampoerna - IT Quality Assurance**

Tool otomatis untuk generate dokumen UAT Result dan Lampiran 7C (Berita Acara) dari UAT Script Excel yang sudah diisi oleh mitra, sesuai standar ASPI SNAP.

---

## Fitur

- Membaca UAT Script Excel (.xlsx) yang diisi mitra
- Otomatis validasi `responseCode` di response vs expected result
- Generate **UAT Result .docx** (dokumen detail hasil pengujian)
- Generate **Lampiran 7C .docx** (Berita Acara format ASPI)
- Support pola `xx` (contoh: `401xx01` match dengan `4014701`)
- Auto-detect skenario yang di-skip dari kolom Notes

---

## Cara Pakai

### Opsi 1: Web Version (Tanpa Install)

1. Buka folder `web/`
2. Double-click `index.html` (buka di Chrome/Edge)
3. Drag & drop file UAT Script Excel ke area upload
4. Lihat hasil validasi otomatis
5. Klik tombol **Generate** untuk download dokumen .docx

> **Penting:** Semua file di folder `web/` harus tetap dalam 1 folder yang sama.

### Opsi 2: Python + BAT (Windows)

#### Prasyarat:
- Python 3.x terinstall ([download](https://www.python.org/downloads/))
- Saat install Python, centang **"Add Python to PATH"**

#### Cara pakai:
1. Taruh file berikut di 1 folder:
   - `run_generator.bat`
   - `generate_uat_docs.py`
   - File UAT Script Excel (.xlsx)
2. Double-click `run_generator.bat`
3. Output dokumen muncul di folder `output/`

---

## Struktur File

```
QRIS-UAT-Generator/
├── web/                      # Versi Web (offline, tanpa install)
│   ├── index.html            # Halaman utama - buka di browser
│   ├── app.js                # Logic aplikasi
│   ├── xlsx.full.min.js      # Library baca Excel
│   ├── docx.min.js           # Library generate Word
│   └── FileSaver.min.js      # Library download file
├── generate_uat_docs.py      # Script Python
├── run_generator.bat         # Launcher Windows (double-click)
├── UAT_Script_QRIS.xlsx      # Contoh input Excel
└── README.md                 # Dokumentasi ini
```

---

## Kriteria PASS / NOT PASS

| Kondisi | Hasil |
|---------|-------|
| `responseCode` di response = expected result | **PASS** |
| `responseCode` di response != expected result | **NOT PASS** |
| Pola `xx` (misal `401xx01`) match digit | **PASS** |
| Tidak ada response / belum diisi | **NOT TESTED** |
| Notes berisi "tidak dites" | **N/A** |

---

## Output

1. **UAT_Result_QRIS_Merchant_Aggregator_[tanggal].docx**
   - Title page, Daftar Isi, Detail per skenario, Ringkasan

2. **Lampiran_7C_QRIS_[tanggal].docx**
   - Format landscape, tabel 8 kolom sesuai template ASPI
   - Footer notes ketentuan ASPI

---

## Produk: QRIS Merchant Aggregator

Sesuai standar ASPI SNAP untuk API:
- QR MPM - Generate QR
- QR MPM - Decode QR
- QR MPM - Payment (Redirect / Host-to-Host)
- Query Payment
- Payment Notification
- Cancel Payment
- Refund Payment
- Transaction Status Inquiry
