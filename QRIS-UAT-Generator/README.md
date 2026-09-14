# QRIS Document Generator

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

#### Opsi lanjutan (CLI):

```
python generate_uat_docs.py <uat_script.xlsx> [--template [<uat_result.docx>]]
```

- Tanpa `--template`: perilaku default, generate **UAT Result** dan **Lampiran 7C** dari nol (seperti sebelumnya).
- Dengan `--template <uat_result.docx>`: jadikan **template UAT Result .docx** milik Anda (yang sudah berisi **screenshot** tiap skenario) sebagai dasar, lalu sisipkan hasil UAT Script (Expected Result, Request, Response) **tepat di bawah screenshot** tiap skenario. Output ditulis ke `output/UAT_Result_from_template_<tanggal>.docx`, dan **Lampiran 7C** tetap dihasilkan dari nol.
- `--template` boleh dipakai **tanpa path**; jika begitu, tool memakai template bawaan di repo (`QRIS-UAT-Generator/UAT Result Penambahan Layanan QRIS Merchant Aggregator.docx`) bila file itu ada.

> **Catatan:** Fitur `--template` (upload + merge template UAT Result ber-screenshot) **hanya tersedia di versi Python**. Versi Web hanya bisa membuat dokumen baru dan tidak bisa mengedit `.docx` existing yang berisi gambar.
>
> **Dasar pencocokan = nama skenario.** Konten UAT Script dicocokkan ke skenario di template berdasarkan **kecocokan NAMA skenario**, bukan nomor ASPI. Ini agar skenario yang **tidak bernomor** di template (mis. `Melakukan cek status QR`, `Melakukan refund transaksi issuer BSS`) juga ikut tercocokkan.
>
> **Cara pencocokan (section-aware + berurutan + fuzzy):**
> - Judul skenario template = paragraf ber-style **Heading 2**; pengelompokan **section** = paragraf ber-style **Heading 1** (mis. `Balance Services`, `QR MPM`, `Pengecekan Mutasi Dan Jurnal`).
> - Nama dinormalkan lebih dulu: ambil baris pertama, buang prefiks nomor (`18,1 `), lowercase, rapatkan spasi ganda, samakan tanda kutip/elipsis.
> - Karena banyak nama **identik** lintas section (mis. `Access Token Invalid` di 3.x/4.x/18.x, `Melakukan pengecekan mutasi dan jurnal` yang berulang), pencocokan dilakukan **per section** dan **berurut maju**: skenario Excel diproses sesuai urutan, tiap heading template dipakai **maksimal sekali** sehingga tidak salah tempel.
> - Beda kecil teks ditoleransi dengan **fuzzy match** (`difflib`, stdlib) + containment (ambang ~0.82) dan penyamaan `QRIS` vs `QR` (mis. Excel `Melakukan transaksi QRIS sukses` cocok ke template `Melakukan transaksi QR sukses`).
>
> **Titik sisip:** konten disisipkan **setelah** blok skenario yang sudah ada di template (setelah screenshot) dan **sebelum** heading skenario berikutnya (atau di akhir dokumen untuk skenario terakhir). Screenshot/gambar/paragraf existing tidak dihapus atau digeser.
>
> **Yang tak berpasangan:** skenario Excel tanpa heading pasangan di template (mis. `Query Successful Transaction`, `Notification for Successful/Failed Transaction`) **dilewati dengan peringatan bernama** (proses tidak gagal), dan heading template tanpa pasangan Excel dibiarkan apa adanya. Di akhir dicetak ringkasan **berapa skenario cocok tersisip dan berapa terlewat**.

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
