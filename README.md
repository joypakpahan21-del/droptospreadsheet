# 📊 Excel → AI → Google Sheets Automation

Sistem otomatis yang membaca file Excel, menganalisisnya dengan **Claude AI (Anthropic)**, lalu menyimpan hasilnya ke **Google Sheets** — dijalankan melalui **GitHub Actions**.

---

## 🗂️ Struktur Folder

```
excel-to-sheets/
├── .github/
│   └── workflows/
│       └── excel_to_sheets.yml   ← Konfigurasi GitHub Actions
├── data/
│   └── input/                    ← Taruh file Excel di sini
├── scripts/
│   ├── process_excel.py          ← Script utama
│   └── buat_contoh_excel.py      ← Script pembuat data contoh
├── requirements.txt
├── .gitignore
└── README.md
```

---

## ⚙️ Cara Setup (Ikuti Urutan Ini)

### 1. Fork / Clone Repo Ini

```bash
git clone https://github.com/USERNAME/excel-to-sheets.git
cd excel-to-sheets
```

---

### 2. Dapatkan API Key Anthropic

1. Buka [console.anthropic.com](https://console.anthropic.com)
2. Klik **API Keys** → **Create Key**
3. Salin key-nya (simpan, hanya tampil sekali)

---

### 3. Buat Google Service Account

1. Buka [console.cloud.google.com](https://console.cloud.google.com)
2. Buat project baru (atau pakai yang sudah ada)
3. Aktifkan **Google Sheets API** dan **Google Drive API**
4. Buka **IAM & Admin → Service Accounts → Create Service Account**
5. Beri nama, lalu klik **Create and Continue**
6. Di bagian **Keys**, klik **Add Key → Create New Key → JSON**
7. File `.json` akan otomatis terunduh — **simpan baik-baik!**

---

### 4. Bagikan Google Sheets ke Service Account

1. Buka Google Sheets tujuan Anda
2. Salin **ID Spreadsheet** dari URL:
   ```
   https://docs.google.com/spreadsheets/d/[INI_SPREADSHEET_ID]/edit
   ```
3. Klik tombol **Share** di Sheets
4. Tambahkan email service account (format: `nama@project-id.iam.gserviceaccount.com`)
5. Beri akses **Editor**

---

### 5. Tambahkan Secrets di GitHub

Buka repo GitHub Anda → **Settings → Secrets and variables → Actions → New repository secret**

| Secret Name | Isi |
|---|---|
| `ANTHROPIC_API_KEY` | API key dari langkah 2 |
| `SPREADSHEET_ID` | ID Google Sheets dari langkah 4 |
| `GOOGLE_CREDS_JSON` | **Isi seluruh teks** file `.json` service account dari langkah 3 |

> 💡 Untuk `GOOGLE_CREDS_JSON`: buka file `.json` dengan teks editor, lalu copy-paste semua isinya ke kolom secret.

---

### 6. Upload File Excel

Taruh file `.xlsx` atau `.xls` Anda di folder `data/input/`, lalu push ke GitHub:

```bash
cp /path/ke/file_anda.xlsx data/input/
git add data/input/
git commit -m "tambah file Excel untuk diproses"
git push
```

GitHub Actions akan **otomatis berjalan** karena ada file baru di `data/input/`.

---

## 🚀 Cara Menjalankan

### Otomatis Terjadwal
Script berjalan setiap hari jam **08:00 WIB**. Ubah jadwal di file workflow:
```yaml
- cron: "0 1 * * *"   # 01:00 UTC = 08:00 WIB
```
Format cron: `menit jam hari bulan hari-minggu`

### Manual via GitHub
1. Buka tab **Actions** di repo GitHub
2. Pilih workflow **"📊 Excel → AI → Google Sheets"**
3. Klik **Run workflow**
4. Isi nama worksheet (opsional) dan klik **Run**

### Lokal (untuk testing)
```bash
pip install -r requirements.txt

# Set environment variables
export ANTHROPIC_API_KEY="sk-ant-..."
export SPREADSHEET_ID="1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgVE2upms"
export GOOGLE_CREDS_JSON='{"type":"service_account",...}'

python scripts/process_excel.py
```

---

## 🤖 Kustomisasi AI Prompt

Ubah variabel `AI_PROMPT` di GitHub (**Settings → Variables → Actions**) atau di workflow:

```
Kamu adalah analis keuangan. Berikan penilaian risiko singkat dari data berikut:
```

---

## 📋 Hasil di Google Sheets

Script akan menambahkan 2 kolom otomatis:

| Kolom | Keterangan |
|---|---|
| `AI_Analisis` | Hasil analisis Claude AI per baris |
| `Diproses_Pada` | Timestamp kapan data diproses |

Data yang **sudah ada** tidak akan diduplikasi (berdasarkan kolom pertama sebagai ID unik).

---

## ❓ Troubleshooting

| Masalah | Solusi |
|---|---|
| `PERMISSION_DENIED` di Sheets | Pastikan email service account sudah diberi akses Editor di Sheets |
| `ANTHROPIC_API_KEY invalid` | Cek kembali secret di GitHub |
| File Excel tidak terbaca | Pastikan format `.xlsx` atau `.xls`, bukan `.csv` |
| Baris terduplikasi | Pastikan kolom pertama berisi nilai unik (sebagai ID) |
