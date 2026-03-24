"""
Excel → AI Processor → Google Sheets
Dibaca oleh GitHub Actions secara otomatis terjadwal.
"""

import os
import json
import glob
import anthropic
import gspread
import pandas as pd
from google.oauth2.service_account import Credentials
from datetime import datetime

# ─── CONFIG ────────────────────────────────────────────────────────────────────
SPREADSHEET_ID   = os.environ["SPREADSHEET_ID"]       # ID Google Sheets tujuan
ANTHROPIC_API_KEY = os.environ["ANTHROPIC_API_KEY"]   # API key Anthropic
GOOGLE_CREDS_JSON = os.environ["GOOGLE_CREDS_JSON"]   # JSON service account (string)
WORKSHEET_NAME   = os.environ.get("WORKSHEET_NAME", "Sheet1")
INPUT_DIR        = os.environ.get("INPUT_DIR", "data/input")
AI_PROMPT        = os.environ.get(
    "AI_PROMPT",
    "Kamu adalah asisten analis data. Berikan ringkasan singkat satu kalimat "
    "dari baris data berikut dalam Bahasa Indonesia:"
)
# ───────────────────────────────────────────────────────────────────────────────


def connect_google_sheets() -> gspread.Worksheet:
    """Autentikasi dan buka worksheet Google Sheets."""
    creds_dict = json.loads(GOOGLE_CREDS_JSON)
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    client = gspread.authorize(creds)
    sh = client.open_by_key(SPREADSHEET_ID)

    # Buat worksheet baru jika belum ada
    try:
        ws = sh.worksheet(WORKSHEET_NAME)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=WORKSHEET_NAME, rows=1000, cols=30)
        print(f"[INFO] Worksheet '{WORKSHEET_NAME}' dibuat baru.")

    return ws


def analyze_row_with_ai(client: anthropic.Anthropic, row_data: dict) -> str:
    """Kirim satu baris data ke Claude dan dapatkan analisis."""
    row_str = ", ".join([f"{k}: {v}" for k, v in row_data.items()])
    message = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=200,
        messages=[{"role": "user", "content": f"{AI_PROMPT}\n\nData: {row_str}"}],
    )
    return message.content[0].text.strip()


def excel_to_dataframe(filepath: str) -> pd.DataFrame:
    """Baca file Excel menjadi DataFrame, normalisasi header."""
    df = pd.read_excel(filepath, engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]
    df = df.dropna(how="all")           # hapus baris kosong total
    df = df.fillna("")                  # isi NaN dengan string kosong
    return df


def get_existing_ids(ws: gspread.Worksheet, id_col_index: int = 0) -> set:
    """Ambil semua nilai di kolom ID untuk menghindari duplikat."""
    try:
        col_values = ws.col_values(id_col_index + 1)  # gspread pakai 1-based
        return set(col_values[1:])  # skip header
    except Exception:
        return set()


def push_to_sheets(ws: gspread.Worksheet, df: pd.DataFrame, ai_results: list[str]):
    """Tulis header + data ke Google Sheets (append, hindari duplikat)."""
    existing_data = ws.get_all_values()
    is_empty = len(existing_data) == 0

    # Tulis header jika sheet masih kosong
    headers = list(df.columns) + ["AI_Analisis", "Diproses_Pada"]
    if is_empty:
        ws.append_row(headers, value_input_option="USER_ENTERED")
        print(f"[INFO] Header ditulis: {headers}")

    # Ambil baris yang sudah ada (kolom pertama sebagai key)
    existing_keys = get_existing_ids(ws, id_col_index=0)

    appended = 0
    skipped = 0
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    for idx, (_, row) in enumerate(df.iterrows()):
        row_values = [str(v) for v in row.tolist()]
        row_key = row_values[0]  # pakai kolom pertama sebagai unique key

        if row_key in existing_keys and row_key != "":
            skipped += 1
            continue

        full_row = row_values + [ai_results[idx], timestamp]
        ws.append_row(full_row, value_input_option="USER_ENTERED")
        existing_keys.add(row_key)
        appended += 1

    print(f"[INFO] Ditambahkan: {appended} baris | Dilewati (duplikat): {skipped} baris")
    return appended, skipped


def process_all_files():
    """Main runner: proses semua file Excel di INPUT_DIR."""
    excel_files = glob.glob(f"{INPUT_DIR}/*.xlsx") + glob.glob(f"{INPUT_DIR}/*.xls")

    if not excel_files:
        print(f"[WARN] Tidak ada file Excel ditemukan di '{INPUT_DIR}'. Selesai.")
        return

    print(f"[INFO] Ditemukan {len(excel_files)} file Excel.")

    ws = connect_google_sheets()
    ai_client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

    for filepath in excel_files:
        print(f"\n[PROSES] {filepath}")
        try:
            df = excel_to_dataframe(filepath)
            print(f"  → {len(df)} baris, {len(df.columns)} kolom")

            # Proses AI per baris
            ai_results = []
            for _, row in df.iterrows():
                analysis = analyze_row_with_ai(ai_client, row.to_dict())
                ai_results.append(analysis)
                print(f"  ✓ AI: {analysis[:60]}...")

            # Push ke Google Sheets
            added, skipped = push_to_sheets(ws, df, ai_results)
            print(f"  → Sheets diperbarui: +{added} baris baru, {skipped} duplikat diabaikan.")

        except Exception as e:
            print(f"  [ERROR] Gagal memproses {filepath}: {e}")
            raise

    print("\n[SELESAI] Semua file berhasil diproses.")


if __name__ == "__main__":
    process_all_files()
