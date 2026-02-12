import os
import glob
from datetime import datetime
from pathlib import Path

import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

SPREADSHEET_ID = '1TaxVb8GrPndXHGhjNWVzQm8QcLBwBqJ5b_-yYS1EC6g'
SHEET_NAME = 'TTR RESELLER'
DOWNLOADS_FOLDER = str(Path.home() / "Downloads")


def col_idx_to_a1(col_idx: int) -> str:
    result = ""
    while col_idx > 0:
        col_idx, rem = divmod(col_idx - 1, 26)
        result = chr(ord('A') + rem) + result
    return result


def find_latest_download():
    xls_files = glob.glob(os.path.join(DOWNLOADS_FOLDER, "*.xls*"))
    if not xls_files:
        print("❌ Tidak ada file Excel/HTML di folder Downloads")
        return None

    latest_file = max(xls_files, key=os.path.getmtime)
    print(f"✓ File: {latest_file}")
    print(f"  Modified: {datetime.fromtimestamp(os.path.getmtime(latest_file))}")
    return latest_file


def read_excel_data(file_path):
    print(f"\n📖 Membaca: {os.path.basename(file_path)}")
    try:
        with open(file_path, 'rb') as f:
            header = f.read(2000).decode('utf-8', errors='ignore').lower()

        if '<table' in header or '<html' in header:
            print("  🌐 HTML TABLE (.xls TelkomCare) → pandas.read_html()")
            dfs = pd.read_html(file_path)
            df = dfs[0]
        elif file_path.lower().endswith('.xls'):
            print("  📄 XLS asli → pandas.read_excel(engine='xlrd')")
            df = pd.read_excel(file_path, engine='xlrd')
        else:
            print("  📄 XLSX → pandas.read_excel(engine='openpyxl')")
            df = pd.read_excel(file_path, engine='openpyxl')

        df = df.fillna("")
        data = df.values.tolist()
        print(f"✓ {len(data)} baris × {len(data[0]) if data else 0} kolom")

        if data:
            print("📋 Preview:")
            print(f"Header: {data[0][:10]}{'...' if len(data[0]) > 10 else ''}")
            if len(data) > 1:
                print(f"Row 2:  {data[1][:10]}{'...' if len(data[1]) > 10 else ''}")
        return data

    except Exception as e:
        print(f"❌ Error baca file: {e}")
        print("💡 Pastikan sudah install: py -m pip install pandas xlrd openpyxl lxml html5lib")
        return None


def setup_gsheets():
    try:
        scopes = [
            'https://www.googleapis.com/auth/spreadsheets',
            'https://www.googleapis.com/auth/drive'
        ]
        creds = Credentials.from_service_account_file('credentials.json', scopes=scopes)
        gc = gspread.authorize(creds)
        return gc
    except Exception as e:
        print(f"❌ Google Sheets error: {e}")
        print("Pastikan 'credentials.json' ada dan formatnya benar.")
        return None


def upload_to_sheets(data, spreadsheet_id, sheet_name):
    print(f"\n📤 Upload ke sheet '{sheet_name}'...")

    if not data:
        print("❌ Data kosong, tidak ada yang diupload.")
        return False

    try:
        gc = setup_gsheets()
        if not gc:
            return False

        sh = gc.open_by_key(spreadsheet_id)

        try:
            ws = sh.worksheet(sheet_name)
            print(f"✓ Sheet '{sheet_name}' ditemukan, header baris 1 akan dipertahankan.")
        except gspread.exceptions.WorksheetNotFound:
            print(f"⚠️ Sheet '{sheet_name}' tidak ditemukan, membuat baru...")
            ws = sh.add_worksheet(title=sheet_name, rows=len(data)+10, cols=len(data[0])+10)

        max_cols = max(len(r) for r in data)
        normalized = []
        for r in data:
            row = list(r)
            if len(row) < max_cols:
                row += [""] * (max_cols - len(row))
            normalized.append(row)

        last_col_letter = col_idx_to_a1(max_cols)
        clear_range = f"A2:{last_col_letter}100000"
        print(f"🧹 Clear range data lama (tanpa header): {clear_range}")
        ws.batch_clear([clear_range])

        print(f"✓ Upload {len(normalized)} baris × {max_cols} kolom via worksheet.update()")
        ws.update(normalized, 'A2')

        print(f"\n✅ Data berhasil diupload!")
        print(f"   Total: {len(normalized)} baris × {max_cols} kolom (mulai baris 2)")
        print(f"   Header baris 1 dipertahankan.")
        print(f"   Sheet: {sheet_name}")
        return True

    except Exception as e:
        print(f"❌ Upload error: {e}")
        print("💡 Cek: jaringan, ukuran data, dan status API jika ada error lain.")
        return False


def main():
    print("=" * 70)
    print("🚀 IMPORT TELKOMCARE (.xls HTML) KE GOOGLE SHEETS (TTR RESELLER)")
    print("=" * 70)

    file_path = find_latest_download()
    if not file_path:
        return False

    data = read_excel_data(file_path)
    if not data:
        print("❌ Tidak ada data untuk diupload.")
        return False

    success = upload_to_sheets(data, SPREADSHEET_ID, SHEET_NAME)

    print("\n" + "=" * 70)
    print("✅ SELESAI!" if success else "❌ GAGAL!")
    print("=" * 70)
    return success


if __name__ == "__main__":
    main()
