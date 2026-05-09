"""
JPX 後場 取引代金概算 トラッカー
毎日16:10以降に実行し、プライム・スタンダード・グロース市場の取引代金をGoogleスプレッドシートに記録する
"""

import sys
import io
import os
import re
import json
import requests
import pdfplumber
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime, date

# ---- 設定 ----
SPREADSHEET_ID = "1_9Av480JcBfne-s2un62v0eddXnbJAlmNEo6Ou7SNZs"
SHEET_NAME = "シート3"
BASE_URL = "https://www.jpx.co.jp/markets/equities/volume-and-value/tvdivq000000derc-att"

PRIME_LABEL    = "内国株式・プライム市場"
STANDARD_LABEL = "内国株式・スタンダード市場"
GROWTH_LABEL   = "内国株式・グロース市場"

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
HEADERS = ["日付", "プライム市場（百万円）", "スタンダード市場（百万円）", "グロース市場（百万円）"]

# 年ごとに横方向に並べる列レイアウト
# 各年: 4データ列 + 1空列セパレータ = 5列
# 例) 2023=A-D, 空=E, 2024=F-I, 空=J, 2025=K-N, 空=O, 2026=P-S, ...
COLS_PER_YEAR = 5
BASE_YEAR = 2023  # 最初の年。これを起点に列位置を計算
HEADER_LABEL_ROW = 1   # Row 1: 年ラベル ("2023年")
HEADER_COLS_ROW = 2    # Row 2: 列ヘッダー (日付/プライム/...)
DATA_START_ROW = 3     # Row 3〜: 日次データ


def col_letter(n: int) -> str:
    """1=A, 2=B, ... 27=AA, ..."""
    s = ""
    while n > 0:
        m = (n - 1) % 26
        s = chr(65 + m) + s
        n = (n - 1) // 26
    return s


def year_start_col(year: int) -> int:
    """対象年の開始列番号 (1-based)"""
    return (year - BASE_YEAR) * COLS_PER_YEAR + 1


def get_sheet() -> gspread.Worksheet:
    """サービスアカウント認証してシートを返す"""
    creds_json = os.environ.get("GOOGLE_CREDENTIALS")
    if not creds_json:
        # ローカル実行時はJSONファイルから読む
        creds_path = os.path.join(os.path.dirname(__file__), "credentials.json")
        with open(creds_path) as f:
            creds_json = f.read()

    creds_info = json.loads(creds_json)
    creds = Credentials.from_service_account_info(creds_info, scopes=SCOPES)
    gc = gspread.authorize(creds)
    return gc.open_by_key(SPREADSHEET_ID).worksheet(SHEET_NAME)


def download_pdf(target_date: date) -> bytes:
    """後場PDFをダウンロードして返す"""
    date_str = target_date.strftime("%Y%m%d")
    url = f"{BASE_URL}/2_{date_str}.pdf"
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
    resp = requests.get(url, headers=headers, timeout=30)
    if resp.status_code == 404:
        print(f"[SKIP] PDFが存在しません（祝日・休場日の可能性）: {url}")
        sys.exit(0)
    if resp.status_code != 200:
        raise RuntimeError(f"PDFダウンロード失敗: HTTP {resp.status_code} / {url}")
    return resp.content


def extract_trading_values(pdf_bytes: bytes) -> dict:
    """
    PDFから取引代金概算（立会内）を抽出する
    戻り値: {"prime": int, "standard": int, "growth": int}  単位: 百万円
    """
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        words = pdf.pages[0].extract_words()

    prime_ys = [w["top"] for w in words if w["text"] == PRIME_LABEL]
    # 市場名は3回出現（メインテーブル・売買高概算・取引代金概算）
    if len(prime_ys) < 3:
        raise ValueError(f"取引代金概算セクションが見つかりません（出現回数={len(prime_ys)}）")

    values = {}
    for label_text, key in [
        (PRIME_LABEL,    "prime"),
        (STANDARD_LABEL, "standard"),
        (GROWTH_LABEL,   "growth"),
    ]:
        occurrences = [i for i, w in enumerate(words) if w["text"] == label_text]
        if len(occurrences) < 3:
            raise ValueError(f"ラベルが3回見つかりません: {label_text}")
        idx = occurrences[2]
        y = words[idx]["top"]
        num_word = next(
            (w for w in words[idx+1:] if abs(w["top"] - y) < 3 and re.match(r"[\d,]+$", w["text"])),
            None,
        )
        if num_word is None:
            raise ValueError(f"数値が見つかりません: {label_text}")
        values[key] = int(num_word["text"].replace(",", ""))

    return values


def write_to_sheet(ws: gspread.Worksheet, target_date: date, values: dict):
    """スプレッドシートに書き込む (年ごと横方向レイアウト)

    Layout:
      Row 1: 年ラベル ("2023年", "2024年", ...)
      Row 2: 列ヘッダー (日付, プライム, スタンダード, グロース)
      Row 3〜: 日次データ
      列方向: 各年が5列ブロック (4データ列 + 1空列セパレータ)
    """
    year = target_date.year
    start_col = year_start_col(year)
    end_col = start_col + 3  # データ4列の最終列
    date_str = target_date.strftime("%Y/%m/%d")
    new_row = [date_str, values["prime"], values["standard"], values["growth"]]

    # 1. その年のヘッダー (Row 1, 2) が無ければ作成
    label_cell = ws.cell(HEADER_LABEL_ROW, start_col).value
    if not label_cell or str(label_cell).strip() != f"{year}年":
        ws.update_cell(HEADER_LABEL_ROW, start_col, f"{year}年")
        print(f"[INIT] 年ラベル {year}年 を {col_letter(start_col)}{HEADER_LABEL_ROW} に書き込み")
    header_cells = ws.range(HEADER_COLS_ROW, start_col, HEADER_COLS_ROW, end_col)
    if [c.value for c in header_cells] != HEADERS:
        ws.update(
            f"{col_letter(start_col)}{HEADER_COLS_ROW}:{col_letter(end_col)}{HEADER_COLS_ROW}",
            [HEADERS],
        )
        print(f"[INIT] 列ヘッダー を {col_letter(start_col)}{HEADER_COLS_ROW}〜{col_letter(end_col)}{HEADER_COLS_ROW} に書き込み")

    # 2. その年の列ブロックの日付列 (start_col) を読み出し
    last_row_used = ws.row_count
    date_col_values = ws.col_values(start_col)  # 1-based全行 (空セルは末尾trim)
    # date_col_values[0] = Row 1 の値 (年ラベル), [1] = Row 2 (日付ヘッダー), [2..] = Row 3〜のデータ

    # 同日付があれば上書き
    for idx, dv in enumerate(date_col_values[DATA_START_ROW - 1:], start=DATA_START_ROW):
        if str(dv).strip() == date_str:
            ws.update(
                f"{col_letter(start_col)}{idx}:{col_letter(end_col)}{idx}",
                [new_row],
            )
            print(f"[UPDATE] {date_str} を上書き: プライム={values['prime']:,} / スタンダード={values['standard']:,} / グロース={values['growth']:,}")
            return

    # 3. 追記先 = その年のデータ列の最初の空行
    # date_col_values の長さがデータ末尾を示す (gspreadは末尾の空セルをtrim)
    next_row = max(DATA_START_ROW, len(date_col_values) + 1)

    ws.update(
        f"{col_letter(start_col)}{next_row}:{col_letter(end_col)}{next_row}",
        [new_row],
    )
    print(f"[OK] {date_str} を記録: {col_letter(start_col)}-{col_letter(end_col)}{next_row} / プライム={values['prime']:,} / スタンダード={values['standard']:,} / グロース={values['growth']:,} （百万円）")


def main(target_date: date = None):
    if target_date is None:
        target_date = date.today()

    print(f"[INFO] {target_date} の後場データを取得中...")
    try:
        pdf_bytes = download_pdf(target_date)
        values = extract_trading_values(pdf_bytes)
        ws = get_sheet()
        write_to_sheet(ws, target_date, values)
    except Exception as e:
        print(f"[ERROR] {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    if len(sys.argv) > 1:
        d = datetime.strptime(sys.argv[1], "%Y%m%d").date()
    else:
        d = date.today()
    main(d)
