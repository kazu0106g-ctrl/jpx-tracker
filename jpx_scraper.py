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
    """スプレッドシートに書き込む"""
    all_values = ws.get_all_values()

    # ヘッダーがなければ追加
    if not all_values or all_values[0] != HEADERS:
        ws.insert_row(HEADERS, 1)
        all_values = ws.get_all_values()

    date_str = target_date.strftime("%Y/%m/%d")
    new_row = [date_str, values["prime"], values["standard"], values["growth"]]

    # 同日付が既存なら上書き
    for i, row in enumerate(all_values[1:], start=2):
        if row and row[0] == date_str:
            ws.update(f"A{i}:D{i}", [new_row])
            print(f"[UPDATE] {date_str} を上書きしました: プライム={values['prime']:,} / スタンダード={values['standard']:,} / グロース={values['growth']:,}")
            return

    # 年替わりチェック: 最終データ行と年が違えば空行を挿入
    data_rows = [r for r in all_values[1:] if r and r[0]]
    if data_rows:
        last_date_str = data_rows[-1][0]
        try:
            last_year = datetime.strptime(last_date_str, "%Y/%m/%d").year
            if last_year != target_date.year:
                ws.append_row([])
        except ValueError:
            pass

    ws.append_row(new_row)
    print(f"[OK] {date_str} を記録: プライム={values['prime']:,} / スタンダード={values['standard']:,} / グロース={values['growth']:,} （百万円）")


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
