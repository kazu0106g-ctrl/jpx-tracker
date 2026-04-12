"""
JPX 過去データ一括取得スクリプト
jpx-tracker リポジトリの jpx_scraper.py と同じフォルダに置いて実行する

Usage:
  python backfill.py 20250101 20251231
  python backfill.py 20250101          # 開始日のみ指定 → 今日まで
"""

import sys
from datetime import date, timedelta, datetime
from jpx_scraper import download_pdf, extract_trading_values, get_sheet, write_to_sheet


def iter_weekdays(start: date, end: date):
    """start〜end の平日（月〜金）を順に返す"""
    d = start
    while d <= end:
        if d.weekday() < 5:
            yield d
        d += timedelta(days=1)


def backfill(start: date, end: date):
    print(f"[INFO] {start} 〜 {end} の範囲で取得します（平日のみ）")
    ws = get_sheet()
    ok = skip = err = 0

    for d in iter_weekdays(start, end):
        try:
            pdf_bytes = download_pdf(d)
            values = extract_trading_values(pdf_bytes)
            write_to_sheet(ws, d, values)
            ok += 1
        except SystemExit:
            # download_pdf が 404 のとき sys.exit(0) を呼ぶ → 祝日・休場日
            print(f"[SKIP] {d}: 非営業日（PDF不在）")
            skip += 1
        except Exception as e:
            print(f"[ERROR] {d}: {e}", file=sys.stderr)
            err += 1

    print(f"\n完了: 取得={ok}件 / スキップ={skip}件 / エラー={err}件")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python backfill.py YYYYMMDD [YYYYMMDD]")
        sys.exit(1)

    start = datetime.strptime(sys.argv[1], "%Y%m%d").date()
    end   = datetime.strptime(sys.argv[2], "%Y%m%d").date() if len(sys.argv) > 2 else date.today()

    if start > end:
        print("[ERROR] 開始日が終了日より後になっています")
        sys.exit(1)

    backfill(start, end)
