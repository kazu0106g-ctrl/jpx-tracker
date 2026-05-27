"""
直近 DAYS_BACK 日の平日について、スプレッドシートにデータが入っているか確認し、
欠損があれば JPX PDF から自動補完する。

毎日 02:00 UTC (11:00 JST) に実行することで、前日 cron が GitHub Actions の
スキップで失敗していても翌朝に自動回復する。
"""

import sys
from datetime import date, timedelta
from jpx_scraper import (
    get_sheet,
    download_pdf,
    extract_trading_values,
    write_to_sheet,
    year_start_col,
)

DAYS_BACK = 7


def safe_download(d: date):
    """download_pdf は 404 で sys.exit(0) するので SystemExit を握りつぶす"""
    try:
        return download_pdf(d)
    except SystemExit:
        return None  # 祝日・休場日 (PDFが存在しない)


def main():
    ws = get_sheet()
    today = date.today()
    print(f"[INFO] 過去 {DAYS_BACK} 日の欠損日をチェック中 (today={today})")

    # 過去 DAYS_BACK 日の平日を列挙
    weekdays_to_check = []
    for i in range(1, DAYS_BACK + 1):
        d = today - timedelta(days=i)
        if d.weekday() < 5:  # 月〜金
            weekdays_to_check.append(d)

    # 各年のデータ列に対象日付があるかチェック
    missing = []
    for d in weekdays_to_check:
        start_col = year_start_col(d.year)
        col_values = ws.col_values(start_col)
        date_str = d.strftime("%Y/%m/%d")
        if date_str not in col_values:
            missing.append(d)

    if not missing:
        print(f"[OK] 直近 {DAYS_BACK} 日 ({len(weekdays_to_check)} 平日) に欠損なし")
        return

    print(f"[WARN] {len(missing)} 日の欠損を検出: {[d.isoformat() for d in missing]}")
    ok = err = skip = 0
    for d in missing:
        try:
            pdf = safe_download(d)
            if pdf is None:
                print(f"[SKIP] {d}: PDF が存在しない (祝日/休場日)")
                skip += 1
                continue
            values = extract_trading_values(pdf)
            write_to_sheet(ws, d, values)
            ok += 1
        except Exception as e:
            print(f"[ERROR] {d}: {e}")
            err += 1

    print(f"[DONE] 補完={ok} / スキップ={skip} / エラー={err}")
    if err > 0:
        sys.exit(1)


if __name__ == "__main__":
    main()
