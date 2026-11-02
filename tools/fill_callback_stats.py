# -*- coding: utf-8 -*-
"""
電訪回收率統計自動填表（小於20%統計表）

用法（在你的 Windows 電腦上執行）：
    python fill_callback_stats.py "D:\\電訪追蹤情形\\11510"
    python fill_callback_stats.py "D:\\電訪追蹤情形\\11510" --dry-run   # 只預覽，不寫入

規則：
  1. 掃描月份資料夾（例如 D:\\電訪追蹤情形\\11510）內所有 .xlsx / .csv 名單。
  2. 從檔名取得日期（支援 1151015、20261015、2026-10-15 等格式）。
  3. 在名單中找出「醫師」欄位，統計每位醫師的筆數。
  4. 填入統計表「115年回收率」：
       - A2~A21 為醫師名字，依名字對應列
       - 日期對應到標題列（第 1 列）的日期欄
       - 格子內容 = 該醫師當天的筆數（1 筆就填 1）
  5. 寫入前會先備份統計表（同資料夾，檔名加上時間戳記）。
"""

import argparse
import csv
import datetime as dt
import re
import shutil
import sys
from collections import Counter, defaultdict
from pathlib import Path

import openpyxl

# ===================== 設定區（依你的實際檔案調整） =====================
STATS_FILE = Path(r"\\10.2.49.124\Volume_1\20 TQM QCC N4\112 提升DM收案\小於20%統計.xlsx")
STATS_SHEET = "115年回收率"
DOCTOR_COL = 1                 # 統計表 A 欄 = 醫師名字
DOCTOR_ROWS = range(2, 22)     # A2 ~ A21
DATE_HEADER_ROW = 1            # 統計表日期標題所在的列
DOCTOR_HEADER_KEYWORDS = ("醫師",)  # 名單中用來找醫師欄位的標題關鍵字
# =====================================================================

DATE_PATTERNS = [
    (re.compile(r"(?<!\d)(\d{3})(\d{2})(\d{2})(?!\d)"), "roc"),      # 1151015 (民國)
    (re.compile(r"(?<!\d)(\d{4})(\d{2})(\d{2})(?!\d)"), "ad"),       # 20261015
    (re.compile(r"(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})"), "ad"),      # 2026-10-15
]


def normalize(value):
    """去除前後與中間空白，統一成字串，方便比對名字。"""
    if value is None:
        return ""
    return re.sub(r"\s+", "", str(value))


def parse_date_from_text(text):
    """從檔名或字串解析日期，失敗回傳 None。"""
    for pattern, kind in DATE_PATTERNS:
        m = pattern.search(text)
        if not m:
            continue
        y, mo, d = (int(x) for x in m.groups())
        if kind == "roc":
            y += 1911
        try:
            return dt.date(y, mo, d)
        except ValueError:
            continue
    return None


def parse_date_cell(value):
    """統計表標題列的日期可能是 datetime 或文字。"""
    if isinstance(value, dt.datetime):
        return value.date()
    if isinstance(value, dt.date):
        return value
    if value is None:
        return None
    return parse_date_from_text(str(value))


def read_source_file(path):
    """讀取單一名單檔，回傳 Counter：{醫師名字: 筆數}。"""
    counts = Counter()
    suffix = path.suffix.lower()

    if suffix == ".csv":
        with open(path, encoding="utf-8-sig", newline="") as f:
            rows = list(csv.reader(f))
        sheets = [rows]
    elif suffix in (".xlsx", ".xlsm"):
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        sheets = [list(ws.iter_rows(values_only=True)) for ws in wb.worksheets]
        wb.close()
    else:
        return counts

    for rows in sheets:
        doctor_idx = None
        start = 0
        # 找標題列與醫師欄位
        for i, row in enumerate(rows[:10]):
            for j, cell in enumerate(row):
                if cell and any(k in str(cell) for k in DOCTOR_HEADER_KEYWORDS):
                    doctor_idx = j
                    start = i + 1
                    break
            if doctor_idx is not None:
                break
        if doctor_idx is None:
            continue
        for row in rows[start:]:
            if doctor_idx < len(row):
                name = normalize(row[doctor_idx])
                if name:
                    counts[name] += 1
    return counts


def collect_by_date(folder):
    """依檔名日期彙總所有名單，回傳 {date: Counter}。"""
    by_date = defaultdict(Counter)
    for path in sorted(folder.iterdir()):
        if not path.is_file() or path.suffix.lower() not in (".xlsx", ".xlsm", ".csv"):
            continue
        if path.name.startswith("~$"):  # Excel 暫存檔，略過
            continue
        date = parse_date_from_text(path.stem)
        if date is None:
            print(f"[略過] 檔名看不出日期：{path.name}")
            continue
        counts = read_source_file(path)
        print(f"[讀取] {path.name} → {date}，醫師 {len(counts)} 位，共 {sum(counts.values())} 筆")
        by_date[date] += counts
    return by_date


def fill_stats(by_date, stats_file, dry_run=False):
    wb = openpyxl.load_workbook(stats_file)
    if STATS_SHEET not in wb.sheetnames:
        sys.exit(f"找不到工作表「{STATS_SHEET}」，現有：{wb.sheetnames}")
    ws = wb[STATS_SHEET]

    # 醫師名字 → 列號
    doctor_rows = {}
    for r in DOCTOR_ROWS:
        name = normalize(ws.cell(r, DOCTOR_COL).value)
        if name:
            doctor_rows[name] = r

    # 日期 → 欄號
    date_cols = {}
    for c in range(1, ws.max_column + 1):
        d = parse_date_cell(ws.cell(DATE_HEADER_ROW, c).value)
        if d:
            date_cols[d] = c

    unmatched_names = set()
    for date in sorted(by_date):
        counts = by_date[date]
        col = date_cols.get(date)
        if col is None:
            print(f"[警告] 統計表找不到日期 {date} 的欄位，略過這一天（請確認第 {DATE_HEADER_ROW} 列標題）")
            continue
        for name, n in counts.items():
            row = doctor_rows.get(name)
            if row is None:
                unmatched_names.add(name)
                continue
            print(f"  {date}  {name}  →  第{row}列 / 第{col}欄 = {n}")
            if not dry_run:
                ws.cell(row, col).value = n  # 設定為當天筆數（不累加，重跑不會重複計算）

    if unmatched_names:
        print("[警告] 以下醫師不在 A2~A21，未填入：", "、".join(sorted(unmatched_names)))

    if dry_run:
        print("\n預覽模式：未寫入任何檔案。")
        return

    backup = stats_file.with_name(
        f"{stats_file.stem}_備份_{dt.datetime.now():%Y%m%d_%H%M%S}{stats_file.suffix}"
    )
    shutil.copy2(stats_file, backup)
    print(f"[備份] {backup}")
    try:
        wb.save(stats_file)
    except PermissionError:
        sys.exit("無法寫入：統計表可能正被 Excel 開啟，請關閉後再執行。")
    print(f"[完成] 已寫入 {stats_file}")


def main():
    parser = argparse.ArgumentParser(description="電訪名單彙總填入小於20%統計表")
    parser.add_argument("folder", help=r"月份資料夾，例如 D:\電訪追蹤情形\11510")
    parser.add_argument("--stats", default=str(STATS_FILE), help="統計表路徑（預設見程式設定區）")
    parser.add_argument("--dry-run", action="store_true", help="只預覽，不寫入")
    args = parser.parse_args()

    folder = Path(args.folder)
    if not folder.is_dir():
        sys.exit(f"找不到資料夾：{folder}")
    stats_file = Path(args.stats)
    if not stats_file.is_file():
        sys.exit(f"找不到統計表：{stats_file}")

    by_date = collect_by_date(folder)
    if not by_date:
        sys.exit("資料夾內沒有可讀取的名單檔。")
    fill_stats(by_date, stats_file, dry_run=args.dry_run)


if __name__ == "__main__":
    main()
