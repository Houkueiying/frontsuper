# -*- coding: utf-8 -*-
"""
電訪回收率統計自動填表（小於20%統計表）

用法（在你的 Windows 電腦上執行）：
    python fill_callback_stats.py "D:\\電訪追蹤情形\\11510"
    python fill_callback_stats.py "D:\\電訪追蹤情形\\11510" --dry-run   # 只預覽，不寫入

統計表格式（工作表「115年回收率」）：
    第 1 列  日期標題（B1、C1…，例如「1月2日」「10月1日」或日期格式）
    第 2 列  實際看診人數      ＝ 當日名單總筆數 − AE 欄含「[未看診]」的筆數
    第 4~21 列 A 欄醫師名字    ＝ 當日該醫師的名單筆數（同醫師兩個不同個案 ＝ 2）
    第 22 列 病人實際衛教人數  ＝ 當日黃色底（已電訪）的總筆數

名單格式（月份資料夾內的 .xlsx / .csv）：
    G 欄  醫師名字
    AE 欄 含「[未看診]」表示未看診
    姓名欄（標題含「姓名」或「病人」）有黃色底圖 ＝ 已電訪
    檔名含日期，例如 1151015.xlsx（民國）或 20261015.xlsx
寫入前會先備份統計表。
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
from openpyxl.utils import column_index_from_string

# ===================== 設定區（依你的實際檔案調整） =====================
STATS_FILE = Path(r"\\10.2.49.124\Volume_1\20 TQM QCC N4\112 提升DM收案\小於20%統計.xlsx")
STATS_SHEET = "115年回收率"
STATS_DOCTOR_COL = 1               # 統計表 A 欄 = 醫師名字
STATS_DOCTOR_ROWS = range(4, 22)   # A4 ~ A21
STATS_HEADER_ROW = 1               # 日期標題列
STATS_TOTAL_ROW = 2                # 實際看診人數
STATS_CALLBACK_ROW = 22            # 病人實際衛教人數

SRC_DOCTOR_COL = "G"               # 名單中醫師名字欄
SRC_NOTSEEN_COL = "AE"             # 名單中「[未看診]」所在欄
SRC_NOTSEEN_TEXT = "[未看診]"
SRC_PATIENT_HEADERS = ("姓名", "病人")  # 依序尋找標題，用來判斷黃底（電訪）的欄位
# =====================================================================

FILE_DATE_PATTERNS = [
    (re.compile(r"(?<!\d)(\d{3})(\d{2})(\d{2})(?!\d)"), "roc"),   # 1151015 (民國)
    (re.compile(r"(?<!\d)(\d{4})(\d{2})(\d{2})(?!\d)"), "ad"),    # 20261015
    (re.compile(r"(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})"), "ad"),   # 2026-10-15
]
HEADER_MD_PATTERN = re.compile(r"(\d{1,2})\s*月\s*(\d{1,2})\s*日?")


def normalize(value):
    """去除所有空白，統一成字串，方便比對名字。"""
    if value is None:
        return ""
    return re.sub(r"\s+", "", str(value))


def md_from_file_name(text):
    """從檔名解析 (月, 日)，失敗回傳 None。"""
    for pattern, kind in FILE_DATE_PATTERNS:
        m = pattern.search(text)
        if not m:
            continue
        y, mo, d = (int(x) for x in m.groups())
        try:
            dt.date(y + 1911 if kind == "roc" else y, mo, d)
        except ValueError:
            continue
        return (mo, d)
    return None


def md_from_header(value):
    """統計表標題列 → (月, 日)。支援 datetime 與「10月1日」文字。"""
    if isinstance(value, (dt.datetime, dt.date)):
        return (value.month, value.day)
    if value is None:
        return None
    m = HEADER_MD_PATTERN.search(str(value))
    if m:
        return (int(m.group(1)), int(m.group(2)))
    return None


def is_yellow(cell):
    """判斷儲存格是否為黃色底圖（RGB 接近 FFFF00）。"""
    fill = cell.fill
    if fill is None or fill.fill_type != "solid":
        return False
    color = fill.start_color
    if color is None or color.type != "rgb" or not isinstance(color.rgb, str):
        return False
    rgb = color.rgb[-6:]
    try:
        r, g, b = int(rgb[0:2], 16), int(rgb[2:4], 16), int(rgb[4:6], 16)
    except ValueError:
        return False
    return r >= 220 and g >= 200 and b <= 120


def read_source_file(path):
    """
    讀取單一名單檔，回傳 (總筆數, 未看診筆數, 電訪筆數 Counter{醫師: n}, 電訪總數)。
    """
    doctor_idx = column_index_from_string(SRC_DOCTOR_COL) - 1
    notseen_idx = column_index_from_string(SRC_NOTSEEN_COL) - 1

    total = notseen = callback_total = 0
    records = Counter()   # 每位醫師的全部名單筆數

    if path.suffix.lower() == ".csv":
        with open(path, encoding="utf-8-sig", newline="") as f:
            sheets = [[(row, None) for row in csv.reader(f)]]
    else:
        wb = openpyxl.load_workbook(path)  # 需要讀取底色，不用 read_only
        sheets = []
        for ws in wb.worksheets:
            rows = []
            for r in ws.iter_rows():
                values = [c.value for c in r]
                rows.append((values, r))
            sheets.append(rows)

    for rows in sheets:
        if not rows:
            continue
        header = rows[0][0]
        patient_idx = None
        for keyword in SRC_PATIENT_HEADERS:
            for j, v in enumerate(header):
                if v and keyword in str(v):
                    patient_idx = j
                    break
            if patient_idx is not None:
                break

        for values, cells in rows[1:]:
            if doctor_idx >= len(values):
                continue
            name = normalize(values[doctor_idx])
            if not name:
                continue
            total += 1
            records[name] += 1

            if notseen_idx < len(values) and values[notseen_idx] and SRC_NOTSEEN_TEXT in str(values[notseen_idx]):
                notseen += 1

            # 電訪：病人欄為黃底（找不到病人欄時，任一儲存格黃底即視為電訪）
            if cells is not None:
                if patient_idx is not None:
                    called = is_yellow(cells[patient_idx])
                else:
                    called = any(is_yellow(c) for c in cells)
            else:
                called = False
            if called:
                callback_total += 1

    return total, notseen, records, callback_total


def collect_by_date(folder):
    """依檔名 (月, 日) 彙總所有名單。"""
    agg = defaultdict(lambda: {"total": 0, "notseen": 0, "records": Counter(), "callback_total": 0})
    for path in sorted(folder.iterdir()):
        if not path.is_file() or path.suffix.lower() not in (".xlsx", ".xlsm", ".csv"):
            continue
        if path.name.startswith("~$"):
            continue
        md = md_from_file_name(path.stem)
        if md is None:
            print(f"[略過] 檔名看不出日期：{path.name}")
            continue
        total, notseen, records, cb_total = read_source_file(path)
        print(f"[讀取] {path.name} → {md[0]}/{md[1]}：總 {total}、未看診 {notseen}、電訪 {cb_total}")
        a = agg[md]
        a["total"] += total
        a["notseen"] += notseen
        a["records"] += records
        a["callback_total"] += cb_total
    return agg


def fill_stats(agg, stats_file, dry_run=False):
    wb = openpyxl.load_workbook(stats_file)
    if STATS_SHEET not in wb.sheetnames:
        sys.exit(f"找不到工作表「{STATS_SHEET}」，現有：{wb.sheetnames}")
    ws = wb[STATS_SHEET]

    doctor_rows = {}
    for r in STATS_DOCTOR_ROWS:
        name = normalize(ws.cell(r, STATS_DOCTOR_COL).value)
        if name:
            doctor_rows[name] = r

    date_cols = {}
    for c in range(2, ws.max_column + 1):
        md = md_from_header(ws.cell(STATS_HEADER_ROW, c).value)
        if md is None:
            continue
        if md in date_cols:
            print(f"[警告] 統計表日期 {md[0]}/{md[1]} 重複出現（第 {date_cols[md]} 欄與第 {c} 欄），使用第 {date_cols[md]} 欄")
            continue
        date_cols[md] = c

    unmatched = set()
    for md in sorted(agg):
        a = agg[md]
        col = date_cols.get(md)
        if col is None:
            print(f"[警告] 統計表找不到 {md[0]}月{md[1]}日 的欄位，略過（請確認第 {STATS_HEADER_ROW} 列標題）")
            continue

        seen = a["total"] - a["notseen"]
        print(f"  {md[0]}/{md[1]} → 第{col}欄：實際看診 {seen}，病人實際衛教 {a['callback_total']}")
        if not dry_run:
            ws.cell(STATS_TOTAL_ROW, col).value = seen
            ws.cell(STATS_CALLBACK_ROW, col).value = a["callback_total"]

        for name, n in a["records"].items():
            row = doctor_rows.get(name)
            if row is None:
                unmatched.add(name)
                continue
            print(f"      {name} → 第{row}列 = {n}")
            if not dry_run:
                ws.cell(row, col).value = n

    if unmatched:
        print("[警告] 以下醫師不在 A4~A21，未填入：", "、".join(sorted(unmatched)))

    if dry_run:
        print("\n預覽模式：未寫入任何檔案。")
        return

    backup = stats_file.with_name(f"{stats_file.stem}_備份_{dt.datetime.now():%Y%m%d_%H%M%S}{stats_file.suffix}")
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

    agg = collect_by_date(folder)
    if not agg:
        sys.exit("資料夾內沒有可讀取的名單檔。")
    fill_stats(agg, stats_file, dry_run=args.dry_run)


if __name__ == "__main__":
    main()
