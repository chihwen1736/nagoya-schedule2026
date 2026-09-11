#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
2026 名古屋亞運｜中華隊每日賽程查詢系統
Excel -> JSON 轉換工具

用途：
  將官方 Excel 賽程表（每個日期一個分頁）轉換為網站使用的
  data/schedule.json，並輸出資料驗證報告（validation_report.json / .txt）。

使用方式：
  python3 convert_excel_to_json.py <輸入Excel路徑> [輸出資料夾]

原則（務必遵守，勿更動）：
  1. 不自行補齊、推測、修改任何賽程內容（時間、對手、場館、選手等）。
  2. 不因看起來重複而刪除資料。
  3. 只進行「顯示 / 搜尋標準化」（例如：去除多餘空白、統一運動種類的
     搜尋別名），原始文字一律完整保留在 raw 欄位中。
  4. 無法解析的時間，保留原始文字，且必須顯示，只是排序時排在
     「可解析時間」之後。
"""

import re
import json
import datetime
from pathlib import Path

import openpyxl

# ---------------------------------------------------------------------------
# 設定
# ---------------------------------------------------------------------------

# 欄位順序（對應 Excel 第 5 列標題，第 6 列起為資料）
COLUMNS = [
    "sport",      # 運動種類
    "time",       # 時間
    "event",      # 項目(量級)及階段
    "athletes",   # 參賽選手
    "opponent",   # 對手國(姓名)
    "score",      # 成績(比數)
    "venue",      # 地點
    "rank",       # 名次
    "note",       # 備註
]

HEADER_ROW = 5
DATA_START_ROW = 6
DATE_CELL = "B3"
EXCLUDED_SHEETS = {"範例"}

# 已知的運動種類別名（僅用於「搜尋 / 篩選」比對，不影響原始顯示文字）
SPORT_ALIASES = {
    "田俓": "田徑",
}

FULLWIDTH_SPACE = "\u3000"


def is_blank(value) -> bool:
    """判斷儲存格是否視為空白（含全形空白、純換行/空白字元）。"""
    if value is None:
        return True
    s = str(value).replace(FULLWIDTH_SPACE, "").strip()
    return s == ""


def clean_display(value) -> str:
    """顯示用清理：去除頭尾空白與全形空白，統一換行為 \n，不改變文字內容本身。"""
    if value is None:
        return ""
    s = str(value)
    s = s.replace(FULLWIDTH_SPACE, "").replace("\r\n", "\n").replace("\r", "\n")
    # 去除每行頭尾空白，但保留行與行之間的斷行（多筆時間/多位對手等資訊）
    lines = [ln.strip() for ln in s.split("\n")]
    lines = [ln for ln in lines if ln != ""]
    return "\n".join(lines).strip()


def clean_search(value) -> str:
    """搜尋用清理：完全去除空白與換行，方便比對。"""
    if value is None:
        return ""
    s = str(value).replace(FULLWIDTH_SPACE, "")
    s = re.sub(r"\s+", "", s)
    return s


def normalize_sport(raw_sport: str) -> str:
    key = raw_sport.strip()
    return SPORT_ALIASES.get(key, key)


TIME_START_RE = re.compile(r"(\d{1,2}):?(\d{2})")


def parse_start_minutes(raw_time_display: str):
    """
    嘗試從時間文字中解析出「開始時間」，回傳 (是否成功, 分鐘數或None)。
    只取第一行、且只取起始時間，不做任何猜測；解析失敗一律回傳 None。
    （此為原有邏輯，維持不變，供 timeSortMinutes / timeParseable 使用）
    """
    if not raw_time_display:
        return False, None
    first_line = raw_time_display.split("\n")[0].strip()
    if first_line == "":
        return False, None
    m = TIME_START_RE.match(first_line)
    if not m:
        return False, None
    hour = int(m.group(1))
    minute = int(m.group(2))
    if hour > 23 or minute > 59:
        return False, None
    return True, hour * 60 + minute


def parse_all_start_minutes(raw_time_display: str):
    """
    解析「多時段」欄位中，每一行可靠解析出的開始時間，回傳分鐘數陣列。
    例如 "09:45-11:00\\n11:45-13:00" -> [585, 705]
        "15:00\\n19:30" -> [900, 1170]
    無法解析的行會被略過（不會讓整筆資料失敗），純粹用於「接下來」判斷，
    不影響 timeSortMinutes / timeParseable 既有邏輯，也不改變原始顯示文字。
    """
    if not raw_time_display:
        return []
    result = []
    for line in raw_time_display.split("\n"):
        line = line.strip()
        if not line:
            continue
        m = TIME_START_RE.match(line)
        if not m:
            continue
        hour = int(m.group(1))
        minute = int(m.group(2))
        if hour > 23 or minute > 59:
            continue
        result.append(hour * 60 + minute)
    return result


def time_cell_to_display(value) -> str:
    """把 Excel 儲存格的時間值轉成顯示字串（保留原始資訊，僅做必要格式化）。"""
    if value is None:
        return ""
    if isinstance(value, datetime.time):
        return f"{value.hour:02d}:{value.minute:02d}"
    if isinstance(value, datetime.datetime):
        return f"{value.hour:02d}:{value.minute:02d}"
    return clean_display(value)


def extract_sheet(ws, sheet_name: str, warnings: list):
    date_val = ws[DATE_CELL].value
    if not isinstance(date_val, (datetime.datetime, datetime.date)):
        warnings.append(f"[{sheet_name}] 日期儲存格 {DATE_CELL} 無法辨識為日期：{date_val!r}")
        return None, []

    date_str = date_val.strftime("%Y-%m-%d")

    # 驗證分頁名稱與日期是否一致（僅記錄，不更動資料）
    m = re.search(r"(\d{4})$", sheet_name)
    if m:
        mmdd = m.group(1)
        expect = f"{date_val.month:02d}{date_val.day:02d}"
        if mmdd != expect:
            warnings.append(
                f"[{sheet_name}] 分頁名稱日期({mmdd}) 與儲存格日期({expect}) 不一致，"
                f"已採用儲存格日期 {date_str}"
            )

    records = []
    max_row = ws.max_row
    for row_idx in range(DATA_START_ROW, max_row + 1):
        row_values = [ws.cell(row=row_idx, column=c + 1).value for c in range(len(COLUMNS))]
        if all(is_blank(v) for v in row_values):
            continue

        raw = dict(zip(COLUMNS, row_values))

        sport_raw = clean_display(raw["sport"])
        if sport_raw == "":
            warnings.append(f"[{sheet_name}] 第{row_idx}列缺少「運動種類」，已保留原始列但標記為異常")

        time_display = time_cell_to_display(raw["time"])
        time_ok, time_minutes = parse_start_minutes(time_display)
        time_start_minutes_all = parse_all_start_minutes(time_display)

        record = {
            "id": f"{date_val.strftime('%Y%m%d')}-{row_idx:03d}",
            "date": date_str,
            "sourceSheet": sheet_name,
            "sourceRow": row_idx,
            "sport": sport_raw,
            "normalizedSport": normalize_sport(sport_raw) if sport_raw else "",
            "time": time_display,
            "timeSortMinutes": time_minutes,
            "timeParseable": time_ok,
            "timeStartMinutes": time_start_minutes_all,
            "event": clean_display(raw["event"]),
            "athletes": clean_display(raw["athletes"]),
            "opponent": clean_display(raw["opponent"]),
            "score": clean_display(raw["score"]),
            "venue": clean_display(raw["venue"]),
            "rank": clean_display(raw["rank"]),
            "note": clean_display(raw["note"]),
        }
        # 搜尋索引（不顯示，僅供搜尋比對用）
        record["_search"] = clean_search(
            " ".join(
                [
                    record["sport"],
                    record["normalizedSport"],
                    record["event"],
                    record["athletes"],
                    record["opponent"],
                    record["venue"],
                    record["note"],
                ]
            )
        )
        records.append(record)

    return date_str, records


def format_chinese_date(d: datetime.date) -> str:
    return f"{d.year}年{d.month}月{d.day}日"


def main():
    import argparse

    parser = argparse.ArgumentParser(
        description="2026 名古屋亞運中華台北代表團每日賽程查詢系統 - Excel 轉換工具"
    )
    parser.add_argument("excel_path", help="輸入的 Excel 賽程表路徑")
    parser.add_argument(
        "out_dir", nargs="?", default=".", help="輸出資料夾（預設為目前資料夾，通常填 data）"
    )
    parser.add_argument(
        "--date",
        dest="update_date",
        default=None,
        help=(
            "手動指定「資料更新日期」文字（例如：2026年9月15日）。"
            "若不指定，預設自動使用執行轉換工具當下的日期。"
        ),
    )
    args = parser.parse_args()

    input_path = Path(args.excel_path)
    out_dir = Path(args.out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    update_date_str = args.update_date or format_chinese_date(datetime.date.today())

    wb = openpyxl.load_workbook(input_path, data_only=True)

    warnings = []
    all_records = []
    per_date_count = {}
    sheets_imported = []
    sheets_skipped = []

    date_sheets = [s for s in wb.sheetnames if s not in EXCLUDED_SHEETS]

    for sheet_name in date_sheets:
        ws = wb[sheet_name]
        date_str, records = extract_sheet(ws, sheet_name, warnings)
        if date_str is None:
            sheets_skipped.append(sheet_name)
            continue
        sheets_imported.append(sheet_name)
        per_date_count[date_str] = len(records)
        all_records.extend(records)

    # 排序：先日期，再開始時間（可解析在前），再運動種類
    def sort_key(r):
        return (
            r["date"],
            0 if r["timeParseable"] else 1,
            r["timeSortMinutes"] if r["timeParseable"] else 0,
            r["normalizedSport"],
        )

    all_records.sort(key=sort_key)

    # 統計
    total_excel_sheets = len(wb.sheetnames)
    total_date_sheets = len(date_sheets)
    total_imported_dates = len(per_date_count)
    total_records = len(all_records)

    sport_counts = {}
    unparseable_times = []
    missing_venue = 0
    missing_athletes = 0
    normalized_groups = {}
    multi_time_samples = []

    for r in all_records:
        sport_counts[r["normalizedSport"]] = sport_counts.get(r["normalizedSport"], 0) + 1
        if not r["timeParseable"]:
            unparseable_times.append({"id": r["id"], "date": r["date"], "time": r["time"], "event": r["event"]})
        if r["venue"] == "":
            missing_venue += 1
        if r["athletes"] == "":
            missing_athletes += 1
        if r["sport"] != r["normalizedSport"]:
            normalized_groups.setdefault(r["normalizedSport"], set()).add(r["sport"])
        if len(r["timeStartMinutes"]) > 1:
            multi_time_samples.append(
                {
                    "id": r["id"],
                    "date": r["date"],
                    "time": r["time"],
                    "timeStartMinutes": r["timeStartMinutes"],
                    "event": r["event"],
                }
            )

    multi_time_record_count = len(multi_time_samples)

    # 可能重複資料的簡單偵測（同日期+同運動+同時間+同項目 完全相同才視為疑似重複）
    dup_check = {}
    for r in all_records:
        key = (r["date"], r["sport"], r["time"], r["event"], r["athletes"], r["venue"])
        dup_check.setdefault(key, []).append(r["id"])
    suspected_duplicates = [
        {"date": k[0], "sport": k[1], "time": k[2], "event": k[3], "athletes": k[4], "venue": k[5], "ids": v}
        for k, v in dup_check.items()
        if len(v) > 1
    ]

    # 輸出 schedule.json（原始資料交換格式，供工具 / 除錯使用）
    schedule_path = out_dir / "schedule.json"
    with open(schedule_path, "w", encoding="utf-8") as f:
        json.dump(all_records, f, ensure_ascii=False, indent=2)

    # 輸出 schedule.js（網站實際載入使用；純前端、不需 fetch，避免本機開啟時的 CORS 問題）
    # 同時寫入 SCHEDULE_META，讓「資料更新日期」只集中在這一個地方由工具自動產生，
    # 不需要在 HTML / JS 多處手動修改。
    schedule_js_path = out_dir / "schedule.js"
    with open(schedule_js_path, "w", encoding="utf-8") as f:
        f.write("// 本檔案由 scripts/convert_excel_to_json.py 自動產生，請勿手動編輯。\n")
        f.write("// 如需更新賽程，請重新執行轉換工具，見 README.md。\n")
        f.write("window.SCHEDULE_DATA = ")
        json.dump(all_records, f, ensure_ascii=False)
        f.write(";\n")
        f.write("window.SCHEDULE_META = ")
        json.dump({"updateDate": update_date_str, "totalRecords": total_records}, f, ensure_ascii=False)
        f.write(";\n")

    # 輸出 meta.json（與 schedule.js 內容相同，僅供工具 / 除錯查看用）
    meta_path = out_dir / "meta.json"
    with open(meta_path, "w", encoding="utf-8") as f:
        json.dump({"updateDate": update_date_str, "totalRecords": total_records}, f, ensure_ascii=False, indent=2)

    # 輸出驗證報告
    report = {
        "generatedAt": datetime.datetime.now().isoformat(),
        "updateDateShownOnSite": update_date_str,
        "sourceFile": str(input_path.name),
        "excelTotalSheets": total_excel_sheets,
        "excelDateSheets": total_date_sheets,
        "sheetsImported": sheets_imported,
        "sheetsSkipped": sheets_skipped,
        "totalDatesImported": total_imported_dates,
        "totalRecords": total_records,
        "recordsPerDate": per_date_count,
        "recordsPerSport": sport_counts,
        "normalizedSportGroups": {k: sorted(v) for k, v in normalized_groups.items()},
        "missingVenueCount": missing_venue,
        "missingAthletesCount": missing_athletes,
        "unparseableTimeCount": len(unparseable_times),
        "unparseableTimeSamples": unparseable_times,
        "multiTimeRecordCount": multi_time_record_count,
        "multiTimeSamples": multi_time_samples,
        "suspectedDuplicateGroups": suspected_duplicates,
        "warnings": warnings,
    }
    report_path = out_dir / "validation_report.json"
    with open(report_path, "w", encoding="utf-8") as f:
        json.dump(report, f, ensure_ascii=False, indent=2)

    # 印出摘要
    print("=== 轉換完成 ===")
    print(f"Excel 分頁總數：{total_excel_sheets}（扣除範例分頁後日期分頁：{total_date_sheets}）")
    print(f"成功匯入日期數：{total_imported_dates}")
    print(f"系統資料筆數：{total_records}")
    print(f"無法解析時間筆數：{len(unparseable_times)}")
    print(f"多時段資料筆數：{multi_time_record_count}")
    print(f"疑似重複群組數：{len(suspected_duplicates)}")
    print(f"資料更新日期（將顯示於網站頁尾）：{update_date_str}")
    print(f"已輸出：{schedule_path}")
    print(f"已輸出：{schedule_js_path}")
    print(f"已輸出：{meta_path}")
    print(f"已輸出：{report_path}")


if __name__ == "__main__":
    main()
