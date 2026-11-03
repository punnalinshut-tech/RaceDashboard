import time
from pathlib import Path

import xlwings as xw

BASE_DIR = Path(__file__).resolve().parent
SOURCE_FILE = BASE_DIR / "data" / "race_5KM.xlsx"
TARGET_FILE = Path(r"C:\RaceDashboardV3\data\race_results.xlsx")

CHECK_INTERVAL = 1.0

SYNC_SHEETS = [
    "M19_5KM",
    "F19_5KM",
    "M20_5KM",
    "F20_5KM",
    "M30_5KM",
    "F30_5KM",
    "M40_5KM",
    "F40_5KM",
    "M50_5KM",
    "F50_5KM",
    "M60_5KM",
    "F60_5KM",
]

def get_last_data_row(sheet):
    last_rows = []
    for col in ("A", "B", "C", "D"):
        row = sheet.range(f"{col}{sheet.cells.last_cell.row}").end("up").row
        last_rows.append(row)
    return max(last_rows)

def sync_workbook():
    if not SOURCE_FILE.exists():
        raise FileNotFoundError(f"ไม่พบไฟล์ต้นทาง: {SOURCE_FILE}")
    if not TARGET_FILE.exists():
        raise FileNotFoundError(f"ไม่พบไฟล์กลาง: {TARGET_FILE}")

    app = None
    src_wb = None
    dst_wb = None

    try:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = False

        src_wb = app.books.open(
            str(SOURCE_FILE),
            update_links=False,
            read_only=True,
            ignore_read_only_recommended=True,
        )

        dst_wb = app.books.open(
            str(TARGET_FILE),
            update_links=False,
            read_only=False,
            ignore_read_only_recommended=True,
        )

        src_sheet_names = [s.name for s in src_wb.sheets]
        dst_sheet_names = [s.name for s in dst_wb.sheets]

        for sheet_name in SYNC_SHEETS:
            if sheet_name not in src_sheet_names:
                print(f"[WARN] ไม่พบ Sheet ต้นทาง: {sheet_name}")
                continue
            if sheet_name not in dst_sheet_names:
                print(f"[WARN] ไม่พบ Sheet ปลายทาง: {sheet_name}")
                continue

            src_ws = src_wb.sheets[sheet_name]
            dst_ws = dst_wb.sheets[sheet_name]

            last_row = get_last_data_row(src_ws)

            dst_ws.range("A2:D10000").clear_contents()

            if last_row >= 2:
                values = src_ws.range(f"A2:D{last_row}").value
                if last_row == 2 and isinstance(values, list):
                    values = [values]
                dst_ws.range("A2").value = values

            print(f"[SYNC] {sheet_name} rows={max(last_row - 1, 0)}")

        dst_wb.save()
        print("[OK] 5 KM -> race_results.xlsx อัปเดตเรียบร้อย")

    finally:
        if src_wb is not None:
            try:
                src_wb.close()
            except Exception:
                pass
        if dst_wb is not None:
            try:
                dst_wb.close()
            except Exception:
                pass
        if app is not None:
            try:
                app.quit()
            except Exception:
                pass

def main():
    print("=" * 64)
    print("Race Dashboard V3 - 5 KM Sync (ONE-PC TEST)")
    print(f"Source : {SOURCE_FILE}")
    print(f"Target : {TARGET_FILE}")
    print("รอการ Save จาก race_5KM.xlsx ...")
    print("หยุดโปรแกรมด้วย Ctrl+C")
    print("=" * 64)

    last_mtime = None

    while True:
        try:
            if not SOURCE_FILE.exists():
                print(f"[WAIT] ยังไม่พบ {SOURCE_FILE.name}")
                time.sleep(CHECK_INTERVAL)
                continue

            current_mtime = SOURCE_FILE.stat().st_mtime

            if last_mtime is None or current_mtime != last_mtime:
                time.sleep(0.8)
                print("\n[CHANGE] ตรวจพบการ Save ของ race_5KM.xlsx")
                sync_workbook()
                last_mtime = SOURCE_FILE.stat().st_mtime

            time.sleep(CHECK_INTERVAL)

        except KeyboardInterrupt:
            print("\n[STOP] หยุดระบบ Sync แล้ว")
            break

        except Exception as e:
            print(f"[ERROR] {type(e).__name__}: {e}")
            print("จะลองใหม่อีกครั้งใน 2 วินาที...")
            time.sleep(2)

if __name__ == "__main__":
    main()
