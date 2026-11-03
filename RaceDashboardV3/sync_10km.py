import time
from pathlib import Path

import xlwings as xw


# =========================================================
# Race Dashboard V3 - Notebook 1 Sync
# Source: race_10_5KM.xlsx
# Target: race_results.xlsx
# =========================================================

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"

SOURCE_FILE = DATA_DIR / "race_10_5KM.xlsx"
TARGET_FILE = DATA_DIR / "race_results.xlsx"

CHECK_INTERVAL = 1.0

# Notebook 1 owns ONLY these sheets.
SYNC_SHEETS = [
    "M_Overall",
    "F_Overall",
    "M19_10.5KM",
    "F19_10.5KM",
    "M20_10.5KM",
    "F20_10.5KM",
    "M30_10.5KM",
    "F30_10.5KM",
    "M40_10.5KM",
    "F40_10.5KM",
    "M50_10.5KM",
    "F50_10.5KM",
    "M60_10.5KM",
    "F60_10.5KM",
]


def get_last_data_row(sheet):
    """Find last used row from columns A:D. Header is row 1."""
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
        # Separate hidden Excel instance.
        # Source is opened read-only, so the user can work in the normal Excel window.
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

            # Clear old race data only.
            # Row 1 (Rank/Bib/Name/Time) remains untouched.
            dst_ws.range("A2:D10000").clear_contents()

            # Copy current data from source into the matching target sheet.
            if last_row >= 2:
                values = src_ws.range(f"A2:D{last_row}").value

                # xlwings returns a flat list for one row; make it 2D.
                if last_row == 2 and isinstance(values, list):
                    values = [values]

                dst_ws.range("A2").value = values

            print(f"[SYNC] {sheet_name}  rows={max(last_row - 1, 0)}")

        dst_wb.save()
        print("[OK] race_results.xlsx อัปเดตเรียบร้อย")

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
    print("=" * 60)
    print("Race Dashboard V3 - 10.5 KM Sync")
    print(f"Source : {SOURCE_FILE}")
    print(f"Target : {TARGET_FILE}")
    print("รอการ Save จาก race_10_5KM.xlsx ...")
    print("หยุดโปรแกรมด้วย Ctrl+C")
    print("=" * 60)

    last_mtime = None

    while True:
        try:
            if not SOURCE_FILE.exists():
                print(f"[WAIT] ยังไม่พบ {SOURCE_FILE.name}")
                time.sleep(CHECK_INTERVAL)
                continue

            current_mtime = SOURCE_FILE.stat().st_mtime

            # Sync once at startup, then again whenever the source file is saved.
            if last_mtime is None or current_mtime != last_mtime:
                # Let Excel finish writing the file before we open it.
                time.sleep(0.8)

                print("\n[CHANGE] ตรวจพบการ Save ของ race_10_5KM.xlsx")
                sync_workbook()

                # Re-read mtime after sync in case the file timestamp changed while saving.
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
