import os
import sys

# =========================================================
# Windows CMD / Terminal UTF-8
# ทำให้ข้อความภาษาไทยใน Console แสดงผลได้ถูกต้อง
# =========================================================
if os.name == "nt":
    try:
        os.system("chcp 65001 > nul")
    except Exception:
        pass

try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass

import json
import time
from datetime import date, datetime, time as dt_time
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen

import xlwings as xw


# =========================================================
# Race Dashboard V3 - Notebook 2 / 5 KM Sender
#
# ONE-PC TEST:
#   SERVER_URL = http://127.0.0.1:8765/sync5km
#
# When Notebook 2 is ready, change only 127.0.0.1 to the LAN IP
# of Notebook 1, e.g. http://192.168.1.10:8765/sync5km
# =========================================================

BASE_DIR = Path(__file__).resolve().parent
SOURCE_FILE = BASE_DIR / "data" / "race_5KM.xlsx"

SERVER_URL = "http://192.168.10.67:8765/sync5km"
CHECK_INTERVAL = 1.0
REQUEST_TIMEOUT = 15

SYNC_SHEETS = [
    "M10_5KM",
    "F10_5KM",
    "M11_5KM",
    "F11_5KM",
    "M16_5KM",
    "F16_5KM",
    "M40_5KM",
    "F40_5KM",
]


def get_last_data_row(sheet):
    last_rows = []
    for col in ("A", "B", "C", "D"):
        row = sheet.range(f"{col}{sheet.cells.last_cell.row}").end("up").row
        last_rows.append(row)
    return max(last_rows)


def encode_value(value):
    """Make xlwings values safe for JSON without losing date/time types."""
    if isinstance(value, datetime):
        return {"__type__": "datetime", "value": value.isoformat()}
    if isinstance(value, date):
        return {"__type__": "date", "value": value.isoformat()}
    if isinstance(value, dt_time):
        return {"__type__": "time", "value": value.isoformat()}
    if value is None or isinstance(value, (str, int, float, bool)):
        return value
    return str(value)


def encode_matrix(rows):
    return [[encode_value(v) for v in row] for row in rows]


def read_5km_data():
    if not SOURCE_FILE.exists():
        raise FileNotFoundError(f"ไม่พบไฟล์ต้นทาง: {SOURCE_FILE}")

    app = None
    wb = None

    try:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = False

        wb = app.books.open(
            str(SOURCE_FILE),
            update_links=False,
            read_only=True,
            ignore_read_only_recommended=True,
        )

        sheet_names = [s.name for s in wb.sheets]
        payload_sheets = {}

        for sheet_name in SYNC_SHEETS:
            if sheet_name not in sheet_names:
                print(f"[WARN] ไม่พบ Sheet: {sheet_name}")
                continue

            ws = wb.sheets[sheet_name]
            last_row = get_last_data_row(ws)

            rows = []
            if last_row >= 2:
                values = ws.range(f"A2:D{last_row}").value

                if last_row == 2 and isinstance(values, list):
                    values = [values]

                # Ensure every record is 4 columns.
                rows = []
                for row in values:
                    row = list(row) if isinstance(row, (list, tuple)) else [row]
                    row = (row + [None, None, None, None])[:4]
                    rows.append(row)

            payload_sheets[sheet_name] = encode_matrix(rows)
            print(f"[READ] {sheet_name} rows={len(rows)}")

        return {"sheets": payload_sheets}

    finally:
        if wb is not None:
            try:
                wb.close()
            except Exception:
                pass
        if app is not None:
            try:
                app.quit()
            except Exception:
                pass


def send_to_notebook1():
    payload = read_5km_data()
    body = json.dumps(payload, ensure_ascii=False).encode("utf-8")

    request = Request(
        SERVER_URL,
        data=body,
        headers={"Content-Type": "application/json; charset=utf-8"},
        method="POST",
    )

    try:
        with urlopen(request, timeout=REQUEST_TIMEOUT) as response:
            result = response.read().decode("utf-8")
            if response.status == 200:
                print(f"[OK] ส่งข้อมูล 5 KM สำเร็จ -> {SERVER_URL}")
            else:
                print(f"[ERROR] Server status={response.status}: {result}")

    except HTTPError as e:
        detail = e.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"Server ตอบ HTTP {e.code}: {detail}") from e

    except URLError as e:
        raise ConnectionError(
            f"เชื่อมต่อ Notebook 1 ไม่ได้: {SERVER_URL} ({e.reason})"
        ) from e


def main():
    print("=" * 68)
    print("Race Dashboard V3 - NOTEBOOK 2 / 5 KM SENDER")
    print(f"Source : {SOURCE_FILE}")
    print(f"Server : {SERVER_URL}")
    print("รอการ Save จาก race_5KM.xlsx ...")
    print("หยุดระบบด้วย Ctrl+C")
    print("=" * 68)

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
                send_to_notebook1()
                last_mtime = SOURCE_FILE.stat().st_mtime

            time.sleep(CHECK_INTERVAL)

        except KeyboardInterrupt:
            print("\n[STOP] หยุด 5 KM Sender แล้ว")
            break

        except Exception as e:
            print(f"[ERROR] {type(e).__name__}: {e}")
            print("จะลองใหม่เมื่อมีการ Save ครั้งถัดไป / ตรวจใหม่ใน 2 วินาที...")
            time.sleep(2)


if __name__ == "__main__":
    main()
