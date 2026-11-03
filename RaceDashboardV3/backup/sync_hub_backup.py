import json
import threading
import time
from datetime import date, datetime, time as dt_time
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

import xlwings as xw


# =========================================================
# Race Dashboard V3 - CENTRAL SYNC HUB (Notebook 1)
#
# Responsibilities:
# 1) Watch local race_10_5KM.xlsx (Overall + 10.5 KM)
# 2) Receive 5 KM data from Notebook 2 over LAN/HTTP
# 3) Be the ONLY process that writes race_results.xlsx
# =========================================================

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"

LOCAL_SOURCE_FILE = DATA_DIR / "race_10_5KM.xlsx"
TARGET_FILE = DATA_DIR / "race_results.xlsx"

HOST = "0.0.0.0"
PORT = 8765
CHECK_INTERVAL = 1.0

LOCAL_SHEETS = [
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

REMOTE_5KM_SHEETS = [
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

# One lock = one writer at a time, even if both notebooks save together.
WRITE_LOCK = threading.Lock()


def get_last_data_row(sheet):
    last_rows = []
    for col in ("A", "B", "C", "D"):
        row = sheet.range(f"{col}{sheet.cells.last_cell.row}").end("up").row
        last_rows.append(row)
    return max(last_rows)


def decode_value(value):
    """Restore date/time values encoded by sender_5km_lan.py."""
    if isinstance(value, dict) and "__type__" in value:
        kind = value.get("__type__")
        raw = value.get("value")
        if kind == "datetime":
            return datetime.fromisoformat(raw)
        if kind == "date":
            return date.fromisoformat(raw)
        if kind == "time":
            return dt_time.fromisoformat(raw)
    return value


def decode_matrix(rows):
    return [[decode_value(v) for v in row] for row in rows]


def sync_local_10km():
    """Copy Notebook 1 source sheets into the central workbook."""
    if not LOCAL_SOURCE_FILE.exists():
        raise FileNotFoundError(f"ไม่พบไฟล์ต้นทาง: {LOCAL_SOURCE_FILE}")
    if not TARGET_FILE.exists():
        raise FileNotFoundError(f"ไม่พบไฟล์กลาง: {TARGET_FILE}")

    with WRITE_LOCK:
        app = None
        src_wb = None
        dst_wb = None

        try:
            app = xw.App(visible=False, add_book=False)
            app.display_alerts = False
            app.screen_updating = False

            src_wb = app.books.open(
                str(LOCAL_SOURCE_FILE),
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

            src_names = [s.name for s in src_wb.sheets]
            dst_names = [s.name for s in dst_wb.sheets]

            for sheet_name in LOCAL_SHEETS:
                if sheet_name not in src_names:
                    print(f"[WARN-10KM] ไม่พบ Sheet ต้นทาง: {sheet_name}")
                    continue
                if sheet_name not in dst_names:
                    print(f"[WARN-10KM] ไม่พบ Sheet ปลายทาง: {sheet_name}")
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

                print(f"[10KM] {sheet_name} rows={max(last_row - 1, 0)}")

            dst_wb.save()
            print("[OK-10KM] race_results.xlsx อัปเดต Overall + 10.5 KM แล้ว")

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


def write_remote_5km(sheets_data):
    """Write received 5 KM sheets into the central workbook."""
    if not TARGET_FILE.exists():
        raise FileNotFoundError(f"ไม่พบไฟล์กลาง: {TARGET_FILE}")

    invalid = [name for name in sheets_data if name not in REMOTE_5KM_SHEETS]
    if invalid:
        raise ValueError(f"พบ Sheet ที่ไม่ได้รับอนุญาต: {invalid}")

    with WRITE_LOCK:
        app = None
        dst_wb = None

        try:
            app = xw.App(visible=False, add_book=False)
            app.display_alerts = False
            app.screen_updating = False

            dst_wb = app.books.open(
                str(TARGET_FILE),
                update_links=False,
                read_only=False,
                ignore_read_only_recommended=True,
            )

            dst_names = [s.name for s in dst_wb.sheets]

            for sheet_name in REMOTE_5KM_SHEETS:
                if sheet_name not in sheets_data:
                    continue

                if sheet_name not in dst_names:
                    print(f"[WARN-5KM] ไม่พบ Sheet ปลายทาง: {sheet_name}")
                    continue

                rows = sheets_data[sheet_name]
                if rows is None:
                    rows = []
                rows = decode_matrix(rows)

                dst_ws = dst_wb.sheets[sheet_name]
                dst_ws.range("A2:D10000").clear_contents()

                if rows:
                    dst_ws.range("A2").value = rows

                print(f"[5KM] {sheet_name} rows={len(rows)}")

            dst_wb.save()
            print("[OK-5KM] race_results.xlsx อัปเดต 5 KM แล้ว")

        finally:
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


class SyncHandler(BaseHTTPRequestHandler):
    def do_POST(self):
        if self.path != "/sync5km":
            self.send_response(404)
            self.end_headers()
            return

        try:
            content_length = int(self.headers.get("Content-Length", "0"))
            raw = self.rfile.read(content_length)
            payload = json.loads(raw.decode("utf-8"))

            sheets_data = payload.get("sheets")
            if not isinstance(sheets_data, dict):
                raise ValueError("payload ต้องมี key 'sheets' เป็น object")

            write_remote_5km(sheets_data)

            response = json.dumps(
                {"ok": True, "message": "5 KM synced"},
                ensure_ascii=False,
            ).encode("utf-8")

            self.send_response(200)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Content-Length", str(len(response)))
            self.end_headers()
            self.wfile.write(response)

        except Exception as e:
            print(f"[ERROR-5KM] {type(e).__name__}: {e}")

            response = json.dumps(
                {"ok": False, "error": str(e)},
                ensure_ascii=False,
            ).encode("utf-8")

            self.send_response(500)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Content-Length", str(len(response)))
            self.end_headers()
            self.wfile.write(response)

    def log_message(self, format, *args):
        # Hide standard HTTP access logs; keep console easy to read.
        return


def start_server():
    server = ThreadingHTTPServer((HOST, PORT), SyncHandler)
    print(f"[SERVER] รอรับข้อมูล 5 KM ที่ port {PORT}")
    server.serve_forever()


def main():
    print("=" * 68)
    print("Race Dashboard V3 - CENTRAL SYNC HUB")
    print(f"10.5 Source : {LOCAL_SOURCE_FILE}")
    print(f"Central     : {TARGET_FILE}")
    print(f"5 KM Server : http://127.0.0.1:{PORT}/sync5km (ทดสอบเครื่องเดียว)")
    print("หยุดระบบด้วย Ctrl+C")
    print("=" * 68)

    server_thread = threading.Thread(target=start_server, daemon=True)
    server_thread.start()

    last_mtime = None

    while True:
        try:
            if not LOCAL_SOURCE_FILE.exists():
                print(f"[WAIT-10KM] ยังไม่พบ {LOCAL_SOURCE_FILE.name}")
                time.sleep(CHECK_INTERVAL)
                continue

            current_mtime = LOCAL_SOURCE_FILE.stat().st_mtime

            if last_mtime is None or current_mtime != last_mtime:
                time.sleep(0.8)
                print("\n[CHANGE-10KM] ตรวจพบการ Save ของ race_10_5KM.xlsx")
                sync_local_10km()
                last_mtime = LOCAL_SOURCE_FILE.stat().st_mtime

            time.sleep(CHECK_INTERVAL)

        except KeyboardInterrupt:
            print("\n[STOP] หยุด Central Sync Hub แล้ว")
            break

        except Exception as e:
            print(f"[ERROR-10KM] {type(e).__name__}: {e}")
            print("จะลองใหม่อีกครั้งใน 2 วินาที...")
            time.sleep(2)


if __name__ == "__main__":
    main()
