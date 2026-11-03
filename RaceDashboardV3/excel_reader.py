import time
from pathlib import Path

import xlwings as xw


# ==========================================
# CONFIGURATION
# ==========================================

BASE_DIR = Path(__file__).resolve().parent
EXCEL_PATH = BASE_DIR / "data" / "race_results.xlsx"

REQUIRED_HEADERS = [
    "Rank",
    "Bib",
    "Name",
    "Time"
]


# ==========================================
# อ่านค่าอย่างปลอดภัย
# ==========================================

def safe_value(row, index, default=""):

    if index >= len(row):
        return default

    value = row[index]

    if value is None:
        return default

    return value


# ==========================================
# จัดรูปแบบ BIB
# ==========================================

def normalize_bib(value):
    """
    แปลง BIB ให้อยู่ในรูปข้อความ 3 หลัก

    ตัวอย่าง:
    5    -> "005"
    19   -> "019"
    125  -> "125"
    """

    if value is None or value == "":
        return ""

    # ถ้า Excel ส่งมาเป็นตัวเลข
    if isinstance(value, (int, float)):

        if float(value).is_integer():

            return str(
                int(value)
            ).zfill(3)

    # ถ้าเป็นข้อความอยู่แล้ว
    text = str(value).strip()

    # รองรับกรณีข้อความเป็นตัวเลข เช่น "19"
    if text.isdigit():

        return text.zfill(3)

    return text


# ==========================================
# จัดการข้อความ Cell ทั่วไป
# ==========================================

def clean_cell_text(value):
    """
    แปลงค่าจาก Excel cell เป็นข้อความที่ตัดช่องว่างแล้ว
    รองรับกรณีตัวเลข เช่น 19.0 -> "19"
    """
    if value is None:
        return ""

    if isinstance(value, float) and value.is_integer():
        return str(int(value)).strip()

    if isinstance(value, int):
        return str(value).strip()

    return str(value).strip()


# ==========================================
# อ่านข้อมูลการตั้งค่าหัวข้อจากชีท CONFIG
# ==========================================

def read_config_sheet(wb):
    """
    อ่านข้อมูลหัวข้อการแข่งขันจากชีท CONFIG
    ดึงคอลัมน์ SheetName, Type, DisplayName, Order
    """
    config_list = []

    # ค้นหาชีท CONFIG (case-insensitive)
    config_sheet = None
    for sheet in wb.sheets:
        if sheet.name.strip().upper() == "CONFIG":
            config_sheet = sheet
            break

    if not config_sheet:
        return config_list

    try:
        raw_data = config_sheet.range("A1").expand().value
    except Exception:
        return config_list

    if not raw_data or not isinstance(raw_data, list):
        return config_list

    if not isinstance(raw_data[0], list):
        return config_list

    first_row = raw_data[0]
    headers = [
        str(header).strip() if header is not None else ""
        for header in first_row
    ]

    header_map = {}
    for idx, header in enumerate(headers):
        clean_key = header.lower().replace(" ", "").replace("_", "")
        if clean_key in ("sheetname", "sheet"):
            header_map["sheet_name"] = idx
        elif clean_key == "type":
            header_map["type"] = idx
        elif clean_key in ("displayname", "display"):
            header_map["display_name"] = idx
        elif clean_key == "name" and "display_name" not in header_map:
            header_map["display_name"] = idx
        elif clean_key == "order":
            header_map["order"] = idx
        elif clean_key in ("layoutmode", "layout", "mode"):
            header_map["layout_mode"] = idx

    for row in raw_data[1:]:
        if not row:
            continue

        if not isinstance(row, list):
            row = [row]

        sheet_name_val = safe_value(row, header_map.get("sheet_name"), "")
        sheet_key = clean_cell_text(sheet_name_val)

        type_val = ""
        if "type" in header_map:
            type_val = clean_cell_text(safe_value(row, header_map["type"], ""))

        display_name_val = ""
        if "display_name" in header_map:
            display_name_val = clean_cell_text(safe_value(row, header_map["display_name"], ""))

        order_val = None
        if "order" in header_map:
            raw_order = safe_value(row, header_map["order"], None)
            if raw_order is not None:
                try:
                    order_val = int(float(raw_order))
                except (ValueError, TypeError):
                    order_val = None

        layout_mode_val = "ranking"
        if "layout_mode" in header_map:
            val = clean_cell_text(safe_value(row, header_map["layout_mode"], "")).strip().lower()
            if val:
                layout_mode_val = val

        if not sheet_key and not display_name_val and not type_val:
            continue

        config_list.append({
            "order": order_val,
            "sheet_name": sheet_key,
            "type": type_val,
            "display_name": display_name_val,
            "layout_mode": layout_mode_val
        })

    return config_list


# ==========================================
# จับคู่ข้อมูลหัวข้อกับชีทการแข่งขัน
# ==========================================

def find_category_config(sheet_name, competition_index, config_list, sample_bib=""):
    """
    จับคู่ข้อมูลหัวข้อการแข่งขันจาก CONFIG อย่างชาญฉลาดและยืดหยุ่น:
    1. ตรงกับ SheetName ใน CONFIG (ตรงเป๊ะ หรือตัดช่องว่าง / เล็กใหญ่)
    2. ตรงกับ DisplayName ใน CONFIG
    3. DisplayName ใน CONFIG ปรากฏอยู่ในชื่อชีท (เช่น 'อายุ 20–29 ปี' ใน '10.5 KM อายุ 20–29 ปี')
    4. SheetName ใน CONFIG เป็นคำย่อยของชื่อชีท หรือชื่อชีทเป็นคำย่อย
    5. รหัส BIB (เช่น M29 จาก M29-2005) ตรงกับ SheetName ใน CONFIG (เช่น M29_10KM)
    6. ลำดับ Order ใน CONFIG ตรงกับลำดับชีทการแข่งขัน (competition_index)
    """
    if not config_list:
        return None

    clean_name = str(sheet_name).strip()
    clean_name_lower = clean_name.lower()

    # 1. ตรงกับ SheetName
    for cfg in config_list:
        s_name = cfg.get("sheet_name", "")
        if s_name and s_name.lower() == clean_name_lower:
            return cfg

    # 2. ตรงกับ DisplayName
    for cfg in config_list:
        d_name = cfg.get("display_name", "")
        if d_name and d_name.lower() == clean_name_lower:
            return cfg

    # 3. DisplayName อยู่ในชื่อชีท (เช่น 'อายุ 20–29 ปี' อยู่ใน '10.5 KM อายุ 20–29 ปี')
    for cfg in config_list:
        d_name = cfg.get("display_name", "").lower()
        if d_name and d_name in clean_name_lower:
            return cfg

    # 4. SheetName เป็นคำย่อย หรือชื่อชีทเป็นคำย่อย
    for cfg in config_list:
        s_name = cfg.get("sheet_name", "").lower()
        if s_name and (s_name in clean_name_lower or clean_name_lower in s_name):
            return cfg

    # 5. เทียบ Prefix จาก Bib (ถ้ามี) เช่น M29 ใน M29-2005 ตรงกับ M29 ใน M29_10KM
    if sample_bib:
        clean_bib = str(sample_bib).strip().upper()
        bib_prefix = clean_bib.split("-")[0] if "-" in clean_bib else clean_bib[:3]
        if bib_prefix:
            for cfg in config_list:
                s_name = cfg.get("sheet_name", "").upper()
                if bib_prefix in s_name:
                    return cfg

    # 6. ลำดับ Order ตรงกับลำดับชีทแข่งขัน
    for cfg in config_list:
        if cfg.get("order") == competition_index:
            return cfg

    return None


# ==========================================
# อ่านข้อมูลการแข่งขันทั้งหมด
# ==========================================

def _read_race_data_from_workbook(wb):
    race_data = []

    warnings = []


    # ======================================
    # อ่านการตั้งค่าจากชีท CONFIG (ถ้ามี)
    # ======================================

    config_list = read_config_sheet(wb)

    competition_index = 0


    # ======================================
    # อ่านทุก Sheet ตามลำดับ Tab ใน Excel
    # ======================================

    for sheet_index, sheet in enumerate(
        wb.sheets,
        start=1
    ):

        sheet_name = sheet.name

        # ข้ามชีท CONFIG ไม่นำมาอ่านเป็นหมวดแข่งขัน
        if sheet_name.strip().upper() == "CONFIG":
            continue


        # ==================================
        # อ่านข้อมูลใน Sheet
        # ==================================

        sheet_data = (
            sheet
            .range("A1")
            .expand()
            .value
        )


        # Sheet ว่าง
        if not sheet_data:
            continue


        # ==================================
        # ตรวจ Header
        # ==================================

        first_row = sheet_data[0]


        # กรณี Sheet มีแค่ Cell เดียว
        if not isinstance(first_row, list):
            continue


        headers = [
            str(header).strip()
            if header is not None
            else ""
            for header in first_row
        ]


        missing_headers = [
            header
            for header in REQUIRED_HEADERS
            if header not in headers
        ]


        # ==================================
        # ถ้า Header ไม่ครบ
        # ไม่ถือว่าเป็น Competition Sheet
        # ==================================

        if missing_headers:

            warnings.append(
                f"{sheet_name}: Ignored - Missing headers: "
                + ", ".join(missing_headers)
            )

            continue


        # ==================================
        # หา Index ของ Column
        # ==================================

        header_map = {
            header: headers.index(header)
            for header in REQUIRED_HEADERS
        }


        rows = sheet_data[1:]

        results = []


        # ==================================
        # อ่านแต่ละ Result
        # ==================================

        for excel_row_number, row in enumerate(
            rows,
            start=2
        ):

            if not row:
                continue


            rank_value = safe_value(
                row,
                header_map["Rank"],
                None
            )


            # ไม่มี Rank = ข้ามแถว
            if rank_value in (
                None,
                ""
            ):
                continue


            # ==================================
            # ตรวจ Rank
            # ==================================

            try:

                rank = int(
                    rank_value
                )

            except (TypeError, ValueError):

                warnings.append(
                    f"{sheet_name} row "
                    f"{excel_row_number}: "
                    f"Invalid Rank"
                )

                continue


            if rank <= 0:

                warnings.append(
                    f"{sheet_name} row "
                    f"{excel_row_number}: "
                    f"Invalid Rank: {rank}"
                )

                continue


            # ==================================
            # อ่าน Column อื่น
            # ==================================

            bib = safe_value(
                row,
                header_map["Bib"]
            )

            name = safe_value(
                row,
                header_map["Name"]
            )

            time = safe_value(
                row,
                header_map["Time"]
            )


            # ==================================
            # เพิ่ม Result
            # ==================================

            results.append({
                "rank": rank,

                # แปลง BIB เช่น 19 -> "019"
                "bib": normalize_bib(bib),

                "name": str(name).strip(),

                "time": time
            })


        # ==================================
        # ตรวจ Rank ซ้ำ
        # ==================================

        seen_ranks = set()

        duplicate_ranks = set()


        for result in results:

            rank = result["rank"]

            if rank in seen_ranks:

                duplicate_ranks.add(
                    rank
                )

            seen_ranks.add(
                rank
            )


        if duplicate_ranks:

            duplicate_text = ", ".join(
                str(rank)
                for rank
                in sorted(
                    duplicate_ranks
                )
            )

            warnings.append(
                f"{sheet_name}: "
                f"Duplicate Rank: "
                f"{duplicate_text}"
            )


        # ==================================
        # เรียงตาม Rank
        # ==================================

        results.sort(
            key=lambda item: item["rank"]
        )


        competition_index += 1

        sample_bib = results[0]["bib"] if results else ""

        # ==================================
        # กำหนดหัวข้อจาก CONFIG
        # ==================================

        cfg_entry = find_category_config(
            sheet_name=sheet_name,
            competition_index=competition_index,
            config_list=config_list,
            sample_bib=sample_bib
        )

        # คงพฤติกรรมเดิมไว้:
        # บรรทัด 1 (type) ว่างไว้
        # บรรทัด 2 (display_name) ใช้ชื่อ SheetName
        category_type = ""
        display_name = sheet_name

        if cfg_entry:
            category_type = cfg_entry.get("type", "")
            cfg_display = cfg_entry.get("display_name", "")
            if cfg_display:
                display_name = cfg_display


        # ค่า layout_mode จาก CONFIG (เช่น "ranking" หรือ "overall")
        layout_mode = "ranking"
        if cfg_entry and cfg_entry.get("layout_mode"):
            layout_mode = cfg_entry.get("layout_mode")


        # ==================================
        # เพิ่ม Category
        # ==================================

        race_data.append({
            "order": sheet_index,

            # ชื่อ Sheet จริง (สำหรับ tracking / matching)
            "sheet_name": sheet_name,

            # บรรทัด 1 = Type
            "type": category_type,

            # บรรทัด 2 = DisplayName (ถ้าไม่มีใน CONFIG ให้คงชื่อเดิมไว้)
            "display_name": display_name,

            # โหมด Layout จาก CONFIG ("ranking" หรือ "overall")
            "layout_mode": layout_mode,

            "status": "ready",

            # มีเท่าไรส่งเท่านั้น
            "results": results
        })


    # ======================================
    # RESPONSE
    # ======================================

    return {
        "success": True,

        "total_categories": len(
            race_data
        ),

        "categories": race_data,

        "warnings": warnings
    }


def read_race_data():
    """
    อ่าน race_results.xlsx แบบเบื้องหลัง

    - ไม่เปิดหน้าต่าง Excel ให้ User เห็น
    - เปิดไฟล์แบบ Read Only
    - ปิด Workbook และ Excel instance ทุกครั้งหลังอ่านเสร็จ
    - Path อ้างอิงจากโฟลเดอร์โปรเจกต์ปัจจุบัน จึงใช้กับ V3 ได้
    """

    if not EXCEL_PATH.exists():
        return {
            "success": False,
            "total_categories": 0,
            "categories": [],
            "warnings": [
                f"ไม่พบไฟล์ race_results.xlsx: {EXCEL_PATH}"
            ]
        }

    last_error = None

    # เผื่อ sync_hub.py กำลัง Save ไฟล์อยู่พอดี
    for attempt in range(1, 4):

        app = None
        wb = None

        try:
            app = xw.App(
                visible=False,
                add_book=False
            )

            app.display_alerts = False
            app.screen_updating = False

            wb = app.books.open(
                str(EXCEL_PATH),
                update_links=False,
                read_only=True,
                ignore_read_only_recommended=True
            )

            return _read_race_data_from_workbook(wb)

        except Exception as exc:
            last_error = exc

            if attempt < 3:
                time.sleep(0.25)

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

    return {
        "success": False,
        "total_categories": 0,
        "categories": [],
        "warnings": [
            "อ่าน race_results.xlsx ไม่สำเร็จ: "
            + str(last_error)
        ]
    }
