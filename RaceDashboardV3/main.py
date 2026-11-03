from fastapi import FastAPI
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from contextlib import asynccontextmanager

from excel_reader import read_race_data

from datetime import datetime
from pathlib import Path

import copy
import hashlib
import json
import threading
import pythoncom


# ==========================================
# PATH
# ==========================================

BASE_DIR = Path(__file__).resolve().parent

# หน้า Dashboard ตัวจริง
INDEX_PATH = BASE_DIR / "index.html"

# หน้าสำหรับทดสอบ Realtime
REALTIME_TEST_PATH = BASE_DIR / "realtime.html"

# Folder หลักของรูปภาพ
IMAGES_DIR = BASE_DIR / "images"

# Folder ภาพ Fix ของระบบ (พื้นหลัง, Default, Logo)
IMAGES_SYSTEM_DIR = IMAGES_DIR / "system"

# Folder ภาพนักวิ่งที่ User อัปโหลด
IMAGES_RUNNERS_DIR = IMAGES_DIR / "runners"


# ==========================================
# CONFIG
# ==========================================

# Python ตรวจข้อมูล Excel ทุก 1 วินาที
EXCEL_POLL_SECONDS = 1


# ==========================================
# CACHE
# ==========================================

# ป้องกันการอ่าน/เขียน Cache พร้อมกัน
cache_lock = threading.Lock()

# ใช้สำหรับหยุด Background Worker
stop_event = threading.Event()


# ข้อมูลล่าสุดที่เก็บไว้ใน Memory
current_data = {
    "success": False,
    "message": "Waiting for Excel data...",
    "version": 0,
    "updated_at": None,
    "categories": [],
    "warnings": []
}


# Hash ของข้อมูลล่าสุด
last_data_hash = None

# Version ของข้อมูล
data_version = 0


# สถานะ Excel
excel_status = "starting"

# Error ล่าสุด
last_error = None


# ==========================================
# อ่าน Excel และ Update Cache
# ==========================================

def refresh_excel_cache():

    global current_data
    global last_data_hash
    global data_version
    global excel_status
    global last_error

    try:

        # ==================================
        # อ่านข้อมูลจาก Excel
        # ==================================

        new_data = read_race_data()


        # ==================================
        # แปลงข้อมูลเป็น String
        # เพื่อใช้สร้าง Hash
        # ==================================

        data_string = json.dumps(
            new_data,
            sort_keys=True,
            ensure_ascii=False,
            default=str
        )


        # ==================================
        # สร้าง Hash
        # ==================================

        current_hash = hashlib.sha256(
            data_string.encode("utf-8")
        ).hexdigest()


        # ==================================
        # Lock ก่อน Update Cache
        # ==================================

        with cache_lock:

            # ------------------------------
            # ข้อมูลมีการเปลี่ยนแปลง
            # ------------------------------

            if current_hash != last_data_hash:

                data_version += 1

                last_data_hash = current_hash

                updated_at = datetime.now().isoformat(
                    timespec="seconds"
                )


            # ------------------------------
            # ข้อมูลเหมือนเดิม
            # ------------------------------

            else:

                updated_at = current_data.get(
                    "updated_at"
                )


            # ==================================
            # เพิ่ม Metadata
            # ==================================

            new_data["version"] = data_version

            new_data["updated_at"] = updated_at


            # ==================================
            # Update Memory Cache
            # ==================================

            current_data = new_data


            # ==================================
            # Update Status
            # ==================================

            excel_status = "connected"

            last_error = None


    except Exception as error:

        # ==================================
        # ถ้า Excel Busy หรือมี Error
        #
        # จะไม่ลบ Cache เดิม
        # Dashboard ยังใช้ข้อมูลล่าสุดได้
        # ==================================

        with cache_lock:

            excel_status = "error"

            last_error = str(error)


# ==========================================
# BACKGROUND WORKER
# ==========================================

def excel_worker():

    # xlwings บน Windows ใช้ COM
    # Thread ใหม่จึงต้อง Initialize COM
    pythoncom.CoInitialize()

    try:

        while not stop_event.is_set():

            # อ่าน Excel และ Update Cache
            refresh_excel_cache()


            # รอ 1 วินาที
            stop_event.wait(
                EXCEL_POLL_SECONDS
            )

    finally:

        pythoncom.CoUninitialize()


# ==========================================
# FASTAPI STARTUP / SHUTDOWN
# ==========================================

@asynccontextmanager
async def lifespan(app: FastAPI):

    # ======================================
    # START
    # ======================================

    stop_event.clear()


    # สร้าง Background Thread
    worker = threading.Thread(
        target=excel_worker,
        daemon=True
    )


    # เริ่มตรวจ Excel
    worker.start()


    # FastAPI เริ่มทำงาน
    yield


    # ======================================
    # STOP
    # ======================================

    stop_event.set()


    # รอ Worker ปิด
    worker.join(
        timeout=2
    )


# ==========================================
# FASTAPI
# ==========================================

app = FastAPI(
    title="Race Dashboard API",
    version="1.0.0",
    lifespan=lifespan
)


@app.middleware("http")
async def add_no_cache_for_images(request, call_next):
    response = await call_next(request)
    if request.url.path.startswith("/images/runners/"):
        # ภาพนักวิ่ง: ไม่ Cache เพราะ User อัปโหลดใหม่ตลอด
        response.headers["Cache-Control"] = "no-cache, no-store, must-revalidate"
        response.headers["Pragma"] = "no-cache"
        response.headers["Expires"] = "0"
    return response


# ==========================================
# STATIC IMAGES
# ==========================================

# ภาพ Fix ของระบบ เช่น
# http://127.0.0.1:8000/images/system/background-portrait.jpg
# http://127.0.0.1:8000/images/system/default.jpg
#
app.mount(
    "/images/system",
    StaticFiles(
        directory=str(IMAGES_SYSTEM_DIR)
    ),
    name="images_system"
)


# ภาพนักวิ่งที่ User อัปโหลด เช่น
# http://127.0.0.1:8000/images/runners/M20-2005.png
#
app.mount(
    "/images/runners",
    StaticFiles(
        directory=str(IMAGES_RUNNERS_DIR)
    ),
    name="images_runners"
)


# ==========================================
# HOME
# ==========================================

@app.get("/")
def root():

    return {
        "message": "Race Dashboard Backend",
        "dashboard": "/dashboard",
        "realtime_test": "/realtime",
        "api_results": "/api/results",
        "api_status": "/api/status"
    }


# ==========================================
# DASHBOARD
# ==========================================

@app.get("/dashboard")
def dashboard():

    if not INDEX_PATH.exists():

        return {
            "error": "index.html not found",
            "expected_path": str(INDEX_PATH)
        }

    return FileResponse(
        str(INDEX_PATH),
        media_type="text/html"
    )

# ==========================================
# REALTIME TEST PAGE
# ==========================================

@app.get("/realtime")
def realtime_test():

    return FileResponse(
        REALTIME_TEST_PATH,
        media_type="text/html"
    )


# ==========================================
# RESULTS API
# ==========================================

@app.get("/api/results")
def get_results():

    # ======================================
    # API อ่านข้อมูลจาก Memory Cache
    #
    # ไม่เข้าไปอ่าน Excel โดยตรง
    # ทำให้ตอบข้อมูลได้เร็ว
    # ======================================

    with cache_lock:

        return copy.deepcopy(
            current_data
        )


# ==========================================
# STATUS API
# ==========================================

@app.get("/api/status")
def get_status():

    with cache_lock:

        return {

            "backend": "online",

            "excel": excel_status,

            "version": current_data.get(
                "version",
                0
            ),

            "updated_at": current_data.get(
                "updated_at"
            ),

            "last_error": last_error
        }