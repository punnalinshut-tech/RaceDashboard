@echo off
setlocal EnableExtensions
chcp 65001 >nul
title Race Dashboard - Setup

echo ============================================================
echo   RACE DASHBOARD - ONE CLICK SETUP
echo ============================================================
echo.
echo ไฟล์นี้ใช้ได้ทั้ง Notebook 1 และ Notebook 2
echo จะตรวจ Python และติดตั้ง Package ที่ระบบ Race Dashboard ใช้
echo.

REM ------------------------------------------------------------
REM 1) Find Python
REM ------------------------------------------------------------
set "PYEXE="

where py >nul 2>&1
if %errorlevel%==0 (
    set "PYEXE=py -3"
    goto :PY_FOUND
)

where python >nul 2>&1
if %errorlevel%==0 (
    set "PYEXE=python"
    goto :PY_FOUND
)

echo [INFO] ยังไม่พบ Python ในเครื่อง
echo [INFO] กำลังลองติดตั้ง Python 3.12 ผ่าน Windows Package Manager...
echo.

where winget >nul 2>&1
if not %errorlevel%==0 (
    echo [ERROR] ไม่พบทั้ง Python และ winget
    echo.
    echo กรุณาติดตั้ง Python 3 ก่อน แล้วเปิดไฟล์นี้อีกครั้ง
    echo ตอนติดตั้ง Python ให้เลือก Add Python to PATH
    echo.
    pause
    exit /b 1
)

winget install --id Python.Python.3.12 -e --scope user --silent --accept-package-agreements --accept-source-agreements

if exist "%LocalAppData%\Programs\Python\Python312\python.exe" (
    set "PYEXE=%LocalAppData%\Programs\Python\Python312\python.exe"
    goto :PY_FOUND
)

where py >nul 2>&1
if %errorlevel%==0 (
    set "PYEXE=py -3"
    goto :PY_FOUND
)

where python >nul 2>&1
if %errorlevel%==0 (
    set "PYEXE=python"
    goto :PY_FOUND
)

echo.
echo [ERROR] ติดตั้ง Python แล้ว แต่ CMD นี้ยังไม่พบ Python
echo ให้ปิดหน้าต่างนี้ แล้วดับเบิลคลิก SETUP_RACE_SYSTEM.bat ใหม่อีกครั้ง
echo.
pause
exit /b 1


:PY_FOUND
echo [OK] พบ Python
%PYEXE% --version
echo.

REM ------------------------------------------------------------
REM 2) Prepare pip
REM ------------------------------------------------------------
echo [1/4] เตรียม pip...
%PYEXE% -m ensurepip --upgrade >nul 2>&1
%PYEXE% -m pip install --upgrade pip setuptools wheel
if errorlevel 1 goto :INSTALL_ERROR

echo.

REM ------------------------------------------------------------
REM 3) Install Race Dashboard dependencies
REM ------------------------------------------------------------
echo [2/4] ติดตั้ง FastAPI / Uvicorn...
%PYEXE% -m pip install fastapi "uvicorn[standard]"
if errorlevel 1 goto :INSTALL_ERROR

echo.
echo [3/4] ติดตั้ง Excel libraries...
%PYEXE% -m pip install xlwings openpyxl
if errorlevel 1 goto :INSTALL_ERROR

echo.

REM ------------------------------------------------------------
REM 4) Verify
REM ------------------------------------------------------------
echo [4/4] ตรวจสอบ Package...
%PYEXE% -c "import fastapi, uvicorn, xlwings, openpyxl; print('[OK] FastAPI      :', fastapi.__version__); print('[OK] Uvicorn      :', uvicorn.__version__); print('[OK] xlwings      :', xlwings.__version__); print('[OK] openpyxl     :', openpyxl.__version__)"
if errorlevel 1 goto :VERIFY_ERROR

echo.
echo ============================================================
echo   ติดตั้งสำเร็จ
echo ============================================================
echo.
echo เครื่องนี้พร้อมสำหรับ Race Dashboard แล้ว
echo.
echo Package หลัก:
echo   - FastAPI
echo   - Uvicorn
echo   - xlwings
echo   - openpyxl
echo.
echo หมายเหตุ:
echo   - เครื่องที่ทำงานกับ Excel ต้องติดตั้ง Microsoft Excel ด้วย
echo   - การติดตั้ง Package ครั้งแรกต้องมี Internet
echo   - ใช้ไฟล์ SETUP นี้ไฟล์เดียวกันได้ทั้ง Notebook 1 และ 2
echo.
pause
exit /b 0


:INSTALL_ERROR
echo.
echo ============================================================
echo   ติดตั้งไม่สำเร็จ
echo ============================================================
echo.
echo ตรวจสอบว่าเครื่องเชื่อมต่อ Internet อยู่
echo แล้วลองเปิด SETUP_RACE_SYSTEM.bat อีกครั้ง
echo.
pause
exit /b 1


:VERIFY_ERROR
echo.
echo ============================================================
echo   Package บางตัวติดตั้งไม่สมบูรณ์
echo ============================================================
echo.
echo กรุณาส่งภาพหน้าจอนี้มาให้ตรวจสอบ
echo.
pause
exit /b 1
