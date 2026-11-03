
@echo off
setlocal EnableExtensions
title Race Dashboard - Notebook 1 Launcher

cd /d "%~dp0"

echo ==========================================
echo RACE DASHBOARD - NOTEBOOK 1
echo ==========================================

set "PY_CMD="

where py >nul 2>&1
if %errorlevel%==0 set "PY_CMD=py -3"

if not defined PY_CMD (
    where python >nul 2>&1
    if %errorlevel%==0 set "PY_CMD=python"
)

if not defined PY_CMD (
    echo ERROR: Python not found.
    pause
    exit /b 1
)

if not exist "sync_hub.py" (
    echo ERROR: sync_hub.py not found.
    pause
    exit /b 1
)

if not exist "main.py" (
    echo ERROR: main.py not found.
    pause
    exit /b 1
)

echo Starting Central Sync Hub...
start "Race Sync Hub" cmd /k "%PY_CMD% sync_hub.py"

timeout /t 2 /nobreak >nul

echo Starting Dashboard Server...
start "Race Dashboard Server" cmd /k "%PY_CMD% -m uvicorn main:app --host 0.0.0.0 --port 8000"

timeout /t 3 /nobreak >nul

echo Opening Dashboard...
start "" "http://127.0.0.1:8000"

echo.
echo Launcher finished.
pause
