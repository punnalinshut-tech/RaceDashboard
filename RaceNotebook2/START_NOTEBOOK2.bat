
@echo off
setlocal
title Race Dashboard - Notebook 2

cd /d "%~dp0"

echo ==========================================
echo RACE DASHBOARD V3 - NOTEBOOK 2
echo ==========================================

if not exist "sender_5km_lan.py" (
    echo ERROR: sender_5km_lan.py not found
    pause
    exit /b 1
)

echo Starting 5KM Sender...
start "Race 5KM Sender" cmd /k "py -3 sender_5km_lan.py"

timeout /t 3 /nobreak >nul

if exist "data\race_5KM.xlsx" (
    echo Opening Excel...
    start "" "data\race_5KM.xlsx"
) else (
    echo WARNING: Excel file not found
)

echo.
echo Notebook 2 startup completed.
pause