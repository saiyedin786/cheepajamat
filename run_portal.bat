@echo off
title Chhipa Samaj Khanda Falsa Jodhpur - Portal Launcher
echo ====================================================================
echo   CHHIPA SAMAJ KHANDA FALSA, JODHPUR - CENSUS ^& DIRECTORY PORTAL
echo ====================================================================
echo.
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [INFO] Python was not detected in PATH. Opening Standalone Browser Edition...
    start "" "chhipa_samaj_standalone.html"
    exit /b 0
)

echo [INFO] Checking Flask installation...
python -c "import flask" >nul 2>&1
if %errorlevel% neq 0 (
    echo [INFO] Installing required dependencies...
    pip install -r requirements.txt
)

echo [INFO] Starting Local Server at http://127.0.0.1:5000 ...
echo [INFO] Public Portal: http://127.0.0.1:5000
echo [INFO] Admin Panel:   http://127.0.0.1:5000/admin (Default: admin / admin123)
start "" "http://127.0.0.1:5000"
python app.py
pause
