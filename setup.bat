@echo off
echo ============================================
echo   SETUP IDLE LAMBDA FINDER (sekali jalan saja)
echo ============================================

where python >nul 2>nul
if errorlevel 1 (
    echo [ERROR] Python tidak ditemukan. Install Python 3.10+ dari python.org
    echo         Saat install, centang "Add Python to PATH".
    pause
    exit /b 1
)

echo [INFO] Membuat virtual environment ...
python -m venv .venv
call .venv\Scripts\activate.bat

echo [INFO] Install dependency ...
pip install --upgrade pip
pip install -r requirements.txt

echo.
echo ============================================
echo   SETUP SELESAI. Sekarang pakai run.bat
echo ============================================
pause
