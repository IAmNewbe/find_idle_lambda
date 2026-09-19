@echo off
cd /d "%~dp0"

if not exist ".venv\Scripts\activate.bat" (
    echo [ERROR] Virtual environment belum ada.
    echo         Jalankan setup.bat dulu ^(sekali saja^), baru run.bat.
    pause
    exit /b 1
)

call .venv\Scripts\activate.bat

echo [INFO] Mencari channel idle dari file di input\raw\ ...
python src\find_idle_lambda.py --config config\config.yaml

if errorlevel 1 (
    echo.
    echo [GAGAL] Ada error di atas. Baca pesan [ERROR] untuk tahu penyebabnya.
) else (
    echo.
    echo [SELESAI] Cek folder output\ untuk hasil mapping-nya.
)

pause
