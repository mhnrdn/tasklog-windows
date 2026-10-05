@echo off
REM ============================================================
REM  Build TaskLog.exe  (klik dua kali file ini)
REM  Hasil ada di folder  dist\TaskLog.exe
REM ============================================================
chcp 65001 >nul
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
    echo [X] Python tidak ditemukan. Pasang dulu dari https://www.python.org/downloads/
    echo     Jangan lupa centang "Add python.exe to PATH".
    pause
    exit /b 1
)

echo [1/4] Pasang dependensi...
python -m pip install --upgrade pip >nul
python -m pip install -r requirements.txt
if errorlevel 1 goto gagal

echo [2/4] Pasang PyInstaller...
python -m pip install pyinstaller
if errorlevel 1 goto gagal

echo [3/4] Buat ikon...
python make_icon.py

echo [4/4] Build exe (butuh 1-3 menit)...
python -m PyInstaller --noconfirm --clean --onefile --noconsole ^
  --name TaskLog --icon icon.ico ^
  --hidden-import win32com ^
  --hidden-import win32com.client ^
  --hidden-import pythoncom ^
  --hidden-import pywintypes ^
  --collect-submodules win32com app.py
if errorlevel 1 goto gagal

echo.
echo ============================================================
echo  SELESAI. File:  %~dp0dist\TaskLog.exe
echo ============================================================
pause
exit /b 0

:gagal
echo.
echo [X] Build gagal. Baca pesan error di atas.
pause
exit /b 1
