@echo off
chcp 65001 >nul
cd /d "%~dp0"
title SEMAR - Aplikasi Web

where python >nul 2>nul
if errorlevel 1 (
  echo Python belum terpasang atau belum masuk PATH.
  echo Pasang dari python.org/downloads dan CENTANG "Add python.exe to PATH".
  pause
  exit /b 1
)

if not exist "%USERPROFILE%\.streamlit" mkdir "%USERPROFILE%\.streamlit"
if not exist "%USERPROFILE%\.streamlit\credentials.toml" (
  > "%USERPROFILE%\.streamlit\credentials.toml" echo [general]
  >> "%USERPROFILE%\.streamlit\credentials.toml" echo email = ""
)

echo Memasang perlengkapan. Pertama kali bisa 2-5 menit, mohon tunggu...
python -m pip install -r requirements.txt
if errorlevel 1 (
  echo.
  echo Pemasangan gagal. Cek koneksi internet lalu jalankan lagi file ini.
  pause
  exit /b 1
)

echo.
echo Menjalankan SEMAR. Browser akan terbuka sendiri.
echo JANGAN tutup jendela hitam ini selama SEMAR dipakai. Tutup jendela ini untuk mematikan SEMAR.
echo.
python -m streamlit run app_web.py --browser.gatherUsageStats false --theme.base light --theme.primaryColor "#F58220" --theme.backgroundColor "#FFF8EC" --theme.secondaryBackgroundColor "#FFEFD2" --theme.textColor "#1B2440"
pause
