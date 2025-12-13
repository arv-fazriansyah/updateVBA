@echo off
setlocal enabledelayedexpansion

set FILE=MASTER_RBK2026_rev5.xlsb
set ZIP=7-Zip.exe
set OUT=EktraksVBA

:: Buat folder tujuan jika belum ada
if not exist "%OUT%" mkdir "%OUT%"

:: Ekstrak vbaproject.bin
"%ZIP%" e "%FILE%" xl\vbaproject.bin -o"%OUT%" -y

:: Ekstrak folder customUI (pakai struktur folder)
"%ZIP%" x "%FILE%" customUI -o"%OUT%" -y

echo.
echo Ekstraksi selesai.
echo File tersimpan di folder "%OUT%"
pause
exit /b
