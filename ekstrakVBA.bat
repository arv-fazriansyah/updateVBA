@echo off
setlocal enabledelayedexpansion

set FILE=MASTER_RBK2026_rev7.xlsb
set ZIP=7-Zip.exe
set OUT=EktraksVBA

:: Buat folder tujuan
if not exist "%OUT%" mkdir "%OUT%"

:: Ekstrak vbaproject.bin
"%ZIP%" x "%FILE%" xl\vbaproject.bin -o"%OUT%" -y

:: Ekstrak folder customUI BESERTA foldernya
"%ZIP%" x "%FILE%" customUI -o"%OUT%" -y

echo.
echo Ekstraksi selesai!
pause
exit /b
