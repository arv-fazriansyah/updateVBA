@echo off
setlocal enabledelayedexpansion

set "FILE=MASTER_RBK2026_rev7.xlsb"
set "ZIP=7-Zip.exe"
set "OUT=EktraksVBA"

echo =====================================
echo             EKSTRAK VBA 
echo =====================================

:: Cek file sumber
if not exist "%FILE%" (
    echo [ERROR] File %FILE% tidak ditemukan!
    goto END
)

:: Cek 7-Zip
if not exist "%ZIP%" (
    echo [ERROR] 7-Zip.exe tidak ditemukan!
    goto END
)

:: Buat folder tujuan
if not exist "%OUT%" mkdir "%OUT%"

:: Hapus hasil lama
if exist "%OUT%\xl" (
    echo [INFO] Menghapus xl lama...
    rmdir /s /q "%OUT%\xl"
)

if exist "%OUT%\customUI" (
    echo [INFO] Menghapus customUI lama...
    rmdir /s /q "%OUT%\customUI"
)

:: ===============================
:: Ekstraksi SENYAP
:: ===============================

echo [1/2] Ekstrak VBA Project...
"%ZIP%" x "%FILE%" xl\vbaproject.bin -o"%OUT%" -y >nul 2>&1
if errorlevel 1 (
    echo [GAGAL] Ekstrak VBA Project
    goto END
) else (
    echo [OK] VBA Project berhasil diekstrak
)

echo [2/2] Ekstrak Custom UI...
"%ZIP%" x "%FILE%" customUI -o"%OUT%" -y >nul 2>&1
if errorlevel 1 (
    echo [GAGAL] Ekstrak Custom UI
    goto END
) else (
    echo [OK] Custom UI berhasil diekstrak
)

echo =====================================
echo          SELESAI TANPA ERROR
echo =====================================

:END
echo.
pause
exit /b
