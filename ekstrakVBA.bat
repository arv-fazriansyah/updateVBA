@echo off
setlocal EnableExtensions EnableDelayedExpansion
title EKSTRAK VBA

set "ZIP=7-Zip.exe"
set "OUT=EktraksVBA"
set "ZIPURL=https://github.com/arv-fazriansyah/updateVBA/raw/main/temp/zip/portable/7-Zip.exe"

:: =====================================================
:: BANNER
:: =====================================================
echo =====================================================
echo                    EKSTRAK VBA
echo =====================================================
echo.

:: =====================================================
:: CEK 7-ZIP
:: =====================================================
if not exist "%ZIP%" (
    echo [INFO] 7-Zip.exe tidak ditemukan. Mengunduh...
    
    :: Gunakan bitsadmin (Windows lama dan baru)
    bitsadmin /transfer myDownloadJob /download /priority normal "%ZIPURL%" "%CD%\%ZIP%" >nul 2>&1

    if not exist "%ZIP%" (
        echo [ERROR] Gagal mengunduh 7-Zip.exe!
        goto END
    )
    echo [OK] 7-Zip.exe berhasil diunduh.
) else (
    echo [INFO] 7-Zip.exe ditemukan.
)
timeout /t 1 >nul

:: =====================================================
:: DETEKSI FILE XLSB
:: =====================================================
set i=0
for %%F in (*.xlsb) do (
    set /a i+=1
    set "FILE[!i!]=%%F"
)

if %i%==0 (
    echo [ERROR] Tidak ada file ditemukan di folder ini!
    goto END
)

echo.
echo Pilih file:
for /L %%N in (1,1,%i%) do echo %%N. !FILE[%%N]!
echo.

set /p PILIH=Input nomor: 

if not defined FILE[%PILIH%] (
    echo [ERROR] Pilihan tidak valid!
    goto END
)

set "FILE=!FILE[%PILIH%]!"
echo.
echo [INFO] File dipilih: %FILE%
timeout /t 1 >nul

:: =====================================================
:: HAPUS HASIL EKSTRAKSI LAMA
:: =====================================================
echo [INFO] Menghapus hasil ekstraksi lama...
if exist "%OUT%\xl" rmdir /s /q "%OUT%\xl"
if exist "%OUT%\customUI" rmdir /s /q "%OUT%\customUI"
if not exist "%OUT%" mkdir "%OUT%"
echo.
timeout /t 1 >nul

:: =====================================================
:: EKSTRAK FILE
:: =====================================================
echo [1/2] Ekstrak VBA Project...
"%ZIP%" x "%FILE%" xl\vbaproject.bin -o"%OUT%" -y >nul 2>&1
if errorlevel 1 goto GAGAL
timeout /t 1 >nul

echo [2/2] Ekstrak Custom UI...
"%ZIP%" x "%FILE%" customUI -o"%OUT%" -y >nul 2>&1
if errorlevel 1 goto GAGAL
timeout /t 1 >nul

:: =====================================================
:: RENAME FILE — FORMAT [vN]
:: =====================================================
for %%A in ("%FILE%") do (
    set "NAME=%%~nA"
    set "EXT=%%~xA"
)

set "BASE=%NAME%"
set "NEXT=1"

:: Ambil versi terakhir dari format [vN] jika ada
for /f "tokens=1,2 delims=[]" %%a in ("%NAME%") do (
    set "PART1=%%a"
    set "PART2=%%b"
)

if defined PART2 (
    set "NUM=!PART2:v=!"
    set /a NEXT=NUM+1
    set "BASE=!PART1!"
)

set "NEWNAME=%BASE%[v%NEXT%]%EXT%"

ren "%FILE%" "%NEWNAME%" || (
    echo [ERROR] Gagal mengganti nama file!
    goto END
)

echo.

echo =====================================================
echo                SEMUA PROSES SELESAI
echo =====================================================
goto END

:: =====================================================
:: GAGAL
:: =====================================================
:GAGAL
echo.
echo [GAGAL] Terjadi kesalahan saat ekstraksi!
echo [INFO] File tidak di-rename
goto END

:: =====================================================
:: AKHIR
:: =====================================================
:END
echo.
pause
exit /b
