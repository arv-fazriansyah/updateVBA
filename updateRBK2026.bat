@echo off
setlocal enabledelayedexpansion
color a
echo.

::=============================================================
::  Banner
::=============================================================
echo ########     ########     ##    ##     #######    #####    #######  ######## 
echo ##     ##    ##     ##    ##   ##     ##     ##  ##   ##  ##     ## ##       
echo ##     ##    ##     ##    ##  ##             ## ##     ##        ## ##       
echo ########     ########     #####        #######  ##     ##  #######  #######  
echo ##   ##      ##     ##    ##  ##      ##        ##     ## ##              ## 
echo ##    ##     ##     ##    ##   ##     ##         ##   ##  ##        ##    ## 
echo ##     ##    ########     ##    ##    #########   #####   #########  ######  
echo.
timeout /t 2 >nul

::=============================================================
::  Definisi direktori dan variabel
::=============================================================
echo [1/10] Menyiapkan variabel dan direktori kerja...
set "download_dir=%temp%"
set "install_dir=%CD%"
set "source=%download_dir%\temp\home"
set "exe=%download_dir%\temp\zip\portable\7-Zip.exe"
set "backup_dir=%install_dir%\backup"
set "download_url=https://github.com/arv-fazriansyah/updateVBA/archive/refs/heads/main.zip"
set "download_path=%download_dir%\updateVBA.zip"
set "file="
set "original_name="
set "message="
timeout /t 2 >nul

::=============================================================
::  Cek koneksi internet
::=============================================================
echo [2/10] Mengecek koneksi internet...
ping -n 1 google.com >nul 2>nul
if errorlevel 1 (
    set "message=Tidak ada koneksi internet. Silakan periksa koneksi Anda."
    call :msg
    exit /b
)
timeout /t 2 >nul

::=============================================================
::  Tutup semua instance Excel tanpa pesan
::=============================================================
echo [3/10] Menutup semua instance Excel...
taskkill /f /im excel.exe >nul 2>nul
timeout /t 2 >nul

::=============================================================
::  Bersihkan file/folder temp lama
::=============================================================
echo [4/10] Membersihkan file sementara lama...
if exist "%download_dir%\temp" rmdir /s /q "%download_dir%\temp"
if exist "%download_path%" del /f /q "%download_path%"
timeout /t 2 >nul

::=============================================================
::  Cari file Excel (.xlsb) di direktori instalasi
::=============================================================
echo [5/10] Mendeteksi file Excel (*.xlsb) di direktori ini...
for %%i in ("%install_dir%\*.xlsb") do (
    set "file=%install_dir%\%%~nxi"
    set "original_name=%%~nxi"
    goto :file_found
)
set "message=Tidak ada file Excel (.xlsb) ditemukan di folder ini. Simpan file RBK Anda di sini."
call :msg
exit /b

:file_found
timeout /t 2 >nul

::=============================================================
::  Unduh file update dari GitHub
::=============================================================
echo [6/10] Mengunduh file update dari server...
curl -L -s "%download_url%" -o "%download_path%" >nul 2>nul || (
    set "message=Gagal mengunduh file."
    call :msg
    exit /b
)
timeout /t 2 >nul

::=============================================================
::  Ekstrak file ZIP ke folder temp
::=============================================================
echo [7/10] Mengekstrak file update...
tar -xf "%download_path%" --strip-components=1 -C "%download_dir%" "updateVBA-main/*" || (
    set "message=Gagal mengekstrak file."
    call :msg
    exit /b
)
del "%download_path%"
timeout /t 2 >nul

::=============================================================
::  Backup file Excel lama
::=============================================================
echo [8/10] Membuat backup file lama...
if not exist "%backup_dir%" mkdir "%backup_dir%"
xcopy "%file%" "%backup_dir%\" /Y >nul 2>nul
timeout /t 2 >nul

::=============================================================
::  Update file menggunakan 7-Zip
::=============================================================
echo [9/10] Memperbarui file...
start /min "" "%exe%" a "%file%" "%source%\*" || (
    set "message=Gagal memperbarui file."
    call :msg
    exit /b
)
timeout /t 2 >nul

::=============================================================
::  [10/10] Ganti nama file hasil update
::=============================================================
echo [10/10] Mengganti nama file hasil update...

:: Ambil nama file batch tanpa ekstensi (misal: v25.10.2025)
set "batname=%~n0"

:: Hapus prefix versi lama dari nama file, kalau sebelumnya sudah di-update dengan versi lain
set "basename=%original_name%"
for /f "tokens=1,* delims=_" %%a in ("%basename%") do (
    if /i "%%a"=="%batname%" (
        set "basename=%%b"
    )
)

:: Buat nama baru dengan prefix versi batch
set "new_name=%batname%_%basename%"

ren "%file%" "%new_name%" || (
    set "message=Gagal mengganti nama file."
    call :msg
    exit /b
)
timeout /t 2 >nul

::=============================================================
::  Hapus folder temp
::=============================================================
if exist "%download_dir%\temp" rmdir /s /q "%download_dir%\temp"
timeout /t 2 >nul

::=============================================================
::  Selesai
::=============================================================

set "message=Proses update selesai!"
call :msg
exit /b

::=============================================================
::  Fungsi Pesan
::=============================================================
:msg
echo.
echo ============================================
echo     PROSES UPDATE SELESAI DENGAN SUKSES!  
echo ============================================
echo.
pause
exit /b
