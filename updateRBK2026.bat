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

timeout /t 3 >nul

::=============================================================
::  Definisi direktori dan variabel
::=============================================================
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

::=============================================================
::  Cek koneksi internet
::=============================================================
ping -n 1 google.com >nul 2>nul
if errorlevel 1 (
    set "message=Tidak ada koneksi internet. Silakan periksa koneksi Anda."
    call :msg
    exit /b
)
echo Koneksi internet OK.
timeout /t 2 >nul

::=============================================================
::  Tutup semua instance Excel tanpa pesan
::=============================================================
taskkill /f /im excel.exe >nul 2>nul

::=============================================================
::  Bersihkan file/folder temp lama
::=============================================================
if exist "%download_dir%\temp" rmdir /s /q "%download_dir%\temp"
if exist "%download_path%" del /f /q "%download_path%"

::=============================================================
::  Cari file Excel (.xlsb) di direktori instalasi
::=============================================================
for %%i in ("%install_dir%\*.xlsb") do (
    set "file=%install_dir%\%%~nxi"
    set "original_name=%%~nxi"
    goto :file_found
)

:: Jika tidak ditemukan file Excel
set "message=Tidak ada file Excel (.xlsb) ditemukan di folder ini. Simpan file RBK Anda di sini."
call :msg
exit /b

:file_found
timeout /t 2 >nul

::=============================================================
::  Unduh file update dari GitHub
::=============================================================
echo.
echo Proses update RBK...
curl -L "%download_url%" -o "%download_path%" || (
    set "message=Gagal mengunduh file."
    call :msg
    exit /b
)

timeout /t 2 >nul

::=============================================================
::  Ekstrak file ZIP ke folder temp
::=============================================================
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
if not exist "%backup_dir%" mkdir "%backup_dir%"
xcopy "%file%" "%backup_dir%\" /Y >nul 2>nul
timeout /t 2 >nul

::=============================================================
::  Update file menggunakan 7-Zip
::=============================================================
start /min "" "%exe%" a "%file%" "%source%\*" || (
    set "message=Gagal memperbarui file."
    call :msg
    exit /b
)

timeout /t 2 >nul

::=============================================================
::  Ganti nama file setelah update
::=============================================================
set "new_name=update_%original_name%"
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
echo ======================================
echo Message: %message%
echo ======================================
echo.
pause
exit /b
