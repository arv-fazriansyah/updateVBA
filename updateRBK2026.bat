@echo off
setlocal enabledelayedexpansion
color a

::=============================================================
::  Ambil nama file batch tanpa ekstensi (misal: v25.11.2025[MASTER_RBK2026])
::=============================================================
set "batname=%~n0"

:: Pisahkan versi dan nama target di antara tanda []
for /f "tokens=1,2 delims=[]" %%a in ("%batname%") do (
    set "version=%%a"
    set "target=%%b"
)

:: Jika tidak ada tanda [], keluar dengan pesan error
if "%target%"=="" (
    echo Format nama file batch salah.
    echo Gunakan format: vDD.MM.YYYY[TARGET_NAME].bat
    pause
    exit /b
)

::=============================================================
::  Banner
::=============================================================
echo.
echo ########     ########     ##    ##     #######    #####    #######  ######## 
echo ##     ##    ##     ##    ##   ##     ##     ##  ##   ##  ##     ## ##       
echo ##     ##    ##     ##    ##  ##             ## ##     ##        ## ##       
echo ########     ########     #####        #######  ##     ##  #######  #######  
echo ##   ##      ##     ##    ##  ##      ##        ##     ## ##              ## 
echo ##    ##     ##     ##    ##   ##     ##         ##   ##  ##        ##    ## 
echo ##     ##    ########     ##    ##    #########   #####   #########  ######  
echo.
echo Versi: %version%
echo Target: %target%.xlsb
echo.
timeout /t 2 >nul

::=============================================================
::  Definisi direktori dan variabel
::=============================================================
echo [1/10] Menyiapkan variabel dan direktori kerja...
:: set "install_dir=%CD%"
set "install_dir=%~dp0"
set "download_dir=%temp%"
set "source=%download_dir%\temp\home2026"
set "exe=%download_dir%\temp\zip\portable\7-Zip.exe"
set "backup_dir=%install_dir%\backup"
set "download_url=https://github.com/arv-fazriansyah/updateVBA/archive/refs/heads/main.zip"
set "download_path=%download_dir%\updateVBA.zip"
set "file=%install_dir%\%target%.xlsb"
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
::  Tutup file Excel target saja
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
::  Pastikan file target ada
::=============================================================
echo [5/10] Mengecek file target...
if not exist "%file%" (
    set "message=File target %target%.xlsb tidak ditemukan di folder ini!"
    call :msg
    exit /b
)
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
start /wait "" "%exe%" a "%file%" "%source%\*" || (
    set "message=Gagal memperbarui file."
    call :msg
    exit /b
)
timeout /t 2 >nul

::=============================================================
::  [10/10] Ganti nama file hasil update
::=============================================================
echo [10/10] Mengganti nama file hasil update...

:: Ambil nama file batch tanpa ekstensi (misal: v25.11.2025)
set "batname=%version%"

:: Ambil nama file Excel asli
set "basename=%target%.xlsb"

:: Jika nama file sudah diawali dengan versi lama (vDD.MM.YYYY_), hapus dulu versi lamanya
for /f "tokens=1,* delims=_" %%a in ("%basename%") do (
    if /i "%%a"=="%batname%" (
        set "basename=%%b"
        goto :version_done
    )
    rem Jika diawali dengan pola versi lama vDD.MM.YYYY (cek pola awal v dan titik)
    echo %%a | findstr /r /c:"^v[0-9][0-9]*\.[0-9][0-9]*\.[0-9][0-9][0-9][0-9]*" >nul
    if not errorlevel 1 (
        set "basename=%%b"
        goto :version_done
    )
)
:version_done

:: Buat nama baru dengan prefix versi batch (selalu satu kali)
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
echo     %message%
echo ============================================
echo.
pause
exit /b
