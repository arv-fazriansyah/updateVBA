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
    echo [ERROR] GUNAKAN FORMAT: vDD.MM.YYYY[TARGET_NAME]
    pause
    call :cleanup
    exit /b
)

title PATCH ARB [v2025]

call :Banner
"%SystemRoot%\System32\timeout.exe" /t 2 >nul

::=============================================================
::  Definisi direktori dan variabel
::=============================================================
echo   [1/10] Menyiapkan update manager...
set "source=%temp%\updateARB-main\ARB2026"
set "exe=%temp%\7-Zip.exe"
set "backup_dir=%~dp0backup"
set "download_path=%temp%\updateARB.zip"
set "file=%~dp0%target%.xlsb"
set "message="
set "download_url=https://github.com/arv-fazriansyah/updateARB/archive/refs/heads/main.zip"
set "zip_url=https://raw.githubusercontent.com/arv-fazriansyah/updateVBA/main/temp/zip/portable/7-Zip.exe"
"%SystemRoot%\System32\timeout.exe" /t 2 >nul

::=============================================================
::  Bersihkan file/folder temp lama
::=============================================================
echo   [2/10] Membersihkan file sementara...
if exist "%temp%\updateARB-main" rmdir /s /q "%temp%\updateARB-main"
if exist "%download_path%" del /f /q "%download_path%"
"%SystemRoot%\System32\timeout.exe" /t 2 >nul

::=============================================================
::  Cek koneksi internet
::=============================================================
echo   [3/10] Mengecek koneksi internet...
"%SystemRoot%\System32\ping.exe" -n 1 google.com >nul 2>nul
if errorlevel 1 (
    set "message=[ERROR] TIDAK ADA KONEKSI INTERNET."
    call :msg
    call :cleanup
    exit /b
)
"%SystemRoot%\System32\timeout.exe" /t 2 >nul

::=============================================================
::  Pastikan file target ada
::=============================================================
echo   [4/10] Mengecek file ARB...
if not exist "%file%" (
    set "message=[ERROR] FILE TIDAK DITEMUKAN."
    call :msg
    call :cleanup
    exit /b
)
"%SystemRoot%\System32\timeout.exe" /t 2 >nul

::=============================================================
::  Tutup file Excel target saja
::=============================================================
echo   [5/10] Menutup file ARB...
"%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -command "$xl=[Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application'); $wb=$xl.Workbooks | Where-Object {$_.Name -like '*%target%*'}; if($wb){$wb.Close($false)}" >nul 2>&1
"%SystemRoot%\System32\timeout.exe" /t 2 >nul

"%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -command "$xl=[Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application'); if($xl.Workbooks | Where-Object {$_.Name -like '*%target%*'}){ exit 1 } else { exit 0 }" >nul 2>&1
if %errorlevel% equ 1 (
    taskkill /f /im excel.exe >nul 2>nul
    "%SystemRoot%\System32\timeout.exe" /t 1 >nul
)
"%SystemRoot%\System32\timeout.exe" /t 2 >nul

::=============================================================
::  Backup file Excel lama
::=============================================================
echo   [6/10] Membuat backup file ARB...
if not exist "%backup_dir%" mkdir "%backup_dir%"
xcopy "%file%" "%backup_dir%\" /Y >nul 2>nul
"%SystemRoot%\System32\timeout.exe" /t 2 >nul

::=============================================================
::  Unduh file update dari GitHub
::=============================================================
echo   [7/10] Mengunduh file update...
"%SystemRoot%\System32\curl.exe" -L -s "%zip_url%" -o "%exe%" || (
    set "message=[ERROR] GAGAL MENGUNDUH 7-Zip."
    call :msg
    call :cleanup
    exit /b
)
"%SystemRoot%\System32\timeout.exe" /t 2 >nul
"%SystemRoot%\System32\curl.exe" -L -s "%download_url%" -o "%download_path%" >nul 2>nul || (
    set "message=[ERROR] GAGAL MENGUNDUH FILE UPDATE."
    call :msg
    call :cleanup
    exit /b
)
"%SystemRoot%\System32\timeout.exe" /t 2 >nul

::=============================================================
::  Ekstrak file ZIP ke folder temp
::=============================================================
echo   [8/10] Mengekstrak file update...
"%exe%" x "%download_path%" -o"%temp%" -y >nul || (
    set "message=[ERROR] GAGAL MENEKSTRAK FILE UPDATE."
    call :msg
    call :cleanup
    exit /b
)

del "%download_path%"
"%SystemRoot%\System32\timeout.exe" /t 2 >nul

::=============================================================
::  Update file menggunakan 7-Zip
::=============================================================
echo   [9/10] Memperbarui file ARB...
start /wait "" "%exe%" a "%file%" "%source%\*" || (
    set "message=[ERROR] GAGAL MEMPERBARUI FILE."
    call :msg
    call :cleanup
    exit /b
)
"%SystemRoot%\System32\timeout.exe" /t 2 >nul

::=============================================================
::  [10/10] Ganti nama file hasil update
::=============================================================
echo   [10/10] Mengganti versi terbaru %version%...
set "batname=%version%"
set "basename=%target%.xlsb"
for /f "tokens=1,* delims=_" %%a in ("%basename%") do (
    if /i "%%a"=="%batname%" (
        set "basename=%%b"
        goto :version_done
    )
    echo %%a | findstr /r /c:"^v[0-9][0-9]*\.[0-9][0-9]*\.[0-9][0-9][0-9][0-9]*" >nul
    if not errorlevel 1 (
        set "basename=%%b"
        goto :version_done
    )
)
:version_done
set "new_name=%batname%_%basename%"
ren "%file%" "%new_name%" || (
    set "message=[ERROR] GAGAL MENGGANTI NAMA FILE."
    call :msg
    call :cleanup
    exit /b
)
"%SystemRoot%\System32\timeout.exe" /t 2 >nul

::=============================================================
::  Hapus folder temp
::=============================================================
if exist "%temp%\updateARB-main" rmdir /s /q "%temp%\updateARB-main"
"%SystemRoot%\System32\timeout.exe" /t 2 >nul

::=============================================================
::  Selesai
::=============================================================
set "message=[OK] PREOSES UPDATE SELESAI."
call :msg
call :cleanup
exit /b

::=============================================================
::  Banner
::=============================================================
:Banner
cls
echo ========================================================================
echo    ###    ########  ########      #######    #####    #######   #######  
echo   ## ##   ##     ## ##     ##    ##     ##  ##   ##  ##     ## ##     ## 
echo  ##   ##  ##     ## ##     ##           ## ##     ##        ## ##        
echo ##     ## ########  ########      #######  ##     ##  #######  ########  
echo ######### ##   ##   ##     ##    ##        ##     ## ##        ##     ## 
echo ##     ## ##    ##  ##     ##    ##         ##   ##  ##        ##     ## 
echo ##     ## ##     ## ########     #########   #####   #########  #######   
echo ========================================================================
echo Versi : %version%
echo File  : %target%.xlsb
echo ========================================================================
echo.
exit /b

::=============================================================
::  Fungsi Pesan
::=============================================================
:msg
echo.
echo ========================================================================
echo %message%
echo ========================================================================
echo.
pause
exit /b

::=============================================================
::  Fungsi Cleanup (hapus diri sendiri)
::=============================================================
:cleanup
"%SystemRoot%\System32\ping.exe" 127.0.0.1 -n 2 >nul
(del "%~f0") >nul 2>&1
exit /b
