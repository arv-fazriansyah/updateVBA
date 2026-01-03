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
echo Versi	: %version%
echo File	: %target%.xlsb
echo ========================================================================
echo.
"%SystemRoot%\System32\timeout.exe" /t 2 >nul

::=============================================================
::  Definisi direktori dan variabel
::=============================================================
echo   [1/10] Menyiapkan update manager...
set "detected_path="
for /f "delims=" %%i in ('%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe -command "$xl=[Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application'); $xl.Workbooks | Where-Object {$_.Name -like '*%target%*'} | Select-Object -ExpandProperty FullName" 2^>nul') do (
    set "detected_path=%%i"
)
if not defined detected_path (
    for %%F in ("%~dp0*%target%*.xlsb") do (
        set "detected_path=%%~fF"
    )
)
if not defined detected_path (
    set "message=[ERROR] FILE TIDAK DITEMUKAN."
    call :msg
    call :cleanup
	exit /b
)
echo   [OK] PATH: %detected_path%
for %%A in ("%detected_path%") do set "parent_dir=%%~dpA"
set "appid=ARB 2026"
set "server=%appid%\2. UPDATE"
set "backup_dir=%parent_dir%BACKUP APLIKASI %appid%"
set "folder=%temp%\%server%"
:: --- TAMBAHKAN LOGIKA INI ---
if not exist "%folder%" mkdir "%folder%"
:: ----------------------------
set "source=%folder%\updateARB-main\%appid%"
set "exe=%folder%\7za.exe"
set "download_path=%folder%\update.zip"
set "file=%detected_path%"
set "download_url=https://codeload.github.com/arv-fazriansyah/updateARB/zip/refs/heads/main"
set "zip_url=https://raw.githubusercontent.com/arv-fazriansyah/updateVBA/refs/heads/main/tools/7za.exe"
"%SystemRoot%\System32\timeout.exe" /t 2 >nul

::=============================================================
::  Bersihkan file/folder temp lama
::=============================================================
echo   [2/10] Membersihkan file sementara...
if exist "%folder%\updateARB-main" rmdir /s /q "%folder%\updateARB-main"
if exist "%download_path%" del /f /q "%download_path%"
:: Loop semua file .bat di parent_dir
for %%F in ("!parent_dir!*.bat") do (
    :: Ambil nama file saja tanpa path
    set "current_file=%%~nxF"
    
    :: Jika nama file tidak sama dengan nama script ini, maka hapus
    if /i not "!current_file!"=="!batname!.bat" (
        del /f /q "%%F"
    )
)
"%SystemRoot%\System32\timeout.exe" /t 2 >nul
for %%F in ("%~dp0*.bat") do (
    set "current_file=%%~nxF"
    
    :: Bandingkan nama file yang ditemukan dengan nama script ini (%~nx0)
    if /i not "%%~nxF"=="%~nx0" (
        del /f /q "%%~F"
    )
)
"%SystemRoot%\System32\timeout.exe" /t 2 >nul

::=============================================================
::  Cek koneksi internet
::=============================================================
echo   [3/10] Mengecek koneksi internet...
"%SystemRoot%\System32\ping.exe" -n 3 google.com >nul 2>nul
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
::%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe -command "Get-CimInstance Win32_Process -Filter \"Name='excel.exe'\" | Where-Object { $_.CommandLine -like '*%target%*' } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force }" >nul 2>&1
::"%SystemRoot%\System32\timeout.exe" /t 2 >nul

"%SystemRoot%\System32\taskkill.exe" /f /im excel.exe >nul 2>nul
"%SystemRoot%\System32\timeout.exe" /t 2 >nul

::"%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -command "$xl=[Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application'); if($xl.Workbooks | Where-Object {$_.Name -like '*%target%*'}){ exit 1 } else { exit 0 }" >nul 2>&1
::if %errorlevel% equ 1 (
::    "%SystemRoot%\System32\taskkill.exe" /f /im excel.exe >nul 2>nul
::    "%SystemRoot%\System32\timeout.exe" /t 1 >nul
::)
::"%SystemRoot%\System32\timeout.exe" /t 2 >nul

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
"%exe%" x "%download_path%" -o"%folder%" -y >nul || (
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
echo   [10/10] Cleaning ^& Finishing...
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
if exist "%folder%\updateARB-main" rmdir /s /q "%folder%\updateARB-main"
"%SystemRoot%\System32\timeout.exe" /t 2 >nul

"%SystemRoot%\System32\taskkill.exe" /f /im powershell.exe >nul 2>nul
"%SystemRoot%\System32\timeout.exe" /t 2 >nul
::=============================================================
::  Selesai
::=============================================================
set "message=[OK] PROSES UPDATE SELESAI."
call :msg
call :cleanup
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
