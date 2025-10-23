@echo off
setlocal enabledelayedexpansion
color A
echo.

echo ########     ########     ##    ##     #######    #####    #######  ######## 
echo ##     ##    ##     ##    ##   ##     ##     ##  ##   ##  ##     ## ##       
echo ##     ##    ##     ##    ##  ##             ## ##     ##        ## ##       
echo ########     ########     #####        #######  ##     ##  #######  #######  
echo ##   ##      ##     ##    ##  ##      ##        ##     ## ##              ## 
echo ##    ##     ##     ##    ##   ##     ##         ##   ##  ##        ##    ## 
echo ##     ##    ########     ##    ##    #########   #####   #########  ######  

echo.

:: ==============================
:: KONFIGURASI DASAR
:: ==============================
set "download_dir=%temp%"
set "install_dir=%CD%"
set "source=%download_dir%\temp\home2026"
set "exe=%download_dir%\temp\zip\portable\7-Zip.exe"
set "backup_dir=%install_dir%\backup"
set "download_url=https://github.com/arv-fazriansyah/updateVBA/archive/refs/heads/main.zip"
set "download_path=%download_dir%\updateVBA.zip"
set "file="
set "original_name="
set "message="

:: ==============================
:: MATIKAN EXCEL DAHULU
:: ==============================
echo Menutup semua instance Excel...
taskkill /f /im excel.exe >nul 2>nul

:: ==============================
:: CEK INTERNET
:: ==============================
ping -n 1 google.com >nul 2>nul
if errorlevel 1 (
    set message=Tidak ada koneksi internet. Silakan periksa koneksi Anda.
    call :msg
    exit /b
)

:: ==============================
:: BERSIHKAN FOLDER SEMENTARA
:: ==============================
if exist "%download_dir%\temp" rmdir /s /q "%download_dir%\temp"
if exist "%download_path%" del /f /q "%download_path%"

:: ==============================
:: CARI FILE XLSB
:: ==============================
for %%i in ("%install_dir%\*.xlsb") do (
    set "file=%install_dir%\%%~nxi"
    set "original_name=%%~nxi"
    goto :file_found
)

set message=Simpan terlebih dahulu file RBK di folder ini.
call :msg
exit /b

:file_found

:: ==============================
:: DOWNLOAD UPDATE
:: ==============================
echo Mengunduh update...
curl -L "%download_url%" -o "%download_path%" || (set message=Gagal mengunduh file. & call :msg & exit /b)

:: ==============================
:: EKSTRAK FILE
:: ==============================
echo Mengekstrak file...
tar -xf "%download_path%" --strip-components=1 -C "%download_dir%" "updateVBA-main/*" || (set message=Gagal mengekstrak file. & call :msg & exit /b)

del "%download_path%"

:: ==============================
:: BACKUP FILE LAMA
:: ==============================
echo Mengbackup file lama...
if not exist "%backup_dir%" mkdir "%backup_dir%"
xcopy "%file%" "%backup_dir%\" /Y >nul 2>nul

:: ==============================
:: UPDATE FILE
:: ==============================
echo Proses update file RBK...
start /min "" "%exe%" a "%file%" "%source%\*" || (set message=Gagal memperbarui file. & call :msg & exit /b)

:: ==============================
:: PESAN SUKSES
:: ==============================
set message=File berhasil diupdate!
call :msg

:: ==============================
:: TUNGGU HINGGA FILE TERBUKA BEBAS
:: ==============================
call :wait_until_unlocked

:: ==============================
:: RENAME FILE
:: ==============================
set "new_name=v20.12.2025_%original_name%"
ren "%file%" "%new_name%" || (set message=Gagal mengganti nama file. & call :msg & exit /b)

:: ==============================
:: BUKA KEMBALI FILE EXCEL
:: ==============================
echo Membuka kembali file %new_name%...
start "" "%install_dir%\%new_name%"

:: ==============================
:: HAPUS DIR TEMP
:: ==============================
if exist "%download_dir%\temp" rmdir /s /q "%download_dir%\temp"

:: ==============================
:: HAPUS DIRI SENDIRI
:: ==============================
echo Menghapus file updater...
(
    ping 127.0.0.1 -n 3 >nul
    del "%~f0"
) >nul 2>&1 & exit /b

:: ==============================
:: SUBROUTINES
:: ==============================
:msg
echo.
echo Message: %message%
echo.
timeout /t 2 >nul
exit /b

:wait_until_unlocked
:: Tunggu sampai file tidak terkunci
:loop
>nul 2>nul (
    >>"%file%" (
        rem do nothing
    )
) || (
    timeout /t 1 >nul
    goto loop
)
exit /b
