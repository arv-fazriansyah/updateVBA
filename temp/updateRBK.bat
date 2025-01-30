@echo off
setlocal enabledelayedexpansion
color c
:: Tampilkan nama besar di awal
echo.

:: Definisikan direktori dan variabel
set "download_dir=%temp%"
set "install_dir=%CD%"
set "source=%download_dir%\temp\home"
set "exe=%download_dir%\temp\zip\7-Zip.exe"
set "backup_dir=%install_dir%\backup"
set "download_url=https://github.com/arv-fazriansyah/updateVBA/archive/refs/heads/main.zip"
set "download_path=%download_dir%\updateVBA.zip"
set "file="
set "original_name="
set "message="

:: Menambahkan path 7-Zip jika belum ada
set PATH=%ProgramFiles%\7-Zip;%ProgramFiles(x86)%\7-Zip;%PATH%

:: Mengecek koneksi internet
ping -n 1 google.com >nul 2>nul
if errorlevel 1 (
    set message=Tidak ada koneksi internet. Silakan periksa koneksi Anda.
    call :msg
    exit /b
)

:: Mengecek dan menghapus folder temp jika sudah ada
if exist "%download_dir%\temp" rmdir /s /q "%download_dir%\temp"

:: Mengecek dan menghapus file downloadPath jika sudah ada
if exist "%download_path%" del /f /q "%download_path%"

:: Mencari file Excel (.xlsb) di direktori instalasi
for %%i in ("%install_dir%\*.xlsb") do (
    set "file=%install_dir%\%%~nxi"
    set "original_name=%%~nxi"
    goto :file_found
)

:: Jika tidak ditemukan file Excel
set message=Simpan terlebih dahulu file RBK disini.
call :msg
exit /b

:file_found

:: Unduh file updateVBA.zip
curl -L "%download_url%" -o "%download_path%" || (set message=Gagal mengunduh file. & call :msg & exit /b)

:: Ekstrak file ZIP ke folder temp
tar -xf "%download_path%" --strip-components=1 -C "%download_dir%" "updateVBA-main/*" || (set message=Gagal mengekstrak file. & call :msg & exit /b)

del "%download_path%"

:: Mengecek apakah 7-Zip terpasang
if not exist "%ProgramFiles%\7-Zip\7z.exe" (
    echo 7-Zip belum terpasang. Sedang menginstal...
    "%exe%" /S || (echo Gagal menginstal 7-Zip. & exit /b)
    echo 7-Zip telah terinstal.
)

:: Membuat folder backup jika belum ada
if not exist "%backup_dir%" mkdir "%backup_dir%"

:: Membackup file Excel ke folder backup
xcopy "%file%" "%backup_dir%\" /Y >nul 2>nul

:: Proses kompresi file menggunakan 7-Zip
start /min "" "%ProgramFiles%\7-Zip\7z.exe" a "%file%" "%source%\*" || (set message=Gagal memperbarui file. & call :msg & exit /b)

:: Berhasil memperbarui file
set message=File berhasil diupdate!
call :msg

:: Rename file setelah update
set "new_name=update_%original_name%"
ren "%file%" "%new_name%" || (set message=Gagal mengganti nama file. & call :msg & exit /b)

:: Cleanup
if exist "%download_dir%\temp" rmdir /s /q "%download_dir%\temp"

exit /b

:msg
:: Menampilkan pesan menggunakan msg (lebih sederhana)
echo.
echo Message: %message%
echo.
pause
exit /b
