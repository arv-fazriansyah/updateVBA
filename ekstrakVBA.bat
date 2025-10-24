@echo off
setlocal enabledelayedexpansion

:: Ekstrak vbaproject.bin tanpa konfirmasi
7-Zip.exe e MASTER_RBK2026.xlsb xl\vbaproject.bin -y

echo vbaproject.bin berhasil diekstrak
pause
exit /b
