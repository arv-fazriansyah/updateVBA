@echo off
set "URL=http://localhost:8535"
set "PS_PATH=C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
set "ENV_PATH=C:\newappraporsd2025\wwwroot\.env"
:: Menyimpan file di folder Temp pengguna
set "CF_EXE=%TEMP%\cloudflared.exe"

cls
echo ======================================================
echo                  CLOUDFLARE TUNNEL
echo ======================================================

:: 1. Cek apakah cloudflared.exe sudah ada di folder Temp
if not exist "%CF_EXE%" (
    echo [SYSTEM] cloudflared tidak ditemukan.
    echo [SYSTEM] Mohon tunggu, mengunduh cloudflared...
    
    "%PS_PATH%" -Command "(New-Object Net.WebClient).DownloadFile('https://github.com/cloudflare/cloudflared/releases/latest/download/cloudflared-windows-amd64.exe', '%CF_EXE%')"
    
    if exist "%CF_EXE%" (
        echo [SUCCESS] Berhasil mengunduh.
    ) else (
        echo [ERROR] Gagal mengunduh. Pastikan Anda terhubung ke internet.
        pause
        exit /b
    )
)

:: 2. Proses penghapusan .env lama
if exist "%ENV_PATH%" del /f /q "%ENV_PATH%"
echo [SYSTEM] Menghubungkan ke Cloudflare...
echo.

:: 3. Menjalankan Tunnel dari folder Temp
"%PS_PATH%" -NoProfile -ExecutionPolicy Bypass -Command "$envP='%ENV_PATH%'; & '%CF_EXE%' tunnel --url %URL% 2>&1 | ForEach-Object { if ($_ -match 'https://[a-z0-9-]+\.trycloudflare\.com') { $url = $matches[0]; $url | clip; \"session.driver = 'CodeIgniter\Session\Handlers\FileHandler'`napp.baseURL = '$url'\" | Out-File $envP -Encoding UTF8; Clear-Host; Write-Host ' +--------------------------------------------------------------------------------------+' -Fore Cyan; Write-Host ' |                            CLOUDFARE TUNNEL BERHASIL AKTIF                           |' -Fore Cyan; Write-Host ' +--------------------------------------------------------------------------------------+' -Fore Cyan; Write-Host '  URL PUBLIK   : ' -NoNewline; Write-Host $url -Fore White -Back Blue; Write-Host '  LOKAL TARGET : %URL%' -Fore Gray; Write-Host ' +--------------------------------------------------------------------------------------+' -Fore Cyan; Write-Host '  [V] URL telah disalin ke Clipboard' -Fore Green; Write-Host '  [!] JANGAN TUTUP jendela ini selama tunnel digunakan.' -Fore Yellow; Write-Host ' +--------------------------------------------------------------------------------------+' -Fore Cyan } }"

pause
