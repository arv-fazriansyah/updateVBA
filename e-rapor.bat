@echo off
set "URL=http://localhost:8535"
set "PS_PATH=C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
set "ENV_PATH=C:\newappraporsd2025\wwwroot\.env"

cls
echo.
if exist "%ENV_PATH%" del /f /q "%ENV_PATH%"
echo  [SYSTEM] Menghubungkan ke Cloudflare...
echo.

"%PS_PATH%" -NoProfile -ExecutionPolicy Bypass -Command "$envP='%ENV_PATH%'; cloudflared tunnel --url %URL% 2>&1 | ForEach-Object { if ($_ -match 'https://[a-z0-9-]+\.trycloudflare\.com') { $url = $matches[0]; \"session.driver = 'CodeIgniter\Session\Handlers\FileHandler'`napp.baseURL = '$url'\" | Out-File $envP -Encoding UTF8; Clear-Host; Write-Host ' +--------------------------------------------------------------------------------------+' -Fore Cyan; Write-Host ' |                            CLOUDFARE TUNNEL BERHASIL AKTIF                           |' -Fore Cyan; Write-Host ' +--------------------------------------------------------------------------------------+' -Fore Cyan; Write-Host '  URL PUBLIK  : ' -NoNewline; Write-Host $url -Fore White -Back Blue; Write-Host '  LOKAL TARGET: %URL%' -Fore Gray; Write-Host ' +--------------------------------------------------------------------------------------+' -Fore Cyan; Write-Host '  [V] URL telah disalin ke Clipboard.' -Fore Green; Write-Host '  [!] JANGAN TUTUP jendela ini selama tunnel digunakan.' -Fore Yellow; Write-Host ' +--------------------------------------------------------------------------------------+' -Fore Cyan; $url | clip } }"

pause
