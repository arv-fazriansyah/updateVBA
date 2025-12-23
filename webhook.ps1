# --- CONFIGURATION ---
$port = "8080"
$urlLocal = "http://127.0.0.1:$port"
$folderARB = Join-Path $env:TEMP "ARB2026"
$cfExe = Join-Path $folderARB "cloudflared.exe"
$logFile = Join-Path $folderARB "vba_webhook_log.txt"

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

function Write-Log($pesan) {
    $waktu = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "[$waktu] $pesan" | Out-File -FilePath $logFile -Append -Encoding UTF8
}

# 1. Download Cloudflared
if (-not (Test-Path $cfExe)) {
    Write-Log "INFO: Memulai download cloudflared..."
    try {
        Invoke-WebRequest -Uri 'https://github.com/cloudflare/cloudflared/releases/latest/download/cloudflared-windows-amd64.exe' -OutFile $cfExe -ErrorAction Stop
        Write-Log "INFO: Download selesai. Menunggu sistem melepas file lock..."
        Start-Sleep -Seconds 3 # JEDA PENTING: Memberi waktu sistem/Antivirus selesai mengecek file
    } catch {
        Write-Log "ERROR: GAGAL Download: $($_.Exception.Message)"
        exit
    }
}

# 2. Jalankan Listener HTTP
$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("$urlLocal/")
try {
    $listener.Start()
    Write-Log "INFO: Listener aktif di $urlLocal"
} catch {
    Write-Log "ERROR: GAGAL menjalankan listener: $($_.Message)"
    exit
}

# 3. Jalankan Cloudflare Tunnel (Background Job)
$jobScript = {
    param($cfExe, $urlLocal, $logFile)
    
    # Gunakan file log sementara untuk Cloudflared agar tidak berebutan akses dengan script utama
    $tempCfLog = $logFile.Replace(".txt", "_cf.tmp")
    
    # Jalankan tunnel
    $process = Start-Process -FilePath $cfExe -ArgumentList "tunnel --url $urlLocal --no-autoupdate" `
               -NoNewWindow -PassThru -RedirectStandardError $tempCfLog

    $found = $false
    for ($i=0; $i -lt 40; $i++) {
        if (Test-Path $tempCfLog) {
            $content = Get-Content $tempCfLog -ErrorAction SilentlyContinue
            foreach ($line in $content) {
                # Tulis ulang log cloudflared ke log utama (Opsional, agar satu file)
                "[$i] [Tunnel Output] $line" | Out-File -FilePath $logFile -Append -Encoding UTF8
                
                if ($line -match 'https://[a-z0-9-]+\.trycloudflare\.com') {
                    $urlPublik = $matches[0]
                    try {
                        $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
                        $excel.Sheets("DEV").Range("F10").Value = $urlPublik
                        $excel.Run("TampilkanToast", "Tunnel Aktif", "Koneksi Berhasil", "")
                        $found = $true
                        break
                    } catch { Start-Sleep -Seconds 1 }
                }
            }
        }
        if ($found) { break }
        Start-Sleep -Seconds 2
    }
    # Hapus log sementara setelah URL ditemukan
    if (Test-Path $tempCfLog) { Remove-Item $tempCfLog -Force }
}
Start-Job -ScriptBlock $jobScript -ArgumentList $cfExe, $urlLocal, $logFile

# 4. Loop Utama
Write-Log "INFO: Menunggu pesan inbound..."

try {
    while ($listener.IsListening) {
        $context = $listener.GetContext()
        $pesan = $context.Request.QueryString["teks"]
        
        if ($pesan) {
            Write-Log "WEBHOOK: Pesan diterima -> $pesan"
            try {
                $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
                $excel.Run("TampilkanToast", "APLIKASI ARB 2026", $pesan, "")
            } catch { Write-Log "WARN: Excel sibuk." }
        }
        
        $buffer = [System.Text.Encoding]::UTF8.GetBytes("OK")
        $context.Response.ContentLength64 = $buffer.Length
        $context.Response.OutputStream.Write($buffer, 0, $buffer.Length)
        $context.Response.Close()
    }
} finally {
    $listener.Stop()
    Write-Log "INFO: Script dihentikan."
}
