# --- CONFIGURATION ---
$port = "8080"
$urlLocal = "http://127.0.0.1:$port"
$currentDir = $PSScriptRoot
if (-not $currentDir) { $currentDir = Get-Location }

$cfExe = Join-Path $currentDir "cloudflared.exe"
$logFile = Join-Path $currentDir "vba_webhook_log.txt"

# Memastikan TLS 1.2 untuk download aman
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

function Write-Log($pesan) {
    $waktu = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$waktu] $pesan"
    Write-Host $logEntry # Tampilkan di console juga
    $logEntry | Out-File -FilePath $logFile -Append -Encoding UTF8
}

# 1. Cek Koneksi Excel di Awal
try {
    $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
    Write-Log "INFO: Koneksi ke Excel berhasil dideteksi."
} catch {
    Write-Host "PERINGATAN: Excel tidak ditemukan. Buka Excel terlebih dahulu!" -ForegroundColor Yellow
}

# 2. Download Cloudflared (Hanya jika belum ada)
if (-not (Test-Path $cfExe)) {
    Write-Log "INFO: Mendownload cloudflared ke $currentDir..."
    try {
        $uri = 'https://github.com/cloudflare/cloudflared/releases/latest/download/cloudflared-windows-amd64.exe'
        Invoke-WebRequest -Uri $uri -OutFile $cfExe -ErrorAction Stop
        Unblock-File -Path $cfExe
        Write-Log "INFO: Download sukses."
    } catch {
        Write-Log "ERROR: Gagal download: $($_.Exception.Message)"
        exit
    }
}

# 3. Jalankan Listener HTTP
$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("$urlLocal/")
try {
    $listener.Start()
    Write-Log "INFO: Listener aktif di $urlLocal"
} catch {
    Write-Log "ERROR: Port $port mungkin sedang digunakan: $($_.Exception.Message)"
    exit
}

# 4. Jalankan Cloudflare Tunnel (Background Job)
$jobScript = {
    param($cfExe, $urlLocal, $logFile)
    $tempCfLog = $logFile.Replace(".txt", "_cf.tmp")
    
    # Menjalankan cloudflared
    $proc = Start-Process -FilePath $cfExe -ArgumentList "tunnel --url $urlLocal --no-autoupdate" `
               -NoNewWindow -PassThru -RedirectStandardError $tempCfLog

    # Cari URL publik selama 60 detik
    for ($i = 0; $i -lt 60; $i++) {
        if (Test-Path $tempCfLog) {
            $content = Get-Content $tempCfLog -ErrorAction SilentlyContinue
            $urlLine = $content | Select-String -Pattern "https://[a-z0-9-]+\.trycloudflare\.com" | Select-Object -First 1
            
            if ($urlLine) {
                $urlPublik = $urlLine.Matches.Value
                try {
                    $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
                    $excel.Sheets("DEV").Range("F10").Value = $urlPublik
                    $excel.Run("TampilkanToast", "Tunnel Aktif", "URL: $urlPublik", "")
                    
                    "[(Get-Date)] TUNNEL: $urlPublik" | Out-File -FilePath $logFile -Append -Encoding UTF8
                    break
                } catch { }
            }
        }
        Start-Sleep -Seconds 1
    }
}
$tunnelJob = Start-Job -ScriptBlock $jobScript -ArgumentList $cfExe, $urlLocal, $logFile

# 5. Loop Utama (Listener)
Write-Log "INFO: Menunggu Webhook... (Tekan Ctrl+C untuk berhenti)"
try {
    while ($listener.IsListening) {
        if ($listener.GetContextAsync) {
            $context = $listener.GetContext()
            $pesan = $context.Request.QueryString["teks"]
            
            if ($pesan) {
                Write-Log "WEBHOOK: Menerima pesan: $pesan"
                try {
                    $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
                    $excel.Run("TampilkanToast", "Pesan Baru", $pesan, "")
                } catch { 
                    Write-Log "ERROR: Tidak bisa mengirim ke Excel (Excel tertutup?)"
                }
            }
            
            # Respon OK ke pengirim
            $buffer = [System.Text.Encoding]::UTF8.GetBytes("OK")
            $context.Response.ContentLength64 = $buffer.Length
            $context.Response.OutputStream.Write($buffer, 0, $buffer.Length)
            $context.Response.Close()
        }
        Start-Sleep -Milliseconds 100 # Mengurangi beban CPU
    }
} finally {
    # 6. CLEANUP (Pembersihan)
    Write-Log "INFO: Menghentikan semua proses..."
    $listener.Stop()
    Stop-Job $tunnelJob
    Get-Process "cloudflared" -ErrorAction SilentlyContinue | Stop-Process -Force
    Write-Log "INFO: Bersih. Script selesai."
}
