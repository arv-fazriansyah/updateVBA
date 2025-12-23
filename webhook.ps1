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

# 1. Download Cloudflared jika belum ada
if (-not (Test-Path $cfExe)) {
    Write-Log "INFO: Memulai download cloudflared..."
    try {
        Invoke-WebRequest -Uri 'https://github.com/cloudflare/cloudflared/releases/latest/download/cloudflared-windows-amd64.exe' -OutFile $cfExe -ErrorAction Stop
        Unblock-File -Path $cfExe
        Start-Sleep -Seconds 2
        Write-Log "INFO: Download selesai."
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
    Write-Log "ERROR: GAGAL listener: $($_.Exception.Message)"
    exit
}

# 3. Jalankan Cloudflare Tunnel (Background Job)
$jobScript = {
    param($cfExe, $urlLocal, $logFile)
    
    $tempCfLog = $logFile.Replace(".txt", "_cf.tmp")
    # Jalankan tunnel, log awal masuk ke file temp saja agar tidak mengotori log utama
    $process = Start-Process -FilePath $cfExe -ArgumentList "tunnel --url $urlLocal --no-autoupdate" `
               -NoNewWindow -PassThru -RedirectStandardError $tempCfLog

    $found = $false
    for ($i = 0; $i -lt 60; $i++) {
        if (Test-Path $tempCfLog) {
            # Cari hanya baris yang mengandung URL trycloudflare
            $content = Get-Content $tempCfLog -ErrorAction SilentlyContinue
            $urlLine = $content | Select-String -Pattern "https://[a-z0-9-]+\.trycloudflare\.com" | Select-Object -First 1
            
            if ($urlLine) {
                $urlPublik = $urlLine.Matches.Value
                try {
                    $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
                    $excel.Sheets("DEV").Range("F10").Value = $urlPublik
                    $excel.Run("TampilkanToast", "Tunnel Aktif", "Koneksi Berhasil", "")
                    
                    # Tulis ke log utama HANYA saat URL ditemukan (sekali saja)
                    $t = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                    "[$t] TUNNEL: Terhubung -> $urlPublik" | Out-File -FilePath $logFile -Append -Encoding UTF8
                    
                    $found = $true
                    break
                } catch { }
            }
        }
        Start-Sleep -Seconds 1
    }
    # Hapus file temp log setelah selesai
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
            } catch { }
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
