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
    Write-Log "INFO: Cloudflared tidak ditemukan. Memulai download..."
    try {
        Invoke-WebRequest -Uri 'https://github.com/cloudflare/cloudflared/releases/latest/download/cloudflared-windows-amd64.exe' -OutFile $cfExe -ErrorAction Stop
        
        # --- PERBAIKAN 1: Unblock File & Jeda ---
        Unblock-File -Path $cfExe  # Lepas proteksi keamanan Windows
        Start-Sleep -Seconds 2     # Beri waktu OS mencatat file baru
        
        Write-Log "INFO: Download selesai dan file telah di-unblock."
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
    Write-Log "ERROR: GAGAL menjalankan listener: $($_.Exception.Message)"
    exit
}

# 3. Jalankan Cloudflare Tunnel (Background Job)
$jobScript = {
    param($cfExe, $urlLocal, $logFile)
    
    # --- PERBAIKAN 2: Gunakan File Log Terpisah Sementara ---
    # Jika cloudflared langsung menulis ke log utama saat log utama sedang ditulis oleh script, 
    # dia akan gagal start. Kita gunakan temp log lalu satukan.
    $tempCfLog = $logFile.Replace(".txt", "_cf.tmp")

    $process = Start-Process -FilePath $cfExe -ArgumentList "tunnel --url $urlLocal --no-autoupdate" `
               -NoNewWindow -PassThru -RedirectStandardError $tempCfLog

    $found = $false
    for ($i = 0; $i -lt 60; $i++) {
        if (Test-Path $tempCfLog) {
            $content = Get-Content $tempCfLog -ErrorAction SilentlyContinue
            foreach ($line in $content) {
                # Pindahkan isi temp log ke log utama (untuk monitoring)
                $line | Out-File -FilePath $logFile -Append -Encoding UTF8
                
                if ($line -match 'https://[a-z0-9-]+\.trycloudflare\.com') {
                    $urlPublik = $matches[0]
                    try {
                        $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
                        $excel.Sheets("DEV").Range("F10").Value = $urlPublik
                        $excel.Run("TampilkanToast", "Tunnel Aktif", "Koneksi Berhasil", "")
                        $found = $true
                        break
                    } catch { }
                }
            }
            if ($found) { break }
        }
        Start-Sleep -Seconds 1
    }
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
