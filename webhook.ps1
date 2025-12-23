# --- CONFIGURATION ---
$port = "8080"
$urlLocal = "http://127.0.0.1:$port"
$folderARB = Join-Path $env:TEMP "ARB2026"
$cfExe = Join-Path $folderARB "cloudflared.exe"
$logFile = Join-Path $folderARB "vba_webhook_log.txt"

# Memastikan protokol keamanan untuk download (TLS 1.2)
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Fungsi Log Tunggal
function Write-Log($pesan) {
    $waktu = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "[$waktu] $pesan" | Out-File -FilePath $logFile -Append -Encoding UTF8
}

# 1. Download Cloudflared jika belum ada
if (-not (Test-Path $cfExe)) {
    Write-Log "INFO: Cloudflared tidak ditemukan. Memulai download..."
    try {
        Invoke-WebRequest -Uri 'https://github.com/cloudflare/cloudflared/releases/latest/download/cloudflared-windows-amd64.exe' -OutFile $cfExe -ErrorAction Stop
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
    Write-Log "ERROR: GAGAL menjalankan listener: $($_.Exception.Message)"
    exit
}

# 3. Jalankan Cloudflare Tunnel (Background Job)
$jobScript = {
    param($cfExe, $urlLocal, $logFile)
    
    # Jalankan tunnel dan arahkan SEMUA output (stderr & stdout) langsung ke satu file log
    # --no-autoupdate ditambahkan untuk mengurangi baris log yang tidak perlu
    $process = Start-Process -FilePath $cfExe -ArgumentList "tunnel --url $urlLocal --no-autoupdate" `
               -NoNewWindow -PassThru -RedirectStandardError $logFile -RedirectStandardOutput $logFile

    $found = $false
    $attempts = 0
    while (-not $found -and $attempts -lt 60) {
        if (Test-Path $logFile) {
            # Membaca log yang sedang diisi oleh cloudflared
            $content = Get-Content $logFile -ErrorAction SilentlyContinue
            foreach ($line in $content) {
                if ($line -match 'https://[a-z0-9-]+\.trycloudflare\.com') {
                    $urlPublik = $matches[0]
                    try {
                        $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
                        $excel.Sheets("DEV").Range("F10").Value = $urlPublik
                        $excel.Run("TampilkanToast", "Tunnel Aktif", "Koneksi Berhasil", "")
                        $found = $true
                    } catch { 
                        Start-Sleep -Seconds 1 
                    }
                }
            }
        }
        $attempts++
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
            } catch {
                Write-Log "WARN: Excel sibuk."
            }
        }
        
        $buffer = [System.Text.Encoding]::UTF8.GetBytes("OK")
        $context.Response.ContentLength64 = $buffer.Length
        $context.Response.OutputStream.Write($buffer, 0, $buffer.Length)
        $context.Response.Close()
    }
} catch {
    Write-Log "ERROR: Loop utama terhenti: $($_.Exception.Message)"
} finally {
    $listener.Stop()
    Write-Log "INFO: Script dihentikan."
}
