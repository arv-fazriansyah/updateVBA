# --- CONFIGURATION ---
$port = "8080"
$urlLocal = "http://127.0.0.1:$port"
$folderARB = Join-Path $env:TEMP "ARB2026"
$cfExe = Join-Path $folderARB "cloudflared.exe"
$logFile = Join-Path $folderARB "vba_webhook_log.txt"

# Memastikan protokol keamanan untuk download (TLS 1.2)
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Fungsi Log yang efisien
function Write-Log($pesan) {
    $waktu = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "[$waktu] $pesan" | Out-File -FilePath $logFile -Append -Encoding UTF8
}

# 1. Download Cloudflared jika belum ada
if (-not (Test-Path $cfExe)) {
    Write-Log "Cloudflared tidak ditemukan. Memulai download..."
    try {
        Invoke-WebRequest -Uri 'https://github.com/cloudflare/cloudflared/releases/latest/download/cloudflared-windows-amd64.exe' -OutFile $cfExe -ErrorAction Stop
        Write-Log "Download selesai."
    } catch {
        Write-Log "GAGAL Download: $($_.Exception.Message)"
        exit
    }
}

# 2. Jalankan Listener HTTP
$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("$urlLocal/")
try {
    $listener.Start()
    Write-Log "Listener aktif di $urlLocal"
} catch {
    Write-Log "Gagal menjalankan listener: $($_.Exception.Message)"
    exit
}

# 3. Jalankan Cloudflare Tunnel (Background Job)
$jobScript = {
    param($cfExe, $urlLocal, $logFile)
    
    function Job-Log($txt) {
        $t = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        "[$t] [Tunnel] $txt" | Out-File -FilePath $logFile -Append -Encoding UTF8
    }

    # Jalankan tunnel
    $process = Start-Process -FilePath $cfExe -ArgumentList "tunnel --url $urlLocal" `
               -NoNewWindow -PassThru -RedirectStandardError $logFile.Replace(".txt", "_error.log") -RedirectStandardOutput $logFile.Replace(".txt", "_tunnel.log")

    # Loop pengecekan URL dari file log yang dihasilkan cloudflared
    # Cloudflared menulis URL ke stderr, kita monitor log file saja untuk mencari URL
    $found = $false
    while (-not $found) {
        if (Test-Path $logFile.Replace(".txt", "_error.log")) {
            $content = Get-Content $logFile.Replace(".txt", "_error.log") -ErrorAction SilentlyContinue
            foreach ($line in $content) {
                if ($line -match 'https://[a-z0-9-]+\.trycloudflare\.com') {
                    $urlPublik = $matches[0]
                    try {
                        $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
                        $excel.Sheets("DEV").Range("F10").Value = $urlPublik
                        # $excel.Run("TampilkanToast", "Tunnel Aktif", "URL: $urlPublik", "")
                        $found = $true
                        Job-Log "URL Berhasil dikirim ke Excel: $urlPublik"
                    } catch { Start-Sleep -Seconds 1 }
                }
            }
        }
        Start-Sleep -Seconds 2
    }
}
Start-Job -ScriptBlock $jobScript -ArgumentList $cfExe, $urlLocal, $logFile

# 4. Loop Utama (Menangani Request Masuk)
Write-Log "Menunggu pesan inbound..."
try {
    while ($listener.IsListening) {
        $context = $listener.GetContext()
        $request = $context.Request
        $pesan = $request.QueryString["teks"]
        
        if ($pesan) {
            Write-Log "Pesan diterima: $pesan"
            try {
                $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
                $excel.Run("TampilkanToast", "APLIKASI ARB 2026", $pesan, "")
            } catch {
                Write-Log "Excel sibuk atau tidak ditemukan."
            }
        }
        
        # Respon ke pengirim
        $buffer = [System.Text.Encoding]::UTF8.GetBytes("OK")
        $context.Response.ContentLength64 = $buffer.Length
        $context.Response.OutputStream.Write($buffer, 0, $buffer.Length)
        $context.Response.Close()
    }
} catch {
    Write-Log "Error pada Loop Utama: $($_.Exception.Message)"
} finally {
    $listener.Stop()
    Write-Log "Script dihentikan secara aman."
}
