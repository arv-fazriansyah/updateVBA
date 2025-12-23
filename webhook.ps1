# --- CONFIGURATION ---
$port = "8080"
$urlLocal = "http://127.0.0.1:$port"
$currentDir = $PSScriptRoot
if (-not $currentDir) { $currentDir = Get-Location }

$cfExe = Join-Path $currentDir "cloudflared.exe"
$logFile = Join-Path $currentDir "vba_webhook_log.txt"

function Write-Log($pesan, $level = "INFO") {
    $waktu = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$waktu] [$level] $pesan"
    Write-Host $logEntry -ForegroundColor (Switch($level) { "ERROR" {"Red"}; "WARN" {"Yellow"}; Default {"White"} })
    $logEntry | Out-File -FilePath $logFile -Append -Encoding UTF8
}

# --- 1. CEK KONEKSI INTERNET ---
Write-Log "Memeriksa koneksi internet..."
try {
    $ping = Test-Connection -ComputerName google.com -Count 1 -ErrorAction Stop
    Write-Log "Koneksi internet tersedia."
} catch {
    Write-Log "Tidak ada koneksi internet! Script dihentikan." "ERROR"
    exit
}

# --- 2. CEK KONEKSI EXCEL ---
$excel = $null
try {
    $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
    Write-Log "Excel terdeteksi dan siap."
} catch {
    Write-Log "Excel tidak terbuka. Script akan tetap berjalan, tapi data tidak akan terkirim." "WARN"
}

# --- 3. DOWNLOAD & VALIDASI CLOUDFLARED ---
if (-not (Test-Path $cfExe)) {
    Write-Log "Mendownload cloudflared..."
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        Invoke-WebRequest -Uri 'https://github.com/cloudflare/cloudflared/releases/latest/download/cloudflared-windows-amd64.exe' -OutFile $cfExe -ErrorAction Stop
        Unblock-File -Path $cfExe
        Write-Log "Download selesai."
    } catch {
        Write-Log "Gagal download: $($_.Exception.Message)" "ERROR"
        exit
    }
}

# --- 4. JALANKAN LISTENER ---
$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("$urlLocal/")
try {
    $listener.Start()
} catch {
    Write-Log "Gagal memulai listener! Port $port mungkin dipakai aplikasi lain." "ERROR"
    exit
}

# --- 5. JALANKAN TUNNEL (JOB) ---
$jobScript = {
    param($cfExe, $urlLocal, $logFile)
    $tempCfLog = $logFile.Replace(".txt", "_cf.tmp")
    Start-Process -FilePath $cfExe -ArgumentList "tunnel --url $urlLocal --no-autoupdate" `
                  -NoNewWindow -RedirectStandardError $tempCfLog
    
    # Loop mencari URL
    for ($i = 0; $i -lt 30; $i++) {
        if (Test-Path $tempCfLog) {
            $urlLine = Get-Content $tempCfLog | Select-String -Pattern "https://[a-z0-9-]+\.trycloudflare\.com" | Select-Object -First 1
            if ($urlLine) {
                $urlPublik = $urlLine.Matches.Value
                try {
                    $ex = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
                    $ex.Sheets("DEV").Range("F10").Value = $urlPublik
                    $ex.Run("TampilkanToast", "Tunnel Aktif", $urlPublik, "")
                } catch {}
                break
            }
        }
        Start-Sleep -Seconds 2
    }
}
$tunnelJob = Start-Job -ScriptBlock $jobScript -ArgumentList $cfExe, $urlLocal, $logFile

# --- 6. LOOP UTAMA ---
Write-Log "Sistem Aktif. Menunggu Webhook di port $port..."
try {
    while ($listener.IsListening) {
        $context = $listener.GetContext()
        $pesan = $context.Request.QueryString["teks"]
        
        if ($pesan) {
            Write-Log "Pesan masuk: $pesan"
            try {
                $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
                $excel.Run("TampilkanToast", "Pesan Masuk", $pesan, "")
            } catch {
                Write-Log "Gagal kirim ke Excel (Mungkin Excel sibuk/diedit user)." "WARN"
            }
        }
        
        $buffer = [System.Text.Encoding]::UTF8.GetBytes("OK")
        $context.Response.ContentLength64 = $buffer.Length
        $context.Response.OutputStream.Write($buffer, 0, $buffer.Length)
        $context.Response.Close()
    }
} finally {
    # --- 7. CLEANUP (SANGAT PENTING) ---
    Write-Log "Membersihkan proses sebelum keluar..."
    $listener.Stop()
    Stop-Job $tunnelJob
    Get-Process "cloudflared" -ErrorAction SilentlyContinue | Stop-Process -Force
    if (Test-Path $logFile.Replace(".txt", "_cf.tmp")) { Remove-Item $logFile.Replace(".txt", "_cf.tmp") -Force }
    Write-Log "Selesai."
}
