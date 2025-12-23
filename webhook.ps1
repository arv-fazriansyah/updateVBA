# --- CONFIGURATION ---
$port = "8080"
$urlLocal = "http://127.0.0.1:$port"

# Mengambil lokasi folder tempat script ini disimpan secara otomatis
$currentDir = $PSScriptRoot

# Jika script belum disimpan (running in memory), gunakan folder kerja saat ini
if (-not $currentDir) { $currentDir = Get-Location }

$cfExe = Join-Path $currentDir "cloudflared.exe"
$logFile = Join-Path $currentDir "vba_webhook_log.txt"

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

function Write-Log($pesan) {
    $waktu = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "[$waktu] $pesan" | Out-File -FilePath $logFile -Append -Encoding UTF8
}

# 1. Download Cloudflared jika belum ada di folder saat ini
if (-not (Test-Path $cfExe)) {
    Write-Log "INFO: Memulai download cloudflared ke $currentDir..."
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
    
    # Jalankan tunnel tanpa jendela baru
    $process = Start-Process -FilePath $cfExe -ArgumentList "tunnel --url $urlLocal --no-autoupdate" `
               -NoNewWindow -PassThru -RedirectStandardError $tempCfLog

    $found = $false
    for ($i = 0; $i -lt 60; $i++) {
        if (Test-Path $tempCfLog) {
            $content = Get-Content $tempCfLog -ErrorAction SilentlyContinue
            # Mencari URL publik yang dihasilkan Cloudflare
            $urlLine = $content | Select-String -Pattern "https://[a-z0-9-]+\.trycloudflare\.com" | Select-Object -First 1
            
            if ($urlLine) {
                $urlPublik = $urlLine.Matches.Value
                try {
                    $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
                    $excel.Sheets("DEV").Range("F10").Value = $urlPublik
                    $excel.Run("TampilkanToast", "Tunnel Aktif", "Koneksi Berhasil", "")
                    
                    $t = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                    "[$t] TUNNEL: Terhubung -> $urlPublik" | Out-File -FilePath $logFile -Append -Encoding UTF8
                    
                    $found = $true
                    break
                } catch { }
            }
        }
        Start-Sleep -Seconds 1
    }
    if (Test-Path $tempCfLog) { Remove-Item $tempCfLog -Force }
}
Start-Job -ScriptBlock $jobScript -ArgumentList $cfExe, $urlLocal, $logFile

# 4. Loop Utama untuk menerima Webhook
Write-Log "INFO: Menunggu pesan inbound..."
try {
    while ($listener.IsListening) {
        $context = $listener.GetContext()
        $pesan = $context.Request.QueryString["teks"]
        
        if ($pesan) {
            Write-Log "WEBHOOK: Pesan diterima -> $pesan"
            try {
                $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
                $excel.Run("TampilkanToast", "Pesan Masuk", $pesan, "")
            } catch { }
        }
        
        # Kirim respon balik ke pengirim webhook
        $buffer = [System.Text.Encoding]::UTF8.GetBytes("OK")
        $context.Response.ContentLength64 = $buffer.Length
        $context.Response.OutputStream.Write($buffer, 0, $buffer.Length)
        $context.Response.Close()
    }
} finally {
    $listener.Stop()
    Write-Log "INFO: Script dihentikan."
}
