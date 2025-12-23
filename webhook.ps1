# --- CONFIGURATION ---
$port = "8080"
$urlLocal = "http://127.0.0.1:$port"
$currentDir = $PSScriptRoot
if (-not $currentDir) { $currentDir = Get-Location }

$cfExe = Join-Path $currentDir "cloudflared.exe"
$logFile = Join-Path $currentDir "vba_webhook_log.txt"
$pathTargetTxt = Join-Path $currentDir "target_excel.txt"

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
$job = Start-Job -ScriptBlock {
    param($cfExe, $urlLocal, $logFile, $pathTargetTxt)
    $tempCfLog = $logFile.Replace(".txt", "_cf.tmp")
    
    # Jalankan tunnel
    Start-Process -FilePath $cfExe -ArgumentList "tunnel --url $urlLocal --no-autoupdate" `
                  -NoNewWindow -PassThru -RedirectStandardError $tempCfLog

    # Cari URL publik
    for ($i = 0; $i -lt 60; $i++) {
        if (Test-Path $tempCfLog) {
            $content = Get-Content $tempCfLog -ErrorAction SilentlyContinue
            $urlLine = $content | Select-String -Pattern "https://[a-z0-9-]+\.trycloudflare\.com" | Select-Object -First 1
            if ($urlLine) {
                $urlPublik = $urlLine.Matches.Value
                try {
                    # Ambil path excel dari file txt
                    $fullPath = (Get-Content $pathTargetTxt -Raw).Trim()
                    $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
                    foreach ($wb in $excel.Workbooks) {
                        if ($wb.FullName -eq $fullPath) {
                            $wb.Sheets("DEV").Range("F10").Value = $urlPublik
                            $excel.Run("TampilkanToast", "Tunnel Aktif", "Koneksi Berhasil", "")
                            break
                        }
                    }
                } catch { }
                break
            }
        }
        Start-Sleep -Seconds 1
    }
    if (Test-Path $tempCfLog) { Remove-Item $tempCfLog -Force }
} -ArgumentList $cfExe, $urlLocal, $logFile, $pathTargetTxt

# 4. Loop Utama
Write-Log "INFO: Menunggu pesan inbound..."
try {
    while ($listener.IsListening) {
        $context = $listener.GetContext()
        $request = $context.Request
        $response = $context.Response
        
        # --- CEK ENDPOINT STOP ---
        if ($request.Url.LocalPath -eq "/stop") {
            Write-Log "INFO: Perintah STOP diterima."
            $buffer = [System.Text.Encoding]::UTF8.GetBytes("STOPPING")
            $response.ContentLength64 = $buffer.Length
            $response.OutputStream.Write($buffer, 0, $buffer.Length)
            $response.Close()
            break # Keluar dari loop untuk trigger 'finally'
        }

        # --- LOGIKA WEBHOOK BIASA ---
        $pesan = $request.QueryString["teks"]
        if ($pesan) {
            Write-Log "WEBHOOK: Pesan diterima -> $pesan"
            try {
                $fullPath = (Get-Content $pathTargetTxt -Raw).Trim()
                $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
                foreach ($wb in $excel.Workbooks) {
                    if ($wb.FullName -eq $fullPath) {
                        $excel.Run("TampilkanToast", "Pesan Masuk", $pesan, "")
                        break
                    }
                }
            } catch { }
        }
        
        $buffer = [System.Text.Encoding]::UTF8.GetBytes("OK")
        $response.ContentLength64 = $buffer.Length
        $response.OutputStream.Write($buffer, 0, $buffer.Length)
        $response.Close()
    }
} finally {
    # MEMBERSIHKAN PROSES
    Write-Log "INFO: Mematikan Listener..."
    $listener.Stop()
    $listener.Close()
    
    # Hentikan Cloudflare Job
    Get-Job | Stop-Job
    Get-Job | Remove-Job
    Stop-Process -Name "cloudflared" -Force -ErrorAction SilentlyContinue
    
    Write-Log "INFO: Script dihentikan secara sempurna."
}
