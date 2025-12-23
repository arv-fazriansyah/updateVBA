# --- CONFIGURATION ---
$port = "8080"
$urlLocal = "http://127.0.0.1:$port"
$currentDir = $PSScriptRoot
if (-not $currentDir) { $currentDir = Get-Location }

$cfExe = Join-Path $currentDir "cloudflared.exe"
$logFile = Join-Path $currentDir "vba_webhook_log.txt"
$pathTargetTxt = Join-Path $currentDir "target_excel.txt" # File penyimpan path Excel

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Fungsi untuk mencatat log
function Write-Log($pesan) {
    $waktu = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "[$waktu] $pesan" | Out-File -FilePath $logFile -Append -Encoding UTF8
}

# Fungsi untuk mendapatkan objek Workbook yang spesifik
function Get-SpecificWorkbook {
    if (Test-Path $pathTargetTxt) {
        $fullPath = Get-Content $pathTargetTxt -Raw
        $fullPath = $fullPath.Trim()
        
        try {
            $excelApp = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
            foreach ($wb in $excelApp.Workbooks) {
                if ($wb.FullName -eq $fullPath) {
                    return $wb
                }
            }
        } catch {
            Write-Log "ERROR: Excel tidak ditemukan atau tidak merespon."
        }
    } else {
        Write-Log "WARNING: File target_excel.txt tidak ditemukan."
    }
    return $null
}

# 1. Download Cloudflared jika belum ada
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
    param($cfExe, $urlLocal, $logFile, $pathTargetTxt)
    
    $tempCfLog = $logFile.Replace(".txt", "_cf.tmp")
    $process = Start-Process -FilePath $cfExe -ArgumentList "tunnel --url $urlLocal --no-autoupdate" `
               -NoNewWindow -PassThru -RedirectStandardError $tempCfLog

    for ($i = 0; $i -lt 60; $i++) {
        if (Test-Path $tempCfLog) {
            $content = Get-Content $tempCfLog -ErrorAction SilentlyContinue
            $urlLine = $content | Select-String -Pattern "https://[a-z0-9-]+\.trycloudflare\.com" | Select-Object -First 1
            
            if ($urlLine) {
                $urlPublik = $urlLine.Matches.Value
                
                # Cari workbook tujuan dari job background
                if (Test-Path $pathTargetTxt) {
                    $fullPath = (Get-Content $pathTargetTxt -Raw).Trim()
                    try {
                        $excelApp = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
                        foreach ($wb in $excelApp.Workbooks) {
                            if ($wb.FullName -eq $fullPath) {
                                $wb.Sheets("DEV").Range("F10").Value = $urlPublik
                                $excelApp.Run("TampilkanToast", "Tunnel Aktif", "Koneksi Berhasil", "")
                                break
                            }
                        }
                    } catch { }
                }
                break
            }
        }
        Start-Sleep -Seconds 1
    }
    if (Test-Path $tempCfLog) { Remove-Item $tempCfLog -Force }
}
Start-Job -ScriptBlock $jobScript -ArgumentList $cfExe, $urlLocal, $logFile, $pathTargetTxt

# 4. Loop Utama untuk menerima Webhook
Write-Log "INFO: Menunggu pesan inbound..."
try {
    while ($listener.IsListening) {
        $context = $listener.GetContext()
        $pesan = $context.Request.QueryString["teks"]
        
        if ($pesan) {
            Write-Log "WEBHOOK: Pesan diterima -> $pesan"
            $targetWb = Get-SpecificWorkbook
            if ($targetWb) {
                try {
                    # Jalankan macro pada workbook yang spesifik
                    $targetWb.Parent.Run("TampilkanToast", "", $pesan, "")
                } catch {
                    Write-Log "ERROR: Gagal menjalankan Macro."
                }
            } else {
                Write-Log "ERROR: Excel tujuan tidak ditemukan."
            }
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
