# --- KONFIGURASI ---
$port = 8080  # Port awal
$currentDir = $PSScriptRoot
if (-not $currentDir) { $currentDir = Get-Location }

$cfExe = Join-Path $currentDir "cloudflared.exe"
$logFile = Join-Path $currentDir "log.txt"
$pathTargetTxt = Join-Path $currentDir "target.txt"
$startTime = Get-Date
$pesanTerhitung = 0

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Fungsi Log
function Write-Log($pesan) {
    $waktu = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "[$waktu] $pesan" | Out-File -FilePath $logFile -Append -Encoding UTF8
}

# Fungsi Helper untuk kirim data ke Excel
function Kirim-Ke-Excel($judul, $isi) {
    try {
        if (Test-Path $pathTargetTxt) {
            $targetPath = (Get-Content $pathTargetTxt -Raw).Trim()
            $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
            foreach ($wb in $excel.Workbooks) {
                if ($wb.FullName -eq $targetPath) {
                    $excel.Run("TampilkanToast", $judul, $isi, "")
                    return $true
                }
            }
        }
    } catch { 
        Write-Log "DEBUG: Gagal kirim ke Excel (Mungkin Excel sibuk/tutup)"
    }
    return $false
}

# 1. Cek & Download Cloudflared
if (-not (Test-Path $cfExe)) {
    Write-Log "INFO: Mendownload cloudflared..."
    try {
        Invoke-WebRequest -Uri 'https://github.com/cloudflare/cloudflared/releases/latest/download/cloudflared-windows-amd64.exe' -OutFile $cfExe -ErrorAction Stop
        Unblock-File -Path $cfExe
    } catch {
        Write-Log "ERROR: Gagal download: $($_.Exception.Message)"
        exit
    }
}

# 2. Inisialisasi Listener HTTP dengan Auto-Increment Port
$listener = New-Object System.Net.HttpListener
$berhasilStatus = $false

while (-not $berhasilStatus -and $port -lt 8100) {
    try {
        $urlLocal = "http://127.0.0.1:$port/"
        $listener.Prefixes.Clear()
        $listener.Prefixes.Add($urlLocal)
        $listener.Start()
        $berhasilStatus = $true
        $port | Out-File -FilePath (Join-Path $currentDir "port.txt") -Encoding ASCII
        Write-Log "INFO: Listener aktif di $urlLocal"
    } catch {
        Write-Log "WARN: Port $port sibuk, mencoba port $($port + 1)..."
        $port++
    }
}

if (-not $listener.IsListening) {
    Write-Log "ERROR: Tidak menemukan port kosong."
    exit
}

# 3. Jalankan Cloudflare Tunnel di Background
# Kita kirimkan $urlLocal yang sudah fix portnya ke background job
$job = Start-Job -ScriptBlock {
    param($cfExe, $urlLocal, $logFile, $pathTargetTxt)
    $tempCfLog = $logFile.Replace(".txt", "_cf.tmp")
    
    # Jalankan Cloudflare Tunnel dengan parameter otomatis
    Start-Process -FilePath $cfExe -ArgumentList "tunnel --url $urlLocal --no-autoupdate --grace-period 1s" `
                  -NoNewWindow -PassThru -RedirectStandardError $tempCfLog`

    for ($i = 0; $i -lt 60; $i++) {
        if (Test-Path $tempCfLog) {
            $content = Get-Content $tempCfLog -ErrorAction SilentlyContinue
            $urlLine = $content | Select-String -Pattern "https://[a-z0-9-]+\.trycloudflare\.com" | Select-Object -First 1
            if ($urlLine) {
                $urlPublik = $urlLine.Matches.Value
                try {
                    $targetPath = (Get-Content $pathTargetTxt -Raw).Trim()
                    $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
                    foreach ($wb in $excel.Workbooks) {
                        if ($wb.FullName -eq $targetPath) {
                            $wb.Sheets("DEV").Range("F10").Value = $urlPublik
                            #$excel.Run("TampilkanToast", "Tunnel Online", "URL: $urlPublik", "")
                            break
                        }
                    }
                } catch {}
                break
            }
        }
        Start-Sleep -Seconds 1
    }
    if (Test-Path $tempCfLog) { Remove-Item $tempCfLog -Force }
} -ArgumentList $cfExe, $urlLocal, $logFile, $pathTargetTxt

# 4. Loop Utama
try {
    while ($listener.IsListening) {
        $context = $listener.GetContext()
        $req = $context.Request
        $res = $context.Response
        $path = $req.Url.LocalPath.ToLower()
        $stopLoop = $false

        switch ($path) {
            "/" { $responTeks = "OK" }
            "/stop" {
                $responTeks = "STOPPING"
                Write-Log "INFO: Perintah STOP diterima."
                $stopLoop = $true
            }
            "/ping" { $responTeks = "PONG" }
            "/status" {
                $uptime = (Get-Date) - $startTime
                $responTeks = "Uptime: $($uptime.ToString('hh\:mm\:ss')) | Pesan: $pesanTerhitung"
            }
            "/pesan" {
                $pesan = $req.QueryString["teks"]
                $judul = $req.QueryString["judul"]
                
                # Cek apakah judul ATAU teks ada (salah satu boleh kosong)
                if ($pesan -ne $null -or $judul -ne $null) {
                    $pesanTerhitung++
                    
                    # Jika judul kosong, beri default "Notifikasi"
                    if (-not $judul) { $judul = "" }
                    # Jika pesan kosong, beri string kosong agar tidak error di Excel
                    if (-not $pesan) { $pesan = "" }
                    
                    Write-Log "WEBHOOK: [$judul] $pesan"
                    $success = Kirim-Ke-Excel $judul $pesan
                    $responTeks = if ($success) { "Diterima Excel" } else { "Excel Sedang Sibuk" }
                } else {
                    $responTeks = "Error: Parameter 'teks' atau 'judul' diperlukan."
                }
            }
            default { $responTeks = "Error: Endpoint $path tidak tersedia." }
        }

        $buffer = [System.Text.Encoding]::UTF8.GetBytes($responTeks)
        $res.ContentLength64 = $buffer.Length
        $res.OutputStream.Write($buffer, 0, $buffer.Length)
        $res.Close()

        if ($stopLoop) { break }
    }
} finally {
    Write-Log "INFO: Menutup semua proses..."
    $listener.Stop(); $listener.Close()
    Get-Job | Stop-Job | Remove-Job
    Stop-Process -Name "cloudflared" -Force -ErrorAction SilentlyContinue
    Write-Log "INFO: Selesai."
}
