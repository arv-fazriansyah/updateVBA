# --- KONFIGURASI ---
$port = "8080"
$urlLocal = "http://127.0.0.1:$port"
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
            # Mencoba menyambung ke aplikasi Excel yang sedang terbuka
            $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
            foreach ($wb in $excel.Workbooks) {
                if ($wb.FullName -eq $targetPath) {
                    # Memanggil Macro 'TampilkanToast' di Excel
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

# 1. Cek & Download Cloudflared jika belum ada
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

# 2. Inisialisasi Listener HTTP
$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("$urlLocal/")
try {
    $listener.Start()
    Write-Log "INFO: Listener aktif di $urlLocal"
} catch {
    Write-Log "ERROR: Port $port sudah digunakan aplikasi lain."
    exit
}

# 3. Jalankan Cloudflare Tunnel di Background
$job = Start-Job -ScriptBlock {
    param($cfExe, $urlLocal, $logFile, $pathTargetTxt)
    $tempCfLog = $logFile.Replace(".txt", "_cf.tmp")
    
    Start-Process -FilePath $cfExe -ArgumentList "tunnel --url $urlLocal --no-autoupdate" `
                  -NoNewWindow -PassThru -RedirectStandardError $tempCfLog

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
                            $excel.Run("TampilkanToast", "Tunnel Online", "URL: $urlPublik", "")
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

# 4. Loop Utama Penanganan Request
Write-Log "INFO: Siap menerima webhook..."
try {
    while ($listener.IsListening) {
        $context = $listener.GetContext()
        $req = $context.Request
        $res = $context.Response
        $path = $req.Url.LocalPath.ToLower()
        # Baris inisialisasi default dihapus agar lebih bersih
        $stopLoop = $false

        switch ($path) {
            "/" {
                $responTeks = "OK"
            }
            "/stop" {
                $responTeks = "STOPPING"
                Write-Log "INFO: Perintah STOP diterima."
                $stopLoop = $true
            }
            "/ping" {
                $responTeks = "PONG"
            }
            "/status" {
                $uptime = (Get-Date) - $startTime
                $responTeks = "Uptime: $($uptime.ToString('hh\:mm\:ss')) | Pesan: $pesanTerhitung"
            }
            "/pesan" {
                $pesan = $req.QueryString["teks"]
                $judul = $req.QueryString["judul"]
                
                if ($pesan) {
                    $pesanTerhitung++
                    if (-not $judul) { $judul = "Notifikasi" }
                    
                    Write-Log "WEBHOOK: [$judul] $pesan"
                    $success = Kirim-Ke-Excel $judul $pesan
                    $responTeks = if ($success) { "Diterima Excel" } else { "Excel Sedang Sibuk" }
                } else {
                    $responTeks = "Error: Parameter 'teks' diperlukan."
                }
            }
            default {
                $responTeks = "Error: Endpoint $path tidak tersedia."
            }
        }

        # Kirim Respon Balik ke Pengirim
        $buffer = [System.Text.Encoding]::UTF8.GetBytes($responTeks)
        $res.ContentLength64 = $buffer.Length
        $res.OutputStream.Write($buffer, 0, $buffer.Length)
        $res.Close()

        if ($stopLoop) { break }
    }
} finally {
    # Pembersihan saat script berhenti
    Write-Log "INFO: Menutup semua proses..."
    $listener.Stop()
    $listener.Close()
    Get-Job | Stop-Job | Remove-Job
    Stop-Process -Name "cloudflared" -Force -ErrorAction SilentlyContinue
    Write-Log "INFO: Selesai."
}
