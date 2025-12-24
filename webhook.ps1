# --- KONFIGURASI ---
$port = 8080 
$currentDir = if ($PSScriptRoot) { $PSScriptRoot } else { Get-Location }

# Lokasi File
$parentDir = Split-Path $currentDir -Parent
$cfExe = Join-Path $parentDir "cloudflared.exe"
$logFile = Join-Path $currentDir "log.txt"
$pathTargetTxt = Join-Path $currentDir "target.txt"
$pidPath = Join-Path $currentDir "pid.txt"
$portPath = Join-Path $currentDir "port.txt"

$startTime = Get-Date
$pesanTerhitung = 0

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# --- FUNGSI ---

function Write-Log($pesan) {
    $waktu = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "[$waktu] $pesan" | Out-File -FilePath $logFile -Append -Encoding UTF8
}

function Kirim-Ke-Excel($judul, $isi) {
    if (-not (Test-Path $pathTargetTxt)) { return $false }
    try {
        $targetPath = (Get-Content $pathTargetTxt -Raw).Trim()
        # Menggunakan ComObject dengan cara yang lebih aman
        $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
        foreach ($wb in $excel.Workbooks) {
            if ($wb.FullName -eq $targetPath) {
                # Pastikan macro "TampilkanToast" ada di Excel Anda
                $excel.Run("TampilkanToast", $judul, $isi, "")
                return $true
            }
        }
    } catch { 
        return $false 
    }
    return $false
}

# --- PERSIAPAN ---

# 1. Download Cloudflared jika belum ada
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

# 2. Setup Listener (Cari Port Kosong)
$listener = New-Object System.Net.HttpListener
$statusTerhubung = $false

while (-not $statusTerhubung -and $port -lt 8100) {
    try {
        $urlLocal = "http://127.0.0.1:$port/"
        $listener.Prefixes.Clear()
        $listener.Prefixes.Add($urlLocal)
        $listener.Start()
        $statusTerhubung = $true
        $port | Out-File -FilePath $portPath -Encoding ASCII
        Write-Log "INFO: Listener aktif di $urlLocal"
    } catch {
        $port++
    }
}

if (-not $listener.IsListening) {
    Write-Log "ERROR: Tidak ada port tersedia."
    exit
}

# 3. Jalankan Cloudflare Tunnel
$tempCfLog = Join-Path $currentDir "cf.tmp"
$cfProc = Start-Process -FilePath $cfExe -ArgumentList "tunnel --url $urlLocal --no-autoupdate --grace-period 1s" `
          -NoNewWindow -PassThru -RedirectStandardError $tempCfLog

$cfPid = $cfProc.Id
$cfPid | Out-File -FilePath $pidPath -Encoding ASCII

# Job untuk mengambil URL publik dan memasukkannya ke Excel
$job = Start-Job -ScriptBlock {
    param($tempCfLog, $pathTargetTxt)
    $found = $false
    for ($i = 0; $i -lt 30; $i++) { # Tunggu maksimal 30 detik
        if (Test-Path $tempCfLog) {
            $content = Get-Content $tempCfLog -Raw
            if ($content -match "https://[a-z0-9-]+\.trycloudflare\.com") {
                $urlPublik = $matches[0]
                try {
                    $targetPath = (Get-Content $pathTargetTxt -Raw).Trim()
                    $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
                    foreach ($wb in $excel.Workbooks) {
                        if ($wb.FullName -eq $targetPath) {
                            $wb.Sheets("DEV").Range("F10").Value = $urlPublik
                            $found = $true; break
                        }
                    }
                } catch {}
            }
        }
        if ($found) { break }
        Start-Sleep -Seconds 1
    }
} -ArgumentList $tempCfLog, $pathTargetTxt

# --- LOOP UTAMA ---

try {
    Write-Log "INFO: Menunggu Webhook..."
    while ($listener.IsListening) {
        $context = $listener.GetContext()
        $req = $context.Request
        $res = $context.Response
        $path = $req.Url.LocalPath.ToLower()
        $responTeks = ""

        switch ($path) {
            "/"        { $responTeks = "Server Online" }
            "/ping"    { $responTeks = "PONG" }
            "/status"  { 
                $uptime = (Get-Date) - $startTime
                $responTeks = "Uptime: $($uptime.ToString('hh\:mm\:ss')) | Pesan: $pesanTerhitung" 
            }
            "/pesan"   {
                $judul = $req.QueryString["judul"]
                $teks = $req.QueryString["teks"]
                if ($judul -or $teks) {
                    $pesanTerhitung++
                    $success = Kirim-Ke-Excel $judul $teks
                    $responTeks = $success ? "OK: Terkirim ke Excel" : "Error: Excel tidak siap"
                    Write-Log "WEBHOOK: $judul - $teks ($responTeks)"
                } else {
                    $responTeks = "Error: Parameter kurang"
                }
            }
            "/stop"    {
                $responTeks = "Server Berhenti..."
                $buffer = [System.Text.Encoding]::UTF8.GetBytes($responTeks)
                $res.ContentLength64 = $buffer.Length
                $res.OutputStream.Write($buffer, 0, $buffer.Length)
                $res.Close()
                break # Keluar dari loop
            }
            default    { $responTeks = "404: Not Found" }
        }

        if ($path -ne "/stop") {
            $buffer = [System.Text.Encoding]::UTF8.GetBytes($responTeks)
            $res.ContentLength64 = $buffer.Length
            $res.OutputStream.Write($buffer, 0, $buffer.Length)
            $res.Close()
        }
    }
} finally {
    Write-Log "INFO: Pembersihan sistem..."
    # Tutup Listener
    if ($null -ne $listener) { $listener.Stop(); $listener.Close() }
    
    # Matikan Cloudflared berdasarkan PID yang disimpan
    if (Test-Path $pidPath) {
        $savedPid = Get-Content $pidPath -Raw
        Stop-Process -Id $savedPid -Force -ErrorAction SilentlyContinue
        Remove-Item $pidPath
    }

    # Hapus file sampah
    Remove-Item $tempCfLog -ErrorAction SilentlyContinue
    Get-Job | Remove-Job -Force
    Write-Log "INFO: Selesai."
}
