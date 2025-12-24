# ==========================================
# --- KONFIGURASI UTAMA (Ubah di Sini) ---
# ==========================================
$port = 8080                                 # Port awal (akan auto-increment jika sibuk)
$cfDownloadUrl = 'https://github.com/cloudflare/cloudflared/releases/latest/download/cloudflared-windows-amd64.exe'

# Lokasi File & Folder
$currentDir = if ($PSScriptRoot) { $PSScriptRoot } else { Get-Location }
$parentDir  = Split-Path $currentDir -Parent
$cfExe      = Join-Path $parentDir "cloudflared.exe"
$logFile    = Join-Path $currentDir "log.txt"
$pathTargetTxt = Join-Path $currentDir "target.txt"
$pidPath    = Join-Path $currentDir "pid.txt"
$portPath   = Join-Path $currentDir "port.txt"
$tempCfLog  = Join-Path $currentDir "cf.tmp"

# ==========================================
# --- LOGIKA INTERNAL (Jangan Diubah) ---
# ==========================================

$startTime = Get-Date
$pesanTerhitung = 0
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

function Write-Log($pesan) {
    $waktu = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "[$waktu] $pesan" | Out-File -FilePath $logFile -Append -Encoding UTF8
}

function Kirim-Ke-Excel($judul, $isi) {
    try {
        if (Test-Path $pathTargetTxt) {
            $targetPath = (Get-Content $pathTargetTxt -Raw).Trim()
            $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
            foreach ($wb in $excel.Workbooks) {
                if ($wb.FullName -eq $targetPath) {
                    # HARDCODED: Nama Macro
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

# 1. Persiapan Cloudflared & Firewall
if (-not (Test-Path $cfExe)) {
    Write-Log "INFO: Mendownload cloudflared..."
    try {
        Invoke-WebRequest -Uri $cfDownloadUrl -OutFile $cfExe -ErrorAction Stop
        Unblock-File -Path $cfExe
    } catch {
        Write-Log "ERROR: Gagal download: $($_.Exception.Message)"; exit
    }
}

try {
    $principal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    if ($principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        if (-not (Get-NetFirewallRule -DisplayName "Allow Cloudflared" -ErrorAction SilentlyContinue)) {
            New-NetFirewallRule -DisplayName "Allow Cloudflared" -Direction Inbound -Program $cfExe -Action Allow -ErrorAction Stop
        }
    }
} catch { Write-Log "WARN: Gagal setup Firewall." }

# 2. Start Listener
$listener = New-Object System.Net.HttpListener
$berhasilStatus = $false

while (-not $berhasilStatus -and $port -lt 8100) {
    try {
        $urlLocal = "http://127.0.0.1:$port/"
        $listener.Prefixes.Clear()
        $listener.Prefixes.Add($urlLocal)
        $listener.Start()
        $berhasilStatus = $true
        $port | Out-File -FilePath $portPath -Encoding ASCII
        Write-Log "INFO: Listener aktif di $urlLocal"
    } catch {
        $port++
    }
}

# 3. Jalankan Cloudflare Tunnel
$cfProc = Start-Process -FilePath $cfExe -ArgumentList "tunnel --url $urlLocal --no-autoupdate --grace-period 1s" `
          -NoNewWindow -PassThru -RedirectStandardError $tempCfLog
$cfProc.Id | Out-File -FilePath $pidPath -Encoding ASCII

# Job untuk Update URL ke Excel (HARDCODED Sheet & Range)
$job = Start-Job -ScriptBlock {
    param($tempCfLog, $pathTargetTxt)
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
                            # HARDCODED: Nama Sheet "DEV" dan Range "F10"
                            $wb.Sheets("DEV").Range("F10").Value = $urlPublik
                            break
                        }
                    }
                } catch {}
                break
            }
        }
        Start-Sleep -Seconds 1
    }
} -ArgumentList $tempCfLog, $pathTargetTxt

# 4. Loop Utama
try {
    while ($listener.IsListening) {
        $context = $listener.GetContext(); $req = $context.Request; $res = $context.Response
        $path = $req.Url.LocalPath.ToLower(); $stopLoop = $false

        switch ($path) {
            "/"        { $responTeks = "OK" }
            "/stop"    { $responTeks = "STOPPING"; $stopLoop = $true }
            "/status"  { $responTeks = "Pesan Terkirim: $pesanTerhitung" }
            "/pesan"   {
                $p = $req.QueryString["teks"]; $j = $req.QueryString["judul"]
                if ($p -or $j) {
                    $pesanTerhitung++
                    $success = Kirim-Ke-Excel ($j ? $j : "") ($p ? $p : "")
                    $responTeks = $success ? "Diterima Excel" : "Excel Sibuk"
                }
            }
            default    { $responTeks = "Endpoint tidak ditemukan" }
        }

        $buffer = [System.Text.Encoding]::UTF8.GetBytes($responTeks)
        $res.ContentLength64 = $buffer.Length
        $res.OutputStream.Write($buffer, 0, $buffer.Length)
        $res.Close()
        if ($stopLoop) { break }
    }
} finally {
    Write-Log "INFO: Cleanup..."
    if ($listener) { $listener.Stop(); $listener.Close() }
    if (Test-Path $pidPath) {
        $savedPid = (Get-Content $pidPath -Raw).Trim()
        Stop-Process -Id $savedPid -Force -ErrorAction SilentlyContinue
        Remove-Item $pidPath
    }
    Get-Job | Remove-Job -Force
}
