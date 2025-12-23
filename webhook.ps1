# --- CONFIGURATION ---
$port = "8080"
$urlLocal = "http://127.0.0.1:$port"
$folderARB = Join-Path $env:TEMP "ARB2026"
$cfExe = Join-Path $folderARB "cloudflared.exe"
$logFile = Join-Path $folderARB "vba_webhook_log.txt"

# --- BAGIAN PEMBUATAN FOLDER SUDAH DIHAPUS (DITANGANI VBA) ---

# 1. Download Cloudflared jika belum ada
if (-not (Test-Path $cfExe)) {
    Write-Output "Downloading cloudflared..."
    (New-Object Net.WebClient).DownloadFile('https://github.com/cloudflare/cloudflared/releases/latest/download/cloudflared-windows-amd64.exe', $cfExe)
}

# Fungsi pembantu untuk menulis log
function Write-Log($pesan) {
    $waktu = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "[$waktu] $pesan" | Out-File -FilePath $logFile -Append
}

Write-Log "Script dimulai."

# 2. Jalankan Listener HTTP (Background)
$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("$urlLocal/")
$listener.Start()
Write-Log "HTTP Listener aktif di $urlLocal"

# 3. Jalankan Cloudflare Tunnel & Kirim URL ke Excel
Start-Job -ScriptBlock {
    param($cfExe, $urlLocal, $logFile)
    
    # Fungsi log di dalam job karena scope berbeda
    function Job-Log($txt) {
        $t = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        "[$t] [Tunnel] $txt" | Out-File -FilePath $logFile -Append
    }

    & $cfExe tunnel --url $urlLocal 2>&1 | ForEach-Object {
        $line = $_.ToString()
        Job-Log $line # Catat semua aktivitas tunnel ke file log
        
        if ($line -match 'https://[a-z0-9-]+\.trycloudflare\.com') {
            $urlPublik = $matches[0]
            try {
                $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
                $excel.Sheets("DEV").Range("F10").Value = $urlPublik
                $excel.Run("TampilkanToast", "Tunnel Aktif", "URL: $urlPublik", "")
            } catch {
                Job-Log "Gagal koneksi ke Excel: $($_.Exception.Message)"
            }
        }
    }
} -ArgumentList $cfExe, $urlLocal, $logFile

# 4. Loop Listener Utama
try {
    Write-Log "Menunggu pesan inbound..."
    while ($listener.IsListening) {
        $context = $listener.GetContext()
        $pesan = $context.Request.QueryString["teks"]
        
        if ($pesan) {
            Write-Log "Pesan diterima: $pesan"
            try {
                $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
                $excel.Run("TampilkanToast", "", $pesan, "")
            } catch {
                Write-Log "Gagal kirim pesan ke Excel."
            }
        }
        
        $buffer = [System.Text.Encoding]::UTF8.GetBytes("OK")
        $context.Response.OutputStream.Write($buffer, 0, $buffer.Length)
        $context.Response.Close()
    }
} finally {
    Write-Log "Script dihentikan."
    $listener.Stop()
}
