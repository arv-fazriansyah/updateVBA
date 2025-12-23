# --- CONFIGURATION ---
$port = "8080"
$urlLocal = "http://127.0.0.1:$port"
$folderARB = Join-Path $env:TEMP "ARB2026"
$cfExe = Join-Path $folderARB "cloudflared.exe"
$logFile = Join-Path $folderARB "vba_webhook_log.txt"

# --- BAGIAN PEMBUATAN FOLDER SUDAH DIHAPUS (SUDAH DITANGANI VBA) ---

# 1. Download Cloudflared jika belum ada
if (-not (Test-Path $cfExe)) {
    (New-Object Net.WebClient).DownloadFile('https://github.com/cloudflare/cloudflared/releases/latest/download/cloudflared-windows-amd64.exe', $cfExe)
}

# 2. Jalankan Listener HTTP (Background)
$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("$urlLocal/")
$listener.Start()

# 3. Jalankan Cloudflare Tunnel & Kirim URL ke Excel
Start-Job -ScriptBlock {
    param($cfExe, $urlLocal)
    & $cfExe tunnel --url $urlLocal 2>&1 | ForEach-Object {
        if ($_ -match 'https://[a-z0-9-]+\.trycloudflare\.com') {
            $urlPublik = $matches[0]
            $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
            $excel.Sheets("DEV").Range("F10").Value = $urlPublik
            $excel.Run("TampilkanToast", "Tunnel Aktif", "URL: $urlPublik", "")
        }
    }
} -ArgumentList $cfExe, $urlLocal

# 4. Loop Listener Utama
try {
    while ($listener.IsListening) {
        $context = $listener.GetContext()
        $pesan = $context.Request.QueryString["teks"]
        if ($pesan) {
            $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
            $excel.Run("TampilkanToast", "Webhook Inbound", $pesan, "")
        }
        $buffer = [System.Text.Encoding]::UTF8.GetBytes("OK")
        $context.Response.OutputStream.Write($buffer, 0, $buffer.Length)
        $context.Response.Close()
    }
} finally {
    $listener.Stop()
}
