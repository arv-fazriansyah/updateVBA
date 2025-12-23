# --- CONFIGURATION ---
$port = "8080"
$urlLocal = "http://127.0.0.1:$port"
$folderARB = Join-Path $env:TEMP "ARB2026"
$cfExe = Join-Path $folderARB "cloudflared.exe"
$logFile = Join-Path $folderARB "vba_webhook_log.txt"

# Pastikan folder ada
if (-not (Test-Path $folderARB)) { New-Item -ItemType Directory -Path $folderARB }

# 1. Download Cloudflared jika belum ada
if (-not (Test-Path $cfExe)) {
    (New-Object Net.WebClient).DownloadFile('https://github.com/cloudflare/cloudflared/releases/latest/download/cloudflared-windows-amd64.exe', $cfExe)
}

# 2. Jalankan Listener HTTP (Background)
$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("$urlLocal/")
$listener.Start()

# 3. Jalankan Cloudflare Tunnel & Kirim URL ke Excel (Tanpa Toast)
Start-Job -ScriptBlock {
    param($cfExe, $urlLocal)
    & $cfExe tunnel --url $urlLocal 2>&1 | ForEach-Object {
        if ($_ -match 'https://[a-z0-9-]+\.trycloudflare\.com') {
            $urlPublik = $matches[0]
            try {
                $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
                # Hanya isi nilai ke Cell F10 tanpa memanggil TampilkanToast
                $excel.Sheets("DEV").Range("F10").Value = $urlPublik
            } catch {
                # Diam jika Excel sedang sibuk
            }
        }
    }
} -ArgumentList $cfExe, $urlLocal

# 4. Loop Listener Utama (Tetap kirim Toast untuk pesan masuk)
try {
    while ($listener.IsListening) {
        $context = $listener.GetContext()
        # Mengambil parameter ?teks= dari URL
        $pesan = $context.Request.QueryString["teks"]
        
        if ($pesan) {
            try {
                $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
                # Notifikasi Toast hanya muncul di sini saat ada data masuk
                $excel.Run("TampilkanToast", "", $pesan, "")
            } catch {
                # Excel mungkin sedang dalam mode edit cell
            }
        }
        
        # Kirim respon balik ke browser agar tidak timeout
        $buffer = [System.Text.Encoding]::UTF8.GetBytes("OK")
        $context.Response.ContentLength64 = $buffer.Length
        $context.Response.OutputStream.Write($buffer, 0, $buffer.Length)
        $context.Response.Close()
    }
} finally {
    $listener.Stop()
}
