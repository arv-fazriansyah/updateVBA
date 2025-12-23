# --- CONFIGURATION ---
$port = "8080"
$urlLocal = "http://127.0.0.1:$port"
$folderARB = Join-Path $env:TEMP "ARB2026"
$cfExe = Join-Path $folderARB "cloudflared.exe"

if (-not (Test-Path $folderARB)) { New-Item -ItemType Directory -Path $folderARB }

# 1. Download Cloudflared jika belum ada
if (-not (Test-Path $cfExe)) {
    (New-Object Net.WebClient).DownloadFile('https://github.com/cloudflare/cloudflared/releases/latest/download/cloudflared-windows-amd64.exe', $cfExe)
}

# 2. Jalankan Cloudflare Tunnel di Background (Tanpa Toast)
Start-Job -ScriptBlock {
    param($cfExe, $urlLocal)
    & $cfExe tunnel --url $urlLocal 2>&1 | ForEach-Object {
        if ($_ -match 'https://[a-z0-9-]+\.trycloudflare\.com') {
            $urlPublik = $matches[0]
            try {
                $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
                $excel.Sheets("DEV").Range("F10").Value = $urlPublik
            } catch {}
        }
    }
} -ArgumentList $cfExe, $urlLocal

# 3. Listener Utama (Penerima Data)
$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("$urlLocal/")
$listener.Start()

try {
    while ($listener.IsListening) {
        $context = $listener.GetContext()
        $pesan = $context.Request.QueryString["teks"]
        
        if ($pesan) {
            try {
                $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
                $excel.Run("TampilkanToast", "Webhook Masuk", $pesan, "")
            } catch {}
        }
        
        # Kirim respon agar browser tidak putih/error
        $buffer = [System.Text.Encoding]::UTF8.GetBytes("Data Berhasil Diterima!")
        $context.Response.ContentLength64 = $buffer.Length
        $context.Response.OutputStream.Write($buffer, 0, $buffer.Length)
        $context.Response.Close()
    }
} finally {
    $listener.Stop()
}
