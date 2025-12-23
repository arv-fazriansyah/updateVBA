# --- CONFIGURATION ---
$port = "8080"
$urlLocal = "http://127.0.0.1:$port"
$folderARB = Join-Path $env:TEMP "ARB2026"
$cfExe = Join-Path $folderARB "cloudflared.exe"

# 1. Jalankan Cloudflare Tunnel (Tanpa Toast)
Start-Job -ScriptBlock {
    param($cfExe, $urlLocal)
    & $cfExe tunnel --url $urlLocal 2>&1 | ForEach-Object {
        if ($_ -match 'https://[a-z0-9-]+\.trycloudflare\.com') {
            $urlPublik = $matches[0]
            try {
                $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
                # HANYA mengisi cell, tidak memanggil macro Toast
                $excel.Sheets("DEV").Range("F10").Value = $urlPublik
            } catch {}
        }
    }
} -ArgumentList $cfExe, $urlLocal

# 2. Listener Utama (Tetap kirim Toast untuk pesan)
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
                # Toast HANYA dipicu di sini saat ada parameter 'teks'
                $excel.Run("TampilkanToast", "tes", $pesan, "")
            } catch {}
        }
        $context.Response.Close()
    }
} finally { $listener.Stop() }
