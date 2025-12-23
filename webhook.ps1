# --- CONFIGURATION ---
$port = "8080"
$localUrl = "http://127.0.0.1:$port"
$folderARB = Join-Path $env:TEMP "ARB2026"
$cfExe = Join-Path $folderARB "cloudflared.exe"
$logTunnel = Join-Path $folderARB "tunnel.log"

if (-not (Test-Path $folderARB)) { New-Item -ItemType Directory -Path $folderARB }

# 1. Unduh Cloudflared jika belum ada
if (-not (Test-Path $cfExe)) {
    (New-Object Net.WebClient).DownloadFile(
        'https://github.com/cloudflare/cloudflared/releases/latest/download/cloudflared-windows-amd64.exe',
        $cfExe
    )
}

# 2. Jalankan Listener HTTP (Background)
$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("$localUrl/")
$listener.Start()

# 3. Jalankan Cloudflare Tunnel dan ambil URL-nya
if (Test-Path $logTunnel) { Remove-Item $logTunnel }

Start-Process `
    -FilePath $cfExe `
    -ArgumentList "tunnel --url $localUrl" `
    -RedirectStandardError $logTunnel `
    -WindowStyle Hidden

# Ambil URL publik (tanpa kirim pesan)
$urlFound = $false
$retryCount = 0
while (-not $urlFound -and $retryCount -lt 20) {
    if (Test-Path $logTunnel) {
        $content = Get-Content $logTunnel
        if ($content -match 'https://[a-z0-9-]+\.trycloudflare\.com') {
            $urlPublik = $matches[0]
            try {
                $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
                $excel.Sheets("DEV").Range("F10").Value = $urlPublik
                $urlFound = $true
            } catch {
                # Excel mungkin sibuk
            }
        }
    }
    Start-Sleep -Seconds 2
    $retryCount++
}

# 4. Listener Webhook (?teks=)
try {
    while ($listener.IsListening) {
        $context = $listener.GetContext()
        $teks = $context.Request.QueryString["teks"]
        if ($teks) {
            $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
            $excel.Run("TampilkanToast", "Webhook Inbound", $teks, "")
        }
        $buffer = [System.Text.Encoding]::UTF8.GetBytes("OK")
        $context.Response.OutputStream.Write($buffer, 0, $buffer.Length)
        $context.Response.Close()
    }
} finally {
    $listener.Stop()
}
