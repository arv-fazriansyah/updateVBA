# --- CONFIGURATION ---
$port = "8080"
$urlLocal = "http://127.0.0.1:$port"
$currentDir = $PSScriptRoot
if (-not $currentDir) { $currentDir = Get-Location }

$cfExe = Join-Path $currentDir "cloudflared.exe"
$logFile = Join-Path $currentDir "vba_webhook_log.txt"

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

function Write-Log($pesan) {
    $waktu = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "[$waktu] $pesan" | Out-File -FilePath $logFile -Append -Encoding UTF8
}

# 1. Download Cloudflared (Logika tetap sama)
if (-not (Test-Path $cfExe)) {
    Write-Log "INFO: Memulai download cloudflared..."
    try {
        Invoke-WebRequest -Uri 'https://github.com/cloudflare/cloudflared/releases/latest/download/cloudflared-windows-amd64.exe' -OutFile $cfExe -ErrorAction Stop
        Unblock-File -Path $cfExe
    } catch { exit }
}

# 2. Jalankan Listener HTTP
$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("$urlLocal/")
try {
    $listener.Start()
    Write-Log "INFO: Listener aktif di $urlLocal"
} catch {
    Write-Log "ERROR: GAGAL listener: $($_.Exception.Message)"
    exit
}

# 3. Jalankan Cloudflare Tunnel (Background Job)
$job = Start-Job -ScriptBlock {
    param($cfExe, $urlLocal, $logFile)
    $tempCfLog = $logFile.Replace(".txt", "_cf.tmp")
    $process = Start-Process -FilePath $cfExe -ArgumentList "tunnel --url $urlLocal --no-autoupdate" -NoNewWindow -PassThru -RedirectStandardError $tempCfLog
    # (Logika pencarian URL publik tetap sama seperti sebelumnya...)
} -ArgumentList $cfExe, $urlLocal, $logFile

# 4. Loop Utama
Write-Log "INFO: Menunggu pesan inbound..."
try {
    while ($listener.IsListening) {
        $context = $listener.GetContext()
        $request = $context.Request
        $response = $context.Response
        
        # --- LOGIKA STOP & WEBHOOK ---
        $urlPath = $request.Url.LocalPath # Mengambil path setelah port
        $pesan = $request.QueryString["teks"]
        
        if ($urlPath -eq "/stop") {
            Write-Log "INFO: Menerima perintah STOP dari URL."
            $buffer = [System.Text.Encoding]::UTF8.GetBytes("Listener ditutup.")
            $response.ContentLength64 = $buffer.Length
            $response.OutputStream.Write($buffer, 0, $buffer.Length)
            $response.Close()
            break # Keluar dari loop untuk memicu blok finally
        }
        
        if ($pesan) {
            Write-Log "WEBHOOK: Pesan diterima -> $pesan"
            try {
                $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
                $excel.Run("TampilkanToast", "Pesan Masuk", $pesan, "")
            } catch { }
        }
        
        $buffer = [System.Text.Encoding]::UTF8.GetBytes("OK")
        $response.ContentLength64 = $buffer.Length
        $response.OutputStream.Write($buffer, 0, $buffer.Length)
        $response.Close()
    }
} finally {
    Write-Log "INFO: Membersihkan proses..."
    $listener.Stop()
    $listener.Close()
    # Mematikan Cloudflare Job
    Get-Job | Stop-Job
    Get-Job | Remove-Job
    # Pastikan proses cloudflared benar-benar mati
    Stop-Process -Name "cloudflared" -Force -ErrorAction SilentlyContinue
    Write-Log "INFO: Script dihentikan sempurna."
}
