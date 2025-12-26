[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$currentDir = if ($PSScriptRoot) { $PSScriptRoot } else { Get-Location }
$id = Split-Path $currentDir -Leaf
$targetFile = [System.IO.Path]::Combine($currentDir, "target.txt")
$logFile = [System.IO.Path]::Combine($currentDir, "log.txt")
$wsUrl = "wss://ntfy.sh/arb2026-$id/ws"

$utf8 = [System.Text.Encoding]::UTF8
$global:xls = $null

function Write-Log {
    param($entry)
    $msg = "[{0}] $entry" -f (Get-Date -Format "HH:mm:ss")
    Write-Host $msg
    try { [System.IO.File]::AppendAllText($logFile, $msg + [System.Environment]::NewLine) } catch {}
}

function Send-ToExcel {
    param($judul, $pesan)
    try {
        if ($null -eq $global:xls) {
            $global:xls = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
        }
        $path = [System.IO.File]::ReadAllText($targetFile).Trim()
        $fileName = [System.IO.Path]::GetFileName($path)
        $wb = $global:xls.Workbooks.Item($fileName)
        
        if ($wb.FullName -eq $path) {
            # FIX: Gunakan array objek eksplisit dan panggil Run melalui InvokeMember agar tidak 'ambiguous'
            $args = [object[]]@($judul, $pesan, "")
            $null = $global:xls.GetType().InvokeMember("Run", [System.Reflection.BindingFlags]::InvokeMethod, $null, $global:xls, @("TampilkanToast") + $args)
        }
    } catch { $global:xls = $null }
}

Write-Log "--- HYPER LISTENER READY ---"

while ($true) {
    $ws = New-Object System.Net.WebSockets.ClientWebSocket
    $ws.Options.KeepAliveInterval = [TimeSpan]::FromSeconds(30)
    $cts = New-Object System.Threading.CancellationTokenSource
    
    try {
        $uri = New-Object System.Uri($wsUrl)
        [void]$ws.ConnectAsync($uri, $cts.Token).GetAwaiter().GetResult()
        Write-Log "Connected to $id"
        
        # Kirim notif koneksi berhasil (Opsional)
        Send-ToExcel "Sistem" "Listener Berhasil Terhubung"

        $buffer = New-Object Byte[] 8192 # Buffer lebih besar
        
        while ($ws.State -eq "Open") {
            $segment = New-Object ArraySegment[Byte] @(,$buffer)
            $res = $ws.ReceiveAsync($segment, $cts.Token).GetAwaiter().GetResult()
            
            if ($res.MessageType -eq "Close") { break }

            if ($res.Count -gt 0) {
                $raw = $utf8.GetString($buffer, 0, $res.Count)
                if ($raw.Contains('"event":"message"')) {
                    $data = $raw | ConvertFrom-Json
                    Write-Log "Pesan Masuk: $($data.message)"
                    
                    # Jalankan tanpa thread baru untuk stabilitas COM
                    Send-ToExcel $data.title $data.message
                }
            }
        }
    } catch {
        $ex = $_.Exception
        while ($ex.InnerException) { $ex = $ex.InnerException }
        if ($ex.Message -notmatch "closed") { Write-Log "Status: $($ex.Message)" }
    } finally {
        if ($ws) { $ws.Dispose() }
        if ($cts) { $cts.Dispose() }
        Start-Sleep -Seconds 1
    }
}
