# --- KONFIGURASI ---
$workerUrl = "wss://arb.arvib-fazriansyah.workers.dev/ws" # Endpoint Worker kamu
$currentDir = $PSScriptRoot
if (-not $currentDir) { $currentDir = Get-Location }

$logFile = Join-Path $currentDir "log.txt"
$pathTargetTxt = Join-Path $currentDir "target.txt"
$startTime = Get-Date

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
                    $excel.Run("TampilkanToast", $judul, $isi, "")
                    return $true
                }
            }
        }
    } catch { 
        Write-Log "DEBUG: Excel sibuk/tutup."
    }
    return $false
}

# --- MAIN LOGIC: WEBSOCKET CLIENT ---
try {
    Write-Log "INFO: Memulai koneksi WebSocket ke $workerUrl"
    
    # Ambil KodeUnik untuk identitas Room (F11 di Excel)
    $targetPath = (Get-Content $pathTargetTxt -Raw).Trim()
    $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
    $roomID = ""
    foreach ($wb in $excel.Workbooks) {
        if ($wb.FullName -eq $targetPath) {
            $roomID = $wb.Sheets("DEV").Range("F11").Value
            $wb.Sheets("DEV").Range("F10").Value = "WS CONNECTING..."
            break
        }
    }

    $ws = New-Object System.Net.WebSockets.ClientWebSocket
    $uri = New-Object System.Uri("$workerUrl`?room=$roomID")
    $cts = New-Object System.Threading.CancellationTokenSource
    
    # Melakukan Koneksi
    $connTask = $ws.ConnectAsync($uri, $cts.Token)
    while (-not $connTask.IsCompleted) { Start-Sleep -Milliseconds 100 }

    if ($ws.State -eq "Open") {
        Write-Log "INFO: Connected to Worker Room: $roomID"
        $excel.Workbooks | Where-Object { $_.FullName -eq $targetPath } | ForEach-Object {
            $_.Sheets("DEV").Range("F10").Value = "WS ACTIVE"
        }

        # Loop Standby Menerima Pesan
        while ($ws.State -eq "Open") {
            $buffer = New-Object Byte[] 4096
            $segment = New-Object ArraySegment[Byte] -ArgumentList @(,$buffer)
            $receiveTask = $ws.ReceiveAsync($segment, $cts.Token)
            
            while (-not $receiveTask.IsCompleted) { Start-Sleep -Milliseconds 100 }
            
            $result = $receiveTask.Result
            if ($result.MessageType -eq "Close") { break }

            $rawData = [System.Text.Encoding]::UTF8.GetString($buffer, 0, $result.Count)
            Write-Log "RECEIVE: $rawData"

            # Asumsi pesan dari Worker berupa JSON: {"judul":"..", "teks":".."}
            try {
                $data = $rawData | ConvertFrom-Json
                Kirim-Ke-Excel $data.judul $data.teks
            } catch {
                # Jika bukan JSON, kirim sebagai teks biasa
                Kirim-Ke-Excel "Notifikasi" $rawData
            }
        }
    }
} catch {
    Write-Log "ERROR: $($_.Exception.Message)"
} finally {
    if ($null -ne $ws) { $ws.Dispose() }
    Write-Log "INFO: Listener Berhenti."
}
