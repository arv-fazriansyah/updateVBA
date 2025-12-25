$currentDir = $PSScriptRoot
if (-not $currentDir) { $currentDir = Get-Location }

# Mengambil nama folder terakhir sebagai ID
$id = Split-Path $currentDir -Leaf
$path = Join-Path $currentDir "target.txt"

# Setting ntfy menggunakan ID sebagai identitas WebSocket
$wsUrl = "wss://ntfy.sh/$id/ws"

Write-Host "--- LISTENER START ---"

$ws = New-Object System.Net.WebSockets.ClientWebSocket
$token = [System.Threading.CancellationToken]::None
$uri = New-Object System.Uri($wsUrl)

try {
    $ws.ConnectAsync($uri, $token).Wait()
    
    if ($ws.State -eq "Open") {
        Write-Host "CONNECTED TO ID: $id"
        
        # --- KIRIM NOTIFIKASI SAAT BERHASIL KONEK ---
        try {
            if (Test-Path $path) {
                $targetPath = (Get-Content $path -Raw).Trim()
                $xls = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
                foreach ($wb in $xls.Workbooks) {
                    if ($wb.FullName -eq $targetPath) { 
                        $xls.Run("TampilkanToast", "Sistem", "Listener Berhasil Terhubung!", "") 
                    }
                }
            }
        } catch { } 
        # --------------------------------------------

        while ($ws.State -eq "Open") {
            $buf = New-Object Byte[] 1024
            $seg = New-Object ArraySegment[Byte] @(,$buf)
            $res = $ws.ReceiveAsync($seg, $token).Result
            
            if ($res.Count -gt 0) {
                $rawData = [System.Text.Encoding]::UTF8.GetString($buf, 0, $res.Count)
                $data = $rawData | ConvertFrom-Json
                
                if ($data.event -eq "message") {
                    Write-Host "Pesan: $($data.message)"
                    
                    try {
                        if (Test-Path $path) {
                            $targetPath = (Get-Content $path -Raw).Trim()
                            $xls = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
                            
                            foreach ($wb in $xls.Workbooks) {
                                if ($wb.FullName -eq $targetPath) { 
                                    $judul = if ($data.title) { $data.title } else { "Notif" }
                                    $xls.Run("TampilkanToast", $judul, $data.message, "") 
                                    Write-Host "Berhasil kirim ke Excel"
                                }
                            }
                        }
                    } catch { 
                        Write-Host "Excel sibuk atau tidak ditemukan" 
                    }
                }
            }
        }
    }
} catch {
    Write-Host "Error: $($_.Exception.Message)"
} finally {
    if ($ws) { $ws.Dispose() }
}
