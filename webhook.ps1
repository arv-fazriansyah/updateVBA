# --- CONFIGURATION ---
$port = "8080"
$endpoint = "pesan" # URL akan menjadi http://localhost:8080/pesan
$logFile = Join-Path $env:TEMP "vba_webhook_log.txt"

# Initialize Listener
$url = "http://127.0.0.1:$port/"
$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add($url)

try {
    $listener.Start()
    "[" + (Get-Date) + "] Listener Started at $url" | Out-File $logFile -Append

    while ($listener.IsListening) {
        $context = $listener.GetContext()
        $request = $context.Request
        
        # Ambil pesan dari parameter ?teks=...
        $pesan = $request.QueryString["teks"]
        
        if (-not [string]::IsNullOrWhiteSpace($pesan)) {
            try {
                # Hubungkan ke Excel yang sedang terbuka
                $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
                
                # Jalankan macro: NamaMacro, Judul, Pesan, PathGambar
                $excel.Run("TampilkanToast", "Webhook Inbound", $pesan, "")
                
                "[" + (Get-Date) + "] Success: Sent '$pesan' to Excel" | Out-File $logFile -Append
            } catch {
                "[" + (Get-Date) + "] Excel Error: $($_.Exception.Message)" | Out-File $logFile -Append
            }
        }

        # Kirim respon sukses ke pengirim (Browser/Fetch)
        $buffer = [System.Text.Encoding]::UTF8.GetBytes("OK - Data diproses")
        $context.Response.ContentLength64 = $buffer.Length
        $context.Response.OutputStream.Write($buffer, 0, $buffer.Length)
        $context.Response.OutputStream.Close()
    }
} finally {
    $listener.Stop()
}
