# --- Konfigurasi ---
$port = "8080"
$url = "http://127.0.0.1:$port/"
$logFile = Join-Path $env:TEMP "listener_log.txt"

$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add($url)

try {
    $listener.Start()
    "[" + (Get-Date) + "] Listener aktif di $url" | Out-File $logFile -Append

    while ($listener.IsListening) {
        $context = $listener.GetContext()
        $request = $context.Request
        
        # --- AMBIL PARAMETER DARI URL ---
        # Mengambil nilai dari ?teks=...
        $pesanDariUrl = $request.QueryString["teks"]
        
        if ([string]::IsNullOrWhiteSpace($pesanDariUrl)) {
            $pesanDariUrl = "Tidak ada pesan di parameter 'teks'"
        }

        # --- JALANKAN MACRO EXCEL ---
        try {
            $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
            
            # Memanggil Sub TampilkanToast(judul, pesan, imgPath)
            # Parameter teks dari URL dikirim sebagai 'pesan'
            $excel.Run("TampilkanToast", "", $pesanDariUrl, "")
            
            "[" + (Get-Date) + "] Berhasil memicu macro dengan teks: $pesanDariUrl" | Out-File $logFile -Append
        }
        catch {
            "[" + (Get-Date) + "] Excel Error: $($_.Exception.Message)" | Out-File $logFile -Append
        }

        # --- RESPON KE BROWSER ---
        $response = $context.Response
        $responseString = "Pesan '$pesanDariUrl' telah dikirim ke Excel!"
        $buffer = [System.Text.Encoding]::UTF8.GetBytes($responseString)
        $response.ContentLength64 = $buffer.Length
        $response.OutputStream.Write($buffer, 0, $buffer.Length)
        $response.OutputStream.Close()
    }
}
catch {
    $_.Exception.Message | Out-File $logFile -Append
}
finally {
    $listener.Stop()
}
