# --- Konfigurasi ---
$port = "8080"
$url = "http://127.0.0.1:$port/"
$logFile = Join-Path $env:TEMP "listener_log.txt"

# Inisialisasi Listener
$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add($url)

try {
    $listener.Start()
    "[" + (Get-Date) + "] Listener dimulai di $url" | Out-File $logFile -Append

    while ($listener.IsListening) {
        # Menunggu permintaan datang
        $context = $listener.GetContext()
        $request = $context.Request
        
        # Membaca Body/JSON yang dikirim
        $reader = New-Object System.IO.StreamReader($request.InputStream)
        $dataDiterima = $reader.ReadToEnd()
        $reader.Close()

        # Mencatat data ke file log di folder TEMP
        "[" + (Get-Date) + "] Data Diterima: $dataDiterima" | Out-File $logFile -Append

        # Mengirim respon balik ke pengirim
        $response = $context.Response
        $responseString = "Data berhasil diterima oleh Listener VBA!"
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
