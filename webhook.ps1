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
        $context = $listener.GetContext()
        $request = $context.Request
        
        $reader = New-Object System.IO.StreamReader($request.InputStream)
        $dataDiterima = $reader.ReadToEnd()
        $reader.Close()

        # --- BAGIAN MENJALANKAN MACRO EXCEL ---
        try {
            # Mengambil objek Excel yang sedang terbuka
            $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
            
            # Menjalankan Macro 'TampilkanToast' dengan parameter data yang diterima
            # Format: $excel.Run("NamaMacro", "Argumen1", "Argumen2", ...)
            $excel.Run("TampilkanToast", "Notifikasi Webhook", $dataDiterima)
        }
        catch {
            "[" + (Get-Date) + "] Gagal memanggil Macro: $($_.Exception.Message)" | Out-File $logFile -Append
        }
        # ---------------------------------------

        $response = $context.Response
        $responseString = "Macro berhasil dipicu!"
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
