# --- Konfigurasi ---
$port = "8080"
$url = "http://127.0.0.1:$port/"
$logFile = Join-Path $env:TEMP "listener_log.txt"

# Inisialisasi Listener
$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add($url)

try {
    $listener.Start()
    "[" + (Get-Date) + "] Listener aktif. Menunggu fetch ke $url" | Out-File $logFile -Append

    while ($listener.IsListening) {
        # 1. Menunggu Request (Fetch) masuk
        $context = $listener.GetContext()
        $request = $context.Request
        
        # 2. Ambil data jika ada (misal JSON atau teks)
        $reader = New-Object System.IO.StreamReader($request.InputStream)
        $dataDiterima = $reader.ReadToEnd()
        $reader.Close()

        if ([string]::IsNullOrWhiteSpace($dataDiterima)) {
            $dataDiterima = "Tidak ada data body"
        }

        # 3. KONEKSI KE EXCEL & JALANKAN MACRO
        try {
            # Mengambil instance Excel yang sedang terbuka
            $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
            
            # Menjalankan Sub TampilkanToast(judul, pesan, imgPath)
            # Kita kirim data dari fetch ke dalam parameter pesan
            $judul = "Notifikasi Webhook"
            $pesan = "Data Fetch: " + $dataDiterima
            $excel.Run("TampilkanToast", $judul, $pesan, "")
            
            "[" + (Get-Date) + "] Berhasil menjalankan macro Excel" | Out-File $logFile -Append
        }
        catch {
            "[" + (Get-Date) + "] Gagal kontak Excel: $($_.Exception.Message)" | Out-File $logFile -Append
        }

        # 4. Berikan respon balik ke Browser/Fetcher
        $response = $context.Response
        $responseString = "Macro Excel Berhasil Dijalankan!"
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
