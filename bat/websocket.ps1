# --- KONFIGURASI ---
$port = 8080
$currentDir = $PSScriptRoot
if (-not $currentDir) { $currentDir = Get-Location }

$logFile = Join-Path $currentDir "log.txt"
$pathTargetTxt = Join-Path $currentDir "target.txt"

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
        Write-Log "DEBUG: Excel Sibuk"
    }
    return $false
}

# Inisialisasi HTTP Listener
$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("http://127.0.0.1:$port/")
try {
    $listener.Start()
    Write-Log "INFO: Listener aktif di port $port"
} catch {
    Write-Log "ERROR: Port $port diduduki aplikasi lain."
    exit
}

# Loop Utama
while ($listener.IsListening) {
    $context = $listener.GetContext()
    $req = $context.Request
    $res = $context.Response
    
    # Ambil parameter dari URL: ?judul=...&teks=...
    $judul = $req.QueryString["judul"]
    $teks = $req.QueryString["teks"]
    
    if ($teks) {
        $success = Kirim-Ke-Excel $judul $teks
        $responTeks = if ($success) { "OK: Terkirim ke Excel" } else { "Error: Excel Sibuk" }
    } else {
        $responTeks = "Ready. Kirim parameter 'teks' untuk notifikasi."
    }

    $buffer = [System.Text.Encoding]::UTF8.GetBytes($responTeks)
    $res.ContentLength64 = $buffer.Length
    $res.OutputStream.Write($buffer, 0, $buffer.Length)
    $res.Close()
}
