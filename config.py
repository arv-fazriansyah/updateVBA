# ==========================================
# KONFIGURASI UTAMA UPDATER APLIKASI
# ==========================================
# Ubah pengaturan di bawah ini dengan mudah sesuai kebutuhan proyek Anda:

CONFIG = {
    # Tahun aplikasi (digunakan untuk folder sumber di dalam ZIP GitHub)
    "TARGET_YEAR": "2027",

    # Nama folder output satu-satunya (file hasil suntikan update)
    "NEW_VERSION_FOLDER_NAME": "VERSI TERBARU ARB 2027",
    
    # URL repository file ZIP update dari GitHub
    "GITHUB_ZIP_URL": "https://codeload.github.com/arv-fazriansyah/updateARB/zip/refs/heads/main",
    
    # URL untuk mengecek versi terbaru secara live
    "VERSION_URL": "https://raw.githubusercontent.com/arv-fazriansyah/updateVBA/refs/heads/main/version.txt",
    
    # Fallback versi jika gagal terhubung ke GitHub
    "DEFAULT_VERSION": "[Versi Terbaru]"
}
# ==========================================
