import os
import sys
import zipfile
import urllib.request
import tempfile
import shutil
from config import CONFIG

def get_zip_subfolders(zip_url: str) -> dict:
    """Mengunduh ZIP dari GitHub dan memindai semua subfolder utama di dalamnya."""
    print("[1/3] Menghubungkan ke GitHub untuk memindai daftar folder...")
    temp_dir = tempfile.mkdtemp()
    try:
        zip_path = os.path.join(temp_dir, "repo.zip")
        urllib.request.urlretrieve(zip_url, zip_path)
        
        extract_dir = os.path.join(temp_dir, "extracted")
        with zipfile.ZipFile(zip_path, 'r') as zf:
            zf.extractall(extract_dir)
            
        # Cari subfolder utama di dalam root extract (misal: "ARB 2026", "ARB 2027", dll)
        folders = []
        for root, dirs, files in os.walk(extract_dir):
            # Ambil level pertama subfolder di dalam extract_dir utama
            rel_path = os.path.relpath(root, extract_dir)
            if rel_path != "." and rel_path.count(os.sep) == 1:
                folders.append(os.path.basename(root))
                
        # Jika tidak ketemu level 1, cari folder yang memiliki file customUI atau xl
        if not folders:
            for root, dirs, files in os.walk(extract_dir):
                if 'customUI' in dirs or 'xl' in dirs:
                    folders.append(os.path.basename(root))
                    
        return list(set(folders)), extract_dir, temp_dir
    except Exception as e:
        print(f"[ERROR] Gagal memindai folder dari GitHub: {e}")
        shutil.rmtree(temp_dir)
        sys.exit(1)

def interactive_extract():
    print(f"=== EKSTRAKTOR FOLDER ARB (INTERAKTIF) ===")
    
    zip_url = CONFIG["GITHUB_ZIP_URL"]
    available_folders, extract_dir, temp_dir = get_zip_subfolders(zip_url)
    
    if not available_folders:
        print("[ERROR] Tidak ada folder yang ditemukan di dalam arsip GitHub!")
        shutil.rmtree(temp_dir)
        sys.exit(1)
        
    print("\n" + "=" * 60)
    print(" PILIH FOLDER YANG AKAN DIEKSTRAK:")
    print("=" * 60)
    for idx, folder_name in enumerate(available_folders):
        print(f"  [{idx + 1}] {folder_name}")
    print("=" * 60)
    
    while True:
        choice = input(f"Pilih nomor folder [1-{len(available_folders)}] (atau Ctrl+C untuk batal): ").strip()
        try:
            sel_idx = int(choice) - 1
            if 0 <= sel_idx < len(available_folders):
                selected_folder = available_folders[sel_idx]
                print(f"\n[OK] Anda memilih: {selected_folder}\n")
                break
            else:
                print("[ERROR] Nomor pilihan tidak valid. Silakan coba lagi.")
        except ValueError:
            print("[ERROR] Masukkan angka yang valid.")
            
    # Tentukan direktori saat ini (tempat script/exe berada)
    current_dir = os.path.dirname(os.path.abspath(__file__))
    output_dir = os.path.join(current_dir, "ARB 2027")
    
    print(f"[2/3] Menyalin folder '{selected_folder}' ke direktori lokal...")
    print(f"       Tujuan: {output_dir}")
    
    try:
        source_path = None
        for root, dirs, files in os.walk(extract_dir):
            if os.path.basename(root).lower() == selected_folder.lower():
                source_path = root
                break
                
        if not source_path:
            print(f"[ERROR] Path sumber untuk '{selected_folder}' tidak ditemukan.")
            sys.exit(1)
            
        os.makedirs(output_dir, exist_ok=True)
        dest_path = os.path.join(output_dir, selected_folder)
        
        if os.path.exists(dest_path):
            shutil.rmtree(dest_path)
            
        shutil.copytree(source_path, dest_path)
        
        print("\n" + "=" * 50)
        print("[OK] EKSTRAK FOLDER SELESAI DENGAN SUKSES!")
        print(f"Disimpan di: {dest_path}")
        print("=" * 50)
        
    except Exception as e:
        print(f"\n[ERROR] Gagal menyimpan folder: {e}")
    finally:
        shutil.rmtree(temp_dir)
        if sys.stdin.isatty():
            input("\nTekan Enter untuk keluar...")

if __name__ == "__main__":
    interactive_extract()
