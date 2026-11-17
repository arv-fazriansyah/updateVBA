import os
import sys
import zipfile
import urllib.request
import tempfile
import shutil
from config import CONFIG

def extract_specific_folder(zip_url: str, target_folder_name: str, output_dir: str):
    """
    Mengunduh file ZIP dari GitHub dan mengekstrak folder tertentu saja 
    (misal: 'ARB 2027') ke direktori tujuan lokal.
    """
    print(f"=== EKSTRAK FOLDER SPESIFIK DARI GITHUB ===")
    print(f"Target Folder : {target_folder_name}")
    print(f"Direktori Output: {output_dir}")
    print("-" * 50)
    
    temp_dir = tempfile.mkdtemp()
    try:
        print("[1/3] Mengunduh arsip ZIP dari GitHub...")
        zip_path = os.path.join(temp_dir, "repo.zip")
        urllib.request.urlretrieve(zip_url, zip_path)
        
        print("[2/3] Mengekstrak arsip...")
        extract_dir = os.path.join(temp_dir, "extracted")
        with zipfile.ZipFile(zip_path, 'r') as zf:
            zf.extractall(extract_dir)
            
        print(f"[3/3] Mencari dan menyalin folder '{target_folder_name}'...")
        source_folder = None
        for root, dirs, files in os.walk(extract_dir):
            if os.path.basename(root).lower() == target_folder_name.lower():
                source_folder = root
                break
                
        if not source_folder:
            print(f"[ERROR] Folder '{target_folder_name}' tidak ditemukan di dalam ZIP!")
            sys.exit(1)
            
        # Buat direktori output jika belum ada
        os.makedirs(output_dir, exist_ok=True)
        
        # Salin seluruh isi folder target ke output_dir
        dest_path = os.path.join(output_dir, target_folder_name)
        if os.path.exists(dest_path):
            shutil.rmtree(dest_path)
            
        shutil.copytree(source_folder, dest_path)
        
        print("=" * 50)
        print("[OK] EKSTRAK FOLDER BERHASIL DISLESAIKAN!")
        print(f"Hasil ekstrak tersimpan di: {dest_path}")
        print("=" * 50)
        
    except Exception as e:
        print(f"[ERROR] Gagal mengekstrak folder: {e}")
        sys.exit(1)
    finally:
        shutil.rmtree(temp_dir)

if __name__ == "__main__":
    # Contoh pemanggilan: Mengekstrak folder "ARB 2027" ke folder Downloads user
    zip_url = CONFIG["GITHUB_ZIP_URL"]
    folder_to_extract = f"ARB {CONFIG['TARGET_YEAR']}"
    
    # Direktori output tujuan sesuai permintaan user
    default_output = os.path.join(os.path.expanduser("~"), "Downloads", "updateARB-main")
    
    output_target = sys.argv[1] if len(sys.argv) > 1 else default_output
    
    extract_specific_folder(zip_url, folder_to_extract, output_target)
