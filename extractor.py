import os
import sys
import zipfile
import urllib.request
import tempfile
import shutil
from config import CONFIG

def get_zip_excel_files(zip_url: str) -> tuple:
    """Mengunduh ZIP dari GitHub dan memindai semua file Excel (.xlsb / .xlsm) di dalamnya."""
    print("[1/3] Menghubungkan ke GitHub untuk memindai daftar file Excel...")
    temp_dir = tempfile.mkdtemp()
    try:
        zip_path = os.path.join(temp_dir, "repo.zip")
        urllib.request.urlretrieve(zip_url, zip_path)
        
        extract_dir = os.path.join(temp_dir, "extracted")
        with zipfile.ZipFile(zip_path, 'r') as zf:
            zf.extractall(extract_dir)
            
        # Cari semua file .xlsb atau .xlsm di dalam arsip ZIP
        excel_files = []
        for root, dirs, files in os.walk(extract_dir):
            for file in files:
                if file.endswith(('.xlsb', '.xlsm')):
                    full_path = os.path.join(root, file)
                    rel_path = os.path.relpath(full_path, extract_dir)
                    excel_files.append({
                        "name": file,
                        "rel_path": rel_path,
                        "full_path": full_path,
                        "parent_dir": os.path.dirname(full_path)
                    })
                    
        return excel_files, extract_dir, temp_dir
    except Exception as e:
        print(f"[ERROR] Gagal memindai file dari GitHub: {e}")
        shutil.rmtree(temp_dir)
        sys.exit(1)

def interactive_extract_excel():
    print(f"=== EKSTRAKTOR FILE EXCEL ARB ({CONFIG['TARGET_YEAR']}) ===")
    
    zip_url = CONFIG["GITHUB_ZIP_URL"]
    excel_files, extract_dir, temp_dir = get_zip_excel_files(zip_url)
    
    if not excel_files:
        print("[ERROR] Tidak ditemukan file Excel (.xlsb/.xlsm) di dalam arsip GitHub!")
        shutil.rmtree(temp_dir)
        sys.exit(1)
        
    print("\n" + "=" * 70)
    print(" DAFTAR FILE EXCEL YANG DITEMUKAN DI GITHUB:")
    print("=" * 70)
    for idx, ef in enumerate(excel_files):
        print(f"  [{idx + 1}] {ef['name']}  (path: {ef['rel_path']})")
    print("=" * 70)
    
    while True:
        choice = input(f"Pilih nomor file Excel yang ingin diekstrak [1-{len(excel_files)}] (atau Ctrl+C untuk batal): ").strip()
        try:
            sel_idx = int(choice) - 1
            if 0 <= sel_idx < len(excel_files):
                selected_file = excel_files[sel_idx]
                print(f"\n[OK] Anda memilih: {selected_file['name']}\n")
                break
            else:
                print("[ERROR] Nomor pilihan tidak valid. Silakan coba lagi.")
        except ValueError:
            print("[ERROR] Masukkan angka yang valid.")
            
    # Direktori tujuan: folder "ARB 2027" di direktori saat ini
    current_dir = os.path.dirname(os.path.abspath(__file__))
    output_dir = os.path.join(current_dir, f"ARB {CONFIG['TARGET_YEAR']}")
    
    print(f"[2/3] Menyimpan file '{selected_file['name']}' ke direktori lokal...")
    print(f"       Tujuan: {output_dir}")
    
    try:
        os.makedirs(output_dir, exist_ok=True)
        dest_file_path = os.path.join(output_dir, selected_file['name'])
        
        shutil.copy2(selected_file['full_path'], dest_file_path)
        
        print("\n" + "=" * 50)
        print("[OK] EKSTRAK FILE EXCEL SELESAI DENGAN SUKSES!")
        print(f"Disimpan di: {dest_file_path}")
        print("=" * 50)
        
    except Exception as e:
        print(f"\n[ERROR] Gagal menyimpan file: {e}")
    finally:
        shutil.rmtree(temp_dir)
        if sys.stdin.isatty():
            input("\nTekan Enter untuk keluar...")

if __name__ == "__main__":
    interactive_extract_excel()
