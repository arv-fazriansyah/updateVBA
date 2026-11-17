import os
import sys
import zipfile
import shutil
from config import CONFIG

def extract_excel_components(excel_path: str, output_dir: str):
    """
    Mengekstrak hanya folder 'xl' dan 'customUI' dari dalam file .xlsb/.xlsm 
    ke dalam folder tujuan.
    """
    print(f"=== EKSTRAKTOR KOMPONEN EXCEL (xl & customUI) ===")
    print(f"File Sumber : {excel_path}")
    print(f"Tujuan      : {output_dir}")
    print("-" * 50)
    
    if not os.path.exists(excel_path):
        print(f"[ERROR] File Excel tidak ditemukan: {excel_path}")
        sys.exit(1)
        
    try:
        os.makedirs(output_dir, exist_ok=True)
        
        print("[1/2] Membuka arsip internal file Excel...")
        extracted_count = 0
        
        with zipfile.ZipFile(excel_path, 'r') as zf:
            for member in zf.namelist():
                norm_member = member.replace('\\', '/')
                if norm_member.startswith('xl/') or norm_member.startswith('customUI/'):
                    zf.extract(member, output_dir)
                    extracted_count += 1
                    
        print(f"[2/2] Berhasil mengekstrak {extracted_count} item komponen.")
        print("=" * 50)
        print("[OK] EKSTRAK KOMPONEN (xl & customUI) SELESAI!")
        print(f"Disimpan di: {output_dir}")
        print("=" * 50)
        
    except Exception as e:
        print(f"\n[ERROR] Gagal mengekstrak komponen Excel: {e}")
        sys.exit(1)

def find_excel_files() -> list:
    """Mencari file Excel (.xlsb / .xlsm) secara cerdas menggunakan Smart Scan (folder exe, Documents, excel_files, Downloads)."""
    search_dirs = [
        os.path.dirname(os.path.abspath(__file__)),
        os.path.join(os.path.expanduser("~"), "Documents", "excel_files"),
        os.path.join(os.path.expanduser("~"), "Documents"),
        os.path.join(os.path.expanduser("~"), "Downloads")
    ]
    
    found_files = []
    for d in search_dirs:
        if os.path.exists(d):
            for root, dirs, files in os.walk(d):
                # Abaikan folder build, dist, git, backup, dan target folder ARB
                if any(exc in root.lower() for exc in ['.git', 'build', 'dist', 'backup', 'arb 2027']):
                    continue
                if root.count(os.sep) - d.count(os.sep) > 2: 
                    continue
                for f in files:
                    if f.endswith(('.xlsb', '.xlsm')):
                        full_path = os.path.join(root, f)
                        rel_path = os.path.relpath(full_path, d)
                        found_files.append({
                            "name": f,
                            "rel_path": f"{d} ({rel_path})",
                            "full_path": full_path
                        })
                        
    # Hilangkan duplikat berdasarkan full_path
    unique_files = {}
    for ef in found_files:
        unique_files[ef['full_path']] = ef
        
    return list(unique_files.values())

def main():
    print(f"=== SMART SCAN EKSTRAKTOR EXCEL ({CONFIG['TARGET_YEAR']}) ===")
    print("Memindai file Excel di komputer secara otomatis...")
    
    excel_files = find_excel_files()
    
    if not excel_files:
        print(f"[ERROR] Tidak ditemukan file Excel (.xlsb/.xlsm) di komputer ini!")
        if sys.stdin.isatty():
            input("\nTekan Enter untuk keluar...")
        sys.exit(1)
        
    print("\n" + "=" * 80)
    print(" PILIH FILE EXCEL UNTUK DIEKSTRAK KOMPONENNYA (xl & customUI):")
    print("=" * 80)
    for idx, ef in enumerate(excel_files):
        print(f"  [{idx + 1}] {ef['name']}")
        print(f"      Path: {ef['full_path']}")
    print("=" * 80)
    
    while True:
        choice = input(f"Pilih nomor file [1-{len(excel_files)}] (atau Ctrl+C untuk batal): ").strip()
        try:
            sel_idx = int(choice) - 1
            if 0 <= sel_idx < len(excel_files):
                selected_file = excel_files[sel_idx]
                print(f"\n[OK] Anda memilih: {selected_file['full_path']}\n")
                break
            else:
                print("[ERROR] Nomor pilihan tidak valid.")
        except ValueError:
            print("[ERROR] Masukkan angka yang valid.")
            
    # Direktori tujuan: folder "ARB 2027" di tempat file Excel tersebut berada (atau di folder exe)
    target_dir = os.path.dirname(selected_file['full_path'])
    output_dir = os.path.join(target_dir, f"ARB {CONFIG['TARGET_YEAR']}")
    
    extract_excel_components(selected_file['full_path'], output_dir)
    
    if sys.stdin.isatty():
        input("\nTekan Enter untuk keluar...")

if __name__ == "__main__":
    main()
