import os
import sys
import zipfile
import shutil
from config import CONFIG

def extract_excel_components(excel_path: str, output_dir: str):
    """
    Mengekstrak hanya folder 'xl' dan 'customUI' dari dalam file .xlsb/.xlsm 
    (karena file Excel berbasis OpenXML/ZIP) ke dalam folder tujuan.
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
                # Cek apakah file berada di dalam folder 'xl/' atau 'customUI/'
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

def get_local_excel_files(search_dir: str) -> list:
    """Memindai semua file Excel (.xlsb / .xlsm) di dalam direktori lokal saat ini."""
    excel_files = []
    for root, dirs, files in os.walk(search_dir):
        if any(exc in root for exc in ['.git', 'build', 'dist', f"ARB {CONFIG['TARGET_YEAR']}", 'BACKUP']):
            continue
        for file in files:
            if file.endswith(('.xlsb', '.xlsm')):
                full_path = os.path.join(root, file)
                rel_path = os.path.relpath(full_path, search_dir)
                excel_files.append({
                    "name": file,
                    "rel_path": rel_path,
                    "full_path": full_path
                })
    return excel_files

def main():
    current_dir = os.path.dirname(os.path.abspath(__file__))
    excel_files = get_local_excel_files(current_dir)
    
    if not excel_files:
        print(f"[ERROR] Tidak ditemukan file Excel (.xlsb/.xlsm) di folder ini!")
        if sys.stdin.isatty():
            input("\nTekan Enter untuk keluar...")
        sys.exit(1)
        
    print("\n" + "=" * 70)
    print(" PILIH FILE EXCEL UNTUK DIEKSTRAK KOMPONENNYA (xl & customUI):")
    print("=" * 70)
    for idx, ef in enumerate(excel_files):
        print(f"  [{idx + 1}] {ef['name']}  (path: {ef['rel_path']})")
    print("=" * 70)
    
    while True:
        choice = input(f"Pilih nomor file [1-{len(excel_files)}] (atau Ctrl+C untuk batal): ").strip()
        try:
            sel_idx = int(choice) - 1
            if 0 <= sel_idx < len(excel_files):
                selected_file = excel_files[sel_idx]
                print(f"\n[OK] Anda memilih: {selected_file['name']}\n")
                break
            else:
                print("[ERROR] Nomor pilihan tidak valid.")
        except ValueError:
            print("[ERROR] Masukkan angka yang valid.")
            
    # Direktori tujuan: folder "ARB 2027" di direktori saat ini
    output_dir = os.path.join(current_dir, f"ARB {CONFIG['TARGET_YEAR']}")
    
    extract_excel_components(selected_file['full_path'], output_dir)
    
    if sys.stdin.isatty():
        input("\nTekan Enter untuk keluar...")

if __name__ == "__main__":
    main()
