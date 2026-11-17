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
    """Mencari file Excel (.xlsb / .xlsm) HANYA di dalam folder direktori saat ini (tempat .exe/.py berada)."""
    current_dir = os.path.dirname(os.path.abspath(__file__))
    
    found_files = []
    if os.path.exists(current_dir):
        for root, dirs, files in os.walk(current_dir):
            # Abaikan folder build, dist, git, backup, dan target folder ARB
            if any(exc in root.lower() for exc in ['.git', 'build', 'dist', 'backup', 'arb 2027']):
                continue
            for f in files:
                if f.endswith(('.xlsb', '.xlsm')):
                    full_path = os.path.join(root, f)
                    found_files.append({
                        "name": f,
                        "full_path": full_path
                    })
                    
    # Hilangkan duplikat
    unique_files = {}
    for ef in found_files:
        unique_files[ef['full_path']] = ef
        
    return list(unique_files.values())

def main():
    current_dir = os.path.dirname(os.path.abspath(__file__))
    print(f"=== EKSTRAKTOR EXCEL LOKAL ({CONFIG['TARGET_YEAR']}) ===")
    print(f"Memindai file Excel di folder: {current_dir}")
    
    excel_files = find_excel_files()
    
    if not excel_files:
        print(f"[ERROR] Tidak ditemukan file Excel (.xlsb/.xlsm) di folder ini ({current_dir})!")
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
            
    # Direktori tujuan: folder "ARB 2027" di direktori tempat script/exe berada
    output_dir = os.path.join(current_dir, f"ARB {CONFIG['TARGET_YEAR']}")
    
    extract_excel_components(selected_file['full_path'], output_dir)
    
    if sys.stdin.isatty():
        input("\nTekan Enter untuk keluar...")

if __name__ == "__main__":
    main()
