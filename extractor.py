import os
import sys
import zipfile
import shutil
from config import CONFIG

def extract_excel_components(excel_path: str, output_dir: str):
    """
    Mengekstrak HANYA folder 'xl/vbaProject.bin' dan folder 'customUI/' 
    dari dalam file .xlsb/.xlsm ke dalam folder tujuan.
    """
    if not os.path.exists(excel_path):
        print(f"[ERROR] File Excel tidak ditemukan: {excel_path}")
        sys.exit(1)
        
    try:
        os.makedirs(output_dir, exist_ok=True)
        
        extracted_count = 0
        with zipfile.ZipFile(excel_path, 'r') as zf:
            for member in zf.namelist():
                norm_member = member.replace('\\', '/')
                # Ekstrak vbaProject.bin ATAU seluruh isi folder customUI/
                if norm_member.lower() == 'xl/vbaproject.bin' or norm_member.lower().startswith('customui/'):
                    zf.extract(member, output_dir)
                    extracted_count += 1
                    
        if extracted_count > 0:
            print(f"\n[OK] Berhasil mengekstrak komponen (vbaProject.bin & customUI) ke folder: {output_dir}")
        else:
            print(f"\n[WARNING] Komponen vbaProject.bin atau customUI tidak ditemukan di dalam file Excel tersebut.")
        
    except Exception as e:
        print(f"\n[ERROR] Gagal mengekstrak komponen: {e}")
        sys.exit(1)

def get_running_dir() -> str:
    """Mendapatkan direktori sebenarnya dari file .exe atau script .py yang sedang dieksekusi."""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(os.path.abspath(sys.executable))
    else:
        return os.path.dirname(os.path.abspath(__file__))

def find_excel_files(search_dir: str) -> list:
    """Mencari file Excel (.xlsb / .xlsm) HANYA di dalam folder direktori aktif."""
    found_files = []
    if os.path.exists(search_dir):
        for root, dirs, files in os.walk(search_dir):
            if any(exc in root.lower() for exc in ['.git', 'build', 'dist', 'backup', 'arb 2027']):
                continue
            for f in files:
                if f.endswith(('.xlsb', '.xlsm')):
                    full_path = os.path.join(root, f)
                    found_files.append({
                        "name": f,
                        "full_path": full_path
                    })
                    
    unique_files = {}
    for ef in found_files:
        unique_files[ef['full_path']] = ef
        
    return list(unique_files.values())

def main():
    current_dir = get_running_dir()
    excel_files = find_excel_files(current_dir)
    
    if not excel_files:
        print(f"[ERROR] Tidak ditemukan file Excel (.xlsb/.xlsm) di folder ini ({current_dir})!")
        if sys.stdin.isatty():
            input("\nTekan Enter untuk keluar...")
        sys.exit(1)
        
    print("\n" + "=" * 60)
    print(" PILIH FILE EXCEL YANG AKAN DIEKSTRAK (vbaProject.bin & customUI):")
    print("=" * 60)
    for idx, ef in enumerate(excel_files):
        print(f"  [{idx + 1}] {ef['name']}")
    print("=" * 60)
    
    while True:
        choice = input(f"Pilih nomor file [1-{len(excel_files)}] (atau Ctrl+C untuk batal): ").strip()
        try:
            sel_idx = int(choice) - 1
            if 0 <= sel_idx < len(excel_files):
                selected_file = excel_files[sel_idx]
                break
            else:
                print("[ERROR] Nomor pilihan tidak valid.")
        except ValueError:
            print("[ERROR] Masukkan angka yang valid.")
            
    output_dir = os.path.join(current_dir, f"ARB {CONFIG['TARGET_YEAR']}")
    
    extract_excel_components(selected_file['full_path'], output_dir)
    
    if sys.stdin.isatty():
        input("\nTekan Enter untuk keluar...")

if __name__ == "__main__":
    main()
