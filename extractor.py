import os
import sys
import shutil
from config import CONFIG

def get_local_excel_files(search_dir: str) -> list:
    """Memindai semua file Excel (.xlsb / .xlsm) di dalam direktori lokal saat ini."""
    excel_files = []
    for root, dirs, files in os.walk(search_dir):
        # Abaikan folder build, dist, git, dan target folder ARB sendiri agar tidak berputar
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

def interactive_extract_local_excel():
    print(f"=== EKSTRAKTOR FILE EXCEL LOKAL ({CONFIG['TARGET_YEAR']}) ===")
    
    current_dir = os.path.dirname(os.path.abspath(__file__))
    print(f"[1/2] Memindai file Excel di direktori: {current_dir}")
    
    excel_files = get_local_excel_files(current_dir)
    
    if not excel_files:
        print(f"[ERROR] Tidak ditemukan file Excel (.xlsb/.xlsm) di folder ini atau subfoldernya!")
        if sys.stdin.isatty():
            input("\nTekan Enter untuk keluar...")
        sys.exit(1)
        
    print("\n" + "=" * 70)
    print(" DAFTAR FILE EXCEL LOKAL YANG DITEMUKAN:")
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
    output_dir = os.path.join(current_dir, f"ARB {CONFIG['TARGET_YEAR']}")
    
    print(f"[2/2] Menyalin file '{selected_file['name']}' ke direktori tujuan...")
    print(f"       Tujuan: {output_dir}")
    
    try:
        os.makedirs(output_dir, exist_ok=True)
        dest_file_path = os.path.join(output_dir, selected_file['name'])
        
        shutil.copy2(selected_file['full_path'], dest_file_path)
        
        print("\n" + "=" * 50)
        print("[OK] SALIN FILE EXCEL SELESAI DENGAN SUKSES!")
        print(f"Disimpan di: {dest_file_path}")
        print("=" * 50)
        
    except Exception as e:
        print(f"\n[ERROR] Gagal menyalin file: {e}")
    finally:
        if sys.stdin.isatty():
            input("\nTekan Enter untuk keluar...")

if __name__ == "__main__":
    interactive_extract_local_excel()
