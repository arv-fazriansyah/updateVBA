import os
import sys
import zipfile
import urllib.request
import tempfile
import shutil
import pythoncom
import win32com.client as win32
from config import CONFIG

def get_latest_version_from_github() -> str:
    """Mengambil string versi terbaru dari file version.txt di repository GitHub berdasarkan config."""
    try:
        req = urllib.request.Request(CONFIG["VERSION_URL"], headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req, timeout=5) as resp:
            ver = resp.read().decode('utf-8').strip()
            if ver:
                return ver
    except Exception:
        pass
    return CONFIG["DEFAULT_VERSION"]

def get_running_dir() -> str:
    """Mendapatkan direktori sebenarnya dari file .exe atau script .py yang sedang dieksekusi."""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(os.path.abspath(sys.executable))
    else:
        return os.path.dirname(os.path.abspath(__file__))

def find_target_excel(arg_input: str = None) -> str:
    """Mencari file target Excel HANYA di dalam folder direktori saat ini, mengabaikan file temporary Excel (~$)."""
    if arg_input and os.path.exists(arg_input) and arg_input.endswith(('.xlsb', '.xlsm')) and not os.path.basename(arg_input).startswith('~$'):
        return os.path.abspath(arg_input)
        
    current_dir = get_running_dir()
    
    found_files = []
    if os.path.exists(current_dir):
        for root, dirs, files in os.walk(current_dir):
            if any(exc in root.lower() for exc in ['.git', 'build', 'dist', 'backup', 'arb 2027']):
                continue
            for f in files:
                if f.startswith('~$'):
                    continue
                if f.endswith(('.xlsb', '.xlsm')):
                    full_path = os.path.join(root, f)
                    found_files.append(full_path)
                            
    found_files = [f for f in list(set(found_files)) if "BACKUP" not in f and not f.endswith('.bak') and not os.path.basename(f).startswith('~$')]
    
    if not found_files:
        print(f"[ERROR] Tidak ditemukan file aplikasi Excel (.xlsb/.xlsm) di folder ini!")
        print(f"        Lokasi: {current_dir}")
        return None
        
    print("\n" + "─" * 50)
    print(" PILIH FILE EXCEL YANG AKAN DI-UPDATE:")
    print("─" * 50)
    for idx, ff in enumerate(found_files):
        print(f"  [{idx + 1}] {os.path.basename(ff)}")
    print("─" * 50)
    
    while True:
        choice = input(f"Pilih nomor [1-{len(found_files)}] (Ctrl+C untuk batal): ").strip()
        try:
            sel_idx = int(choice) - 1
            if 0 <= sel_idx < len(found_files):
                selected = found_files[sel_idx]
                print(f"\n[OK] Dipilih: {os.path.basename(selected)}\n")
                return selected
            else:
                print("[ERROR] Nomor tidak valid. Silakan coba lagi.")
        except ValueError:
            print("[ERROR] Masukkan angka yang valid.")

def run_update(github_zip_url: str, target_path_input: str, version_arg: str = None):
    print(f"=== UPDATER APLIKASI ARB {CONFIG['TARGET_YEAR']} ===")
    
    version = version_arg if version_arg else get_latest_version_from_github()
    print(f"Versi Target : {version}")
    
    target_file = find_target_excel(target_path_input)
    if not target_file:
        sys.exit(1)
        
    target_dir = os.path.dirname(target_file)
    file_name = os.path.basename(target_file)
    target_year = CONFIG["TARGET_YEAR"]
    backup_dir = os.path.join(target_dir, CONFIG["BACKUP_FOLDER_NAME"])
    
    temp_dir = tempfile.mkdtemp()
    backup_file_target = os.path.join(backup_dir, file_name)
    
    try:
        print("\n[1/4] Memeriksa koneksi & menutup file Excel yang dipilih...")
        urllib.request.urlopen("https://google.com", timeout=5)
        
        try:
            pythoncom.CoInitialize()
            app = win32.Dispatch("Excel.Application")
            
            # Reset paksa settingan aplikasi
            try:
                app.Visible = True
                app.Interactive = True
                app.ScreenUpdating = True
                app.DisplayAlerts = True
                app.EnableEvents = True
                app.Calculation = -4105 # xlCalculationAutomatic
            except Exception:
                pass
                
            target_abs = os.path.abspath(target_file).lower()
            
            for wb in app.Workbooks:
                try:
                    if os.path.abspath(wb.FullName).lower() == target_abs:
                        wb.Close(SaveChanges=False)
                        print(f"      File target berhasil ditutup secara aman.")
                        break
                except Exception:
                    continue
                    
            # SOLUSI PASTI BLANK: Paksa buat jendela baru (Window New) jika workbook tersisa mengalami blank view
            for wb in app.Workbooks:
                try:
                    wb.Windows(1).Visible = True
                    wb.Activate()
                    # Buat window baru untuk workbook tersebut agar merefresh tampilan grid Excel yang blank
                    wb.NewWindow()
                    # Tutup window lama yang blank
                    if wb.Windows.Count > 1:
                        wb.Windows(1).Close()
                except Exception as e:
                    print(f"      [Info Window] {e}")
                    
        except Exception as e:
            print(f"      [Info COM] {e}")
            
        print(f"[2/4] Membuat backup ke folder '{CONFIG['BACKUP_FOLDER_NAME']}'...")
        os.makedirs(backup_dir, exist_ok=True)
        shutil.copy2(target_file, backup_file_target)
        
        print("[3/4] Mengunduh & mengekstrak pembaruan dari GitHub...")
        zip_path = os.path.join(temp_dir, "update.zip")
        urllib.request.urlretrieve(github_zip_url, zip_path)
        
        extract_dir = os.path.join(temp_dir, "extracted")
        with zipfile.ZipFile(zip_path, 'r') as zf:
            zf.extractall(extract_dir)
            
        source_update = None
        target_folder_name = f"ARB {target_year}"
        
        for root, dirs, files in os.walk(extract_dir):
            if os.path.basename(root).lower() == target_folder_name.lower():
                source_update = root
                break
                
        if not source_update:
            for root, dirs, files in os.walk(extract_dir):
                if 'customUI' in dirs or 'xl' in dirs:
                    source_update = root
                    break
            if not source_update:
                source_update = extract_dir

        print("[4/4] Menyuntikkan pembaruan & merapikan file...")
        old_xlsb_dir = os.path.join(temp_dir, "old_xlsb")
        with zipfile.ZipFile(target_file, 'r') as old_zf:
            old_zf.extractall(old_xlsb_dir)
            
        for root, dirs, files in os.walk(source_update):
            for file in files:
                full_path = os.path.join(root, file)
                rel_path = os.path.relpath(full_path, source_update)
                dest_path = os.path.join(old_xlsb_dir, rel_path)
                os.makedirs(os.path.dirname(dest_path), exist_ok=True)
                shutil.copy2(full_path, dest_path)
                
        updated_temp_file = os.path.join(temp_dir, "updated.xlsb")
        with zipfile.ZipFile(updated_temp_file, 'w', zipfile.ZIP_DEFLATED) as new_zf:
            for root, dirs, files in os.walk(old_xlsb_dir):
                for file in files:
                    full_path = os.path.join(root, file)
                    arcname = os.path.relpath(full_path, old_xlsb_dir)
                    new_zf.write(full_path, arcname)
                    
        shutil.move(updated_temp_file, target_file)
        
        base_name_only = file_name
        for prefix in ["[Versi Terbaru]", "v24.", "v25.", "v26.", "v27."]:
            if prefix in base_name_only:
                parts = base_name_only.split("_", 1)
                if len(parts) > 1: base_name_only = parts[1]
                break
                
        new_name = f"{version}_{base_name_only}"
        new_path = os.path.join(target_dir, new_name)
        if target_file != new_path:
            if os.path.exists(new_path): os.remove(new_path)
            os.rename(target_file, new_path)
            
        print("\n" + "═" * 50)
        print(" [OK] UPDATE SELESAI DENGAN SUKSES!")
        print(f" File baru: {new_name}")
        print("═" * 50)
        
        if sys.stdin.isatty():
            input("\nTekan Enter untuk keluar...")
        
    except Exception as e:
        print(f"\n[ERROR] Update gagal: {e}")
        if os.path.exists(backup_file_target):
            shutil.copy2(backup_file_target, target_file)
            print("File dipulihkan dari backup.")
        if sys.stdin.isatty():
            input("\nTekan Enter untuk keluar...")
        sys.exit(1)
    finally:
        shutil.rmtree(temp_dir)

if __name__ == "__main__":
    url = CONFIG["GITHUB_ZIP_URL"]
    
    arg_path = sys.argv[1] if len(sys.argv) > 1 and os.path.exists(sys.argv[1]) else None
    arg_ver = sys.argv[2] if len(sys.argv) > 2 else None
    
    run_update(url, arg_path, arg_ver)
