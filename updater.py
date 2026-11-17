import os
import sys
import zipfile
import urllib.request
import tempfile
import shutil
import threading
import queue
import tkinter as tk
from tkinter import filedialog, messagebox
from config import CONFIG

class UpdaterGUI:
    def __init__(self, root):
        self.root = root
        self.root.title(f"Updater Aplikasi ARB {CONFIG['TARGET_YEAR']}")
        self.root.geometry("640x420")
        self.root.resizable(False, False)

        self.log_queue = queue.Queue()
        self.target_file = None

        # Header
        tk.Label(root, text=f"UPDATER APLIKASI ARB {CONFIG['TARGET_YEAR']}", font=("Segoe UI", 14, "bold")).pack(pady=10)

        # File selection frame
        frame = tk.Frame(root)
        frame.pack(pady=5, padx=15, fill="x")

        tk.Label(frame, text="File Excel Target:", font=("Segoe UI", 10)).pack(anchor="w")
        self.file_entry = tk.Entry(frame, width=70, state="readonly")
        self.file_entry.pack(side="left", fill="x", expand=True, padx=(0, 5))

        self.browse_btn = tk.Button(frame, text="Pilih File", command=self.browse_file, width=12)
        self.browse_btn.pack(side="right")

        # Start button
        self.start_btn = tk.Button(root, text="Mulai Update", command=self.start_update, width=20, height=2,
                                  state="disabled", bg="#2196F3", fg="white", font=("Segoe UI", 11, "bold"))
        self.start_btn.pack(pady=10)

        # Log area
        tk.Label(root, text="Log Proses:", font=("Segoe UI", 10)).pack(anchor="w", padx=15)
        self.log_text = tk.Text(root, height=14, width=78, state="disabled", bg="#1e1e1e", fg="#d4d4d4",
                                font=("Consolas", 9))
        self.log_text.pack(padx=15, pady=5)

        self.root.after(100, self.process_queue)

    def log(self, msg):
        self.log_queue.put(msg)

    def process_queue(self):
        try:
            while True:
                msg = self.log_queue.get_nowait()
                self.log_text.config(state="normal")
                self.log_text.insert("end", msg + "\n")
                self.log_text.see("end")
                self.log_text.config(state="disabled")
        except queue.Empty:
            pass
        self.root.after(100, self.process_queue)

    def browse_file(self):
        path = filedialog.askopenfilename(
            title="Pilih File Excel Aplikasi",
            filetypes=[("Excel Files", "*.xlsb *.xlsm"), ("All Files", "*.*")]
        )
        if path and path.lower().endswith((".xlsb", ".xlsm")):
            self.target_file = os.path.abspath(path)
            self.file_entry.config(state="normal")
            self.file_entry.delete(0, "end")
            self.file_entry.insert(0, self.target_file)
            self.file_entry.config(state="readonly")
            self.start_btn.config(state="normal")
            self.log(f"[OK] File dipilih: {os.path.basename(self.target_file)}")
        elif path:
            messagebox.showerror("Error", "File harus berformat .xlsb atau .xlsm")

    def start_update(self):
        if not self.target_file:
            messagebox.showwarning("Peringatan", "Pilih file Excel terlebih dahulu.")
            return
        self.browse_btn.config(state="disabled")
        self.start_btn.config(state="disabled")
        self.log("\n" + "=" * 60)
        self.log(" MEMULAI PROSES UPDATE")
        self.log("=" * 60)
        threading.Thread(target=self._run_update, daemon=True).start()

    def _run_update(self):
        target_file = self.target_file
        source_dir = os.path.dirname(target_file)
        file_name = os.path.basename(target_file)
        target_year = CONFIG["TARGET_YEAR"]

        old_dir = os.path.join(source_dir, CONFIG["OLD_VERSION_FOLDER_NAME"])
        new_dir = os.path.join(source_dir, CONFIG["NEW_VERSION_FOLDER_NAME"])
        old_file = os.path.join(old_dir, file_name)
        new_file = os.path.join(new_dir, file_name)

        temp_dir = tempfile.mkdtemp()

        try:
            # Step 1: Cek versi terbaru
            self.log("\n[1/4] Mengecek versi terbaru...")
            version = get_latest_version_from_github(self.log)
            self.log(f"      Versi Target : {version}")

            # Step 2: Salin file ke 2 folder (versi lama & versi terbaru)
            self.log("\n[2/4] Menyalin file ke folder backup & versi terbaru...")
            os.makedirs(old_dir, exist_ok=True)
            os.makedirs(new_dir, exist_ok=True)
            shutil.copy2(target_file, old_file)
            shutil.copy2(target_file, new_file)
            self.log(f"      Versi lama  -> {CONFIG['OLD_VERSION_FOLDER_NAME']}")
            self.log(f"      Versi baru  -> {CONFIG['NEW_VERSION_FOLDER_NAME']}")

            # Step 3: Unduh & ekstrak paket update dari GitHub
            self.log("\n[3/4] Mengunduh & mengekstrak pembaruan dari GitHub...")
            github_zip_url = CONFIG["GITHUB_ZIP_URL"]
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
            self.log(f"      Sumber update: {target_folder_name}")

            # Step 4: Suntikkan update ke file di folder VERSI TERBARU
            self.log("\n[4/4] Menyuntikkan pembaruan ke file di folder versi terbaru...")
            old_xlsb_dir = os.path.join(temp_dir, "new_xlsb")
            with zipfile.ZipFile(new_file, 'r') as new_zf:
                new_zf.extractall(old_xlsb_dir)

            for root, dirs, files in os.walk(source_update):
                for file in files:
                    full_path = os.path.join(root, file)
                    rel_path = os.path.relpath(full_path, source_update)
                    dest_path = os.path.join(old_xlsb_dir, rel_path)
                    os.makedirs(os.path.dirname(dest_path), exist_ok=True)
                    shutil.copy2(full_path, dest_path)

            updated_temp_file = os.path.join(temp_dir, "updated.xlsb")
            with zipfile.ZipFile(updated_temp_file, 'w', zipfile.ZIP_DEFLATED) as out_zf:
                for root, dirs, files in os.walk(old_xlsb_dir):
                    for file in files:
                        full_path = os.path.join(root, file)
                        arcname = os.path.relpath(full_path, old_xlsb_dir)
                        out_zf.write(full_path, arcname)

            shutil.move(updated_temp_file, new_file)

            self.log("\n" + "=" * 60)
            self.log(" [OK] UPDATE SELESAI DENGAN SUKSES!")
            self.log(f" File versi terbaru: {new_file}")
            self.log("=" * 60)
            self.root.after(0, lambda: messagebox.showinfo("Sukses", f"Update selesai!\nFile berhasil disimpan di:\n{new_dir}"))

        except Exception as e:
            self.log(f"\n[ERROR] Update gagal: {e}")
            self.root.after(0, lambda: messagebox.showerror("Gagal", f"Update gagal:\n{e}"))
        finally:
            shutil.rmtree(temp_dir, ignore_errors=True)
            self.root.after(0, self._enable_buttons)

    def _enable_buttons(self):
        self.browse_btn.config(state="normal")
        self.start_btn.config(state="normal")

def get_latest_version_from_github(log_func):
    try:
        req = urllib.request.Request(CONFIG["VERSION_URL"], headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req, timeout=5) as resp:
            ver = resp.read().decode('utf-8').strip()
            if ver:
                log_func(f"      [OK] Versi dari GitHub: {ver}")
                return ver
    except Exception:
        pass
    fallback = CONFIG["DEFAULT_VERSION"]
    log_func(f"      [INFO] Menggunakan versi default: {fallback}")
    return fallback

if __name__ == "__main__":
    root = tk.Tk()
    app = UpdaterGUI(root)
    root.mainloop()
