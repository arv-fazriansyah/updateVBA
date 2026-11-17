import os
import zipfile
import urllib.request
import tempfile
import shutil
import threading
import queue
import customtkinter as ctk
from tkinter import filedialog, messagebox
from config import CONFIG

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

class UpdaterGUI(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title(f"Updater ARB {CONFIG['TARGET_YEAR']}")
        self.geometry("480x300")
        self.resizable(False, False)

        self.log_queue = queue.Queue()
        self.target_file = None

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(4, weight=1)

        # Judul ringkas
        ctk.CTkLabel(
            self,
            text=f"UPDATER ARB {CONFIG['TARGET_YEAR']}",
            font=ctk.CTkFont(size=17, weight="bold")
        ).grid(row=0, column=0, sticky="w", padx=20, pady=(16, 10))

        # Pilih file
        file_frame = ctk.CTkFrame(self, corner_radius=10)
        file_frame.grid(row=1, column=0, padx=20, pady=4, sticky="ew")
        file_frame.grid_columnconfigure(0, weight=1)

        self.file_entry = ctk.CTkEntry(
            file_frame,
            placeholder_text="Pilih file Excel (.xlsb / .xlsm)...",
            height=34,
            corner_radius=8,
            state="disabled"
        )
        self.file_entry.grid(row=0, column=0, sticky="ew", padx=(12, 6), pady=10)

        self.browse_btn = ctk.CTkButton(
            file_frame, text="Pilih", width=70, height=34, corner_radius=8,
            command=self.browse_file
        )
        self.browse_btn.grid(row=0, column=1, padx=(0, 12), pady=10)

        # Progress tipis
        self.progress = ctk.CTkProgressBar(self, height=6, corner_radius=4)
        self.progress.grid(row=2, column=0, padx=20, pady=12, sticky="ew")
        self.progress.set(0)

        # Tombol mulai
        self.start_btn = ctk.CTkButton(
            self, text="Mulai Update", height=38, corner_radius=8,
            font=ctk.CTkFont(size=13, weight="bold"),
            state="disabled", command=self.start_update
        )
        self.start_btn.grid(row=3, column=0, padx=20, pady=(0, 10), sticky="ew")

        # Log ringkas
        self.log_text = ctk.CTkTextbox(
            self, corner_radius=8,
            font=ctk.CTkFont(family="Consolas", size=10),
            state="disabled"
        )
        self.log_text.grid(row=4, column=0, padx=20, pady=(0, 16), sticky="nsew")

        self.after(100, self.process_queue)

    def log(self, msg):
        self.log_queue.put(msg)

    def process_queue(self):
        try:
            while True:
                msg = self.log_queue.get_nowait()
                self.log_text.configure(state="normal")
                self.log_text.insert("end", msg + "\n")
                self.log_text.see("end")
                self.log_text.configure(state="disabled")
        except queue.Empty:
            pass
        self.after(100, self.process_queue)

    def browse_file(self):
        path = filedialog.askopenfilename(
            title="Pilih File Excel",
            filetypes=[("Excel Files", "*.xlsb *.xlsm"), ("All Files", "*.*")]
        )
        if path and path.lower().endswith((".xlsb", ".xlsm")):
            self.target_file = os.path.abspath(path)
            self.file_entry.configure(state="normal")
            self.file_entry.delete(0, "end")
            self.file_entry.insert(0, self.target_file)
            self.file_entry.configure(state="disabled")
            self.start_btn.configure(state="normal")
            self.log(f"[OK] {os.path.basename(self.target_file)}")
        elif path:
            messagebox.showerror("Error", "File harus berformat .xlsb atau .xlsm")

    def start_update(self):
        if not self.target_file:
            return
        self.browse_btn.configure(state="disabled")
        self.start_btn.configure(state="disabled")
        self.progress.set(0)
        self.log_text.configure(state="normal")
        self.log_text.delete("0.0", "end")
        self.log_text.configure(state="disabled")
        threading.Thread(target=self._run_update, daemon=True).start()

    def _run_update(self):
        target_file = self.target_file
        source_dir = os.path.dirname(target_file)
        file_name = os.path.basename(target_file)
        target_year = CONFIG["TARGET_YEAR"]

        new_dir = os.path.join(source_dir, CONFIG["NEW_VERSION_FOLDER_NAME"])
        new_file = os.path.join(new_dir, file_name)

        temp_dir = tempfile.mkdtemp()

        try:
            self.log("[1/3] Cek versi terbaru...")
            version = get_latest_version_from_github(self.log)
            self.log(f"      Versi: {version}")
            self.after(0, lambda: self.progress.set(0.33))

            self.log("[2/3] Salin file ke folder versi terbaru...")
            os.makedirs(new_dir, exist_ok=True)
            shutil.copy2(target_file, new_file)
            self.after(0, lambda: self.progress.set(0.66))

            self.log("[3/3] Unduh, ekstrak & suntikkan update...")
            zip_path = os.path.join(temp_dir, "update.zip")
            urllib.request.urlretrieve(CONFIG["GITHUB_ZIP_URL"], zip_path)

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
            work_dir = os.path.join(temp_dir, "new_xlsb")
            with zipfile.ZipFile(new_file, 'r') as zf:
                zf.extractall(work_dir)

            for root, dirs, files in os.walk(source_update):
                for file in files:
                    full_path = os.path.join(root, file)
                    rel_path = os.path.relpath(full_path, source_update)
                    dest_path = os.path.join(work_dir, rel_path)
                    os.makedirs(os.path.dirname(dest_path), exist_ok=True)
                    shutil.copy2(full_path, dest_path)

            updated_temp = os.path.join(temp_dir, "updated.xlsb")
            with zipfile.ZipFile(updated_temp, 'w', zipfile.ZIP_DEFLATED) as out_zf:
                for root, dirs, files in os.walk(work_dir):
                    for file in files:
                        full_path = os.path.join(root, file)
                        arcname = os.path.relpath(full_path, work_dir)
                        out_zf.write(full_path, arcname)

            shutil.move(updated_temp, new_file)

            self.log("")
            self.log("[OK] Update selesai!")
            self.log(f"     {new_file}")
            self.after(0, lambda: self.progress.set(1.0))
            self.after(0, lambda: messagebox.showinfo("Sukses", "Update selesai dengan sukses!"))

        except Exception as e:
            self.log(f"[ERROR] {e}")
            self.after(0, lambda: messagebox.showerror("Gagal", f"Update gagal:\n{e}"))
        finally:
            shutil.rmtree(temp_dir, ignore_errors=True)
            self.after(0, self._enable_buttons)

    def _enable_buttons(self):
        self.browse_btn.configure(state="normal")
        self.start_btn.configure(state="normal")

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
    log_func(f"      [INFO] Versi default: {fallback}")
    return fallback

if __name__ == "__main__":
    app = UpdaterGUI()
    app.mainloop()
