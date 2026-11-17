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
        self.title(f"Updater Aplikasi ARB {CONFIG['TARGET_YEAR']}")
        self.geometry("760x560")
        self.resizable(False, False)

        self.log_queue = queue.Queue()
        self.target_file = None

        # Grid layout
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(4, weight=1)

        # Header card
        header = ctk.CTkFrame(self, corner_radius=16)
        header.grid(row=0, column=0, padx=24, pady=(24, 12), sticky="ew")
        header.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            header,
            text=f"UPDATER APLIKASI ARB {CONFIG['TARGET_YEAR']}",
            font=ctk.CTkFont(size=22, weight="bold")
        ).grid(row=0, column=0, sticky="w", padx=24, pady=(18, 4))

        ctk.CTkLabel(
            header,
            text="Pilih file Excel aplikasi, lalu klik Mulai Update",
            font=ctk.CTkFont(size=13),
            text_color="gray70"
        ).grid(row=1, column=0, sticky="w", padx=24, pady=(0, 18))

        # File selection card
        file_card = ctk.CTkFrame(self, corner_radius=16)
        file_card.grid(row=1, column=0, padx=24, pady=12, sticky="ew")
        file_card.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            file_card,
            text="FILE EXCEL TARGET",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color="gray70"
        ).grid(row=0, column=0, columnspan=2, sticky="w", padx=20, pady=(16, 6))

        self.file_entry = ctk.CTkEntry(
            file_card,
            placeholder_text="Belum ada file dipilih...",
            height=40,
            corner_radius=10,
            state="disabled"
        )
        self.file_entry.grid(row=1, column=0, sticky="ew", padx=(20, 8), pady=(0, 18))

        self.browse_btn = ctk.CTkButton(
            file_card,
            text="Pilih File",
            width=120,
            height=40,
            corner_radius=10,
            command=self.browse_file
        )
        self.browse_btn.grid(row=1, column=1, padx=(0, 20), pady=(0, 18))

        # Progress + status card
        status_card = ctk.CTkFrame(self, corner_radius=16)
        status_card.grid(row=2, column=0, padx=24, pady=12, sticky="ew")
        status_card.grid_columnconfigure(0, weight=1)

        self.status_label = ctk.CTkLabel(
            status_card,
            text="Siap untuk memulai update",
            font=ctk.CTkFont(size=13, weight="bold")
        )
        self.status_label.grid(row=0, column=0, sticky="w", padx=20, pady=(14, 8))

        self.progress = ctk.CTkProgressBar(status_card, height=10, corner_radius=6)
        self.progress.grid(row=1, column=0, sticky="ew", padx=20, pady=(0, 16))
        self.progress.set(0)

        # Start button
        self.start_btn = ctk.CTkButton(
            self,
            text="Mulai Update",
            height=48,
            corner_radius=12,
            font=ctk.CTkFont(size=15, weight="bold"),
            fg_color="#2563EB",
            hover_color="#1D4ED8",
            state="disabled",
            command=self.start_update
        )
        self.start_btn.grid(row=3, column=0, padx=24, pady=(4, 12), sticky="ew")

        # Log card
        log_card = ctk.CTkFrame(self, corner_radius=16)
        log_card.grid(row=4, column=0, padx=24, pady=(0, 24), sticky="nsew")
        log_card.grid_columnconfigure(0, weight=1)
        log_card.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(
            log_card,
            text="LOG PROSES",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color="gray70"
        ).grid(row=0, column=0, sticky="w", padx=20, pady=(16, 6))

        self.log_text = ctk.CTkTextbox(
            log_card,
            corner_radius=10,
            font=ctk.CTkFont(family="Consolas", size=11),
            state="disabled"
        )
        self.log_text.grid(row=1, column=0, sticky="nsew", padx=20, pady=(0, 18))

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
            title="Pilih File Excel Aplikasi",
            filetypes=[("Excel Files", "*.xlsb *.xlsm"), ("All Files", "*.*")]
        )
        if path and path.lower().endswith((".xlsb", ".xlsm")):
            self.target_file = os.path.abspath(path)
            self.file_entry.configure(state="normal")
            self.file_entry.delete(0, "end")
            self.file_entry.insert(0, self.target_file)
            self.file_entry.configure(state="disabled")
            self.start_btn.configure(state="normal")
            self.status_label.configure(text=f"File siap: {os.path.basename(self.target_file)}")
            self.log(f"[OK] File dipilih: {self.target_file}")
        elif path:
            messagebox.showerror("Error", "File harus berformat .xlsb atau .xlsm")

    def start_update(self):
        if not self.target_file:
            messagebox.showwarning("Peringatan", "Pilih file Excel terlebih dahulu.")
            return
        self.browse_btn.configure(state="disabled")
        self.start_btn.configure(state="disabled")
        self.progress.set(0)
        self.status_label.configure(text="Memulai proses update...")
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
            # Step 1
            self.after(0, lambda: self.status_label.configure(text="[1/4] Mengecek versi terbaru..."))
            self.log("\n[1/4] Mengecek versi terbaru...")
            version = get_latest_version_from_github(self.log)
            self.log(f"      Versi Target : {version}")
            self.after(0, lambda: self.progress.set(0.25))

            # Step 2
            self.after(0, lambda: self.status_label.configure(text="[2/4] Menyalin file ke folder backup & versi terbaru..."))
            self.log("\n[2/4] Menyalin file ke folder backup & versi terbaru...")
            os.makedirs(old_dir, exist_ok=True)
            os.makedirs(new_dir, exist_ok=True)
            shutil.copy2(target_file, old_file)
            shutil.copy2(target_file, new_file)
            self.log(f"      Versi lama  -> {CONFIG['OLD_VERSION_FOLDER_NAME']}")
            self.log(f"      Versi baru  -> {CONFIG['NEW_VERSION_FOLDER_NAME']}")
            self.after(0, lambda: self.progress.set(0.5))

            # Step 3
            self.after(0, lambda: self.status_label.configure(text="[3/4] Mengunduh & mengekstrak pembaruan..."))
            self.log("\n[3/4] Mengunduh & mengekstrak pembaruan dari GitHub...")
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
            self.log(f"      Sumber update: {target_folder_name}")
            self.after(0, lambda: self.progress.set(0.75))

            # Step 4
            self.after(0, lambda: self.status_label.configure(text="[4/4] Menyuntikkan pembaruan ke file versi terbaru..."))
            self.log("\n[4/4] Menyuntikkan pembaruan ke file di folder versi terbaru...")
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

            self.log("\n" + "=" * 60)
            self.log(" [OK] UPDATE SELESAI DENGAN SUKSES!")
            self.log(f" File versi terbaru: {new_file}")
            self.log("=" * 60)
            self.after(0, lambda: self.progress.set(1.0))
            self.after(0, lambda: self.status_label.configure(text="Update selesai dengan sukses!"))
            self.after(0, lambda: messagebox.showinfo(
                "Sukses",
                f"Update selesai!\nFile berhasil disimpan di:\n{new_dir}"
            ))

        except Exception as e:
            self.log(f"\n[ERROR] Update gagal: {e}")
            self.after(0, lambda: self.status_label.configure(text="Update gagal."))
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
    log_func(f"      [INFO] Menggunakan versi default: {fallback}")
    return fallback

if __name__ == "__main__":
    app = UpdaterGUI()
    app.mainloop()
