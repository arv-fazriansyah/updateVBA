# Excel Standalone Updater (ARB 2026)

A secure, lightweight, and robust standalone Python-based updater for Excel macro-enabled workbooks (`.xlsb` / `.xlsm`). Designed to update enterprise applications seamlessly without killing unrelated Excel processes or requiring external utilities like 7-Zip.

---

## Key Features

1. **Targeted Excel Termination**: Only closes the specific target workbook being updated. All other open Excel spreadsheets and documents remain completely safe and untouched (no aggressive `taskkill /f /im excel.exe`).
2. **Zero External Dependencies**: Does not require 7-Zip (`7za.exe`) or any external binaries. File extraction and ZIP merging are handled natively by Python's built-in libraries.
3. **Flexible Execution Modes**:
   - **Automated via VBA**: Triggered directly from Excel VBA (passing absolute paths and version tags).
   - **Drag & Drop**: Simply drag and drop your Excel file (`.xlsb` / `.xlsm`) directly onto `updater.exe`.
   - **Smart Scan & Interactive Menu**: When double-clicked manually, it automatically scans common system directories (`Documents`, `Documents\excel_files`, `Downloads`) and presents an interactive prompt to select which application file to update.
4. **Automated Safety Backup**: Automatically generates backups inside a dedicated `BACKUP APLIKASI ARB 2026` folder prior to patching, with automated rollback on failure.
5. **Smart Versioning & Renaming**: Automatically handles version tagging prefixing (e.g., `v25.11.2025_MASTER_RBK2026.xlsb`).

---

## Project Structure

```text
excel-updater-standalone/
│
├── dist/
│   └── updater.exe       # Standalone executable (ready for distribution)
├── build/                # PyInstaller build cache
├── backup_local/         # Local backup archive directory
├── updater.py            # Main Python update engine source code
├── updater.spec          # PyInstaller specification file
└── README.md             # Project documentation
```

---

## How to Build from Source

If you need to recompile `updater.py` into a standalone `.exe` using PyInstaller:

```powershell
# Install PyInstaller (if not already installed)
pip install pyinstaller

# Build standalone executable (onefile, no console window)
python -m PyInstaller --onefile updater.py
```

The resulting executable will be located at `dist/updater.exe`.

---

## VBA Integration Example

To trigger the updater directly from your Excel VBA project:

```vba
Sub JalankanUpdaterExe()
    Dim wsh As Object, exePath As String, versionTag As String, targetFullPath As String
    
    versionTag = Trim(Range("DATAUSER!J2").Value) ' e.g., v25.11.2025
    targetFullPath = ThisWorkbook.FullName        ' Full absolute path of this workbook
    exePath = ThisWorkbook.Path & "\updater.exe"
    
    Dim cmd As String
    cmd = """" & exePath & """ """ & targetFullPath & """ """ & versionTag & """"
    
    Set wsh = CreateObject("WScript.Shell")
    wsh.Run cmd, 1, False
    
    ' Close current workbook safely without saving prompt
    ThisWorkbook.Saved = True
    ThisWorkbook.Close SaveChanges:=False
End Sub
```

---

## License

Internal Enterprise Tool - All rights reserved.
