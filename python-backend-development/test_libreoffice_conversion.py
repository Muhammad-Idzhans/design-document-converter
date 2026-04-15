import os
import shutil
import tempfile
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox

# -----------------------------
# LibreOffice Path Detection
# -----------------------------
def find_soffice():
    """Locate LibreOffice soffice binary. Works on both Linux and Windows."""
    # 1. Check if soffice is on PATH
    found = shutil.which("soffice")
    if found:
        return found

    # 2. Windows: check common installation paths
    if os.name == "nt":
        candidates = [
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        ]
        for candidate in candidates:
            if os.path.isfile(candidate):
                return candidate

    raise RuntimeError(
        "LibreOffice 'soffice' not found. "
        "Install LibreOffice and ensure it is on your PATH or in a standard location."
    )

def make_user_profile():
    """Create a temporary LibreOffice user profile directory."""
    return tempfile.mkdtemp(prefix="lo_profile_test_")

# -----------------------------
# Conversion Logic
# -----------------------------
def convert_docx_to_pdf(docx_path: str):
    print(f"[*] Starting conversion for: {docx_path}")
    abs_docx = os.path.abspath(docx_path)
    out_dir = os.path.dirname(abs_docx)
    base_name = os.path.splitext(os.path.basename(abs_docx))[0]
    expected_pdf = os.path.join(out_dir, base_name + ".pdf")
    
    try:
        soffice = find_soffice()
        print(f"[*] Found LibreOffice at: {soffice}")
    except Exception as e:
        print(f"[!] Error: {e}")
        return

    user_profile = make_user_profile()
    
    print("[*] Running LibreOffice Headless...")
    try:
        # Run the exact same subprocess command as slides-to-doc-fastapi.py
        result = subprocess.run(
            [soffice, "--headless", "--norestore", "--nologo",
             f"-env:UserInstallation=file:///{user_profile.replace(os.sep, '/')}",
             "--convert-to", "pdf",
             "--outdir", out_dir, abs_docx],
            check=True, timeout=180,
            stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True
        )
        
        if os.path.exists(expected_pdf):
            print(f"[+] SUCCESS! PDF saved to: {expected_pdf}")
            # Open the generated file (Windows specific)
            if os.name == "nt":
                os.startfile(expected_pdf)
        else:
            print("[-] LibreOffice ran successfully, but the PDF file was not created.")
            if result.stderr:
                print(f"    LibreOffice Error Output:\n{result.stderr}")
            
    except subprocess.TimeoutExpired:
        print("[-] Conversion timed out after 180 seconds.")
    except subprocess.CalledProcessError as e:
        print(f"[-] Conversion failed with exit code {e.returncode}.")
        if e.stderr:
            print(f"    Error Output:\n{e.stderr}")
    except Exception as e:
        print(f"[-] Unexpected Error: {e}")
    finally:
        shutil.rmtree(user_profile, ignore_errors=True)
        print("[*] Temporary LibreOffice profile cleaned up.\n")


# -----------------------------
# Entry Point (UI)
# -----------------------------
if __name__ == "__main__":
    print(r"""
=============================================
|      LibreOffice Local PDF Test Tool      |
=============================================
This tool perfectly mimics the exact LibreOffice
architecture running inside your Azure Docker.
    """)
    
    # Hide the main tkinter root window
    root = tk.Tk()
    root.withdraw()
    
    # Ask the user to select a DOCX file
    file_path = filedialog.askopenfilename(
        title="Select a Word Document (.docx) to test",
        filetypes=[("Word Documents", "*.docx"), ("All Files", "*.*")]
    )
    
    if file_path:
        convert_docx_to_pdf(file_path)
    else:
        print("[-] No file was selected. Exiting.")
