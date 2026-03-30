import tkinter as tk
from tkinter import filedialog
import requests
import os
import subprocess
import time
from urllib.parse import quote

def main():
    # 1. Select the file
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    root.destroy()
    
    if not file_path:
        return

    # Normalize path for Windows
    file_path = os.path.abspath(file_path).replace("\\", "/")
    
    # 2. Try to send to existing server
    success = False
    try:
        encoded_path = quote(file_path)
        response = requests.get(f"http://127.0.0.1:8000/trigger-file?path={encoded_path}", timeout=1)
        if response.status_code == 200:
            print("Successfully sent to running server.")
            success = True
    except:
        pass

    if not success:
        # 3. If server is not running, start it
        print("Server not responding. Starting it...")
        
        current_dir = os.path.dirname(os.path.abspath(__file__))
        app_exe = os.path.join(current_dir, "Leitor_PDF_Pro.exe")
        app_script = os.path.join(current_dir, "app.py")
        
        if os.path.exists(app_exe):
            # Prefer using the compiled executable if it exists
            print(f"Starting {app_exe}...")
            subprocess.Popen(f'"{app_exe}"', shell=True, creationflags=subprocess.CREATE_NEW_CONSOLE)
        elif os.path.exists(app_script):
            # Fallback to python script
            print(f"Starting {app_script} via python...")
            cmd = f'python "{app_script}"'
            subprocess.Popen(cmd, shell=True, creationflags=subprocess.CREATE_NEW_CONSOLE)
        else:
            print("ERROR: Could not find Leitor_PDF_Pro.exe or app.py")
            return
        
        # Wait for server to start and try to send the file
        for _ in range(15):
            time.sleep(1)
            try:
                encoded_path = quote(file_path)
                response = requests.get(f"http://127.0.0.1:8000/trigger-file?path={encoded_path}", timeout=0.5)
                if response.status_code == 200:
                    print("Successfully sent to new server instance.")
                    break
            except:
                continue

if __name__ == "__main__":
    main()
