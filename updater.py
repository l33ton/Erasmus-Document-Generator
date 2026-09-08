import hashlib
import sys
import tempfile
import subprocess
from pathlib import Path
import requests
from tkinter import messagebox
import functools
import threading

def run_in_thread(func):
    @functools.wraps(func)
    def wrapper(*args, **kwargs):
        thread = threading.Thread(target=func, args=args, kwargs=kwargs, daemon=True)
        thread.start()
        return thread
    return wrapper

def calculate_sha256(file_path):
    sha256_hash = hashlib.sha256()
    with open(file_path, "rb") as f:
        for chunk in iter(lambda: f.read(8192), b""):
            sha256_hash.update(chunk)
    return sha256_hash.hexdigest()

@run_in_thread
def download_and_install(root, status_label, download_url, checksum_url):
  try:
    root.after(
        0,
        lambda: status_label.config(
            text='Installing the new version...', foreground='#003399'
        ),
    )
    temp_setup = Path(tempfile.gettempdir()) / 'Erasmus_Generator_new.exe'
    exe_data = requests.get(download_url, timeout=15).content
    raw_hash_text = requests.get(checksum_url, timeout=10).text.strip()
    expected_hash = raw_hash_text.split()[0].lower().strip()

    with open(temp_setup, 'wb') as f:
      f.write(exe_data)

    if calculate_sha256(temp_setup) != expected_hash:
      temp_setup.unlink(missing_ok=True)
      root.after(
          0,
          lambda: messagebox.showerror(
              'Security Error', 'File failed verification.'
          ),
      )
      return

    def finalize():
          current_exe = Path(sys.executable)

          if not getattr(sys, 'frozen', False):
            subprocess.Popen([str(temp_setup)])
            root.destroy()
            sys.exit()
            return

          bat_path = Path(tempfile.gettempdir()) / 'update_replacer.bat'
          exe_name = current_exe.name
          exe_dir = current_exe.parent

          bat_content = f"""@echo off
                        taskkill /IM "{exe_name}" /F >nul 2>&1
                        timeout /t 3 /nobreak > nul
                        copy /y "{temp_setup}" "{current_exe}"
                        del "{temp_setup}"
                        start "" /d "{exe_dir}" "{current_exe}"
                        del "%~f0"
                        """
          with open(bat_path, 'w', encoding='utf-8') as f:
            f.write(bat_content)

          subprocess.Popen(
              ['cmd.exe', '/c', str(bat_path)],
              creationflags=subprocess.CREATE_NO_WINDOW,
          )

          root.destroy()
          sys.exit()

    root.after(0, finalize)

  except Exception as e:
    root.after(
        0, lambda: messagebox.showerror('Update Failed', f'Error: {e}')
    )

@run_in_thread
def check_for_updates(root, status_label, current_version, username, repo):
  try:
    api_url = (
        f'https://api.github.com/repos/{username}/{repo}/releases/latest'
    )
    response = requests.get(api_url, timeout=3)

    if response.status_code == 200:
      latest_release = response.json()
      latest_version = latest_release.get('tag_name')

      if latest_version != current_version:
        
        assets = latest_release.get('assets', [])
        dl_url = next(
            (
                a['browser_download_url']
                for a in assets
                if a['name'].endswith('.exe')
            ),
            None,
        )
        cs_url = next(
            (
                a['browser_download_url']
                for a in assets
                if a['name'].endswith('.sha256')
            ),
            None,
        )
        
        if dl_url and cs_url:

          def ask_user():
            answer = messagebox.askyesno(
                'New Version Available',
                f'Found new version {latest_version} (current:'
                f' {current_version}). Do you want to download it?',
            )
            if answer:
              download_and_install(root, status_label, dl_url, cs_url)

          
          root.after(0, ask_user)
  
  except Exception as e:
            print(f'Check for updates failed: {e}')