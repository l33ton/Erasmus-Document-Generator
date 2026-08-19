import tempfile
import sys
from pathlib import Path
import pandas as pd
import subprocess
import requests
from docxtpl import DocxTemplate
from pathvalidate import sanitize_filename
import tkinter as tk
from tkinter import filedialog, messagebox
import tkinter.ttk as ttk
import hashlib
import win32com.client
import pythoncom
import functools
import threading

CURRENT_VERSION = 'v1.0.0'
GITHUB_USERNAME = 'l33ton'
GITHUB_REPO = 'Erasmus-Document-Generator'
documents_generated = False

def run_in_thread(func):

    @functools.wraps(func)  
    def wrapper(*args, **kwargs):
        thread = threading.Thread(
            target=func, args=args, kwargs=kwargs, daemon=True
        )
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
def _download_and_install(download_url, checksum_url):
  try:
    root.after(
        0,
        lambda: status_label.config(
            text='Installing the new version', foreground='#003399'
        ),
    )

    temp_setup = Path(tempfile.gettempdir()) / 'Update_Setup.exe'
    exe_data = requests.get(download_url, timeout=15).content
    expected_hash = (
        requests.get(checksum_url, timeout=10).text.strip().split()[0].lower()
    )

    with open(temp_setup, 'wb') as f:
      f.write(exe_data)

    actual_hash = calculate_sha256(temp_setup)

    if actual_hash != expected_hash:
      temp_setup.unlink(missing_ok=True)
      root.after(
          0,
          lambda: messagebox.showerror(
              'Security Error',
              'The file failed verification and was not installed.',
          ),
      )
      return

    def finalize():
      subprocess.Popen([str(temp_setup)])
      root.destroy()
      sys.exit()

    root.after(0, finalize)

  except Exception as e:
    root.after(
        0,
        lambda: messagebox.showerror('Update Failed', f'Error downloading: {e}'),
    )

@run_in_thread
def check_for_updates():
  try:
    api_url = f'https://api.github.com/repos/{GITHUB_USERNAME}/{GITHUB_REPO}/releases/latest'
    response = requests.get(api_url, timeout=3)

    if response.status_code == 200:
      latest_release = response.json()
      latest_version = latest_release.get('tag_name')

      if latest_version != CURRENT_VERSION:

        def ask_user():
          answer = messagebox.askyesno(
              'There is a new version',
              f'We found a new version {latest_version} (you are using'
              f' {CURRENT_VERSION}).\nDo you want to download it?',
          )
          if answer:
            assets = latest_release.get('assets', [])
            download_url = next(
                (a['browser_download_url'] for a in assets if a['name'].endswith('.exe')),
                None,
            )
            checksum_url = next(
                (a['browser_download_url'] for a in assets if a['name'].endswith('.sha256')),
                None,
            )

            if not download_url or not checksum_url:
              messagebox.showwarning(
                  'Error', 'Missing required update files on GitHub.'
              )
              return

            _download_and_install(download_url, checksum_url)

        root.after(0, ask_user)

  except Exception as e:
    print(f'Check for updates failed: {e}')
              
if getattr(sys, 'frozen', False):
    base_dir = Path(sys.executable).parent
else:
    base_dir = Path(__file__).parent

nominations = base_dir / "Erasmus_Nominations_Data.xlsx"
output_dir = base_dir / "generated_documents"

class LetterGenerator:
    def __init__(self, template_path, suffix):
        self.template_path = template_path
        self.suffix = suffix

    def generate(self, record, student_folder):
        doc = DocxTemplate(self.template_path)
        doc.render(record)
        file_name = f"{record['FullName']}'s {self.suffix}.docx"
        doc.save(student_folder / sanitize_filename(file_name))

def select_file(target_var, title, file_types):
    path = filedialog.askopenfilename(title=title, filetypes=file_types)
    if path:
        target_var.set(path)

def select_folder(target_var, title):
    path = filedialog.askdirectory(title=title)
    if path:
        target_var.set(path)

def create_file_row(parent, row, label_text, var, browse_func):
    lbl = ttk.Label(parent, text=label_text)
    lbl.grid(row=row, column=0, padx=5, pady=5, sticky='w')
    
    entry = ttk.Entry(parent, textvariable=var)
    entry.grid(row=row, column=1, padx=5, pady=5, sticky='ew')
    
    btn = ttk.Button(parent, text='Browse', command=browse_func)
    btn.grid(row=row, column=2, padx=5, pady=5)

@run_in_thread
def start_generation_process():
    global documents_generated
    nominations_path = Path(path_noms.get())
    output_dir = Path(path_output.get())
    acs_tpl_path = path_acs.get()
    acm_tpl_path = path_acm.get()
    ga_tpl_path = path_ga.get()

    if not nominations_path.is_file():
        root.after(0, lambda: messagebox.showwarning('Error', 'Please choose a valid Excel file with nominations'))
        return

    if not output_dir.exists():
        root.after(0, lambda: messagebox.showwarning('Error', 'Please choose a valid directory'))
        return
    
    if not (acs_tpl_path and acm_tpl_path and ga_tpl_path):
        root.after(0, lambda: messagebox.showwarning('Error', 'Please provide all the 3 templates'))
        return

    generate_button.config(state='disabled')
    convert_button.config(state='disabled')
    
    try:
        root.after(0, lambda: status_label.config(text='Loading the templates', foreground='blue'))
        acceptance_letter_template = LetterGenerator(acs_tpl_path, 'Acceptance Letter')
        accommodation_letter_template = LetterGenerator(acm_tpl_path, 'Accommodation Letter')
        grant_agreement_template = LetterGenerator(ga_tpl_path, 'Grant Agreement')

        nominations_reader = pd.read_excel(nominations_path, sheet_name=0)
        nominations_reader = nominations_reader.fillna('')
                   
        if 'StartDate' in nominations_reader:
            nominations_reader['StartDate'] = pd.to_datetime(nominations_reader['StartDate']).dt.strftime('%d.%m.%Y')
        if 'EndDate' in nominations_reader:
            nominations_reader['EndDate'] = pd.to_datetime(nominations_reader['EndDate']).dt.strftime('%d.%m.%Y')

        records = nominations_reader.to_dict(orient='records')
        total_participants = len(records)
        root.after(0, lambda: progress_bar.config(maximum=total_participants, value=0))
        successfull = []
        failed = []

        for idx, record in enumerate(records, 1):
            
            full_name = str(record.get('FullName')).strip()
            valid_full_name = sanitize_filename(full_name)    
            country = str(record.get('Country') or 'Unknown').strip() or 'Unknown'
            valid_country = sanitize_filename(country)

            try:                                 
                root.after(0, lambda: status_label.config(text=f'Processing [{idx}/{total_participants}]: {valid_full_name} ({valid_country})'))
                
                student_folder = output_dir / valid_country / valid_full_name
                student_folder.mkdir(parents=True, exist_ok=True)
        
                acceptance_letter_template.generate(record, student_folder)
                accommodation_letter_template.generate(record, student_folder)
                grant_agreement_template.generate(record, student_folder)

                successfull.append(valid_full_name)
                                        
            except Exception as e:                 
                 failed.append({'name': full_name, 'row': idx, 'error': str(e)})                                
                 continue
            finally:
                root.after(0, lambda v=idx: progress_bar.config(value=v))     
        if successfull:
            documents_generated = True
            root.after(0, lambda: convert_button.config(state='normal'))

        if not failed:
            root.after(0, lambda: status_label.config(text='All documents generated successfully!', foreground='green'))
            root.after(0, lambda: messagebox.showinfo('Success', f'Done! Documents generated for {total_participants} participants.')) 
        else:
            root.after(0, lambda: status_label.config(text='Documents generated with errors.', foreground='orange'))
            failed_details = '\n'.join(
            f"Row {item['row']}: {item['name']} - {item['error']}"
            for item in failed)
                                
            root.after(0, lambda: messagebox.showwarning('Completed with errors', f'Generated successfully: {len(successfull)}\n' f'Failed: {len(failed)}\n\n' f'Failed Participants: \n{failed_details}'))                                
            
    except Exception as e:
        root.after(0, lambda: status_label.config(text='Error with processing!', foreground='red'))
        root.after(0, lambda: messagebox.showerror('Error', f'An unexpected error occurred:\n{str(e)}'))

    finally:
        root.after(0, lambda: generate_button.config(state='normal')) 
        root.after(0, lambda: progress_bar.config(value=0))   

def convert_all_to_pdf(root_folder: Path, progress_callback=None):
    pythoncom.CoInitialize()
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False

    converted, errors = [], []
    docx_files = list(root_folder.rglob('*.docx'))
    try:
        for idx, docx_file in enumerate(docx_files, 1):
            try:
                doc = word.Documents.Open(str(docx_file))
                pdf_path = docx_file.with_suffix('.pdf')
                doc.SaveAs(str(pdf_path), FileFormat=17)
                doc.Close()
                converted.append(docx_file.name)
            except Exception as e:
                errors.append({'file': docx_file.name, 'error': str(e)})
            finally:
                if progress_callback:
                    progress_callback(idx)
    finally:
        word.Quit()
        pythoncom.CoUninitialize()

    return converted, errors

@run_in_thread
def start_pdf_conversion():

    if not documents_generated:
        root.after(0, lambda: messagebox.showwarning('Error', 'Please generate the documents first'))
        return

    output_dir = Path(path_output.get())

    if not output_dir.exists():
        root.after(0, lambda: messagebox.showwarning('Error', 'Please choose a valid output directory first'))
        return

    docx_files = list(output_dir.rglob('*.docx'))
    if not docx_files:
        root.after(0, lambda: messagebox.showwarning('Error', 'No .docx files found in the output directory'))
        return

    total_files = len(docx_files)
    
    root.after(0, lambda: progress_bar.config(maximum=total_files, value=0))
    root.after(0, lambda: generate_button.config(state='disabled'))
    root.after(0, lambda: convert_button.config(state='disabled'))
    root.after(0, lambda: status_label.config(text=f'Converting {len(docx_files)} document(s) to PDF...', foreground='blue'))

    def update_progress(current_val):
        root.after(0, lambda v=current_val: progress_bar.config(value=v))

    try:
        converted, pdf_errors = convert_all_to_pdf(output_dir, progress_callback=update_progress)

        if not pdf_errors:
            root.after(0, lambda: status_label.config(text='All PDFs generated successfully!', foreground='green'))
            root.after(0, lambda: messagebox.showinfo('Success', f'Converted {len(converted)} document(s) to PDF.'))
        else:
            root.after(0, lambda: status_label.config(text='PDF conversion completed with errors.', foreground='orange'))
            details = '\n'.join(f"{e['file']}: {e['error']}" for e in pdf_errors)
            root.after(0, lambda: messagebox.showwarning('Completed with errors',
                f'Converted: {len(converted)}\nFailed: {len(pdf_errors)}\n\n{details}'))
            
    except Exception as e:
        root.after(0, lambda: status_label.config(text='PDF conversion failed.', foreground='red'))
        root.after(0, lambda: messagebox.showerror('Error', f'An unexpected error occurred:\n{str(e)}'))

    finally:
        root.after(0, lambda: generate_button.config(state='normal'))
        root.after(0, lambda: convert_button.config(state='normal'))
        root.after(0, lambda: progress_bar.config(value=0))
               
root = tk.Tk()

root.geometry('600x500')
root.title('Erasmus Document Generator')

header = tk.Frame(root, bg='#003399', bd=2, relief=tk.RIDGE)
header.pack(fill='x', padx=20, pady=20)

label = tk.Label(header, 
                    text='Erasmus Document Generator', 
                    bg='#003399',
                    fg='white',
                    font=('Segoe UI', 14, 'bold'))
label.pack(fill='x', 
            pady=20)

path_acs = tk.StringVar()
path_acm = tk.StringVar()
path_ga = tk.StringVar()
path_noms = tk.StringVar()
path_output = tk.StringVar()

notebook = ttk.Notebook(root)
notebook.pack(fill='both', expand=True, padx=15, pady=5)

tab_main = ttk.Frame(notebook)
notebook.add(tab_main, text='Main work')

tab_templates = ttk.Frame(notebook)
notebook.add(tab_templates, text='Templates')

templates = ttk.Labelframe(tab_templates, text='Templates')
templates.columnconfigure(1, weight=1)
templates.pack(fill='x', padx=15, pady=10)

create_file_row(templates, 0, 'Acceptance Letter', path_acs, lambda: select_file(path_acs, 'Choose Acceptance Letter', [("Word", "*.docx")]))
create_file_row(templates, 1, 'Accommodation Letter', path_acm, lambda: select_file(path_acm, 'Choose Accommodation Letter', [("Word", "*.docx")]))
create_file_row(templates, 2, 'Grant Agreement', path_ga, lambda: select_file(path_ga, 'Choose Grant Agreement', [("Word", "*.docx")]))


files_and_dirs = ttk.Labelframe(tab_main, text='Files and Directories')
files_and_dirs.columnconfigure(1, weight=1)
files_and_dirs.pack(fill='x', padx=15, pady=10)

create_file_row(files_and_dirs, 0, 'Nominations', path_noms, lambda: select_file(path_noms, 'Choose Nominations', [("Excel", "*.xls *.xlsx")]))
create_file_row(files_and_dirs, 1, 'Output Folder', path_output, lambda: select_folder(path_output, 'Choose Directory'))

progress_frame = ttk.Labelframe(tab_main, text='Progress and Button')
progress_frame.columnconfigure(1, weight=1)
progress_frame.pack(fill='x', padx=15, pady=10)

progress_bar = ttk.Progressbar(progress_frame, orient='horizontal', mode='determinate')
progress_bar.pack(fill='x', padx=15, pady=5)

status_label = ttk.Label(progress_frame, text='Ready to process..')
status_label.pack(pady=5)

generate_button = ttk.Button(progress_frame, text='Generate', command=start_generation_process)
generate_button.pack(pady=10)

convert_button = ttk.Button(progress_frame, text='Convert to PDF', command=start_pdf_conversion, state='disabled')
convert_button.pack(pady=5)

if (base_dir / 'AcceptanceLetterTemplate.docx').exists():
    path_acs.set(str(base_dir / "AcceptanceLetterTemplate.docx"))
if (base_dir / 'AccommodationLetterTemplate.docx').exists():
    path_acm.set(str(base_dir / "AccommodationLetterTemplate.docx"))
if (base_dir / 'GrantAgreementTemplate.docx').exists():
    path_ga.set(str(base_dir / "GrantAgreementTemplate.docx"))
if (base_dir / 'Erasmus_Nominations_Data.xlsx').exists():
    path_noms.set(str(base_dir / "Erasmus_Nominations_Data.xlsx"))

if __name__ == "__main__":
    root.after(1000, check_for_updates)
    root.mainloop()