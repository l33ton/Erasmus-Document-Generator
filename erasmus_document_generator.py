import sys
from pathlib import Path
import pandas as pd
from docxtpl import DocxTemplate
from pathvalidate import sanitize_filename
import tkinter as tk
from tkinter import filedialog, messagebox
import tkinter.ttk as ttk

base_dir = Path(__file__).parent

if getattr(sys, 'frozen', False):
    base_dir = Path(sys.executable).parent
else:
    base_dir = Path(__file__).parent

nominations = base_dir / "Erasmus_Nominations_Data.xlsx"
output_dir = base_dir / "generated_letters"

class AcceptanceLetterGenerator:
    def __init__(self, template_path):
        self.template_path = template_path       

    def generate(self, record, student_folder):    
        doc = DocxTemplate(self.template_path)
        doc.render(record)
        file_name = f"{record['FullName']}'s Acceptance Letter.docx"
        valid_file_name = sanitize_filename(file_name)
        doc.save(student_folder / valid_file_name)

class AccommodationLetterGenerator:
    def __init__(self, template_path):
        self.template_path = template_path

    def generate(self, record, student_folder):    
        doc = DocxTemplate(self.template_path)
        doc.render(record)
        file_name = f"{record['FullName']}'s Accommodation Letter.docx"
        valid_file_name = sanitize_filename(file_name)
        doc.save(student_folder / valid_file_name)

class GrantAgreementGenerator:
    def __init__(self, template_path):
        self.template_path = template_path

    def generate(self, record, student_folder):    
        doc = DocxTemplate(self.template_path)
        doc.render(record)
        file_name = f"{record['FullName']}'s Grant Agreement.docx"
        valid_file_name = sanitize_filename(file_name)
        doc.save(student_folder / valid_file_name)
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

def start_generation_proccess():
    print("--- Generate is clicked ---")
    nominations_path = Path(path_noms.get())
    output_dir = Path(path_output.get())
    acs_tpl_path = path_acs.get()
    acm_tpl_path = path_acm.get()
    ga_tpl_path = path_ga.get()


    if not nominations_path.is_file():
        messagebox.showwarning('Error', 'Please choose a valid Excel file with nominations')
        return

    if not output_dir.exists():
        messagebox.showwarning('Error', 'Please choose a valid directory')
        return
    if not (acs_tpl_path and acm_tpl_path and ga_tpl_path):
        messagebox.showwarning('Error', 'Please provide all the 3 templates')
        return

    status_label.config(text='Loading the templates', foreground='blue')
    root.update()

    try:
        acceptance_letter_template = AcceptanceLetterGenerator(acs_tpl_path)
        accommodation_letter_template = AccommodationLetterGenerator(acm_tpl_path)
        grant_agreement_template = GrantAgreementGenerator(ga_tpl_path)

        nominations_reader = pd.read_excel(nominations_path, sheet_name="Sheet1")
        nominations_reader = nominations_reader.dropna(how='all')

        if 'StartDate' in nominations_reader:
            nominations_reader['StartDate'] = pd.to_datetime(nominations_reader['StartDate']).dt.strftime('%d.%m.%Y')
        if 'EndDate' in nominations_reader:
            nominations_reader['EndDate'] = pd.to_datetime(nominations_reader['EndDate']).dt.strftime('%d.%m.%Y')

        records = nominations_reader.to_dict(orient='records')
        total_participants = len(records)
        for idx, record in enumerate(records, 1):
              
            full_name = str(record.get('FullName') or 'Unknown').strip() or 'Unknown'
            country = str(record.get('Country') or 'Unknown').strip() or 'Unknown'
            valid_full_name = sanitize_filename(full_name)
            valid_country = sanitize_filename(country)

            status_label.config(text=f'Processing [{idx}/{total_participants}]: {valid_full_name} ({valid_country})')
            root.update()

            try:    
                student_folder = output_dir / valid_country / valid_full_name
                student_folder.mkdir(parents=True, exist_ok=True)
        
                acceptance_letter_template.generate(record, student_folder)
                accommodation_letter_template.generate(record, student_folder)
                grant_agreement_template.generate(record, student_folder)

            except Exception as e:
                 with open(base_dir / "errors.log", "a", encoding="utf-8") as f:
                    f.write(f"Row: {record} -> Error: {e}\n")
                    continue
        status_label.config(text='All documents generated successfully!', foreground='green')
        messagebox.showinfo('Success', f'Done! Documents generated for {total_participants} students.') 

    except Exception as e:
        status_label.config(text='Error with processing!', foreground='red')
        messagebox.showerror('Error', f'An unexpected error occurred:\n{str(e)}')

               
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

status_label = ttk.Label(progress_frame, text='Ready to proccess..')
status_label.pack(pady=5)

generate_button = ttk.Button(progress_frame, text='Generate', command=start_generation_proccess)
generate_button.pack(pady=10)

if (base_dir / "AcceptanceLetterTemplate.docx").exists():
    path_acs.set(str(base_dir / "AcceptanceLetterTemplate.docx"))
if (base_dir / "AccommodationLetterTemplate.docx").exists():
    path_acm.set(str(base_dir / "AccommodationLetterTemplate.docx"))
if (base_dir / "GrantAgreementTemplate.docx").exists():
    path_ga.set(str(base_dir / "GrantAgreementTemplate.docx"))
if (base_dir / "Erasmus_Nominations_Data.xlsx").exists():
    path_noms.set(str(base_dir / "Erasmus_Nominations_Data.xlsx"))
path_output.set(str(base_dir / "generated_letters"))

if __name__ == "__main__":
        root.mainloop()