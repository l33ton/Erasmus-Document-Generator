import tkinter as tk
from tkinter import filedialog, messagebox
import tkinter.ttk as ttk
from pathlib import Path
import pandas as pd
from pathvalidate import sanitize_filename

from services import LetterGenerator, convert_all_to_pdf
from updater import run_in_thread

class ErasmusApp:
    def __init__(self, root, base_dir):
        self.root = root
        self.base_dir = base_dir
        self.documents_generated = False

        self.root.geometry('600x500')
        self.root.title('Erasmus Document Generator')

        
        self.path_acs = tk.StringVar()
        self.path_acm = tk.StringVar()
        self.path_ga = tk.StringVar()
        self.path_noms = tk.StringVar()
        self.path_output = tk.StringVar()

        self.build_ui()
        self.load_defaults()

    def build_ui(self):
        # Header
        header = tk.Frame(self.root, bg='#003399', bd=2, relief=tk.RIDGE)
        header.pack(fill='x', padx=20, pady=20)
        label = tk.Label(header, text='Erasmus Document Generator', bg='#003399', fg='white', font=('Segoe UI', 14, 'bold'))
        label.pack(fill='x', pady=20)

        # Notebook (tabs)
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill='both', expand=True, padx=15, pady=5)

        tab_main = ttk.Frame(notebook)
        notebook.add(tab_main, text='Main work')

        tab_templates = ttk.Frame(notebook)
        notebook.add(tab_templates, text='Templates')

        # Templates Tab
        templates_frame = ttk.Labelframe(tab_templates, text='Templates')
        templates_frame.columnconfigure(1, weight=1)
        templates_frame.pack(fill='x', padx=15, pady=10)

        self.create_file_row(templates_frame, 0, 'Acceptance Letter', self.path_acs, lambda: self.select_file(self.path_acs, 'Choose Acceptance Letter', [("Word", "*.docx")]))
        self.create_file_row(templates_frame, 1, 'Accommodation Letter', self.path_acm, lambda: self.select_file(self.path_acm, 'Choose Accommodation Letter', [("Word", "*.docx")]))
        self.create_file_row(templates_frame, 2, 'Grant Agreement', self.path_ga, lambda: self.select_file(self.path_ga, 'Choose Grant Agreement', [("Word", "*.docx")]))

        # Main Tab - Files & Dirs
        files_frame = ttk.Labelframe(tab_main, text='Files and Directories')
        files_frame.columnconfigure(1, weight=1)
        files_frame.pack(fill='x', padx=15, pady=10)

        self.create_file_row(files_frame, 0, 'Nominations', self.path_noms, lambda: self.select_file(self.path_noms, 'Choose Nominations', [("Excel", "*.xls *.xlsx")]))
        self.create_file_row(files_frame, 1, 'Output Folder', self.path_output, lambda: self.select_folder(self.path_output, 'Choose Directory'))

        # Progress Bar
        progress_frame = ttk.Labelframe(tab_main, text='Progress and Controls')
        progress_frame.pack(fill='x', padx=15, pady=10)

        self.progress_bar = ttk.Progressbar(progress_frame, orient='horizontal', mode='determinate')
        self.progress_bar.pack(fill='x', padx=15, pady=5)

        self.status_label = ttk.Label(progress_frame, text='Ready to process..')
        self.status_label.pack(pady=5)

        self.generate_button = ttk.Button(progress_frame, text='Generate', command=self.start_generation_process)
        self.generate_button.pack(pady=10)

        self.convert_button = ttk.Button(progress_frame, text='Convert to PDF', command=self.start_pdf_conversion, state='disabled')
        self.convert_button.pack(pady=5)

    def create_file_row(self, parent, row, label_text, var, browse_func):
        ttk.Label(parent, text=label_text).grid(row=row, column=0, padx=5, pady=5, sticky='w')
        ttk.Entry(parent, textvariable=var).grid(row=row, column=1, padx=5, pady=5, sticky='ew')
        ttk.Button(parent, text='Browse', command=browse_func).grid(row=row, column=2, padx=5, pady=5)

    def select_file(self, target_var, title, file_types):
        path = filedialog.askopenfilename(title=title, filetypes=file_types)
        if path: target_var.set(path)

    def select_folder(self, target_var, title):
        path = filedialog.askdirectory(title=title)
        if path: target_var.set(path)

    def load_defaults(self):
        if (self.base_dir / 'AcceptanceLetterTemplate.docx').exists():
            self.path_acs.set(str(self.base_dir / "AcceptanceLetterTemplate.docx"))
        if (self.base_dir / 'AccommodationLetterTemplate.docx').exists():
            self.path_acm.set(str(self.base_dir / "AccommodationLetterTemplate.docx"))
        if (self.base_dir / 'GrantAgreementTemplate.docx').exists():
            self.path_ga.set(str(self.base_dir / "GrantAgreementTemplate.docx"))
        if (self.base_dir / 'Erasmus_Nominations_Data.xlsx').exists():
            self.path_noms.set(str(self.base_dir / "Erasmus_Nominations_Data.xlsx"))

    @run_in_thread
    def start_generation_process(self):
        nominations_path = Path(self.path_noms.get())
        output_dir = Path(self.path_output.get())
        acs_tpl_path = self.path_acs.get()
        acm_tpl_path = self.path_acm.get()
        ga_tpl_path = self.path_ga.get()

        if not nominations_path.is_file() or not (acs_tpl_path and acm_tpl_path and ga_tpl_path):
            self.root.after(0, lambda: messagebox.showwarning('Error', 'Please select valid templates and Excel file.'))
            return

        self.generate_button.config(state='disabled')
        self.convert_button.config(state='disabled')

        try:
            self.root.after(0, lambda: self.status_label.config(text='Loading templates...', foreground='blue'))
            acc_tpl = LetterGenerator(acs_tpl_path, 'Acceptance Letter')
            acm_tpl = LetterGenerator(acm_tpl_path, 'Accommodation Letter')
            ga_tpl = LetterGenerator(ga_tpl_path, 'Grant Agreement')

            df = pd.read_excel(nominations_path, sheet_name=0).fillna('')

            if 'StartDate' in df: df['StartDate'] = pd.to_datetime(df['StartDate']).dt.strftime('%d.%m.%Y')
            if 'EndDate' in df: df['EndDate'] = pd.to_datetime(df['EndDate']).dt.strftime('%d.%m.%Y')

            records = df.to_dict(orient='records')
            total = len(records)
            self.root.after(0, lambda: self.progress_bar.config(maximum=total, value=0))

            successful, failed = [], []

            for idx, record in enumerate(records, 1):
                full_name = str(record.get('FullName')).strip()
                valid_name = sanitize_filename(full_name)
                country = sanitize_filename(str(record.get('Country') or 'Unknown').strip())

                try:
                    self.root.after(0, lambda: self.status_label.config(text=f'Processing [{idx}/{total}]: {valid_name}'))
                    folder = output_dir / country / valid_name
                    folder.mkdir(parents=True, exist_ok=True)

                    acc_tpl.generate(record, folder)
                    acm_tpl.generate(record, folder)
                    ga_tpl.generate(record, folder)
                    successful.append(valid_name)
                except Exception as e:
                    failed.append({'name': full_name, 'row': idx, 'error': str(e)})
                finally:
                    self.root.after(0, lambda v=idx: self.progress_bar.config(value=v))

            if successful:
                self.documents_generated = True
                self.root.after(0, lambda: self.convert_button.config(state='normal'))

            if not failed:
                self.root.after(0, lambda: self.status_label.config(text='All documents generated successfully!', foreground='green'))
                self.root.after(0, lambda: messagebox.showinfo('Success', f'Generated documents for {total} participants.'))
            else:
                self.root.after(0, lambda: self.status_label.config(text='Completed with errors.', foreground='orange'))

        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror('Error', f'Unexpected error: {e}'))
        finally:
            self.root.after(0, lambda: self.generate_button.config(state='normal'))
            self.root.after(0, lambda: self.progress_bar.config(value=0))

    @run_in_thread
    def start_pdf_conversion(self):
        output_dir = Path(self.path_output.get())
        docx_files = list(output_dir.rglob('*.docx'))

        if not self.documents_generated or not docx_files:
            self.root.after(0, lambda: messagebox.showwarning('Error', 'No .docx files found to convert.'))
            return

        total = len(docx_files)
        self.root.after(0, lambda: self.progress_bar.config(maximum=total, value=0))
        self.root.after(0, lambda: self.status_label.config(text=f'Converting {total} files to PDF...', foreground='blue'))
        self.generate_button.config(state='disabled')
        self.convert_button.config(state='disabled')

        def update_progress(v):
            self.root.after(0, lambda: self.progress_bar.config(value=v))

        try:
            converted, errors = convert_all_to_pdf(output_dir, progress_callback=update_progress)
            if not errors:
                self.root.after(0, lambda: self.status_label.config(text='PDF conversion complete!', foreground='green'))
                self.root.after(0, lambda: messagebox.showinfo('Success', f'Converted {len(converted)} files.'))
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror('Error', f'PDF Error: {e}'))
        finally:
            self.root.after(0, lambda: self.generate_button.config(state='normal'))
            self.root.after(0, lambda: self.convert_button.config(state='normal'))
            self.root.after(0, lambda: self.progress_bar.config(value=0))