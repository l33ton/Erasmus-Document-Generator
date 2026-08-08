import tkinter as tk
import tkinter.ttk as ttk
from tkinter import filedialog

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
create_file_row(files_and_dirs, 1, 'Outpud Folder', path_output, lambda: select_folder(path_output, 'Choose Directory'))



progress_frame = ttk.Labelframe(tab_main, text='Progress and Button')
progress_frame.columnconfigure(1, weight=1)
progress_frame.pack(fill='x', padx=15, pady=10)

status_label = ttk.Label(progress_frame, text='Ready to proccess..')
status_label.pack(pady=5)

generate_button = ttk.Button(progress_frame, text='Generate')
generate_button.pack(pady=10)

root.mainloop()