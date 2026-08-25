import pandas as pd
from pathlib import Path
from docxtpl import DocxTemplate
from pathvalidate import sanitize_filename
import win32com.client
import pythoncom

class LetterGenerator:
    def __init__(self, template_path, suffix):
        self.template_path = template_path
        self.suffix = suffix

    def generate(self, record, student_folder):
        doc = DocxTemplate(self.template_path)
        doc.render(record)
        file_name = f"{record['FullName']}'s {self.suffix}.docx"
        doc.save(student_folder / sanitize_filename(file_name))

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