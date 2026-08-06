import sys
from pathlib import Path
import pandas as pd
from docxtpl import DocxTemplate
from pathvalidate import sanitize_filename
import tkinter as tk

window = tk.Tk()

window.geometry("1000x1000")
window.title("EDG")
window.iconbitmap('logo.ico')
window.config(background="black")

window.mainloop()

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

def main():
    if not nominations.exists():
        return f"Nominations does not exist{nominations.name}"

    acceptance_letter_template = AcceptanceLetterGenerator(base_dir / 'AcceptanceLetterTemplate.docx')
    accommodation_letter_template = AccommodationLetterGenerator(base_dir / 'AccommodationLetterTemplate.docx')
    grant_agreement_template = GrantAgreementGenerator(base_dir / 'GrantAgreementTemplate.docx')

    nominations_reader = pd.read_excel(nominations, sheet_name="Sheet1")
    nominations_reader = nominations_reader.dropna(how='all')
    nominations_reader['StartDate'] = pd.to_datetime(nominations_reader['StartDate']).dt.strftime('%d.%m.%Y')
    nominations_reader['EndDate'] = pd.to_datetime(nominations_reader['EndDate']).dt.strftime('%d.%m.%Y')

    for record in nominations_reader.to_dict(orient="records"):
        try:
            full_name = str(record.get('FullName') or 'Unknown').strip() or 'Unknown'
            country = str(record.get('Country') or 'Unknown').strip() or 'Unknown'
            valid_full_name = sanitize_filename(full_name)
            valid_country = sanitize_filename(country)
            student_folder = output_dir / valid_country / valid_full_name
            student_folder.mkdir(parents=True, exist_ok=True)
        
            acceptance_letter_template.generate(record, student_folder)
            accommodation_letter_template.generate(record, student_folder)
            grant_agreement_template.generate(record, student_folder)

        except Exception as e:
            with open(base_dir / "errors.log", "a", encoding="utf-8") as f:
                f.write(f"Row: {record} -> Error: {e}\n")
            continue


        root = tk.Tk()

        root.geometry("1000x1000")
        root.title("EDG")
        root.iconbitmap('logo.ico')
        root.config(background="black")

        root.mainloop()
        
if __name__ == "__main__":
    main()