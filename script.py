import datetime
import sys
from pathlib import Path
import pandas as pd
from docxtpl import DocxTemplate

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

    def generate(self, record, template_path, student_folder):    
        doc = DocxTemplate(template_path)
        doc.render(record)
        file_name = f"{record['FullName']}'s Acceptance Letter.docx"
        doc.save(student_folder / file_name)

class AccommodationLetterGenerator:
    def __init__(self, template_path):
        self.template_path = template_path

    def generate(self, record, template_path, student_folder):    
        doc = DocxTemplate(template_path)
        doc.render(record)
        file_name = f"{record['FullName']}'s Accommodation Letter.docx"
        doc.save(student_folder / file_name)

class GrantAgreementGenerator:
    def __init__(self, template_path):
        self.template_path = template_path

    def generate(self, record, template_path, student_folder):    
        doc = DocxTemplate(template_path)
        doc.render(record)
        file_name = f"{record['FullName']}'s Grant Agreement.docx"
        doc.save(student_folder / file_name)


def main():
    if not nominations.exists:
        return f"Nominations does not exist{nominations.name}"

    acceptance_letter_template = AcceptanceLetterGenerator(base_dir / 'AcceptanceLetterTemplate.docx')
    accommodation_letter_template = AcceptanceLetterGenerator(base_dir / 'AccommodationLetterTemplate.docx')
    grant_agreement_template = GrantAgreementGenerator(base_dir / 'GrantAgreementTemplate.docx')

    nominations_reader = pd.read_excel(nominations, sheet_name="Sheet1")
    nominations_reader['StartDate'] = pd.to_datetime(nominations_reader['StartDate']).dt.strftime('%d.%m.%Y')
    nominations_reader['EndDate'] = pd.to_datetime(nominations_reader['EndDate']).dt.strftime('%d.%m.%Y')

    for record in nominations_reader.to_dict(orient="records"):
        student_folder = output_dir / f"{record['FullName']}"
        student_folder.mkdir(exist_ok=True)

        acceptance_letter_template.generate(record, acceptance_letter_template, student_folder)
        accommodation_letter_template.generate(record, accommodation_letter_template, student_folder)
        grant_agreement_template.generate(record, grant_agreement_template, student_folder)

if __name__ == "__main__":
    main()        