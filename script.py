import datetime
from pathlib import Path
import pandas as pd
from docxtpl import DocxTemplate

base_dir = Path(__file__).parent

acceptance_letter_template_path = base_dir / "Acceptance_Letter_Template_Blank.docx"
nominations = base_dir / "Erasmus_Nominations_Data.xlsx"
output_dir = base_dir / "generated letters"

output_dir.mkdir(exist_ok=True)

doc = DocxTemplate(acceptance_letter_template_path)
data_sheet = pd.read_excel(nominations, sheet_name="Sheet1")

data_sheet['StartDate'] = pd.to_datetime(data_sheet['StartDate']).dt.strftime('%d.%m.%Y')
data_sheet['EndDate'] = pd.to_datetime(data_sheet['EndDate']).dt.strftime('%d.%m.%Y')

for record in data_sheet.to_dict(orient="records"):
    doc = DocxTemplate(acceptance_letter_template_path)
    doc.render(record) # type: ignore

    file_name = f"{record['FullName']}'s Acceptance Letter.docx"
    output_path = output_dir / file_name

    doc.save(output_path)