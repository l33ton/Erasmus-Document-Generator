# Erasmus Document Generator

**Erasmus Document Generator** is a professional desktop automation application designed to streamline the processing, generation, and conversion of student documents for **Erasmus+** educational exchange programs.

The application significantly reduces administrative workload by converting raw Excel datasets into formatted, personalized Word (`.docx`) and PDF documents for every participant.

---

## 🌟 Key Features

* 📄 **Automated Document Generation:** Batch-create *Acceptance Letters*, *Accommodation Letters*, and *Grant Agreements* populated with personalized student data.
* 📂 **Smart File Organization:** Automatically structures generated output files into a clean directory hierarchy categorized by country and student name (`Output/Country/Student_Name/`).
* 🔄 **One-Click PDF Conversion:** Integrated conversion engine transforming generated `.docx` templates into ready-to-print `.pdf` documents.
* ⚡ **Multithreaded Architecture:** Keeps the graphical user interface (GUI) responsive and smooth even during heavy batch operations.
* 🛡️ **Secure Auto-Update System:** Automated checks for new releases via GitHub API with SHA-256 cryptographic hash validation.
* 📊 **Real-time Visual Feedback:** Built-in progress bar and status feedback for error tracking and runtime execution logs.

---

## 📋 System Requirements

### Minimum Requirements:
* **Operating System:** Windows 10 / Windows 11 (64-bit)
* **Required Software:** Microsoft Word (necessary for the background `.docx` to `.pdf` rendering engine)
* **Input Data Format:** Valid `.xlsx` or `.xls` spreadsheet containing student nominations

---

## 🛠️ Installation & Developer Setup

To run or modify the source code locally:

### 1. Clone the Repository
```bash
git clone https://github.com/l33ton/Erasmus-Document-Generator.git
cd Erasmus-Document-Generator
```

### 2. Create a Virtual Environment (Recommended)
```bash
python -m venv venv
# On Windows:
venv\Scripts\activate
```

### 3. Install Required Dependencies
```bash
pip install -r requirements.txt
```

#### Dependencies (`requirements.txt`):
```text
pandas
openpyxl
docxtpl
pathvalidate
requests
pywin32
```

---

## 🚀 How to Use

1. **Prepare Templates:**
   Place your Word template files (`.docx`) containing Jinja2 placeholder tags (e.g., `{{ FullName }}`, `{{ StartDate }}`) into the **Templates** tab or rely on the defaults:
   * `AcceptanceLetterTemplate.docx`
   * `AccommodationLetterTemplate.docx`
   * `GrantAgreementTemplate.docx`

2. **Load Data Source:**
   Navigate to the **Main work** tab, select your Excel nominations dataset (`Erasmus_Nominations_Data.xlsx`), and select your desired target output directory.

3. **Generate Documents:**
   Click the **Generate** button. The application will organize folders and generate personalized Word documents.

4. **Convert to PDF:**
   Once generation finishes, click **Convert to PDF** to create print-ready PDF copies.

---

## 📦 Building Standalone `.exe` Executable

To compile a standalone distribution package for end users using `PyInstaller`:

```bash
pyinstaller --noconfirm --onedir --windowed --add-data "AcceptanceLetterTemplate.docx;." --add-data "AccommodationLetterTemplate.docx;." --add-data "GrantAgreementTemplate.docx;." ErasmusGenerator.py
```

---

## 📄 License & Copyright

© 2026 **l33ton**. All rights reserved.  
This software is provided under a proprietary commercial license. Unauthorized copying, distribution, or modification of the source code is strictly prohibited.
