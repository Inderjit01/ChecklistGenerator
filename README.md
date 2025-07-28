# ChecklistGenerator

**ChecklistGenerator** is a desktop application built with PyQt5 that generates customized device checklists and customer documents for various hardware configurations. It provides a user-friendly GUI to select device types, enter serial and configuration information, and produce formatted Word documents based on predefined templates.

## Features

- GUI interface for selecting device type and entering required data
- Dynamic form fields that change based on device selection
- Generates two Word `.docx` documents:
  - Internal checklist
  - Customer information sheet
- Option to **save** or **save and print** both documents
- Automatically moves documents to the appropriate OneDrive folders
- Validates user input for format and completeness
- Sound and popup feedback on success/failure
- Built using PyQt5 and `python-docx`
- Packaged as a standalone executable using PyInstaller

---

## Folder Structure
ChecklistGenerator/
│
├── main.py # Entry point of the application
├── Application_GUI.py # GUI logic (PyQt5)
├── CheckList_Generator.py # Document generation logic
│
├── templates/ # Checklist document templates (.docx)
│ ├── LHC_Template.docx
│ └── ...
│
├── customerInfoTemplates/ # Customer document templates (.docx)
│ ├── LHC_CustomerTemplate.docx
│ └── ...
│
├── ssg_logo.png # App icon
├── leedle.mp3 # Alert sound

---

## Getting Started

### Requirements (for development)

- Python 3.8+
- PyQt5
- python-docx
- pywin32 (for Windows printing)
- PyInstaller (for building executable)

Install dependencies:

```bash
pip install -r requirements.txt
```

## Using the executable (for end users):
Run the generated .exe file.
Select your device type from the dropdown.
Fill in all required fields.
Click Save or Save and Print.
Documents are saved automatically to a OneDrive checklist folder.

## Building the Executable
```bash
pyinstaller --noconfirm --onefile --windowed --add-data "templates;templates" --add-data "customerInfoTemplates;customerInfoTemplates" --add-data "ssg_logo.png;." --add-data "leedle.mp3;." main.py
```
