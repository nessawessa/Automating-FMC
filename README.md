# Automating-FMC
Repository holding separate scripts and FM&amp;C mod template, aiming for script unification to achieve a streamlined CLI execution.

## 📌 Features
- **Excel Integration**: Reads data from multiple FM&amp;C Modification Template (excel file)
- **Bidirectional Updates**: Writes created IDs and statuses back into the Excel template.
- **RV&S Integration**: Executes commands like:
  - Create Failure Modes
  - Create Causes
  - Revise Causes
  - (other functions pending)
- **Validation**:
  - Ensures proper Excel structure and data completeness.
  - Checks RV&S connectivity before execution.
- **Dry-Run Support**: Simulate operations without making changes to RV&S (future implementation).

---

## ✅ Prerequisites
- **Python 3.10+**
- Required libraries:
  ```bash
  pip install openpyxl pandas numpy

## Clone this repository
    git clone https://github.com/<nessawessa>/automating-fmc.git
    
    cd automating-fmc

## Current usage in CLI
Create Fail modes 

    python CreateFailModes.py
          
Create Causes 

    python CreateCauses.py

Revise Causes

    python ReviseCauses.py


# File path mapping - Where to update in code

## for CreateFailModes.py

update excel path in two places:

wb_obj = openpyxl.load_workbook(r"C:\Users\aq34o\Documents\Automating FM&C\python scripts\FM&C Modification Template.xlsx")

wb_obj.save(r"C:\Users\aq34o\Documents\Automating FM&C\python scripts\FM&C Modification Template.xlsx")

## for ReviseCauses.py
update vaiable:

EXCEL_PATH = r"C:\Users\aq34o\Documents\Automating FM&C\python scripts\FM&C Modification Template.xlsx"

## for CreateFailModes.py
update excel path in two places:

wb_obj = openpyxl.load_workbook(r"C:\Users\aq34o\Documents\Automating FM&C\python scripts\FM&C Modification Template.xlsx")
wb_obj.save(r"C:\Users\aq34o\Documents\Automating FM&C\python scripts\FM&C Modification Template.xlsx")

## Known Limitations

Currently supports Windows environment only.

RV&S CLI must be pre-configured.

Excel template structure must not be modified.

    
