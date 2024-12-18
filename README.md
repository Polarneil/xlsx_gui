# Installation Instructions

1. Create an empty folder, open in VS Code or IDE of choice.

2. Create and activate a virtual environment in the folder:
### Mac OS
Create:
```bash
pyhton3 -m venv .venv
```
Activate:
```bash
source .venv/bin/activate
```
<br />

### Windows
Create:
```bash
pyhton3 -m venv .venv
```
Activate:
```bash
.venv/Scripts/Activate
```
<br />

3. Clone the repo:
```bash
git clone https://github.com/Polarneil/xlsx_gui.git
```
<br />

4. Navigate to the directory and install dependencies:
```bash
cd xlsx_gui
```
```bash
pip3 install -r requirements.txt
```
<br />

5. Run the GUI
```bash
python python/add_xlsx.py
```
<br />

6. View Results
    * Open the xlsx file created in the `xlsx` folder named `add_record.xlsx`