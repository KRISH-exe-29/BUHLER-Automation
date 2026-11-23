# ⚙️ Bühler Automation Script

Turns a Bühler electrical equipment list in Excel into feeder types, total load and PLC I/O card counts in one run, instead of working them out by hand. Built during an internship at Bühler Group.

[![Live App](https://img.shields.io/badge/🚀%20Live%20App-Streamlit-FF4B4B?style=for-the-badge&logo=streamlit&logoColor=white)](https://bulhergroupofcompany.streamlit.app/)
[![Python](https://img.shields.io/badge/Python-3.12+-3776AB?style=for-the-badge&logo=python&logoColor=white)](https://www.python.org/)
[![Platform](https://img.shields.io/badge/Platform-Windows%2010%2F11-0078D6?style=for-the-badge&logo=windows&logoColor=white)](https://www.microsoft.com/windows)

> A newer browser version, with junction-box mapping and validation, lives at **[buhler-automation-app-](https://github.com/KRISH-exe-29/buhler-automation-app-)** ([live demo](https://buhler-feeder-planner.vercel.app)).

- Works out the feeder type and designation for every piece of equipment.
- Counts 8 DI, AO and changeover cards, including spare I/O.
- Totals the connected load in kW and amps.
- Writes a `processed_<input name>.xlsx` with the full list plus a summary sheet.
- Packages as a standalone `.exe`, so colleagues can run it without Python.

## Quick start

```bash
git clone https://github.com/KRISH-exe-29/BUHLER-Automation.git && cd BUHLER-Automation
pip install -r buhler_automation_script/requirements.txt
python buhler_automation_script/automation_gui.py
```

Click **Browse Input File & Run**, pick the sample `buhler_automation_script/Prepared_Input_File.xlsx`, choose an output folder and click **OK**. The result is `processed_Prepared_Input_File.xlsx` in that folder.

---

## 📥 Input file format

The input Excel file must contain **3 sheets**:

| Sheet | Required Columns |
|-------|-----------------|
| `Sheet1` | `starters`, `KW` (Starter KW mapping) |
| `Sheet2` | `Number`, `Machine`, `Power`, `Current`, `Voltage`, `Designation`, `Remarks` |
| `Sheet3` | `Number`, `Schrank Aufstellungsort Feldkasten Stockwerk` (JB Location) |

> ⚠️ Sheet names must be exactly: `Sheet1`, `Sheet2`, `Sheet3`

## 📤 Output

| Sheet | Contents |
|-------|----------|
| `Sheet1` | Processed equipment list with Feeders, Types, and calculated IO columns |
| `Summary` | Total Load (kW/A), spare calculations, number of IO cards |

## 🗂️ Project structure

```
BUHLER-Automation/
├── buhler_automation_script/
│   ├── automation_gui.py          # Main GUI application
│   ├── Prepared_Input_File.xlsx   # Sample input file
│   ├── requirements.txt           # Python dependencies
│   └── picture.ico                # App icon
└── README.md
```

## 📦 Build a standalone EXE (optional)

```bash
pyinstaller --onefile --windowed \
  --add-data "buhler-case-study-banner-1024x137 (1).png;." \
  --icon=picture.ico \
  automation_gui.py
```

The compiled `.exe` will be in the `dist/` folder, and it runs without Python installed.

## 🛠️ Tech stack

| Technology | Purpose |
|------------|---------|
| Python 3.12 | Core language |
| Tkinter | Desktop GUI |
| Pandas | Data processing |
| OpenPyXL / XlsxWriter | Excel read/write |
| Streamlit | Web application |
| PyInstaller | EXE packaging |

## 👤 Author

**Krishna Raju S**, intern at Bühler Group. For help or questions, open a GitHub issue.
