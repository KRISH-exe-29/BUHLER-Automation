# ⚙️ Bühler Group — Automation Script

> **Python GUI + Excel Automation Tool for Electrical Engineering Data Processing**

[![Live App](https://img.shields.io/badge/🚀%20Live%20App-Streamlit-FF4B4B?style=for-the-badge&logo=streamlit&logoColor=white)](https://bulhergroupofcompany.streamlit.app/)
[![Python](https://img.shields.io/badge/Python-3.12+-3776AB?style=for-the-badge&logo=python&logoColor=white)](https://www.python.org/)
[![Platform](https://img.shields.io/badge/Platform-Windows%2010%2F11-0078D6?style=for-the-badge&logo=windows&logoColor=white)](https://www.microsoft.com/windows)

---

## 🌐 About

This project is a **Python-based desktop GUI application** developed during an internship at **Bühler Group** — a leading global technology company for food and feed processing.

The tool automates the processing of **electrical engineering Excel data**, specifically handling equipment lists, starter KW mappings, and Junction Box (JB) location information. It calculates IO card requirements, load summaries, and outputs a fully structured Excel report.

🔗 **Live Web App**: [https://bulhergroupofcompany.streamlit.app/](https://bulhergroupofcompany.streamlit.app/)

---

## ✨ Features

- 📂 **GUI File Browser** — Select input Excel files via a clean Tkinter interface
- 📊 **Automated Data Processing** — Handles 3-sheet Excel inputs and calculates:
  - Feeder types and designations
  - 8 DI, AO, c/o card calculations
  - Total Load (kW / Amps)
  - Spare IO calculations
- 📤 **Structured Output** — Generates a `processed_*.xlsx` file with:
  - `Sheet1`: Processed equipment list
  - `Summary`: Total load + IO card summary
- 🖥️ **Standalone EXE** — Can be packaged as `.exe` (no Python required for end users)
- 🌐 **Streamlit Web App** — Interactive web version available online

---

## 🗂️ Project Structure

```
BUHLER/
├── buhler_automation_script/
│   ├── automation_gui.py          # Main GUI application
│   ├── Prepared_Input_File.xlsx   # Sample input file
│   ├── requirements.txt           # Python dependencies
│   ├── automation_gui.spec        # PyInstaller spec file
│   └── buhler-case-study-banner.png
├── automation_gui.exe             # Standalone executable
└── README.md
```

---

## 📥 Input File Format

The input Excel file must contain **3 sheets**:

| Sheet | Required Columns |
|-------|-----------------|
| `Sheet1` | `starters`, `KW` (Starter KW mapping) |
| `Sheet2` | `Number`, `Machine`, `Power`, `Current`, `Voltage`, `Designation`, `Remarks` |
| `Sheet3` | `Number`, `Schrank Aufstellungsort Feldkasten Stockwerk` (JB Location) |

> ⚠️ Sheet names must be exactly: `Sheet1`, `Sheet2`, `Sheet3`

---

## 📤 Output

The generated `processed_*.xlsx` file contains:

| Sheet | Contents |
|-------|----------|
| `Sheet1` | Processed equipment list with Feeders, Types, and calculated IO columns |
| `Summary` | Total Load (kW/A), spare calculations, number of IO cards |

---

## 🚀 Getting Started

### Prerequisites

- Python 3.12 or later → [Download](https://www.python.org/downloads/)

### Installation

```bash
# Clone the repository
git clone https://github.com/KRISH-exe-29/BUHLER-Automation.git
cd BUHLER-Automation

# Install dependencies
pip install -r buhler_automation_script/requirements.txt
```

### Run the App

```bash
python buhler_automation_script/automation_gui.py
```

### Using the GUI

1. Click **"Browse Input File & Run"**
2. Select a valid Excel input file
3. Choose an output folder
4. Click **"OK"** to process and export

---

## 📦 Build Standalone EXE (Optional)

```bash
pyinstaller --onefile --windowed \
  --add-data "buhler-case-study-banner-1024x137 (1).png;." \
  --icon=picture.ico \
  automation_gui.py
```

The compiled `.exe` will be in the `dist/` folder — no Python installation required to run!

---

## 🛠️ Tech Stack

| Technology | Purpose |
|------------|---------|
| Python 3.12 | Core language |
| Tkinter | Desktop GUI |
| Pandas | Data processing |
| OpenPyXL / XlsxWriter | Excel read/write |
| Streamlit | Web application |
| PyInstaller | EXE packaging |

---

## 👤 Author

**Krishna Raju S**
> Intern @ Bühler Group

---

## 📩 Support

For help or questions, contact **Krishna Raju** or open a GitHub Issue.
