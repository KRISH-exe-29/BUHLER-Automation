📄 README.txt — Bühler Automation Script

📁 Folder: buhler_automation_script
👤 Author: Krishna Raju
📦 Tool: Python GUI + Excel Automation
🖥️ OS: Windows 10/11
📂 Requires: Python 3.12+

──────────────────────────────
📌 Contents
──────────────────────────────
1. Requirements
2. How to Run the Python GUI App
3. Optional: Build .exe for Distribution
4. Input File Format
5. Output File Details

──────────────────────────────
1. ✅ Requirements
──────────────────────────────

You must install Python 3.12 or later from:
→ https://www.python.org/downloads/

Then, open VS Code and install required packages:

In the VS Code terminal (Ctrl + `), run:

    pip install pandas openpyxl pillow xlsxwriter pyinstaller

──────────────────────────────
2. ▶️ Run the GUI App (Python)
──────────────────────────────

1. Open the folder:
   C:\Users\...\Documents\VS codes\buhler_automation_script

2. Open the script:
   automation_gui.py

3. Press ▶ Run or use terminal:

    python automation_gui.py

4. In the GUI:
   - Click “Browse Input File & Run”
   - Select a valid input Excel file
   - Choose an output folder
   - Click “OK” to process and export

──────────────────────────────
3. 📦 Build an EXE (Optional)
──────────────────────────────

To bundle the GUI as a standalone .exe:

1. Make sure these files exist in the folder:
   - automation_gui.py
   - buhler-case-study-banner-1024x137 (1).png
   - picture.ico (icon for the app)

2. Open terminal and run:

    pyinstaller --onefile --windowed ^
      --add-data "buhler-case-study-banner-1024x137 (1).png;." ^
      --icon=picture.ico ^
      automation_gui.py

3. After building, go to:
   dist/automation_gui.exe

You can now double-click this file and send it to anyone. No Python required to run!

──────────────────────────────
4. 📥 Input File Format
──────────────────────────────

The input Excel file must contain 3 sheets:

- Sheet1 → Starter KW mapping  
  Required columns: starters, KW

- Sheet2 → Equipment data  
  Required columns: Number, Machine, Power, Current, Voltage, Designation, Remarks, etc.

- Sheet3 → JB Location info  
  Required columns: Number, Schrank Aufstellungsort Feldkasten Stockwerk

Make sure the sheet names are exactly:
Sheet1, Sheet2, Sheet3

──────────────────────────────
5. 📤 Output File Contents
──────────────────────────────

The generated file (processed_*.xlsx) will contain:

✔️ Sheet1 — processed equipment list  
  Includes: Feeders, Types, calculated columns (8 DI, AO, c/o, etc.)

✔️ Summary — total load + IO card summary  
  Includes: Total Load (kW/A), spare calculations, number of cards

──────────────────────────────
📩 Need help?
──────────────────────────────

Contact Krishna Raju or share this folder along with README.txt for support.
