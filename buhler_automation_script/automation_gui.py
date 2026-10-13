import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk
import pandas as pd
import numpy as np
import os
import math
import xlsxwriter

def process_excel(input_path, output_path, open_after):
    try:
        df_starter = pd.read_excel(input_path, sheet_name='Sheet1')
        df_data = pd.read_excel(input_path, sheet_name='Sheet2')
        df_jb = pd.read_excel(input_path, sheet_name='Sheet3')

        df_data = df_data.merge(
            df_jb[['Number', 'Schrank Aufstellungsort Feldkasten Stockwerk']],
            on='Number', how='left'
        )
        df_data.rename(columns={'Schrank Aufstellungsort Feldkasten Stockwerk': 'JB'}, inplace=True)

        ranges = df_starter[df_starter['starters'].isin(['DOL', 'SD', 'SS'])]['KW']
        Range = ranges.tolist() if len(ranges) >= 2 else [0, 10]

        def pow(row):
            k = str(row['Power']).strip()
            if k in ["---", ""]:
                return 0
            if k[-1] == 'W':
                return pd.to_numeric(k[:-1], errors='coerce') / 1000
            elif k[-1] == 'A':
                return pd.to_numeric(k[:-2], errors='coerce') / 1000
            return pd.to_numeric(k, errors='coerce')

        def cat(row):
            n = str(row['Number'])
            m = str(row['Machine'])
            three_char = n[-5:-2]
            if three_char == 'MXZ':
                return 'M'
            if three_char in ['BPS', 'SHE', 'PAZ', 'BLH', 'BQS', 'SCS', 'EXL', 'BFT', 'BSA', 'BZS', 'QXV', 'BLL', 'XCJ']:
                return 'MA' if three_char in ['XCJ', 'EXL', 'QXV'] and m[:4] in ['MVSL', 'MQRH', 'MVRT', 'MQRF'] else 'A'
            if three_char in ['QYS', 'SHC', 'PWC', 'BZA', 'KCU', 'BST', 'BTT', 'BRE', 'KCL', 'KFC', 'RIN', 'BWE']:
                return 'MA'
            return ''

        def cat1(row):
            n = str(row['Number'])
            m = str(row['Machine'])
            power_value = row['X']
            three_char = n[-5:-2]
            remarks = row['Remarks']
            if power_value <= Range[0] and (three_char == 'MXZ' and remarks not in ['SC-S-P=BA','SC-S-P','SC-S-M=BM','SC-S-P=BM']):
                if m[:4] in ['MTRC', 'MTRA', 'MQRH', 'MQRF']:
                    return 'RDDOL'
                elif m[:4] in ['MTSD', 'MPAV','MPAW','MPAU']:
                    return 'RDOL'
                elif m == 'FAN':
                    return 'DOL_F'
                else:
                    return 'DOL'
            elif power_value >= Range[1] and (three_char == 'MXZ' and remarks not in ['SC-S-P=BA','SC-S-P','SC-S-M=BM','SC-S-P=BM']):
                return 'SS'
            elif three_char == 'MXZ' and remarks in ['SC-S-P=BA','SC-S-P','SC-S-M=BM','SC-S-P=BM']:
                return 'VFD_VA=BA' if remarks == 'SC-S-P=BA' else 'VFD_VA=GB'
            return 'NA'

        def cat2(row):
            m = str(row['Machine'])
            n = str(row['Number'])
            z = str(row.get('Designation', ''))
            three_char = n[-5:-2]
            if three_char == 'MXZ' and m[:4] != 'MDDP':
                return '1', '1', '1'
            if three_char == 'MXZ' and m[:4] == 'MDDP' and z[7:9] == 'KA':
                return '1', '1', '1'
            if three_char == 'MXZ' and m[:4] == 'MDDP' and z[7:9] == 'GM':
                return '1', '', ''
            if three_char == 'KCU' and m[:4] in ['LBEB', 'MVRT']:
                return '1', '', ''
            return '', '', ''

        def cat3(row):
            m = str(row['Machine'])
            z = str(row.get('Designation', ''))
            remarks = row['Remarks']
            three_char = row['Number'][-5:-2]
            if three_char == 'MXZ':
                if m[:4] == 'MDDP' and z[7:9] == 'KA':
                    return '', '-'
                elif m[:4] in ['MRRK', 'MRRY'] and z[7:9] == 'KA':
                    return '1', ''
                elif m[:4] in ['MZDH','MZMA','MZDG'] and remarks in ['SC-S-P=BA','SC-S-P','SC-S-M=BM','SC-S-P=BM']:
                    return '-', '1'
                elif m[:4] == 'MPSN' and  remarks in ['SC-S-P=BA','SC-S-P','SC-S-M=BM','SC-S-P=BM']:
                    return '', '1'
            return '', ''

        def cat4(row):
            m = str(row['Machine'])
            three_char = row['Number'][-5:-2]
            if three_char in ['BPS', 'BLH', 'BSA', 'BZS', 'BQS', 'QXV']:
                if three_char == 'QXV' and m[:4] != 'MVRT':
                    return '', '1'
                elif m[:4] == 'MVRT':
                    return '', ''
                else:
                    return '1', ''
            return '', ''

        def cat5(row):
            three_char = row['Number'][-5:-2]
            return '1' if three_char in ['MXZ', 'PWC'] else ''

        df_data['X'] = df_data.apply(pow, axis=1)
        df_data['C'] = df_data.apply(cat, axis=1)
        df_data['D'] = df_data.apply(cat1, axis=1)
        df_data[['P', 'Q', 'V']] = df_data.apply(cat2, axis=1, result_type='expand')
        df_data[['T', 'U']] = df_data.apply(cat3, axis=1, result_type='expand')
        df_data[['R', 'S']] = df_data.apply(cat4, axis=1, result_type='expand')               
        df_data['W'] = df_data.apply(cat5, axis=1)
        df_data['8 DI'] = pd.to_numeric(df_data['P'], errors='coerce').fillna(0)
        df_data['8 DO'] = pd.to_numeric(df_data['Q'], errors='coerce').fillna(0)
        df_data['4 DI'] = pd.to_numeric(df_data['R'], errors='coerce').fillna(0)
        df_data['4 DO'] = pd.to_numeric(df_data['S'], errors='coerce').fillna(0)
        df_data['AI'] = pd.to_numeric(df_data['T'], errors='coerce').fillna(0)
        df_data['AO'] = pd.to_numeric(df_data['U'], errors='coerce').fillna(0)
        df_data['c/o'] = pd.to_numeric(df_data['V'], errors='coerce').fillna(0)
        df_data['S/S'] = pd.to_numeric(df_data['W'], errors='coerce').fillna(0)

        total_power = df_data['X'].sum()
        total_current = pd.to_numeric(df_data['Current'], errors='coerce').fillna(0).sum()

        df_output = df_data[['Sl.No', 'Rev No', 'Number', 'Machine', 'Alarm message', 'Designation', 'Make', 'X', 'Power', 'Current', 'Voltage', 'C', 'D', 'JB', 'Buhler standard', 'Bldg.','Level','Remarks','8 DI','8 DO','4 DI','4 DO','AI','AO','c/o','S/S']]

        df_output.rename(columns={'C': 'Types', 'D': 'Feeders'}, inplace=True)

        summary = []
        signal_map = {
            '8 DI': ('DI-8ch', 8),
            '8 DO': ('DO_8ch', 8),
            '4 DI': ('DI-4ch', 4),
            '4 DO': ('DO-4ch', 4),
            'AI': ('AI-4ch', 4),
            'AO': ('AO-4ch', 4),
            'c/o': ('1C/O', 1),
            'S/S': ('SS', 1)
        }

        for col, (label, ch) in signal_map.items():
            total = df_output[col].sum()
            spare = total * 0.10
            grand = total + spare
            cards = total / ch if ch > 0 else 0
            consider = int(np.ceil(cards)) if ch > 0 else grand
            summary.append([label, total, spare, grand, cards, consider])

        summary_df = pd.DataFrame(summary, columns=["Type", "Total", "Spare 10%", "Grand Total", "Cards", "Consider"])
        load_summary = pd.DataFrame({"Type": ["Total Load", "Incomer"], "Total": [f"{int(total_power)}kW, {int(total_current)}A", "630A"]})

        with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
            df_output.to_excel(writer, sheet_name='Sheet1', index=False)
            load_summary.to_excel(writer, sheet_name='Summary', index=False, startrow=0)
            summary_df.to_excel(writer, sheet_name='Summary', index=False, startrow=4)

        if open_after:
            os.startfile(output_path)

        messagebox.showinfo("Success", f"Processed and saved to:\n{output_path}")

    except Exception as e:
        messagebox.showerror("Error", str(e))


def browse_and_run():
    file_path = filedialog.askopenfilename(title="Select Input Excel File", filetypes=[("Excel files", "*.xlsx")])
    if not file_path:
        return
    output_dir = filedialog.askdirectory(title="Select Output Folder")
    if not output_dir:
        return
    base_name = os.path.basename(file_path)
    output_path = os.path.join(output_dir, f"processed_{base_name}")
    process_excel(file_path, output_path, open_checkbox_var.get())


# GUI Setup
root = tk.Tk()
root.title("Bühler Automation Script")
root.geometry("520x380")
root.configure(bg="#E6F0FA")

# Load Bühler Logo
logo_path = "buhler-case-study-banner-1024x137 (1).png"
if os.path.exists(logo_path):
    logo_img = Image.open(logo_path)
    logo_img = logo_img.resize((480, 60))
    logo_img = ImageTk.PhotoImage(logo_img)
    logo_label = tk.Label(root, image=logo_img, bg="#E6F0FA")
    logo_label.image = logo_img
    logo_label.pack(pady=(10, 5))

tk.Label(root, text="Excel Processing Utility", font=("Segoe UI", 14, "bold"), bg="#E6F0FA", fg="#003366").pack(pady=5)
tk.Button(root, text="Browse Input File & Run", command=browse_and_run, bg="#0072C6", fg="white", font=("Segoe UI", 11), width=30).pack(pady=10)

open_checkbox_var = tk.BooleanVar()
tk.Checkbutton(root, text="Open file after save", variable=open_checkbox_var, bg="#E6F0FA", font=("Segoe UI", 10)).pack()

tk.Label(root, text="© Bühler Automation", font=("Segoe UI", 9), bg="#E6F0FA", fg="#666666").pack(side="bottom", pady=10)
root.mainloop()
