import eel
import pandas as pd
import openpyxl
import tkinter as tk
from tkinter import filedialog

# for .exe creation
# import sys
# logfile = open('program_output.txt', 'w')
# sys.stdout = logfile
# sys.stderr = logfile

file_path = ""
writer = None
df = None
totals = []

def create_sheet(sheet_name, start_idx, end_idx):
    df_slice = df.iloc[start_idx:end_idx]
    df_slice.index = range(1, len(df_slice) + 1)
    df_slice.rename_axis('N°', inplace=True)
    df_slice.reset_index(inplace=True)
    df_slice.to_excel(writer, sheet_name=sheet_name, index=False)
    last_row = len(df_slice) + 2  # Dernière ligne + 1
    total = df_slice['MONTANT ECHEANCE'].sum()
    #if sheet_num+1 != df.shape[0] // 971:
    totals.append(total)
    # Écrire la somme dans la cellule
    writer.sheets[sheet_name].cell(row=last_row, column=4, value=total)
    # Formater la cellule comme un nombre avec deux décimales
    cell = writer.sheets[sheet_name].cell(row=last_row, column=4)
    cell.number_format = '0.00'
    pass


eel.init('web')

@eel.expose
def select_and_send_excel_path():
    global file_path 
    root = tk.Tk()
    root.withdraw()  # Hide the main tkinter window

    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])

    if file_path:
        return file_path
    else:
        return ""  
    
@eel.expose
def division_excel(path, page_name = 'BATCH', line_numbers = 978):
    print(path)
    df = pd.read_excel(path, sheet_name='Feuil1')
    with pd.ExcelWriter("outputT4.xlsx") as writer:
        for sheet_num in range(round(df.shape[0] / line_numbers))[:-1]:
            start_idx = sheet_num * line_numbers
            end_idx = (sheet_num + 1) * line_numbers
            sheet_name = page_name+f'{sheet_num + 1}'  
            create_sheet(sheet_name, start_idx, end_idx)
        sheet_name = page_name+f'{sheet_num + 2}'
        start_idx = (sheet_num + 1) * line_numbers
        end_idx = (sheet_num + 2) * line_numbers
        if 997-((df.shape[0] % line_numbers) + line_numbers+1) > 0:
            create_sheet(sheet_name, start_idx, df.shape[0])
        else:
            start_idx = (sheet_num+1) * line_numbers
            create_sheet(sheet_name, start_idx, end_idx)
            sheet_name = page_name+f'{sheet_num + 3}'
            create_sheet(sheet_name, end_idx)
    print("done")
    pass


eel.start("index.html", size=(1366, 743))
