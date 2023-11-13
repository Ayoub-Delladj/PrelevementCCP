import sys
if hasattr(sys, "_MEIPASS"): # if the script is started from an executable file
    with open("logs.txt", "w") as f_logs:
        sys.stdout = f_logs
        sys.stderr = f_logs
import eel
import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import filedialog
import math
import traceback    
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import re
import os
import shutil


# import sys
# logfile = open('program_output.txt', 'w')
# sys.stdout = logfile
# sys.stderr = logfile

eel.init('web')

@eel.expose
def select_and_send_pdf_path():
    global pdf_path 
    root = tk.Tk()
    root.withdraw()  # Hide the main tkinter window
    root.attributes('-topmost', 1)
    pdf_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])

    if pdf_path:
        return pdf_path
    else:
        return "" 

 
@eel.expose
def select_and_send_excel_path():
    global file_path 
    root = tk.Tk()
    root.withdraw()  # Hide the main tkinter window
    root.attributes('-topmost', 1)
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])

    if file_path:
        return file_path
    else:
        return ""  

@eel.expose
def select_and_send_folder_path():
    global folder_path 
    root = tk.Tk()
    root.withdraw()  # Hide the main tkinter window
    root.attributes('-topmost', 1)
    folder_path = filedialog.askdirectory()

    if folder_path:
        return folder_path
    else:
        return ""  
    
@eel.expose
def telecharger_fichier(fichier, download_path, folder_path):
    if (folder_path == ''):
        if os.path.exists(fichier):
            if os.path.exists(download_path+'/'+fichier):
                os.remove(download_path+'/'+fichier)
                shutil.copy(fichier, download_path)
            else :
                shutil.copy(fichier, download_path)
            return (download_path+'/'+fichier)
    else :
        if os.path.exists(fichier):
            if os.path.exists(folder_path+'/'+fichier):
                os.remove(folder_path+'/'+fichier)
                shutil.copy(fichier, folder_path)
            else :
                shutil.copy(fichier, folder_path)
            return(folder_path+'/'+fichier)

@eel.expose
def verification_CCP (excl_path, pdf_path):
    try :
        data = {
            'Name': ['John', 'Jane', 'Bob','John', 'Jane', 'Bob'],
            'Age': [25, 30, 22,'John', 'Jane', 'Bob'],
            'City': ['New York', 'San Francisco', 'Los Angeles','John', 'Jane', 'Bob'],
            'Pro': ['John', 'Jane', 'Bob','John', 'Jane', 'Bob'],
            'Agge': [25, 30, 22,'John', 'Jane', 'Bob'],
            'Cittttty': ['New York', 'San Francisco', 'Los Angeles','John', 'Jane', 'Bob']
        }
        df = pd.DataFrame(data)
        df = df.to_html(index=False)
        df = df.replace('text-align: right', 'text-align: center')
        eel.close_loading_popup(df)
    except :
        eel.gestion_exception()

@eel.expose
def division_excel (path, page_name = 'BATCH', line_numbers = 971,sheet_name='Feuil1'):
    try :
        line_numbers = int(line_numbers)
        page_name = str(page_name)
        df = pd.read_excel(path, sheet_name)
        column_compte_ccp = ''
        for column in df.columns:
            if column.lower().strip() == 'ccp client':
                column_compte_ccp = column
                print(column)
                break  
        df[column_compte_ccp] = df[column_compte_ccp].astype('str')
        df[column_compte_ccp] = df[column_compte_ccp].str.pad(width=10, fillchar='0')
        totals = []
        file_path = '/'.join(path.split('/')[:-1])
        file_name = path.split('/')[-1].split('.')[0]
        file_extention = path.split('/')[-1].split('.')[-1]
        with pd.ExcelWriter(file_name+' divisé.'+file_extention) as writer:
            df_vide = pd.DataFrame()
            df_vide.to_excel(writer, sheet_name='Total', index=False)
            for sheet_num in range(math.floor(df.shape[0] / line_numbers))[:-1]:
                start_idx = sheet_num * line_numbers
                end_idx = (sheet_num + 1) * line_numbers
                sheet_name = page_name+f'{sheet_num + 1}'
    #           create_sheet(writer, df, sheet_name, start_idx, end_idx)
                df_slice = df.iloc[start_idx:end_idx]
                df_slice.index = range(1, len(df_slice) + 1)
                df_slice.rename_axis('N°', inplace=True)
                df_slice.reset_index(inplace=True)
                total = df_slice['MONTANT ECHEANCE'].sum()
                new_row = pd.DataFrame({'MONTANT ECHEANCE': [total]})
                df_slice = pd.concat([df_slice, new_row], ignore_index=True)
                df_slice.to_excel(writer, sheet_name=sheet_name, index=False)
                #if sheet_num+1 != df.shape[0] // 971:
                totals.append([total, df_slice.shape[0]])
                # Écrire la somme dans la cellule
            sheet_name = page_name+f'{sheet_num + 2}'
            start_idx = (sheet_num + 1) * line_numbers
            end_idx = (sheet_num + 2) * line_numbers
            if 997-((df.shape[0] % line_numbers) + line_numbers+1) > 0:
    #           create_sheet(writer, df, sheet_name, start_idx, end_idx:df.shape[0])
                df_slice = df.iloc[start_idx:df.shape[0]]
                df_slice.index = range(1, len(df_slice) + 1)
                df_slice.rename_axis('N°', inplace=True)
                df_slice.reset_index(inplace=True)
                total = df_slice['MONTANT ECHEANCE'].sum()
                new_row = pd.DataFrame({'MONTANT ECHEANCE': [total]})
                df_slice = pd.concat([df_slice, new_row], ignore_index=True)
                df_slice.to_excel(writer, sheet_name=sheet_name, index=False)
                #if sheet_num+1 != df.shape[0] // 971:
                totals.append([total, df_slice.shape[0]])
                # Écrire la somme dans la cellule
            else:
                start_idx = (sheet_num+1) * line_numbers
                # print(start_idx, end_idx, df.shape[0])
    #             create_sheet(writer, df, sheet_name, start_idx, end_idx)
                df_slice = df.iloc[start_idx:end_idx]
                df_slice.index = range(1, len(df_slice) + 1)
                df_slice.rename_axis('N°', inplace=True)
                df_slice.reset_index(inplace=True)
                total = df_slice['MONTANT ECHEANCE'].sum()
                new_row = pd.DataFrame({'MONTANT ECHEANCE': [total]})
                df_slice = pd.concat([df_slice, new_row], ignore_index=True)
                df_slice.to_excel(writer, sheet_name=sheet_name, index=False)
                #if sheet_num+1 != df.shape[0] // 971:
                totals.append([total, df_slice.shape[0]])
                # Écrire la somme dans la cellule
                sheet_name = page_name+f'{sheet_num + 3}'
    #     create_sheet(writer, df, sheet_name, start_idx:end_idx, end_idx:df.shape[0])
                df_slice = df.iloc[end_idx:df.shape[0]]
                df_slice.index = range(1, len(df_slice) + 1)
                df_slice.rename_axis('N°', inplace=True)
                df_slice.reset_index(inplace=True)
                total = df_slice['MONTANT ECHEANCE'].sum()
                new_row = pd.DataFrame({'MONTANT ECHEANCE': [total]})
                df_slice = pd.concat([df_slice, new_row], ignore_index=True)
                df_slice.to_excel(writer, sheet_name=sheet_name, index=False)
                #if sheet_num+1 != df.shape[0] // 971:
                totals.append([total, df_slice.shape[0]])
                # Écrire la somme dans la cellule
            # create a sheet that have the total
            batch_name = []
            totals = np.array(totals)
            totals = np.append(totals, [[sum(totals[:,0]),sum(totals[:,1])]], axis =0)
            for i in range(1,len(totals)):
                batch_name.append(page_name+f'{i}')
            batch_name.append('Total')
            df_total = pd.DataFrame(list(zip(batch_name, totals[:,0], totals[:,1])), columns =['Batch','Montant echeance total','Nombre de comptes'])
            df_total.index = range(1, df_total.shape[0]+1)
            df_total['Nombre de comptes'] = df_total['Nombre de comptes'].astype('int32')
            # Write the DataFrame to the new sheet
            df_total.to_excel(writer, sheet_name='Total', index=False)
            # sheet = writer.book['Total']
            # # Remove the sheet from its original position
            # writer.book.remove(sheet)
            # # Insert the sheet at the beginning of the sheets
            # writer.book._sheets.insert(0, sheet)

        eel.close_loading_popup(file_path, file_name+' divisé.'+file_extention) #chemin du fichier de retour
    except :
        eel.gestion_exception()

def color_rows(ws, nb_columns, indexes, color):
    alphabet = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z']
    for i in indexes:
        for col in alphabet[:nb_columns]:
            ws[col+str(i+2)].fill = PatternFill(patternType='solid',fgColor=color)


@eel.expose
def verification_compte(excel_path, nom_feuille1, colonne1, situation_compte_path, nom_feuille2, colonne2):
    try:
        output_path = '/'.join(excel_path.split('/')[:-1])
        file_name = excel_path.split('/')[-1].split('.')[0]
        file_extention = excel_path.split('/')[-1].split('.')[-1]

        df_croise = pd.read_excel(excel_path, nom_feuille1)
        df_situation = pd.read_excel(situation_compte_path, nom_feuille2)

        column_compte = ''
        for column in df_croise.columns:
            if column.lower().strip() == colonne1:
                column_compte = column
                break
        column_nb_compte = ''
        for column in df_situation.columns:
            if column.lower().strip() == colonne2:
                column_nb_compte = column
                break  
        
        df_situation = df_situation[[column_nb_compte]]
        index_rib_pb = []
        for value in zip(df_croise[column_compte].values, df_croise[column_compte].index):
            if value[0] is np.nan or len(re.findall('[^0-9]',str(value[0]).strip())) > 0:
                index_rib_pb.append(value[1])
        index_rib_pb = pd.Index(index_rib_pb)
        df_rib_pb = df_croise.iloc[index_rib_pb]
        df_rib_pb.loc[:,['Etat']] = 'PB RIB'
        df_croise.drop(index=index_rib_pb, inplace=True)
        df_croise = df_croise.astype({column_compte: 'int64'})
        df_situation = df_situation.astype({column_nb_compte: 'int64'})
        df_croisement = pd.merge(df_croise,df_situation, left_on=column_compte, right_on=column_nb_compte, how='left')
        df_croisement['Etat'] = np.where((df_croisement[column_nb_compte].isna()),'NON ACTIF','ACTIF')
        df_croisement.drop(columns = [column_nb_compte], inplace = True)
        df_croisement.reset_index(drop=True,inplace=True)
        df_croisement = pd.concat([df_croisement, df_rib_pb], axis=0)
        #df_croisement.sort_index(inplace=True)
    
        df_croisement = df_croisement.astype({column_compte: 'str'})
        df_croisement[column_compte] = df_croisement[column_compte].str.pad(width=10, fillchar='0')
        #df_croisement.sort_index(inplace=True)
        df_croisement.to_excel(file_name+' vérifié.'+file_extention, index=False)
        wb = load_workbook(file_name+' vérifié.'+file_extention)
        ws = wb['Sheet1']
        indexes = df_croisement[df_croisement['Etat'] == 'NON ACTIF'].index
        color_rows(ws, len(df_croisement.columns), indexes, 'ff3333')
    
        indexes = range(df_croise.shape[0], df_croise.shape[0]+len(df_croisement[df_croisement['Etat'] == 'PB RIB']))
        color_rows(ws, len(df_croisement.columns), indexes , 'f5933d')
        wb.save(file_name+' vérifié.'+file_extention)

        eel.close_loading_popup(output_path, file_name+' vérifié.'+file_extention) #chemin du fichier de retour
    except Exception as e:
        eel.gestion_exception()

@eel.expose
def close_python(*args):
    # Lorsque vous avez fini, libérez le verrou en fermant le fichier
    # Récupérer le chemin du répertoire actuel
    repertoire_actuel = os.getcwd()

    # Liste tous les fichiers dans le répertoire actuel
    fichiers_dans_repertoire = os.listdir(repertoire_actuel)

    # Parcourir la liste des fichiers et supprimer les fichiers Excel
    for fichier in fichiers_dans_repertoire:
        if fichier.endswith(".xlsx"):
            chemin_fichier = os.path.join(repertoire_actuel, fichier)
            os.remove(chemin_fichier)
    os.remove(lock_file)
    sys.exit()



def log_exception(exc_type, exc_value, exc_traceback):
    with open('error.log', 'a') as f:
        f.write(f"Exception Type: {exc_type}\n")
        f.write(f"Exception Value: {exc_value}\n")
        f.write(f"Traceback:\n")
        traceback.print_tb(exc_traceback, file=f)

sys.excepthook = log_exception

if __name__ == '__main__':
    
    lock_file = 'my_program.lock'

    if os.path.isfile(lock_file):
        log_exception('exc_type', 'exc_value', 'exc_traceback')
        sys.exit(0)

    try:
        with open(lock_file, 'w') as f:
            pass
    except IOError:
        log_exception('exc_type', 'exc_value', 'exc_traceback')
        sys.exit(1)
    eel.start("index.html", size=(1366, 743), port=8001)