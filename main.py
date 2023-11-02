import eel
import pandas as pd
import base64
# import sys
# logfile = open('program_output.txt', 'w')
# sys.stdout = logfile
# sys.stderr = logfile


eel.init('web')


@eel.expose
def division_excel(excel_data):
    print('hello')
    print(excel_data)
    # try:
    #     # excel_file = formData.get("excelFile")
    #     # df = pd.read_excel(excel_file)
    #     decoded_data = base64.b64decode(data)
    #     print("bien reçu", decoded_data, name)
    #     # Vous pouvez maintenant utiliser df (un DataFrame) pour effectuer des opérations de division, par exemple.
    #     # division_excel(df)
    # except Exception as e:
    #     # Gérer les erreurs éventuelles
    #     print(f"Erreur : {e}")
    #     # print(excel_file)


eel.start("index.html", size=(1366, 743))
