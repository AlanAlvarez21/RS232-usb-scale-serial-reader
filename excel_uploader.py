
import gspread
import pandas as pd
from oauth2client.service_account import ServiceAccountCredentials

scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']

credentials = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)

gc = gspread.authorize(credentials)



file_id = '1Rkyfzr_gWjIUYaMJNhBM9jW7Q-cxhOp3cLETCH-pUhM'  # Reemplaza con el ID de tu archivo en Google Sheets
file_name = './merged.xlsx'  # Reemplaza con el nombre de tu archivo XLSX

# Abre el archivo existente en Google Sheets
xls_file = gc.open_by_key(file_id)

# Carga el contenido del archivo XLSX utilizando pandas
xls_data = pd.read_excel(file_name, sheet_name=None)

# Actualiza los datos en cada hoja de cálculo
for sheet_name, df in xls_data.items():
    worksheet = xls_file.worksheet(sheet_name)
    worksheet.clear()  # Borra el contenido existente

    # Convierte todos los valores a cadenas de texto
    df = df.astype(str)

    # Actualiza los datos en Google Sheets
    worksheet.update([df.columns.values.tolist()] + df.values.tolist())

print('El archivo XLSX se ha sobrescrito exitosamente en Google Sheets.')