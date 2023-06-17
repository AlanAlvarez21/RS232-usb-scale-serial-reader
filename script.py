from openpyxl import Workbook, load_workbook
import os

# Ruta de la carpeta donde se encuentran los archivos de Excel generados
excel_folder = '/Users/brito/Documents/datostablas/excel/'

# Obtén la lista de archivos de Excel en la carpeta
excel_files = [f for f in os.listdir(excel_folder) if f.endswith('.xlsx')]

# Crea un nuevo archivo de Excel
merged_wb = Workbook()

# Itera sobre los archivos de Excel y agrega las hojas al archivo fusionado
for file in excel_files:
    file_path = os.path.join(excel_folder, file)
    wb = load_workbook(file_path)
    for sheet_name in wb.sheetnames:
        sheet = wb[sheet_name]
        new_sheet = merged_wb.create_sheet(title=f"{file[:-5]} - {sheet_name}")  # Utiliza el nombre del archivo como parte del título de la hoja
        for row in sheet.iter_rows(values_only=True):
            new_sheet.append(row)

# Elimina la hoja de inicio vacía
del merged_wb['Sheet']

# Guarda el archivo fusionado
merged_file_path = '/Users/brito/Documents/datostablas/excel/merged4.xlsx'
merged_wb.save(merged_file_path)