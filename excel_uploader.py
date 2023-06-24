import gspread 
from oauth2client.service_account import ServiceAccountCredentials
from pprint import pprint

scope = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/drive.file",
    "https://www.googleapis.com/auth/drive.readonly",
    "https://www.googleapis.com/auth/drive.metadata.readonly",
    "https://www.googleapis.com/auth/spreadsheets"
]

credentials = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
client = gspread.authorize(credentials)

sheet = client.open("tutorial").sheet1

data = sheet.get_all_records()

# row = sheet.row_values(1)
# col = sheet.column_values(1)
cell = sheet.cell(1,1).value
sheet.update_cell(1,1, ')*(')

