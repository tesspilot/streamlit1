import openpyxl
from openpyxl import Workbook

# Open the source Excel file
source_wb = openpyxl.load_workbook('Integrale kosten Wegen v04.0 LIVE.xlsx', data_only=True)
source_sheet = source_wb['Onderhoud']

# Create a new workbook and select the active sheet
new_wb = Workbook()
new_sheet = new_wb.active
new_sheet.title = 'Copied Data'

# Copy both category names and values
for row in range(4, 29):  # 4 to 28 inclusive
    category = source_sheet[f'D{row}'].value
    value = source_sheet[f'E{row}'].value
    print(f"Reading row {row}: {category} - {value}")  # Debug print
    new_sheet[f'A{row-3}'] = category
    new_sheet[f'B{row-3}'] = value

# Save the new workbook
new_wb.save('copied_values.xlsx')

print("Data has been successfully copied to 'copied_values.xlsx'")
