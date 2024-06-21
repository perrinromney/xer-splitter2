import openpyxl
import subprocess
import tkinter as tk
from tkinter import filedialog
import platform
from openpyxl.worksheet.table import Table, TableStyleInfo


# Create a basic file selector GUI
root = tk.Tk()
root.withdraw()  # Hide the main window

# Ask the user to select the input .xer file
xer_file_path = filedialog.askopenfilename(title="Select .xer file", filetypes=[("XER Files", "*.xer")])

# Read the selected .xer file
with open(xer_file_path, 'r') as xer_file:
    lines = xer_file.readlines()

# Initialize workbook
workbook = openpyxl.Workbook()

# Initialize sheet
sheet = workbook.active
sheet.title = 'Table_1'
firstSheet = True
# Initialize row index
row_index = 1
def move_sheet_by_name(wb, sheet_name, to_loc=None):
    sheets = wb._sheets
    try:
        from_loc = next(i for i, worksheet in enumerate(sheets) if worksheet.title == sheet_name)
    except StopIteration:
        print(f"Sheet '{sheet_name}' not found in the workbook.")
        return

    # If no to_loc given, assume first
    if to_loc is None:
        to_loc = 0
    ws = sheets.pop(from_loc)
    sheets.insert(to_loc, ws)
# Process each line
for line in lines:
    if line.startswith('%T'):
        if not firstSheet:
            sheet.delete_cols(1)
            sheet.auto_filter.ref = sheet.dimensions
        # Start a new table
        firstSheet = False
        table_name = line.strip().lstrip('%T').replace(" ", "_").strip()
        sheet = workbook.create_sheet(title=table_name)
        sheet.freeze_panes = 'A2'
        row_index = 1
    elif line.startswith('%F'):
        # Header row
        header = line.strip().lstrip('%F').split('\t')
        sheet.append(header)
    elif line.startswith('%R'):
        # Normal data row
        data = line.strip().lstrip('%R').split('\t')
        sheet.append(data)
    else:
        data = line.split('\t')
        sheet.append(data)
move_sheet_by_name(workbook,'TASK', 1)
move_sheet_by_name(workbook,'PROJWBS', 1)
# Ask the user to select the output .xlsx file
output_file_path = xer_file_path.replace(".xer",'_toXL.xlsx')#filedialog.asksaveasfilename(title="Save as .xlsx file", defaultextension=".xlsx")

# Save the Excel file
workbook.save(output_file_path)

# Open the Excel file based on the operating system
if platform.system() == 'Windows':
    excel_path = r'C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE'
    subprocess.Popen([excel_path, output_file_path])
elif platform.system() == 'Darwin':  # macOS
    subprocess.Popen(['open', output_file_path])
else:
    print("Unsupported operating system. Please open the generated file manually.")

print("Excel file saved at:", output_file_path)