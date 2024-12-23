import os
import clr
clr.AddReference("Microsoft.Office.Interop.Excel")
from Microsoft.Office.Interop import Excel
from pyrevit import forms

# Allow user to select the Excel file
excel_file = forms.pick_file(file_ext="xlsx", title="Select Excel File")
if not excel_file:
    forms.alert("No Excel file selected. Exiting...", exitscript=True)

# Allow user to select the folder with TXT files
txt_folder = forms.pick_folder(title="Select Folder Containing TXT Files")
if not txt_folder:
    forms.alert("No folder selected. Exiting...", exitscript=True)

worksheet_name = "STATUS"

# Color codes for formatting
GREEN = 5296274   # RGB(85, 255, 85)
RED = 255         # RGB(255, 0, 0)
GRAY = 12632256   # RGB(192, 192, 192)

# Read plugin statuses from TXT files
plugin_status = {}
for file in os.listdir(txt_folder):
    if file.endswith("Revit Plugins 2024.txt"):
        pc_name = file.split(" - ")[0]  # Extract PC name
        plugin_status[pc_name] = {}
        with open(os.path.join(txt_folder, file), "r") as f:
            lines = f.readlines()
        for line in lines[1:]:
            parts = line.strip().split(",")
            if len(parts) == 3:
                plugin_name = parts[1]
                status = parts[2]
                plugin_status[pc_name][plugin_name] = status

print("Plugin statuses from TXT files: {}".format(plugin_status))  # Debugging

# Open Excel and update
try:
    excel_app = Excel.ApplicationClass()
    excel_app.Visible = False
    workbook = excel_app.Workbooks.Open(excel_file, ReadOnly=False)
    sheet = workbook.Worksheets(worksheet_name)

    # Get plugin names from row 11
    plugin_columns = {}
    for col in range(7, 14):  # Columns G to M
        plugin_name = sheet.Cells(11, col).Value2
        print("Plugin in column {}: {}".format(col, plugin_name))  # Debugging
        if plugin_name:
            plugin_columns[plugin_name] = col

    print("Plugin columns mapping: {}".format(plugin_columns))  # Debugging

    # Update statuses for each PC in column F
    row = 12  # Starting row for PC names
    empty_count = 0  # Counter for consecutive empty rows

    while True:
        pc_name = sheet.Cells(row, 6).Value2  # Column F (PC names)
        if pc_name:  # Process rows with PC names
            pc_name = pc_name.strip()
            empty_count = 0  # Reset empty counter
            print("Processing PC: {}".format(pc_name))  # Debugging
            if pc_name in plugin_status:
                # Update plugin statuses for this PC
                for plugin_name, col in plugin_columns.items():
                    status = plugin_status[pc_name].get(plugin_name, "NOT INSTALLED")
                    cell = sheet.Cells(row, col)
                    cell.Value2 = status
                    if status == "INSTALLED":
                        cell.Interior.Color = GREEN
                    elif status == "NOT INSTALLED":
                        cell.Interior.Color = RED
                    else:
                        cell.Interior.Color = GRAY
                    print("Writing '{}' for plugin '{}' at row {}, column {}".format(
                        status, plugin_name, row, col))  # Debugging
            else:
                # Write 'NO DATA' if no TXT file exists for this PC
                for plugin_name, col in plugin_columns.items():
                    cell = sheet.Cells(row, col)
                    cell.Value2 = "NO DATA"
                    cell.Interior.Color = GRAY
                    print("Writing 'NO DATA' for plugin '{}' at row {}, column {}".format(
                        plugin_name, row, col))  # Debugging
        else:
            empty_count += 1
            print("Skipping empty row at {}, consecutive empty rows: {}".format(row, empty_count))  # Debugging
            if empty_count > 5:
                print("Encountered more than 5 consecutive empty rows. Stopping.")
                break

        row += 1  # Move to the next row

    # Save the workbook
    workbook.SaveAs(excel_file)
    workbook.Close(SaveChanges=True)
    excel_app.Quit()

    forms.alert("Write operation completed with formatting! Check the 'STATUS' sheet.")
except Exception as e:
    forms.alert("Error while updating Excel: {}".format(e))
