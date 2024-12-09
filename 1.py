import pandas as pd

file_path = r"C:\Users\annu.sharma\Desktop\playwright\timesheet_data.xlsx"

# Load all sheets and print their names
#sheets = pd.read_excel(file_path, sheet_name=None)
#print("Available sheets:", sheets.keys())
sheet_data = pd.read_excel(file_path, sheet_name="akshay_Timesheet")
print(sheet_data)

# Load all sheets
#all_sheets = pd.read_excel(file_path, sheet_name=None)

# Process each sheet
#for sheet_name, data in all_sheets.items():
    #print(f"Processing sheet: {sheet_name}")
    # Example: Print the first row
    #print(data.head(1))



