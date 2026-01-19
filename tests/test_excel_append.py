import pandas as pd
import os

excel_path = "test_append.xlsx"

if os.path.exists(excel_path):
    os.remove(excel_path)

def save_to_excel(name, data):
    if os.path.exists(excel_path):
        writer = pd.ExcelWriter(excel_path, engine='openpyxl', mode='a', if_sheet_exists='replace')
    else:
        writer = pd.ExcelWriter(excel_path, engine='openpyxl')
        pd.DataFrame([["Summary Text"]]).to_excel(writer, sheet_name="Summary", index=False, header=False)
    
    pd.DataFrame(data).to_excel(writer, sheet_name=name, index=False, header=False)
    writer.close()

# 1. Create file with first sheet
save_to_excel("Sheet1", [["Data1"]])
print("Created Sheet1")

# 2. Append second sheet
save_to_excel("Sheet2", [["Data2"]])
print("Appended Sheet2")

# 3. Verify
with pd.ExcelFile(excel_path) as xls:
    sheets = xls.sheet_names
    print(f"Sheets found: {sheets}")
    if "Sheet1" in sheets and "Sheet2" in sheets and "Summary" in sheets:
        print("VERIFICATION SUCCESS: All sheets preserved.")
    else:
        print("VERIFICATION FAILURE: Some sheets are missing.")

if os.path.exists(excel_path):
    os.remove(excel_path)
