import xlwings as xw
import pandas as pd
import time
from datetime import datetime

def debug_infomax():
    fields_path = r'infomax_functions_templetes/mmkt_infomax_fields.xlsx'
    df = pd.read_excel(fields_path, sheet_name='Sheet2', header=1).head(5)
    
    app = xw.App(visible=True)
    wb = app.books.add()
    sheet = wb.sheets[0]
    
    start_date = "2026-01-01"
    end_date = "2026-01-24"
    
    # Try one formula
    row = df.iloc[0]
    # Use 20260101 format just in case
    s_date = start_date.replace("-", "")
    e_date = end_date.replace("-", "")
    
    formula = f'=IMDH("{row.DATA_TYPE}", "{row.DATA_ID}", "{row.FIELD_ID}", "{s_date}", "{e_date}", 100, "Per=D,Headers=0,Orient=V")'
    print(f"Testing formula: {formula}")
    sheet.range("A1").formula = formula
    
    for i in range(10):
        time.sleep(2)
        val = sheet.range("A1:B10").value
        print(f"Attempt {i+1} values: {val[0] if val else 'None'}")
        if val and val[0][0] is not None and not isinstance(val[0][0], str):
            print("Data received!")
            break
        elif val and isinstance(val[0][0], str) and "#" in val[0][0]:
            print(f"Error detected: {val[0][0]}")
            # If it's #NAME?, Infomax is not loaded
            if "#NAME?" in val[0][0]:
                print("Infomax Add-in not loaded in this Excel instance.")
            break

    wb.close()
    app.quit()

if __name__ == "__main__":
    debug_infomax()
