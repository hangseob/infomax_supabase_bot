import xlwings as xw
import pandas as pd
import time
import os

def check_headers():
    app = xw.App(visible=True)
    infomax_xlam_path = r"C:\Infomax\bin\excel\infomaxexcel.xlam"
    if os.path.exists(infomax_xlam_path):
        app.books.open(infomax_xlam_path)
    
    wb = app.books.add()
    sheet = wb.sheets[0]
    
    # Try IR market code
    formula = '=IMDH("IR", "CRST25USDCNH06M", "MID종가", "20260120", "20260124", 1, "Headers=1,Orient=H")'
    sheet.range("A1").formula = formula
    
    print("Waiting for headers...")
    for i in range(10):
        time.sleep(2)
        data = sheet.range("A1:E2").value
        if data and data[0][0] is not None and not (isinstance(data[0][0], str) and "#WAITING" in data[0][0].upper()):
            print("Headers found!")
            print(f"Row 1: {data[0]}")
            print(f"Row 2: {data[1]}")
            break
            
    wb.close()
    app.quit()

if __name__ == "__main__":
    check_headers()
