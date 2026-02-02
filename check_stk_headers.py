import xlwings as xw
import os
import time

def check_stk_headers():
    app = xw.apps.active if xw.apps.count > 0 else xw.App(visible=True)
    infomax_xlam_path = r"C:\Infomax\bin\excel\infomaxexcel.xlam"
    if os.path.exists(infomax_xlam_path):
        app.books.open(infomax_xlam_path)
    
    wb = app.books.add()
    sheet = wb.sheets[0]
    
    # Try Samsung with Headers=1, Orient=H to see field names
    formula = '=IMDH("STK", "005930", "날짜,현재가", "20260120", "20260124", 1, "Headers=1,Orient=H")'
    sheet.range("A1").formula = formula
    
    print("Waiting for STK headers...")
    for i in range(10):
        time.sleep(2)
        data = sheet.range("A1:C2").value
        if data and data[0][0] is not None and not (isinstance(data[0][0], str) and "#WAITING" in data[0][0].upper()):
            print("Headers found!")
            print(f"Row 1: {data[0]}")
            print(f"Row 2: {data[1]}")
            break
        print(f"Status: {data[0][0] if data else 'None'}")

if __name__ == "__main__":
    check_stk_headers()
