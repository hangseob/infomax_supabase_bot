import xlwings as xw
import os
import time

def test_three_columns():
    app = xw.apps.active if xw.apps.count > 0 else xw.App(visible=True)
    wb = app.books.add()
    sheet = wb.sheets[0]
    
    target_date = "20260123"
    # Headers=1, Orient=V
    formula = f'=IMDH("IR", "CRST25USDCNH06M", "일자,MID종가", "{target_date}", "{target_date}", 1, "Headers=1,Orient=V")'
    sheet.range("A1").formula = formula
    
    print("Waiting for 3-column result...")
    time.sleep(5)
    
    data = sheet.range("A1:C3").value
    if data:
        for r_idx, r in enumerate(data):
            print(f"Row {r_idx+1}: {r}")

if __name__ == "__main__":
    test_three_columns()
