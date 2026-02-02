import xlwings as xw
import os
import time

def test_stk_ilja():
    app = xw.apps.active if xw.apps.count > 0 else xw.App(visible=True)
    wb = app.books.add()
    sheet = wb.sheets[0]
    
    formula = '=IMDH("STK", "005930", "일자,현재가", "20260123", "20260123", 1, "Headers=0,Orient=V")'
    sheet.range("A1").formula = formula
    
    print("Waiting for STK result with '일자'...")
    time.sleep(5)
    
    data = sheet.range("A1:C2").value
    if data:
        print(f"STK Result: {data[0]}")

if __name__ == "__main__":
    test_stk_ilja()
