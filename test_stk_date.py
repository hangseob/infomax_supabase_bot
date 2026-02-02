import xlwings as xw
import pandas as pd
import time
from datetime import datetime
import os

def test_stk_date():
    print("STK 시장 날짜 테스트를 시작합니다...")
    
    app = xw.App(visible=True)
    infomax_xlam_path = r"C:\Infomax\bin\excel\infomaxexcel.xlam"
    if os.path.exists(infomax_xlam_path):
        app.books.open(infomax_xlam_path)
    
    wb = app.books.add()
    sheet = wb.sheets[0]
    
    start_date = "20260101"
    end_date = datetime.now().strftime("%Y%m%d")
    
    # 삼성전자 (005930)
    formula = f'=IMDH("STK", "005930", "날짜,현재가", "{start_date}", "{end_date}", 5, "Headers=1,Orient=V,Per=D")'
    print(f"입력 수식: {formula}")
    sheet.range("A1").formula = formula
    
    for i in range(10):
        time.sleep(2)
        data = sheet.range("A1:B6").value
        if data and data[0][0] is not None and not (isinstance(data[0][0], str) and "#WAITING" in data[0][0].upper()):
            print("결과 수신!")
            for r_idx, r in enumerate(data):
                print(f"  Row {r_idx+1}: {r}")
            break
        print(".", end="", flush=True)
    
    wb.close()
    app.quit()

if __name__ == "__main__":
    test_stk_date()
