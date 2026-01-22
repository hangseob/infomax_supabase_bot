import xlwings as xw
import os
import time

def test_market_cap():
    infomax_xlam_path = r"C:\Infomax\bin\excel\infomaxexcel.xlam"
    
    try:
        app = xw.App(visible=True, add_book=False)
        if os.path.exists(infomax_xlam_path):
            app.books.open(infomax_xlam_path)
            time.sleep(2)
            
        wb = app.books.add()
        sheet = wb.sheets[0]
        
        # Test Samsung Electronics (005930) market cap
        # Trying different field names for Market Cap
        test_fields = ["시가총액", "시가총액(억원)", "상장시가총액"]
        stock_code = "005930"
        
        sheet.range("A1").value = "필드명"
        sheet.range("B1").value = "결과"
        
        for i, field in enumerate(test_fields, 2):
            formula = f'=IMDP("STK", "{stock_code}", "{field}")'
            sheet.range(f"A{i}").value = field
            sheet.range(f"B{i}").formula = formula
            
        print("데이터 로딩 대기 (10초)...")
        time.sleep(10)
        
        results = sheet.range("A1:B4").value
        for row in results:
            print(f"{row[0]}: {row[1]}")
            
    except Exception as e:
        print(f"오류: {e}")

if __name__ == "__main__":
    test_market_cap()
