import xlwings as xw
import pandas as pd
import time
from datetime import datetime
import os

def test_formula_variations():
    print("인포맥스 수식 변형 테스트를 시작합니다...")
    
    fields_path = r'infomax_functions_templetes/mmkt_infomax_fields.xlsx'
    df_fields = pd.read_excel(fields_path, sheet_name='Sheet2', header=1)
    df_fields = df_fields.dropna(subset=['RATE_ID', 'DATA_TYPE', 'DATA_ID', 'FIELD_ID'])
    row = df_fields.iloc[0]
    
    market = row.DATA_TYPE
    code = row.DATA_ID
    field = row.FIELD_ID
    
    app = xw.App(visible=True)
    infomax_xlam_path = r"C:\Infomax\bin\excel\infomaxexcel.xlam"
    if os.path.exists(infomax_xlam_path):
        app.books.open(infomax_xlam_path)
    
    wb = app.books.add()
    sheet = wb.sheets[0]
    
    start_date = "20260101"
    end_date = datetime.now().strftime("%Y%m%d")
    
    variations = [
        f'=IMDH("{market}", "{code}", "날짜,{field}", "{start_date}", "{end_date}", 5, "Headers=1,Orient=V,Per=D")',
        f'=IMDH("{market}", "{code}", "{field}", "{start_date}", "{end_date}", 5, "Headers=1,Orient=V,Per=D")',
        f'=IMDH("{market}", "{code}", "날짜,{field}", "{start_date}", "{end_date}", 5, "Headers=0,Orient=V,Per=D")',
    ]
    
    for idx, formula in enumerate(variations):
        col = (idx * 3) + 1
        print(f"\n변형 {idx+1} 테스트 중: {formula}")
        cell = sheet.range(1, col)
        cell.formula = formula
        
        # 각 변형당 20초 대기
        for wait in range(10):
            time.sleep(2)
            data = sheet.range((1, col), (10, col + 1)).value
            if data and data[0][0] is not None and not (isinstance(data[0][0], str) and "#WAITING" in data[0][0].upper()):
                print(f"변형 {idx+1} 결과 수신!")
                for r_idx, r in enumerate(data[:6]):
                    print(f"  Row {r_idx+1}: {r}")
                break
            print(".", end="", flush=True)
            
    print("\n테스트 종료. 엑셀을 확인하세요.")

if __name__ == "__main__":
    test_formula_variations()
