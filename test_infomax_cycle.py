import xlwings as xw
import os
import time
import pandas as pd
from datetime import datetime

def test_infomax_full_cycle():
    test_file = os.path.abspath("test_infomax_cycle.xlsx")
    fields_path = r'infomax_functions_templetes/mmkt_infomax_fields.xlsx'
    infomax_xlam = r"C:\Infomax\bin\excel\infomaxexcel.xlam"
    
    if os.path.exists(test_file):
        try: os.remove(test_file)
        except: pass

    # 1. 대상 종목 하나 선정
    df_fields = pd.read_excel(fields_path, sheet_name='Sheet2', header=1)
    row = df_fields.iloc[0]
    rate_id = row['RATE_ID']
    
    print(f"--- [1단계] 엑셀 및 인포맥스 로드 ---")
    app = xw.App(visible=True, add_book=False)
    try:
        if os.path.exists(infomax_xlam):
            print(f"애드인 로드: {infomax_xlam}")
            app.books.open(infomax_xlam)
            time.sleep(5) # 애드인 초기화 대기
        
        wb = app.books.add()
        scratch = wb.sheets[0]
        scratch.name = "Scratch"
        
        print(f"--- [2단계] 데이터 요청 (IMDH) ---")
        formula = f'=IMDH("{row["DATA_TYPE"]}", "{row["DATA_ID"]}", "일자,{row["FIELD_ID"]}", "20260120", "20260124", 5, "Headers=0,Orient=V,Per=D")'
        scratch.range("A1").formula = formula
        
        fetched_data = None
        for i in range(15):
            time.sleep(2)
            val = scratch.range("A1:C5").value
            if val and (val[0][0] or val[0][1]) and "#WAITING" not in str(val[0][0] or "").upper():
                print(f"데이터 수신 성공! (시도 {i+1})")
                fetched_data = val
                break
            print(".", end="", flush=True)
            
        if not fetched_data:
            print("\n데이터 수신 실패. 종료합니다.")
            return

        print(f"\n--- [3단계] 결과 시트에 쓰기 ---")
        final = wb.sheets.add("FinalTable")
        final.range("A1").value = ["날짜", "코드", "값"]
        
        # 첫 번째 수신 행만 써보기
        r = fetched_data[0]
        final.range("A2").value = [r[1], rate_id, r[2]]
        
        print(f"저장 및 닫기: {test_file}")
        wb.save(test_file)
        wb.close()
        app.quit()
        
        print("\n--- [4단계] 다시 열어서 읽기 확인 ---")
        time.sleep(2)
        app2 = xw.App(visible=True, add_book=False)
        wb2 = app2.books.open(test_file)
        final2 = wb2.sheets["FinalTable"]
        
        result = final2.range("A1:C2").value
        print(f"최종 읽기 결과: {result}")
        
        wb2.close()
        app2.quit()
        
        print("\n[성공] 인포맥스 함수 호출부터 파일 저장/읽기까지 전 과정이 정상입니다!")
        
    except Exception as e:
        print(f"\n[에러 발생] {e}")
        try: app.quit()
        except: pass

if __name__ == "__main__":
    test_infomax_full_cycle()
