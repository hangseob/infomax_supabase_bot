import xlwings as xw
import pandas as pd
import time
from datetime import datetime
import os

def test_single_code():
    print("인포맥스 단일 종목 테스트를 시작합니다...")
    
    # 1. 테스트할 종목 선정 (첫 번째 행 사용)
    fields_path = r'infomax_functions_templetes/mmkt_infomax_fields.xlsx'
    try:
        df_fields = pd.read_excel(fields_path, sheet_name='Sheet2', header=1)
        df_fields = df_fields.dropna(subset=['RATE_ID', 'DATA_TYPE', 'DATA_ID', 'FIELD_ID'])
        row = df_fields.iloc[0]
        rate_id = row.RATE_ID
        market = row.DATA_TYPE
        code = row.DATA_ID
        field = row.FIELD_ID
        print(f"테스트 대상: RATE_ID={rate_id}, Market={market}, Code={code}, Field={field}")
    except Exception as e:
        print(f"필드 정보 읽기 실패: {e}")
        return

    # 2. 엑셀 실행
    try:
        app = xw.App(visible=True)
        # 인포맥스 애드인 로드
        infomax_xlam_path = r"C:\Infomax\bin\excel\infomaxexcel.xlam"
        if os.path.exists(infomax_xlam_path):
            app.books.open(infomax_xlam_path)
            
        wb = app.books.add()
        sheet = wb.sheets[0]
    except Exception as e:
        print(f"엑셀 실행 실패: {e}")
        return

    # 3. IMDH 수식 입력 (날짜와 필드 함께 요청)
    start_date = "20260101"
    end_date = datetime.now().strftime("%Y%m%d")
    
    # 필드 인자에 "날짜,필드명" 형식으로 전달
    formula = f'=IMDH("{market}", "{code}", "날짜,{field}", "{start_date}", "{end_date}", 10, "Headers=1,Orient=V,Per=D")'
    print(f"입력 수식: {formula}")
    
    sheet.range("A1").formula = formula
    
    # 4. 결과 확인
    print("데이터 수신 대기 중 (30초)...")
    for i in range(15):
        time.sleep(2)
        data = sheet.range("A1:B11").value
        if data and data[0][0] is not None:
            first_row = data[0]
            if isinstance(first_row[0], str) and "#WAITING" in first_row[0].upper():
                print(f"[{i+1}/15] 대기 중...", end='\r')
                continue
            
            print("\n데이터 수신 성공!")
            print("상위 5개 결과:")
            for r in data[:5]:
                print(f"  {r}")
            break
        print(f"[{i+1}/15] 데이터 없음...", end='\r')

    print("\n테스트 종료. 엑셀을 확인해 보세요.")

if __name__ == "__main__":
    test_single_code()
