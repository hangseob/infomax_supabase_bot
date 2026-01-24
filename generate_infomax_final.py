import xlwings as xw
import pandas as pd
import time
from datetime import datetime
import os

def create_infomax_excel():
    print("인포맥스 데이터 추출 및 최종 테이블 생성을 시작합니다...")
    
    # 1. 필드 정보 읽기
    fields_path = r'infomax_functions_templetes/mmkt_infomax_fields.xlsx'
    try:
        df_fields = pd.read_excel(fields_path, sheet_name='Sheet2', header=1)
        df_fields = df_fields.dropna(subset=['RATE_ID', 'DATA_TYPE', 'DATA_ID', 'FIELD_ID'])
        print(f"총 {len(df_fields)}개의 RATE_ID를 처리합니다.")
    except Exception as e:
        print(f"필드 정보를 읽는 중 오류 발생: {e}")
        return

    # 2. 엑셀 실행 및 워크북 생성
    try:
        if xw.apps.count > 0:
            app = xw.apps.active
            print("기존 엑셀 인스턴스에 연결했습니다.")
        else:
            app = xw.App(visible=True)
            print("새로운 엑셀 인스턴스를 실행했습니다.")
        
        infomax_xlam_path = r"C:\Infomax\bin\excel\infomaxexcel.xlam"
        if os.path.exists(infomax_xlam_path):
            print(f"인포맥스 애드인을 로드합니다: {infomax_xlam_path}")
            try:
                app.books.open(infomax_xlam_path)
            except Exception as e:
                print(f"애드인 로드 중 오류: {e}")
        
        wb = app.books.add()
        raw_sheet = wb.sheets[0]
        raw_sheet.name = "RawData"
    except Exception as e:
        print(f"엑셀 초기화 중 오류: {e}")
        return

    # 3. 데이터 수집 (Wide Format)
    start_date = "20260101"
    end_date = datetime.now().strftime("%Y%m%d")
    print(f"조회 기간: {start_date} ~ {end_date}")
    
    # 날짜 컬럼 (A열)
    first_row = df_fields.iloc[0]
    date_formula = f'=IMDH("{first_row.DATA_TYPE}", "{first_row.DATA_ID}", "날짜", "{start_date}", "{end_date}", 100, "Per=D,Headers=0,Orient=V")'
    raw_sheet.range("A1").value = "날짜"
    raw_sheet.range("A2").formula = date_formula
    
    # 각 종목별 데이터 컬럼 (B열부터)
    rate_ids = df_fields['RATE_ID'].tolist()
    formulas = []
    for _, row in df_fields.iterrows():
        scale = row.SCALE_FACTOR if not pd.isna(row.SCALE_FACTOR) else 1.0
        f = f'=IMDH("{row.DATA_TYPE}", "{row.DATA_ID}", "{row.FIELD_ID}", "{start_date}", "{end_date}", 100, "Per=D,Headers=0,Orient=V") * {scale}'
        formulas.append(f)
    
    print("수식 입력 중...")
    raw_sheet.range("B1").value = rate_ids
    raw_sheet.range("B2").formula = formulas
    
    print("데이터 수신 대기 중 (최대 5분)...")
    
    # 로드 완료 확인
    max_retries = 30
    loaded = False
    for i in range(max_retries):
        # 엑셀 계산 트리거
        try:
            app.api.Calculate()
        except:
            pass
            
        time.sleep(10)
        
        # A2(첫 날짜)와 B2(첫 데이터) 확인
        check_vals = raw_sheet.range("A2:B2").value
        
        if check_vals and check_vals[0] and check_vals[1]:
            val_a2 = check_vals[0]
            val_b2 = check_vals[1]
            
            # 문자열인 경우 #WAITING 체크
            is_waiting = False
            for v in [val_a2, val_b2]:
                if isinstance(v, str):
                    if "#WAITING" in v.upper():
                        is_waiting = True
                    elif "#NAME?" in v.upper():
                        print("\n[!] #NAME? 에러 발생. 애드인 로드 상태를 확인하세요.")
                        break
            
            if not is_waiting and not (isinstance(val_a2, str) and "#" in val_a2):
                print(f"\n데이터 로드 확인됨. (날짜: {val_a2}, 첫번째 값: {val_b2})")
                loaded = True
                break
        
        print(f"[{i+1}/{max_retries}] 데이터 대기 중... (A2: {check_vals[0] if check_vals else 'None'})", end='\r')
    
    if not loaded:
        print("\n경고: 일부 데이터가 로드되지 않았을 수 있습니다. 현재 상태로 진행합니다.")

    # 4. 데이터 변환 (Wide -> Long)
    print("데이터를 플랫 테이블 형식으로 변환 중...")
    
    # 날짜가 있는 마지막 행 찾기
    if raw_sheet.range("A3").value is None:
        if raw_sheet.range("A2").value is not None:
            last_row = 2
        else:
            print("에러: 로드된 데이터가 없습니다.")
            return
    else:
        last_row = raw_sheet.range("A2").end('down').row
        if last_row > 100000: # 엑셀 끝까지 간 경우 (데이터가 A2 하나뿐일 때 등)
            last_row = 2
    
    print(f"로드된 행 수: {last_row - 1}")
    # 전체 데이터 영역 읽기
    total_cols = len(rate_ids) + 1
    raw_data = raw_sheet.range((1, 1), (last_row, total_cols)).value
    
    if not raw_data or len(raw_data) < 2:
        print("에러: 읽어온 데이터가 없습니다.")
        return

    # Pandas를 이용한 변환
    df_raw = pd.DataFrame(raw_data[1:], columns=raw_data[0])
    # 날짜가 None인 행 제거
    df_raw = df_raw.dropna(subset=['날짜'])
    
    # Unpivot (Melt)
    df_long = df_raw.melt(id_vars=['날짜'], var_name='코드', value_name='값')
    
    # 5. 결과 시트 작성
    try:
        result_sheet = wb.sheets.add("FinalTable")
    except:
        result_sheet = wb.sheets.add()
        result_sheet.name = "FinalTable_" + datetime.now().strftime("%H%M%S")
        
    print(f"최종 테이블 작성 중... (총 {len(df_long)}행)")
    
    # 헤더
    result_sheet.range("A1").value = ["날짜", "코드", "값"]
    # 데이터 (성능을 위해 리스트로 변환 후 한꺼번에 입력)
    result_sheet.range("A2").value = df_long.values.tolist()
    
    # 서식 정리
    result_sheet.range("A1:C1").color = (200, 200, 200) # 회색 헤더
    result_sheet.autofit()
    
    # 저장
    output_filename = f"infomax_final_table_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    wb.save(output_filename)
    print(f"\n모든 작업이 완료되었습니다.")
    print(f"파일명: {output_filename}")

if __name__ == "__main__":
    create_infomax_excel()
