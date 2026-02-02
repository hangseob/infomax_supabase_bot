import xlwings as xw
import pandas as pd
import time
from datetime import datetime
import os

def create_infomax_excel():
    print("인포맥스 데이터 추출 및 엑셀 생성을 시작합니다...")
    
    # 1. 필드 정보 읽기
    fields_path = r'infomax_functions_templetes/mmkt_infomax_fields.xlsx'
    try:
        df = pd.read_excel(fields_path, sheet_name='Sheet2', header=1)
        # 필요한 컬럼만 추출 및 유효한 데이터 필터링
        df = df.dropna(subset=['RATE_ID', 'DATA_TYPE', 'DATA_ID', 'FIELD_ID'])
        print(f"총 {len(df)}개의 RATE_ID를 처리합니다.")
    except Exception as e:
        print(f"필드 정보를 읽는 중 오류 발생: {e}")
        return

    # 2. 엑셀 실행 및 워크북 생성
    try:
        # 실행 중인 엑셀이 있으면 연결, 없으면 새로 실행
        if xw.apps.count > 0:
            app = xw.apps.active
            print("기존 엑셀 인스턴스에 연결했습니다.")
        else:
            app = xw.App(visible=True)
            print("새로운 엑셀 인스턴스를 실행했습니다.")
        
        # 인포맥스 애드인(xlam) 명시적으로 열기
        infomax_xlam_path = r"C:\Infomax\bin\excel\infomaxexcel.xlam"
        if os.path.exists(infomax_xlam_path):
            print(f"인포맥스 애드인을 로드합니다: {infomax_xlam_path}")
            try:
                app.books.open(infomax_xlam_path)
            except Exception as e:
                print(f"애드인 로드 중 오류 (이미 열려있을 수 있음): {e}")
        else:
            print(f"경고: 애드인 파일을 찾을 수 없습니다: {infomax_xlam_path}")

        wb = app.books.add()
        sheet = wb.sheets[0]
        sheet.name = "InfomaxData"
    except Exception as e:
        print(f"엑셀 연결 중 오류 발생: {e}")
        return

    # 3. 데이터 및 수식 작성
    start_date = "20260101"
    end_date = datetime.now().strftime("%Y%m%d")
    
    print(f"조회 기간: {start_date} ~ {end_date}")
    
    headers = ["Date"]
    formulas = []
    
    # 첫 번째 종목으로 날짜 가져오기 위한 수식
    first_row = df.iloc[0]
    date_formula = f'=IMDH("{first_row.DATA_TYPE}", "{first_row.DATA_ID}", "날짜", "{start_date}", "{end_date}", 100, "Per=D,Headers=0,Orient=V")'
    
    for i, (_, row) in enumerate(df.iterrows()):
        headers.append(row.RATE_ID)
        scale = row.SCALE_FACTOR if not pd.isna(row.SCALE_FACTOR) else 1.0
        # 필드명 및 인자들 구성
        formula = f'=IMDH("{row.DATA_TYPE}", "{row.DATA_ID}", "{row.FIELD_ID}", "{start_date}", "{end_date}", 100, "Per=D,Headers=0,Orient=V") * {scale}'
        formulas.append(formula)

    print("헤더 및 수식 작성 중...")
    # 헤더 한꺼번에 작성
    sheet.range("A1").value = headers
    
    # 날짜 수식 작성
    sheet.range("A2").formula = date_formula
    
    # 데이터 수식 한꺼번에 작성 (B2부터 가로로)
    sheet.range("B2").formula = formulas
    
    print(f"모든 수식({len(formulas)}개) 입력 완료. 데이터 수신 대기 중 (최대 5분)...")

    # 4. 데이터 로딩 확인 및 에러 수정 루프
    max_retries = 30 
    retry_count = 0
    
    while retry_count < max_retries:
        time.sleep(10)
        retry_count += 1
        
        # 샘플 체크 (A2:F5 영역)
        check_range = sheet.range("A2:F5").value
        
        waiting = False
        error_found = False
        
        if check_range is None:
            waiting = True
        else:
            for r_idx, row_vals in enumerate(check_range):
                for c_idx, val in enumerate(row_vals):
                    if val is None or val == "":
                        waiting = True
                        break
                    if isinstance(val, str):
                        if "#WAITING" in val.upper():
                            waiting = True
                            break
                        if "#NAME?" in val.upper():
                            print(f"\n[!] #NAME? 에러 발생 ({xw.utils.address(r_idx+2, c_idx+1)}). 애드인이 정상 작동하지 않습니다.")
                            error_found = True
                            break
                        if "#" in val:
                            # 다른 에러 (#VALUE!, #REF! 등)는 일단 기다려보거나 재시도 대상
                            pass
                if waiting or error_found:
                    break
        
        if error_found:
            break
            
        if not waiting:
            # 마지막 행 확인
            last_row = sheet.range("A2").end('down').row
            if last_row > 1 and last_row < 100000:
                print(f"\n성공! 데이터 로드 완료. (총 {last_row - 1}개 행)")
                break
        
        print(f"[{retry_count}/{max_retries}] 데이터 수신 대기 중... (샘플 확인 중)", end='\r')

    # 결과 저장
    output_filename = f"infomax_all_rates_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    wb.save(output_filename)
    print(f"\n작업 완료. 파일 저장됨: {output_filename}")

if __name__ == "__main__":
    create_infomax_excel()
