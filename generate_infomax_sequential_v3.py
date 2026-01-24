import xlwings as xw
import pandas as pd
import time
from datetime import datetime
import os

def create_infomax_sequential_realtime_write():
    print("인포맥스 데이터 순차 추출 및 즉시 쓰기를 시작합니다 (일자-값 정확 매칭)...")
    
    # 1. 필드 정보 읽기
    fields_path = r'infomax_functions_templetes/mmkt_infomax_fields.xlsx'
    try:
        df_fields = pd.read_excel(fields_path, sheet_name='Sheet2', header=1)
        df_fields = df_fields.dropna(subset=['RATE_ID', 'DATA_TYPE', 'DATA_ID', 'FIELD_ID'])
        print(f"총 {len(df_fields)}개의 RATE_ID를 처리합니다.")
    except Exception as e:
        print(f"필드 정보를 읽는 중 오류 발생: {e}")
        return

    # 2. 엑셀 실행 및 준비
    try:
        # 기존 엑셀 앱 사용 또는 새로 열기
        app = xw.apps.active if xw.apps.count > 0 else xw.App(visible=True)
        
        # 인포맥스 애드인 로드 확인
        infomax_xlam_path = r"C:\Infomax\bin\excel\infomaxexcel.xlam"
        if os.path.exists(infomax_xlam_path):
            try:
                app.books.open(infomax_xlam_path)
            except:
                pass
        
        wb = app.books.add()
        
        # 데이터 수집용 임시 시트
        scratch_sheet = wb.sheets[0]
        scratch_sheet.name = "Scratch"
        
        # 최종 결과 시트 생성 및 헤더 작성
        final_sheet = wb.sheets.add("FinalTable")
        final_sheet.range("A1").value = ["날짜", "코드", "값"]
        final_row = 2
        
        print("엑셀 파일 준비 완료. 데이터를 순차적으로 기록합니다.")
    except Exception as e:
        print(f"엑셀 초기화 오류: {e}")
        return

    # 3. 데이터 수집 및 즉시 기록
    start_date = "20260101"
    end_date = datetime.now().strftime("%Y%m%d")
    
    total_count = len(df_fields)
    
    for i, (_, row) in enumerate(df_fields.iterrows()):
        rate_id = row.RATE_ID
        market = row.DATA_TYPE
        code = row.DATA_ID
        field = row.FIELD_ID
        scale = row.SCALE_FACTOR if not pd.isna(row.SCALE_FACTOR) else 1.0
        
        # '일자' 키워드 사용하여 날짜와 값을 동시에 요청
        formula = f'=IMDH("{market}", "{code}", "일자,{field}", "{start_date}", "{end_date}", 100, "Headers=0,Orient=V,Per=D")'
        
        # Scratch 시트 비우고 수식 입력
        scratch_sheet.clear_contents()
        scratch_sheet.range("A1").formula = formula
        
        records_to_write = []
        success = False
        
        # 데이터 로드 대기 (최대 30초)
        for retry in range(15):
            time.sleep(2)
            try:
                # 결과 범위 (Title, 일자, 값 순서로 들어옴)
                data = scratch_sheet.range("A1:C101").value
                if data and data[0][0] is not None:
                    first_cell = data[0][0]
                    if isinstance(first_cell, str) and "#WAITING" in first_cell.upper():
                        continue
                    if isinstance(first_cell, str) and "#NAME?" in first_cell.upper():
                        print(f"[{i+1}/{total_count}] {rate_id}: #NAME? 에러 발생 (인포맥스 메뉴 확인 필요)")
                        break
                    
                    # 데이터 파싱 및 스케일 적용
                    for r in data:
                        if r[0] is None: break # 데이터 끝
                        date_val = r[1]
                        price_val = r[2]
                        if date_val is not None:
                            if price_val is not None and not isinstance(price_val, str):
                                price_val = price_val * scale
                            records_to_write.append([date_val, rate_id, price_val])
                    
                    success = True
                    break
            except:
                # 엑셀이 계산 중이거나 다른 상호작용으로 인해 오류가 날 수 있음
                continue
        
        # 성공 시 즉시 FinalTable에 기록
        if success and records_to_write:
            final_sheet.range(f"A{final_row}").value = records_to_write
            final_row += len(records_to_write)
            print(f"[{i+1}/{total_count}] {rate_id}: 처리 완료 ({len(records_to_write)}행 추가됨)")
        else:
            if not success:
                print(f"[{i+1}/{total_count}] {rate_id}: 데이터 로드 실패 (Time out)")
            else:
                print(f"[{i+1}/{total_count}] {rate_id}: 수신된 데이터가 없음")

    # 4. 마무리 및 저장
    print(f"\n모든 종목의 작업이 완료되었습니다! 총 {final_row - 2}행의 데이터가 수집되었습니다.")
    
    # 헤더 서식 및 열 너비 자동 맞춤
    final_sheet.range("A1:C1").color = (200, 200, 200)
    final_sheet.autofit()
    
    output_filename = f"infomax_sequential_write_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    wb.save(output_filename)
    print(f"최종 파일 저장 완료: {output_filename}")

if __name__ == "__main__":
    create_infomax_sequential_realtime_write()
