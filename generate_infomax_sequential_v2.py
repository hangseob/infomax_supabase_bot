import xlwings as xw
import pandas as pd
import time
from datetime import datetime
import os

def create_infomax_sequential_final():
    print("인포맥스 데이터 순차 추출을 시작합니다 (일자-값 정확 매칭)...")
    
    # 1. 필드 정보 읽기
    fields_path = r'infomax_functions_templetes/mmkt_infomax_fields.xlsx'
    try:
        df_fields = pd.read_excel(fields_path, sheet_name='Sheet2', header=1)
        df_fields = df_fields.dropna(subset=['RATE_ID', 'DATA_TYPE', 'DATA_ID', 'FIELD_ID'])
        print(f"총 {len(df_fields)}개의 RATE_ID를 순차적으로 처리합니다.")
    except Exception as e:
        print(f"필드 정보를 읽는 중 오류 발생: {e}")
        return

    # 2. 엑셀 실행 및 준비
    try:
        app = xw.apps.active if xw.apps.count > 0 else xw.App(visible=True)
        infomax_xlam_path = r"C:\Infomax\bin\excel\infomaxexcel.xlam"
        if os.path.exists(infomax_xlam_path):
            try:
                app.books.open(infomax_xlam_path)
            except:
                pass
        
        wb = app.books.add()
        scratch_sheet = wb.sheets[0]
        scratch_sheet.name = "Scratch"
        
        final_sheet = wb.sheets.add("FinalTable")
        final_sheet.range("A1").value = ["날짜", "코드", "값"]
        final_row = 2
    except Exception as e:
        print(f"엑셀 초기화 오류: {e}")
        return

    # 3. 데이터 수집
    start_date = "20260101"
    end_date = datetime.now().strftime("%Y%m%d")
    
    total_count = len(df_fields)
    batch_records = []
    
    for i, (_, row) in enumerate(df_fields.iterrows()):
        rate_id = row.RATE_ID
        market = row.DATA_TYPE
        code = row.DATA_ID
        field = row.FIELD_ID
        scale = row.SCALE_FACTOR if not pd.isna(row.SCALE_FACTOR) else 1.0
        
        # '일자' 키워드 사용, Headers=0 (그래도 Title은 나옴)
        formula = f'=IMDH("{market}", "{code}", "일자,{field}", "{start_date}", "{end_date}", 100, "Headers=0,Orient=V,Per=D")'
        
        print(f"[{i+1}/{total_count}] {rate_id} 처리 중...", end='\r')
        
        scratch_sheet.clear_contents()
        scratch_sheet.range("A1").formula = formula
        
        # 데이터 로드 대기
        success = False
        for retry in range(15):
            time.sleep(2)
            try:
                # Title(A), 일자(B), 값(C) - 최대 100행까지
                data = scratch_sheet.range("A1:C101").value
                if data and data[0][0] is not None:
                    first_cell = data[0][0]
                    if isinstance(first_cell, str) and "#WAITING" in first_cell.upper():
                        continue
                    if isinstance(first_cell, str) and "#NAME?" in first_cell.upper():
                        print(f"\n[!] {rate_id}: #NAME? 에러")
                        break
                    
                    # 데이터 파싱 (2번째, 3번째 컬럼)
                    for r in data:
                        if r[0] is None: break # 끝
                        date_val = r[1]
                        price_val = r[2]
                        if date_val is not None:
                            if price_val is not None and not isinstance(price_val, str):
                                price_val = price_val * scale
                            batch_records.append([date_val, rate_id, price_val])
                    
                    success = True
                    break
            except:
                continue
        
        if not success:
            print(f"\n[!] {rate_id}: 데이터 수신 실패")

        # 50개 종목마다 최종 시트에 쓰기
        if (i + 1) % 50 == 0:
            if batch_records:
                final_sheet.range(f"A{final_row}").value = batch_records
                final_row += len(batch_records)
                batch_records = []

    # 남은 데이터 쓰기
    if batch_records:
        final_sheet.range(f"A{final_row}").value = batch_records
        final_row += len(batch_records)

    # 4. 마무리
    print(f"\n\n작업 완료! 총 {final_row - 2}행 수집.")
    final_sheet.range("A1:C1").color = (200, 200, 200)
    final_sheet.autofit()
    
    output_filename = f"infomax_flat_table_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    wb.save(output_filename)
    print(f"최종 파일: {output_filename}")

if __name__ == "__main__":
    create_infomax_sequential_final()
