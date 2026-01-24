import xlwings as xw
import pandas as pd
import time
from datetime import datetime
import os
import sys

def create_sample_02():
    # 0. 폴더 위치 및 경로 설정
    base_dir = r'C:\git_repository\infomax_supabase_bot'
    new_file_name = 'infomax_ficc_data_sample_02.xlsx'
    new_file = os.path.join(base_dir, new_file_name)
    fields_path = os.path.join(base_dir, 'infomax_functions_templetes', 'mmkt_infomax_fields.xlsx')
    
    print(f"새로운 샘플 파일 생성을 시작합니다: {new_file_name}", flush=True)
    
    # 기존 파일이 있다면 삭제하고 새로 시작 (사용자가 처음부터 다시하자고 함)
    if os.path.exists(new_file):
        try:
            # 엑셀에서 열려있을 수 있으므로 먼저 닫기 시도
            if xw.apps.count > 0:
                for app in xw.apps:
                    for wb in app.books:
                        if wb.name == new_file_name:
                            wb.close()
            os.remove(new_file)
            print(f"기존 {new_file_name} 파일을 삭제했습니다.")
        except Exception as e:
            print(f"기존 파일 삭제 중 오류 (이미 열려있을 수 있음): {e}")

    # 1. 전체 필드 정보 읽기
    try:
        df_fields = pd.read_excel(fields_path, sheet_name='Sheet2', header=1)
        df_fields = df_fields.dropna(subset=['RATE_ID', 'DATA_TYPE', 'DATA_ID', 'FIELD_ID'])
        
        # 첫 3종목만 선택
        df_to_process = df_fields.head(3)
        print(f"테스트를 위해 상위 3개 종목만 처리를 시작합니다.", flush=True)
    except Exception as e:
        print(f"필드 정보 로드 실패: {e}", flush=True)
        return

    # 2. 엑셀 실행
    try:
        app = xw.App(visible=True)
        wb = app.books.add()
        wb.save(new_file)

        final_sheet = wb.sheets[0]
        final_sheet.name = "FinalTable"
        final_sheet.range("A1").value = ["날짜", "코드", "값"]
        final_row = 2

        scratch_sheet = wb.sheets.add("Scratch")
        
        infomax_xlam = r"C:\Infomax\bin\excel\infomaxexcel.xlam"
        if os.path.exists(infomax_xlam):
            try: app.books.open(infomax_xlam)
            except: pass

    except Exception as e:
        print(f"엑셀 초기화 오류: {e}", flush=True)
        return

    # 3. 데이터 수집 (상위 3개)
    start_date = "20260101"
    end_date = datetime.now().strftime("%Y%m%d")
    to_process_list = df_to_process.to_dict('records')
    total = len(to_process_list)
    
    for i, row in enumerate(to_process_list):
        rate_id = row['RATE_ID']
        scale = row['SCALE_FACTOR'] if not pd.isna(row['SCALE_FACTOR']) else 1.0
        
        print(f"[{i+1}/{total}] {rate_id} 요청 중", end='', flush=True)
        
        # 일자, 값 동시 호출
        formula = f'=IMDH("{row["DATA_TYPE"]}", "{row["DATA_ID"]}", "일자,{row["FIELD_ID"]}", "{start_date}", "{end_date}", 100, "Headers=0,Orient=V,Per=D")'
        
        try:
            scratch_sheet.clear_contents()
            scratch_sheet.range("A1").formula = formula
            
            success = False
            for _ in range(15):
                time.sleep(2)
                print(".", end='', flush=True)
                data = scratch_sheet.range("A1:C101").value # A: Title, B: 일자, C: 값
                
                if data and (data[0][0] or data[0][1]) and "#WAITING" not in str(data[0][0] or "").upper():
                    records = []
                    for r in data:
                        if r[1] is None and r[2] is None:
                            continue
                        if r[1] is not None:
                            val = r[2] * scale if isinstance(r[2], (int, float)) else r[2]
                            records.append([r[1], rate_id, val])
                    
                    if records:
                        final_sheet.range(f"A{final_row}").value = records
                        final_row += len(records)
                        sample = f"[{records[0][0]}, {records[0][2]}]"
                        print(f" 완료! {len(records)}행 {sample}", flush=True)
                        wb.save()
                    else:
                        print(f" 데이터 없음", flush=True)
                    success = True
                    break
            
            if not success:
                print(f" 타임아웃/실패", flush=True)
        except Exception as e:
            print(f" 오류: {e}", flush=True)

    print(f"\n테스트 작업 완료. 'FinalTable' 시트를 확인해 주세요.", flush=True)
    try:
        final_sheet.autofit()
        wb.save()
    except: pass

if __name__ == "__main__":
    create_sample_02()
