import xlwings as xw
import pandas as pd
import time
from datetime import datetime
import os

def create_infomax_sequential():
    print("인포맥스 데이터 순차 추출을 시작합니다 (날짜-값 매칭)...")
    
    # 1. 필드 정보 읽기
    fields_path = r'infomax_functions_templetes/mmkt_infomax_fields.xlsx'
    try:
        df_fields = pd.read_excel(fields_path, sheet_name='Sheet2', header=1)
        df_fields = df_fields.dropna(subset=['RATE_ID', 'DATA_TYPE', 'DATA_ID', 'FIELD_ID'])
        # USE_FLAG이 있는 경우 필터링 (있다면)
        if 'USE_FLAG' in df_fields.columns:
            df_fields = df_fields[df_fields['USE_FLAG'] == 'Y']
        print(f"총 {len(df_fields)}개의 RATE_ID를 순차적으로 처리합니다.")
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
            except:
                pass
        
        wb = app.books.add()
        # 데이터 수집용 임시 시트
        scratch_sheet = wb.sheets[0]
        scratch_sheet.name = "Scratch"
        
        # 최종 결과 시트
        final_sheet = wb.sheets.add("FinalTable")
        final_sheet.range("A1").value = ["날짜", "코드", "값"]
        final_row = 2
        
    except Exception as e:
        print(f"엑셀 초기화 중 오류: {e}")
        return

    # 3. 순차적 데이터 수집
    start_date = "20260101"
    end_date = datetime.now().strftime("%Y%m%d")
    print(f"조회 기간: {start_date} ~ {end_date}")
    
    total_count = len(df_fields)
    collected_records = []
    
    for i, (_, row) in enumerate(df_fields.iterrows()):
        rate_id = row.RATE_ID
        field_id = row.FIELD_ID
        scale = row.SCALE_FACTOR if not pd.isna(row.SCALE_FACTOR) else 1.0
        
        # IMDH(시장, 코드, "날짜,필드", 시작, 종료, 건수, 옵션)
        # Headers=0 (헤더 제외), Orient=V (수직)
        formula = f'=IMDH("{row.DATA_TYPE}", "{row.DATA_ID}", "날짜,{field_id}", "{start_date}", "{end_date}", 100, "Headers=0,Orient=V,Per=D")'
        
        print(f"[{i+1}/{total_count}] {rate_id} 처리 중...", end='\r')
        
        # 임시 시트 청소 및 수식 입력
        scratch_sheet.clear_contents()
        scratch_sheet.range("A1").formula = formula
        
        # 데이터 로드 대기 (개별 종목별로 짧게 반복 체크)
        success = False
        for retry in range(15): # 최대 30초 (2초 * 15)
            time.sleep(2)
            try:
                # A1:B100 영역 확인
                data = scratch_sheet.range("A1:B100").value
                if data and data[0][0] is not None:
                    first_val = data[0][0]
                    if isinstance(first_val, str):
                        if "#WAITING" in first_val.upper():
                            continue
                        if "#NAME?" in first_val.upper():
                            print(f"\n[!] {rate_id}: #NAME? 에러 (애드인 확인 필요)")
                            break
                        if "#" in first_val: # 기타 에러
                            print(f"\n[!] {rate_id}: 에러 발생 ({first_val})")
                            break
                    
                    # 유효한 데이터가 있으면 리스트에 추가
                    for r in data:
                        if r[0] is None: break # 데이터 끝
                        # r[0]은 날짜, r[1]은 값
                        val = r[1]
                        if val is not None and not isinstance(val, str):
                            val = val * scale
                        collected_records.append([r[0], rate_id, val])
                    
                    success = True
                    break
            except Exception as e:
                # 엑셀이 바쁜 경우 (Edit 모드 등)
                continue
        
        if not success:
            print(f"\n[!] {rate_id}: 데이터 로드 실패 또는 데이터 없음")

        # 중간중간 최종 시트에 써주기 (메모리 관리 및 진행 확인용, 50개마다)
        if len(collected_records) >= 500:
            final_sheet.range(f"A{final_row}").value = collected_records
            final_row += len(collected_records)
            collected_records = []

    # 남은 데이터 작성
    if collected_records:
        final_sheet.range(f"A{final_row}").value = collected_records
        final_row += len(collected_records)

    # 4. 마무리
    print(f"\n\n작업 완료! 총 {final_row - 2}행의 데이터가 수집되었습니다.")
    
    # 서식 정리
    final_sheet.range("A1:C1").color = (200, 200, 200)
    final_sheet.autofit()
    
    # 임시 시트 삭제 (선택 사항)
    # scratch_sheet.delete()
    
    output_filename = f"infomax_sequential_table_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    wb.save(output_filename)
    print(f"파일명: {output_filename}")

if __name__ == "__main__":
    create_infomax_sequential()
