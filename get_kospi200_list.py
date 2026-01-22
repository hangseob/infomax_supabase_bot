import xlwings as xw
import os
import time

def get_kospi200_list():
    # 1. 설정
    group_code = "211"  # KOSPI 200 그룹코드
    infomax_xlam_path = r"C:\Infomax\bin\excel\infomaxexcel.xlam"
    
    print(f"KOSPI 200 종목 리스트 조회를 시작합니다. (그룹코드: {group_code})")
    
    try:
        # 2. 엑셀 실행 및 애드인 로드
        app = xw.App(visible=True, add_book=False)
        if os.path.exists(infomax_xlam_path):
            app.books.open(infomax_xlam_path)
            time.sleep(2)
        
        # 3. 새 통합문서 생성
        wb = app.books.add()
        sheet = wb.sheets[0]
        
        # 4. IMDG 함수 입력 (그룹 종목 리스트 조회)
        # 다양한 필드명과 그룹코드 형식을 시도해 봅니다.
        fields = "한글종목명,종목코드"
        options = "Orient=V"
        
        # 그룹코드에 따옴표가 있는 경우와 없는 경우 모두 고려
        formula = f'=IMDG({group_code}, "{fields}", 200, "{options}")'
        
        print(f"수식 입력: {formula}")
        sheet.range("A1").formula = formula
        
        # 5. 데이터 로딩 대기 및 확인 루프 (최대 30초)
        print("데이터 수신 대기 중...")
        valid_rows = []
        for i in range(30):
            time.sleep(1)
            # A열부터 B열까지 데이터 확인
            data = sheet.range("A1:B205").value
            # 오류 코드(#NAME?, #N/A 등)가 아닌 실제 데이터가 있는지 확인
            valid_rows = [r for r in data if r[0] is not None and not str(r[0]).startswith('#')]
            
            if len(valid_rows) > 1: # 헤더 외에 데이터가 들어오기 시작하면 중단
                break
                
            if i % 5 == 0:
                current_val = sheet.range("A1").value
                print(f"{i}초 경과... (A1 현재값: {current_val})")
        
        # 6. 결과 확인 및 저장
        print(f"조회된 종목 수: {len(valid_rows)}개")
        
        if len(valid_rows) > 0:
            print("상위 10개 종목:")
            for row in valid_rows[:10]:
                print(f"  {row[0]} ({row[1]})")
            
            save_path = os.path.abspath("kospi200_items.xlsx")
            wb.save(save_path)
            print(f"파일이 저장되었습니다: {save_path}")
        else:
            # 실패 시 다른 수식 시도 (IMDI)
            print("IMDG 실패, IMDI로 재시도합니다...")
            formula_alt = f'=IMDI("{group_code}", "{fields}", "{options}")'
            sheet.range("D1").formula = formula_alt
            time.sleep(5)
            data_alt = sheet.range("D1:E205").value
            valid_alt = [r for r in data_alt if r[0] is not None and not str(r[0]).startswith('#')]
            if len(valid_alt) > 0:
                print(f"IMDI로 {len(valid_alt)}개 종목을 찾았습니다.")
                # 저장 로직...
            else:
                print("데이터를 가져오지 못했습니다. 엑셀에서 수식 상태를 확인해 주세요.")

    except Exception as e:
        print(f"오류 발생: {e}")

if __name__ == "__main__":
    get_kospi200_list()
