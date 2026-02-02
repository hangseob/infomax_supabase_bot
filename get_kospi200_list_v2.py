import xlwings as xw
import os
import time
import sys

# 인코딩 설정
if sys.stdout.encoding != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except:
        pass

def get_kospi200_list_v2():
    # 1. 설정
    group_code = "211"  # KOSPI 200 그룹코드
    infomax_xlam_path = r"C:\Infomax\bin\excel\infomaxexcel.xlam"
    
    print(f"KOSPI 200 종목 리스트 조회를 시작합니다. (그룹코드: {group_code})")
    
    app = None
    try:
        # 2. 엑셀 실행 및 애드인 로드
        if xw.apps.count > 0:
            app = xw.apps.active
            print("기존 Excel 인스턴스에 연결합니다.")
        else:
            app = xw.App(visible=True, add_book=False)
            print("새로운 Excel 인스턴스를 실행합니다.")

        # 인포맥스 애드인이 없으면 로드 시도
        is_addin_loaded = False
        for wb in app.books:
            if "infomaxexcel" in wb.name.lower():
                is_addin_loaded = True
                break
        
        if not is_addin_loaded and os.path.exists(infomax_xlam_path):
            print(f"인포맥스 애드인을 로드합니다: {infomax_xlam_path}")
            app.books.open(infomax_xlam_path)
            time.sleep(5) # 로딩 대기시간 증가
        
        # 3. 새 워크북 추가
        wb = app.books.add()
        sheet = wb.sheets[0]
        
        # 4. 다양한 수식 시도
        # 인덱스 코드로 시도 (보통 지수는 IDX, 그룹은 숫자로만)
        formulas = [
            f'=IMDG("211", "한글종목명,종목코드", 200, "Orient=V")',
            f'=IMDG(211, "종목명,종목코드", 200, "Orient=V")',
            f'=IMDI("211", "한글종목명,종목코드", "Orient=V")'
        ]
        
        for idx, formula in enumerate(formulas):
            target_cell = sheet.range(1, idx * 3 + 1)
            print(f"수식 입력 ({idx+1}): {formula}")
            target_cell.formula = formula
            
        # 5. 데이터 로딩 대기 (충분히 대기)
        print("데이터 수신 대기 중 (30초)...")
        for i in range(30):
            time.sleep(1)
            if i % 10 == 0:
                print(f"{i}초 경과...")
                
        # 6. 결과 확인
        print("\n[결과 확인]")
        found = False
        for idx in range(len(formulas)):
            col_idx = idx * 3 + 1
            data = sheet.range((1, col_idx), (205, col_idx + 1)).value
            valid_rows = [r for r in data if r[0] is not None and not str(r[0]).startswith('#')]
            
            if len(valid_rows) > 5:
                print(f"수식 {idx+1} 성공! {len(valid_rows)}개 종목 발견.")
                print(f"샘플: {valid_rows[:3]}")
                found = True
                break
        
        if found:
            save_path = os.path.abspath("kospi200_items.xlsx")
            wb.save(save_path)
            print(f"\n파일이 저장되었습니다: {save_path}")
        else:
            print("\n모든 수식이 실패했습니다. 인포맥스 로그인 상태와 엑셀 메뉴의 '인포맥스' 탭 활성화 여부를 확인해주세요.")

    except Exception as e:
        print(f"오류 발생: {e}")

if __name__ == "__main__":
    get_kospi200_list_v2()
