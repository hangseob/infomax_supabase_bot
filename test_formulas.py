import xlwings as xw
import os
import time

def test_multiple_formulas():
    infomax_xlam_path = r"C:\Infomax\bin\excel\infomaxexcel.xlam"
    group_code = "211" # KOSPI 200
    
    try:
        app = xw.App(visible=True, add_book=False)
        if os.path.exists(infomax_xlam_path):
            app.books.open(infomax_xlam_path)
            time.sleep(2)
        
        wb = app.books.add()
        sheet = wb.sheets[0]
        
        # 시도할 공식 목록
        formulas = [
            f'=IMDG("{group_code}", "한글종목명,종목코드", 200, "Orient=V")',
            f'=IMDG({group_code}, "한글종목명,종목코드", 200, "Orient=V")',
            f'=IMDI("{group_code}", "한글종목명,종목코드", "Orient=V")',
            f'=IMDI({group_code}, "한글종목명,종목코드", "Orient=V")',
            f'=IMDG("211", "종목명,코드", 200, "Orient=V")',
            f'=IMDB("211", "한글종목명,종목코드", "Orient=V")' # IMDB는 보통 Basket/Group
        ]
        
        for i, f in enumerate(formulas):
            cell = sheet.range(1, i*3 + 1)
            cell.value = f"Formula {i+1}: {f}"
            sheet.range(2, i*3 + 1).formula = f
            print(f"Testing in column {i*3 + 1}: {f}")
            
        print("데이터 수신 대기 중 (20초)...")
        time.sleep(20)
        
        # 결과 확인
        for i in range(len(formulas)):
            col = i*3 + 1
            val = sheet.range(2, col).value
            print(f"Formula {i+1} result at (2, {col}): {val}")
            if val and not str(val).startswith('#'):
                print(f"  -> SUCCESS! Formula {i+1} worked.")
        
        save_path = os.path.abspath("kospi200_test_results.xlsx")
        wb.save(save_path)
        print(f"테스트 결과 저장 완료: {save_path}")

    except Exception as e:
        print(f"오류: {e}")

if __name__ == "__main__":
    test_multiple_formulas()
