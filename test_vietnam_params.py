import xlwings as xw
import datetime
import time
import os

def fetch_vietnam_history():
    # 1. 날짜 설정 (1년 전)
    # 오늘이 2026-01-14 라고 가정 (사용자 정보 기준)
    target_date = datetime.date(2025, 1, 14) 
    date_str = target_date.strftime("%Y%m%d")
    
    app = None
    wb = None
    try:
        if xw.apps.count > 0:
            app = xw.apps.active
        else:
            app = xw.App(visible=True, add_book=False)
        
        # 인포맥스 애드인 로드
        xlam_path = r"C:\Infomax\bin\excel\infomaxexcel.xlam"
        if os.path.exists(xlam_path):
            try:
                app.books.open(xlam_path)
            except:
                pass
        
        wb = app.books.add()
        sheet = wb.sheets[0]
        
        time.sleep(3)
        
        # 테스트할 조합들
        tests = [
            ("FRN", "VNI:VIDX", "현재가"),
            ("FRN", "VNI:VIDX", "종가"),
            ("IDX", "VNI:VIDX", "현재가"),
            ("VIDX", "VNI", "현재가"),
        ]
        
        sheet.range("A1").value = ["구분", "수식", "결과"]
        
        for i, (m, s, f) in enumerate(tests, 2):
            formula = f'=IMDH("{m}", "{s}", "{f}", "{date_str}", "{date_str}", 1, "Per=D,Orient=V,Headers=0")'
            sheet.range(f"A{i}").value = f"{m}/{s}/{f}"
            sheet.range(f"B{i}").formula = formula
        
        print("데이터 로딩 대기 (15초)...")
        time.sleep(15)
        
        # 결과 출력
        results = sheet.range(f"A1:C{len(tests)+1}").value
        for row in results:
            print(row)
            
        save_path = os.path.abspath("vietnam_test_results.xlsx")
        wb.save(save_path)
        print(f"결과 저장: {save_path}")
        
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    fetch_vietnam_history()
