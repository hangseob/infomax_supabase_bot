import xlwings as xw
import datetime
import time
import os

def get_vietnam_report():
    today = datetime.date(2026, 1, 14)
    one_year_ago = today - datetime.timedelta(days=365)
    
    today_str = today.strftime("%Y%m%d")
    prev_str = one_year_ago.strftime("%Y%m%d")
    
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
            try: app.books.open(xlam_path)
            except: pass
            
        wb = app.books.add()
        sheet = wb.sheets[0]
        
        # 1. 1년 전 주가 (IMDH)
        # 2. 현재 주가 (IMDP)
        
        sheet.range("A1").value = "구분"
        sheet.range("B1").value = "날짜"
        sheet.range("C1").value = "지수값"
        
        sheet.range("A2").value = "1년 전"
        sheet.range("B2").value = prev_str
        sheet.range("C2").formula = f'=IMDH("FRN", "VNI:VIDX", "현재가", "{prev_str}", "{prev_str}", 1, "Per=D,Headers=0")'
        
        sheet.range("A3").value = "현재"
        sheet.range("B3").value = today_str
        sheet.range("C3").formula = f'=IMDP("FRN", "VNI:VIDX", "현재가")'
        
        print("데이터 로딩 중...")
        time.sleep(15)
        
        # IMDH는 수식 셀 옆(D열)에 값이 나올 수 있으므로 범위를 확인
        prev_val = sheet.range("D2").value # IMDH 결과값 (Headers=0 일 때)
        curr_val = sheet.range("C3").value # IMDP 결과값
        
        print(f"\n[베트남 호치민 지수 (VNI:VIDX) 조회 결과]")
        print(f"- 1년 전 ({prev_str}): {prev_val}")
        print(f"- 현재 ({today_str}): {curr_val}")
        
        if isinstance(prev_val, (int, float)) and isinstance(curr_val, (int, float)):
            change = curr_val - prev_val
            pct_change = (change / prev_val) * 100
            print(f"- 수익률: {pct_change:.2f}%")
        
        save_path = os.path.abspath("vietnam_final_report.xlsx")
        wb.save(save_path)
        
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    get_vietnam_report()
