import xlwings as xw
from datetime import datetime, timedelta
import os
import time

def get_sk_hynix_history():
    # 1. 설정
    stock_code = "000660"  # SK하이닉스
    infomax_xlam_path = r"C:\Infomax\bin\excel\infomaxexcel.xlam"
    
    end_date = datetime(2026, 1, 12)  # 오늘 날짜
    start_date = end_date - timedelta(days=31)
    
    start_str = start_date.strftime("%Y%m%d")
    end_str = end_date.strftime("%Y%m%d")
    
    print(f"조회 기간: {start_str} ~ {end_str}")
    
    try:
        # 2. 엑셀 실행
        print("엑셀을 실행합니다...")
        app = xw.App(visible=True, add_book=False)
        
        # 3. 인포맥스 애드인(xlam) 명시적으로 열기
        if os.path.exists(infomax_xlam_path):
            print(f"인포맥스 애드인을 로드합니다: {infomax_xlam_path}")
            app.books.open(infomax_xlam_path)
            time.sleep(2) # 애드인 로딩 대기
        else:
            print(f"경고: 애드인 파일을 찾을 수 없습니다: {infomax_xlam_path}")
        
        # 4. 새 통합문서 생성
        wb = app.books.add()
        sheet = wb.sheets[0]
        
        # 5. 인포맥스 수식 입력
        fields = "날짜,현재가"
        options = "Per=D,Orient=V,sort=1"
        formula = f'=IMDH("STK", "{stock_code}", "{fields}", "{start_str}", "{end_str}", 35, "{options}")'
        
        print(f"수식 입력: {formula}")
        sheet.range("A1").formula = formula
        
        # 6. 데이터 로딩 대기
        print("데이터 로딩 대기 중 (15초)...")
        time.sleep(15)
        
        # 7. 데이터 확인
        data = sheet.range("A1:B30").value
        valid_rows = [r for r in data if r[0] is not None]
        print(f"가져온 데이터 건수: {len(valid_rows)}건")
        
        if len(valid_rows) > 1:
            print("상위 5개 데이터:")
            for row in valid_rows[:5]:
                print(f"  {row}")
        else:
            print("데이터가 아직 로드되지 않았거나 결과가 없습니다. (인포맥스 로그인 확인 필요)")
            
        # 파일 저장
        save_path = os.path.abspath("sk_hynix_price.xlsx")
        wb.save(save_path)
        print(f"파일 저장 완료: {save_path}")

    except Exception as e:
        print(f"오류 발생: {e}")

if __name__ == "__main__":
    get_sk_hynix_history()
