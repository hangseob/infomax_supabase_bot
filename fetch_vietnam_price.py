import xlwings as xw
import datetime
import time
import os

def get_vietnam_price():
    # 1. 날짜 설정 (1년 전)
    today = datetime.date(2026, 1, 14)
    one_year_ago = today - datetime.timedelta(days=365)
    
    start_str = one_year_ago.strftime("%Y%m%d")
    end_str = one_year_ago.strftime("%Y%m%d")
    
    # 2. 엑셀 실행 또는 연결
    app = None
    wb = None
    try:
        if xw.apps.count > 0:
            app = xw.apps.active
            print("기존 엑셀 인스턴스에 연결합니다.")
        else:
            app = xw.App(visible=True, add_book=False)
            print("새로운 엑셀 인스턴스를 실행합니다.")
        
        # 인포맥스 애드인 로드 (이미 열려있을 수도 있지만 안전을 위해)
        xlam_path = r"C:\Infomax\bin\excel\infomaxexcel.xlam"
        if os.path.exists(xlam_path):
            try:
                app.books.open(xlam_path)
                print(f"인포맥스 애드인 로드 완료: {xlam_path}")
            except Exception as e:
                print(f"애드인 로드 중 알림 (이미 열려있을 수 있음): {e}")
        
        wb = app.books.add()
        sheet = wb.sheets[0]
        
        print("인포맥스 데이터 로딩을 대기합니다...")
        time.sleep(5)
        
        # 베트남 호치민 지수 티커: VNI:VIDX
        # 6511__세계주요지수(히스토리).xlsx 에서 확인됨
        market = "IDX"
        symbol = "VNI:VIDX"
        fields = "현재가" # 또는 "종가"
        
        # IMDH(Market, Symbol, Fields, StartDate, EndDate, Count, Options)
        # 1년 전 그날 하루의 데이터를 가져오기 위해 Start/End를 동일하게 설정하고 Count를 1로 설정
        formula = f'=IMDH("{market}", "{symbol}", "{fields}", "{start_str}", "{end_str}", 1, "Per=D,Orient=V")'
        
        print(f"입력할 수식: {formula}")
        sheet.range("A1").value = "항목"
        sheet.range("B1").value = "값"
        sheet.range("A2").value = f"베트남 호치민 지수 ({start_str})"
        sheet.range("B2").formula = formula
        
        print("데이터를 기다리는 중 (10초)...")
        time.sleep(10)
        
        # 결과 확인
        result = sheet.range("B2").value
        print(f"조회 결과: {result}")
        
        # 파일 저장
        save_path = os.path.abspath("vietnam_index_price.xlsx")
        if os.path.exists(save_path):
            os.remove(save_path)
        wb.save(save_path)
        print(f"결과가 {save_path}에 저장되었습니다.")
        
    except Exception as e:
        print(f"오류 발생: {e}")
    # finally:
    #     if wb:
    #         wb.close()
    #     if app and xw.apps.count == 1:
    #         app.quit()

if __name__ == "__main__":
    get_vietnam_price()
