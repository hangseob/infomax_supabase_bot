import os
import xlwings as xw
from datetime import datetime, timedelta

def get_samsung_stock_history():
    # 1. 설정
    stock_code = "005930"  # 삼성전자
    end_date = datetime(2026, 1, 12)  # 오늘 날짜 (시스템 정보 기준)
    start_date = end_date - timedelta(days=31)
    
    start_str = start_date.strftime("%Y%m%d")
    end_str = end_date.strftime("%Y%m%d")
    
    print(f"조회 기간: {start_str} ~ {end_str}")
    
    # 2. 기존 엑셀 인스턴스에 연결
    try:
        if xw.apps.count > 0:
            # 이미 열려 있는 엑셀이 있다면 활성화된 통합문서(Workbook) 사용
            wb = xw.books.active
            sheet = wb.sheets.active
            print(f"활성화된 엑셀 파일 '{wb.name}'의 '{sheet.name}' 시트에 연결했습니다.")
        else:
            # 열려 있는 엑셀이 없다면 새로 생성
            app = xw.App(visible=True)
            wb = app.books.add()
            sheet = wb.sheets[0]
            print("새로운 엑셀 인스턴스를 생성했습니다.")
        
        # 3. 인포맥스 함수 입력
        # IMDH(마켓, 종목코드, 필드, 시작일, 종료일, 건수, 옵션)
        # 필드: 날짜, 현재가
        fields = "날짜,현재가"
        options = "Per=D,Orient=V,sort=1"
        
        # A1 셀에 헤더 작성 (IMDH가 자동으로 헤더를 가져올 수도 있지만 명시적으로 필드를 지정함)
        # 템플릿 참조: =IMDH("STK", "005930", "날짜,현재가", "20251212", "20260112", 30, "Per=D,Orient=V")
        
        formula = f'=IMDH("STK", "{stock_code}", "{fields}", "{start_str}", "{end_str}", 35, "{options}")'
        
        print(f"입력할 수식: {formula}")
        sheet.range("A1").formula = formula
        
        # 4. 결과 확인을 위해 잠시 대기 및 저장
        print("데이터 로딩 중... (인포맥스 터미널이 실행 중이어야 합니다)")
        
        import time
        time.sleep(5) # 데이터 수신 대기
        
        # 파일 저장 시 이미 열려있을 경우를 대비해 예외 처리 및 다른 이름으로 저장 시도
        base_name = "samsung_electronics_price"
        extension = ".xlsx"
        save_path = os.path.abspath(f"{base_name}{extension}")
        
        counter = 1
        while os.path.exists(save_path):
            try:
                # 기존 파일이 있으면 삭제 시도 (열려있으면 에러 발생)
                if os.path.exists(save_path):
                    os.remove(save_path)
                break
            except OSError:
                # 파일이 열려있어서 삭제할 수 없는 경우 이름을 변경
                save_path = os.path.abspath(f"{base_name}_{counter}{extension}")
                counter += 1
        
        wb.save(save_path)
        print(f"파일이 저장되었습니다: {save_path}")
        
    except Exception as e:
        print(f"오류 발생: {e}")
    finally:
        # 에러가 나더라도 앱이 백그라운드에 남지 않도록 함 (필요시 주석 해제)
        # wb.close()
        # app.quit()
        pass

if __name__ == "__main__":
    get_samsung_stock_history()
