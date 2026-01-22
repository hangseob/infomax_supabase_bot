import xlwings as xw
import time

def verify_excel_data():
    print("엑셀 데이터 확인을 시작합니다...")
    try:
        # 1. 활성화된 엑셀 앱 연결
        if xw.apps.count == 0:
            print("실행 중인 엑셀이 없습니다.")
            return

        # 2. 활성화된 통합문서 및 시트 연결
        wb = xw.books.active
        sheet = wb.sheets.active
        print(f"연결된 파일: {wb.name}, 시트: {sheet.name}")

        # 3. 데이터 로딩 대기 (인포맥스 데이터 수신 시간 고려)
        print("데이터 수신 여부를 확인 중입니다 (10초 대기)...")
        time.sleep(10)

        # 4. 데이터 확인 (A1부터 데이터가 들어오므로 A1:B30 영역 확인)
        # IMDH 결과는 보통 헤더를 포함하므로 20개 이상의 데이터가 있는지 체크
        data = sheet.range("A1:B35").value
        
        # None이 아닌 데이터 필터링
        valid_rows = [row for row in data if row[0] is not None and row[1] is not None]
        count = len(valid_rows)

        print(f"확인된 유효 데이터 건수: {count}건")
        
        if count >= 20:
            print("OK: 20 rows or more found.")
            for i, row in enumerate(valid_rows[:5]):
                print(f"  Sample {i+1}: {row}")
        elif count > 0:
            print(f"Warning: Only {count} rows found.")
        else:
            print("Error: No data found in Excel. Check Infomax login and formula.")

    except Exception as e:
        print(f"오류 발생: {e}")

if __name__ == "__main__":
    verify_excel_data()
