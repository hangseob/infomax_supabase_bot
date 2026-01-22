import xlwings as xw
import time
import sys

def fetch_and_print_excel_data():
    print("--- 인포맥스 데이터 터미널 출력 시작 ---")
    
    # 한국어 출력을 위한 인코딩 설정 (윈도우 터미널 대응)
    if sys.platform == 'win32':
        import io
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

    try:
        if xw.apps.count == 0:
            print("실행 중인 엑셀이 없습니다.")
            return

        # 모든 앱과 워크북 탐색
        found = False
        for app in xw.apps:
            for wb in app.books:
                for sheet in wb.sheets:
                    # 데이터가 있을 법한 범위 (A~C열) 읽기
                    # 인포맥스 데이터 수신 시간을 고려하여 수식 결과가 나올 때까지 최대 5회 재시도
                    for retry in range(5):
                        try:
                            data = sheet.range("A1:C35").value
                            # 데이터가 있는지 확인 (A1에 '조회'가 있거나 C열에 숫자가 있는 경우)
                            rows_with_values = []
                            if data:
                                for i, row in enumerate(data):
                                    # 하나라도 None이 아닌 셀이 있는 행 추출
                                    if any(cell is not None for cell in row):
                                        rows_with_values.append((i + 1, row))
                            
                            if len(rows_with_values) > 1: # 헤더 외에 데이터가 더 있는 경우
                                print(f"\n[워크북: {wb.name}] [시트: {sheet.name}]")
                                print("-" * 40)
                                print(f"{'행':<3} | {'A열':<15} | {'B열':<10} | {'C열(가격)':<10}")
                                print("-" * 40)
                                for idx, row in rows_with_values:
                                    a = str(row[0]) if row[0] is not None else ""
                                    b = str(row[1]) if row[1] is not None else ""
                                    c = str(int(row[2])) if (len(row) > 2 and row[2] is not None) else ""
                                    print(f"{idx:<3} | {a:<15} | {b:<10} | {c:<10}")
                                print("-" * 40)
                                found = True
                                break # 데이터 찾았으므로 재시도 중단
                        except Exception as e:
                            if "0x800ac472" in str(e): # OLE busy error
                                time.sleep(1)
                                continue
                            else:
                                raise e
                    if found: break
                if found: break
        
        if not found:
            print("데이터를 찾을 수 없습니다. 엑셀 수식이 로드되었는지 확인해 주세요.")

    except Exception as e:
        print(f"오류 발생: {e}")

if __name__ == "__main__":
    fetch_and_print_excel_data()
