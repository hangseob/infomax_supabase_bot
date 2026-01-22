import xlwings as xw
import datetime
import sys

# 인코딩 설정
if sys.stdout.encoding != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except:
        pass

def search_smt_in_columns():
    try:
        app = xw.apps.active
        wb = None
        for b in app.books:
            if "우리집 가계 금융 현황.종합" in b.name:
                wb = b
                break
        
        if not wb:
            print("파일이 열려있지 않습니다.")
            return
            
        tbl_trade = None
        for sheet in wb.sheets:
            try:
                if "표.거래내역" in [t.name for t in sheet.tables]:
                    tbl_trade = sheet.tables["표.거래내역"]
                    break
            except: continue
            
        if tbl_trade:
            full_data = tbl_trade.range.value
            headers = full_data[0]
            rows = full_data[1:]
            
            print(f"컬럼 목록: {headers}")
            
            # 모든 컬럼에서 'SMT'라는 값이 있는지 전수 조사
            found_locations = []
            for r_idx, row in enumerate(rows):
                for c_idx, cell in enumerate(row):
                    if cell and "SMT" in str(cell).upper():
                        found_locations.append(f"행:{r_idx+2}, 열:'{headers[c_idx]}', 값:{cell}")
                        if len(found_locations) >= 5: break # 일부만 출력
                if len(found_locations) >= 5: break
            
            if found_locations:
                print("\n[SMT 발견 위치]")
                for loc in found_locations:
                    print(loc)
            else:
                print("\n테이블 내 어느 컬럼에서도 'SMT'를 찾을 수 없습니다.")

    except Exception as e:
        print(f"오류: {e}")

if __name__ == "__main__":
    search_smt_in_columns()
