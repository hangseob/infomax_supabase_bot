import xlwings as xw
import datetime
import sys

# 인코딩 설정
if sys.stdout.encoding != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except:
        pass

def debug_smt_values():
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
            
            col_map = {name: i for i, name in enumerate(headers)}
            type_idx = col_map["계좌 분류"]
            
            # 고유한 계좌 분류 값들 확인
            unique_types = set()
            for row in rows:
                if row[type_idx]:
                    unique_types.add(str(row[type_idx]))
            
            print(f"발견된 계좌 분류 목록: {list(unique_types)}")
            
            # SMT가 포함된 행이 있는지 확인 (대소문자 무시)
            smt_count = 0
            for row in rows:
                val = str(row[type_idx]) if row[type_idx] else ""
                if "SMT" in val.upper():
                    smt_count += 1
            print(f"'SMT' 문구를 포함하는 행 수: {smt_count}")

    except Exception as e:
        print(f"오류: {e}")

if __name__ == "__main__":
    debug_smt_values()
