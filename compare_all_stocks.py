import xlwings as xw
import datetime
import sys

# 인코딩 설정
if sys.stdout.encoding != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except:
        pass

def compare_all_stock_balances():
    target_date = datetime.datetime(2025, 12, 31)
    
    try:
        # 1. 엑셀 앱 연결
        if xw.apps.count > 0:
            app = xw.apps.active
            print("기존 Excel 인스턴스에 연결했습니다.")
        else:
            app = xw.App(visible=True)
            print("새로운 Excel 인스턴스를 시작했습니다.")
            
        # 2. 워크북 찾기
        wb = None
        target_name_part = "우리집 가계 금융 현황.종합"
        for b in app.books:
            if target_name_part in b.name:
                wb = b
                break
        
        if not wb:
            print(f"'{target_name_part}' 파일이 열려있지 않습니다.")
            return

        # 3. 테이블 찾기
        tbl_trade = None
        tbl_2512 = None
        
        for sheet in wb.sheets:
            try:
                table_names = [t.name for t in sheet.tables]
                if "표.거래내역" in table_names:
                    tbl_trade = sheet.tables["표.거래내역"]
                if "표.주식.2512" in table_names:
                    tbl_2512 = sheet.tables["표.주식.2512"]
            except: continue
            
        if not tbl_trade or not tbl_2512:
            print(f"테이블 찾기 실패: 거래내역({ 'O' if tbl_trade else 'X' }), 주식.2512({ 'O' if tbl_2512 else 'X' })")
            return

        # 4. 거래내역 데이터 로드 및 합산 (모든 종목)
        trade_data = tbl_trade.range.value
        trade_headers = trade_data[0]
        trade_rows = trade_data[1:]
        trade_col_map = {name: i for i, name in enumerate(trade_headers)}
        
        t_date_idx = trade_col_map["거래일"]
        t_code_idx = trade_col_map["종목코드"]
        t_qty_idx = trade_col_map["주식수 (매도: 마이너스)"]
        
        summary_trade = {} # {(종목코드): 합계수량}
        
        for row in trade_rows:
            t_date = row[t_date_idx]
            t_code = str(row[t_code_idx]).strip().upper() if row[t_code_idx] else ""
            t_qty = row[t_qty_idx] if row[t_qty_idx] is not None else 0
            
            if not t_code: continue
            
            if isinstance(t_date, datetime.datetime) and t_date <= target_date:
                summary_trade[t_code] = summary_trade.get(t_code, 0) + t_qty

        # 5. 표.주식.2512 데이터 로드 및 합산 (모든 종목)
        data_2512 = tbl_2512.range.value
        headers_2512 = data_2512[0]
        rows_2512 = data_2512[1:]
        map_2512 = {name: i for i, name in enumerate(headers_2512)}
        
        s_code_idx = map_2512["종목코드"]
        s_qty_idx = map_2512["주식수"]
        
        summary_2512 = {}
        for row in rows_2512:
            s_code = str(row[s_code_idx]).strip().upper() if row[s_code_idx] else ""
            s_qty = row[s_qty_idx] if row[s_qty_idx] is not None else 0
            
            if not s_code: continue
            
            summary_2512[s_code] = summary_2512.get(s_code, 0) + s_qty

        # 6. 결과 비교 및 출력
        all_stocks = sorted(list(set(summary_trade.keys()) | set(summary_2512.keys())))
        
        print("\n" + "="*85)
        print(f" [전 종목 수량 비교 보고서 (2025-12-31 기준)]")
        print(f" - A: 거래내역 전 기간 합산 (2025.12.31 이전)")
        print(f" - B: 표.주식.2512 (25년말 잔고표)")
        print("="*85)
        print(f" {'종목코드':<15} | {'거래합산(A)':>15} | {'잔고표(B)':>15} | {'차이(B-A)':>15}")
        print("-" * 85)
        
        match_count = 0
        mismatch_count = 0
        
        for code in all_stocks:
            a_qty = summary_trade.get(code, 0)
            b_qty = summary_2512.get(code, 0)
            diff = b_qty - a_qty
            
            # 부동소수점 오차 고려
            if abs(diff) < 0.0001:
                match_count += 1
                # 일치하는 경우는 요약해서 보거나 생략 가능하지만, 일단 모두 출력
                print(f" {code:<15} | {a_qty:>17,.2f} | {b_qty:>17,.2f} | {0:>17}")
            else:
                mismatch_count += 1
                print(f" {code:<15} | {a_qty:>17,.2f} | {b_qty:>17,.2f} | {diff:>17,.2f} ***")

        print("-" * 85)
        print(f" [결과 요약]")
        print(f" - 일치 종목 수: {match_count}개")
        print(f" - 불일치 종목 수: {mismatch_count}개")
        print("="*85)
        if mismatch_count > 0:
            print(" *** 표시가 있는 종목은 수량이 일치하지 않습니다.")

    except Exception as e:
        import traceback
        traceback.print_exc()
        print(f"오류 발생: {e}")

if __name__ == "__main__":
    compare_all_stock_balances()
