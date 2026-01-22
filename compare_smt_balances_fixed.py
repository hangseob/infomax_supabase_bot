import xlwings as xw
import datetime
import sys

# 인코딩 설정
if sys.stdout.encoding != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except:
        pass

def compare_smt_balances():
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

        # 4. 거래내역 합산 (SMT)
        trade_data = tbl_trade.range.value
        trade_headers = trade_data[0]
        trade_rows = trade_data[1:]
        
        trade_col_map = {name: i for i, name in enumerate(trade_headers)}
        # 거래내역의 계좌 분류명 (예: '윤기.미래')
        t_acc_idx = trade_col_map["계좌 분류"]
        t_code_idx = trade_col_map["종목코드"]
        t_qty_idx = trade_col_map["주식수 (매도: 마이너스)"]
        t_date_idx = trade_col_map["거래일"]
        
        summary_trade = {}
        for row in trade_rows:
            t_date = row[t_date_idx]
            t_code = str(row[t_code_idx]) if row[t_code_idx] else ""
            t_qty = row[t_qty_idx] if row[t_qty_idx] is not None else 0
            t_acc = str(row[t_acc_idx]) if row[t_acc_idx] else "분류없음"
            
            if isinstance(t_date, datetime.datetime) and t_date <= target_date:
                if t_code.upper() == "SMT":
                    summary_trade[t_acc] = summary_trade.get(t_acc, 0) + t_qty

        # 5. 표.주식.2512 데이터 읽기 (SMT)
        data_2512 = tbl_2512.range.value
        headers_2512 = data_2512[0]
        rows_2512 = data_2512[1:]
        
        map_2512 = {name: i for i, name in enumerate(headers_2512)}
        # 확인된 컬럼명: '계좌' (거래내역의 '계좌 분류'와 매칭되는지 확인 필요)
        idx_acc = map_2512["계좌"]
        idx_code = map_2512["종목코드"]
        idx_qty = map_2512["주식수"]

        summary_2512 = {}
        for row in rows_2512:
            s_acc = str(row[idx_acc]) if row[idx_acc] else "분류없음"
            s_code = str(row[idx_code]) if row[idx_code] else ""
            s_qty = row[idx_qty] if row[idx_qty] is not None else 0
            
            if s_code.upper() == "SMT":
                summary_2512[s_acc] = summary_2512.get(s_acc, 0) + s_qty

        # 6. 결과 비교 및 출력
        all_accounts = sorted(list(set(summary_trade.keys()) | set(summary_2512.keys())))
        
        print("\n" + "="*85)
        print(f" [SMT 수량 비교 보고서 (2025-12-31 기준)]")
        print(f" - A: 거래내역 합산 (2025.12.31 이전)")
        print(f" - B: 표.주식.2512 (25년말 잔고표)")
        print("="*85)
        print(f" {'계좌 분류':<22} | {'거래합산(A)':>14} | {'잔고표(B)':>14} | {'차이(B-A)':>14}")
        print("-" * 85)
        
        total_a = 0
        total_b = 0
        
        for acc in all_accounts:
            a_qty = summary_trade.get(acc, 0)
            b_qty = summary_2512.get(acc, 0)
            diff = b_qty - a_qty
            
            total_a += a_qty
            total_b += b_qty
            
            # 차이가 있는 경우 강조 (가상)
            print(f" {acc:<22} | {a_qty:>14,.0f} | {b_qty:>14,.0f} | {diff:>14,.0f}")
            
        print("-" * 85)
        print(f" {'합계':<22} | {total_a:>14,.0f} | {total_b:>14,.0f} | {total_b - total_a:>14,.0f}")
        print("="*85)

    except Exception as e:
        print(f"오류 발생: {e}")

if __name__ == "__main__":
    compare_smt_balances()
