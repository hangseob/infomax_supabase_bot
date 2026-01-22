import xlwings as xw
import datetime
import sys

# 인코딩 설정
if sys.stdout.encoding != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except:
        pass

def summarize_smt_final_v3():
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
        
        # 3. '표.거래내역' 테이블 찾기
        tbl_trade = None
        for sheet in wb.sheets:
            try:
                if "표.거래내역" in [t.name for t in sheet.tables]:
                    tbl_trade = sheet.tables["표.거래내역"]
                    break
            except: continue
            
        if not tbl_trade:
            print("테이블 '표.거래내역'을 찾을 수 없습니다.")
            return

        # 4. 데이터 추출
        full_data = tbl_trade.range.value
        headers = full_data[0]
        rows = full_data[1:]
        
        col_map = {name: i for i, name in enumerate(headers)}
        
        # SMT는 '종목코드' 컬럼에 있는 것으로 확인됨
        try:
            date_idx = col_map["거래일"]
            code_idx = col_map["종목코드"] # SMT가 있는 컬럼
            qty_idx = col_map["주식수 (매도: 마이너스)"]
            acc_type_idx = col_map["계좌 분류"] # 합산 기준이 될 분류
        except KeyError as e:
            print(f"컬럼 매핑 오류: {e}")
            return

        # 5. 필터링 및 합산
        summary = {}
        processed_count = 0
        
        for row in rows:
            trade_date = row[date_idx]
            stock_code = str(row[code_idx]) if row[code_idx] else ""
            qty = row[qty_idx] if row[qty_idx] is not None else 0
            # 계좌분류별 합산을 위해 '계좌 분류' 컬럼 사용
            acc_type = str(row[acc_type_idx]) if row[acc_type_idx] else "분류없음"
            
            # 조건: 종목코드가 SMT이고 날짜가 기준일 이전
            if isinstance(trade_date, datetime.datetime):
                if trade_date <= target_date and stock_code.upper() == "SMT":
                    summary[acc_type] = summary.get(acc_type, 0) + qty
                    processed_count += 1

        # 6. 결과 출력
        print("\n" + "="*55)
        print(f" [종목코드 'SMT' 합산 결과 (계좌 분류별)]")
        print(f" - 기준일: 2025-12-31 이전 거래")
        print(f" - 집계된 거래 건수: {processed_count}건")
        print("="*55)
        if summary:
            # 계좌 분류 이름순으로 정렬
            for acc in sorted(summary.keys()):
                print(f" 계좌 분류: {acc:<20} | 합산수량: {summary[acc]:>12,.0f}")
        else:
            print(" 해당 조건(종목코드:SMT & ~2025.12.31)에 맞는 데이터가 없습니다.")
        print("="*55)

    except Exception as e:
        print(f"오류 발생: {e}")

if __name__ == "__main__":
    summarize_smt_final_v3()
