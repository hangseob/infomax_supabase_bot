import xlwings as xw
import datetime
import os
import sys

# 인코딩 문제 해결
if sys.stdout.encoding != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except:
        pass

def summarize_smt_transactions():
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
            # xlwings Book 객체는 .name 속성을 사용 (소문자)
            if target_name_part in b.name:
                wb = b
                print(f"워크북을 찾았습니다: {wb.name}")
                break
        
        if not wb:
            print(f"'{target_name_part}' 파일이 엑셀에 열려있지 않습니다.")
            return
        
        # 3. '표.거래내역' 테이블 데이터 읽기
        tbl_trade = None
        for sheet in wb.sheets:
            try:
                # xlwings의 tables 컬렉션 사용
                if "표.거래내역" in [t.name for t in sheet.tables]:
                    tbl_trade = sheet.tables["표.거래내역"]
                    print(f"'{sheet.name}' 시트에서 테이블을 찾았습니다.")
                    break
            except:
                continue
            
        if not tbl_trade:
            print("테이블 '표.거래내역'을 찾을 수 없습니다.")
            return

        # 4. 데이터 및 헤더 추출
        # range.value는 전체 테이블(헤더 포함)을 가져옵니다.
        full_data = tbl_trade.range.value
        headers = full_data[0]
        rows = full_data[1:]
        
        col_map = {name: i for i, name in enumerate(headers)}
        
        # 필요한 컬럼 인덱스 확인
        try:
            date_idx = col_map["거래일자"]
            type_idx = col_map["자산 분류"]
            qty_idx = col_map["주식수 (매도: 마이너스)"]
            acc_idx = col_map["계좌번호"]
        except KeyError as e:
            print(f"필요한 컬럼이 없습니다: {e}")
            print(f"현재 컬럼 목록: {headers}")
            return

        # 5. 필터링 및 합산
        summary = {}
        processed_count = 0
        
        for row in rows:
            trade_date = row[date_idx]
            asset_type = str(row[type_idx]) if row[type_idx] else ""
            qty = row[qty_idx] if row[qty_idx] is not None else 0
            account = str(row[acc_idx]) if row[acc_idx] else "분류없음"
            
            # 날짜 비교
            if isinstance(trade_date, datetime.datetime):
                if trade_date <= target_date and asset_type == "SMT":
                    summary[account] = summary.get(account, 0) + qty
                    processed_count += 1

        # 6. 결과 출력
        print("\n" + "="*50)
        print(f" [SMT 자산 합산 결과]")
        print(f" - 기준일: 2025-12-31 이전")
        print(f" - 해당 거래 건수: {processed_count}건")
        print("="*50)
        if summary:
            for acc in sorted(summary.keys()):
                print(f" 계좌번호: {acc:<15} | 합산수량: {summary[acc]:>12,.0f}")
        else:
            print(" 조건에 맞는 데이터가 없습니다.")
        print("="*50)

    except Exception as e:
        print(f"오류 발생: {e}")

if __name__ == "__main__":
    summarize_smt_transactions()
