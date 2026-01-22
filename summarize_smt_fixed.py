import xlwings as xw
import datetime
import os
import sys

# 인코딩 문제 해결을 위해 출력 설정
sys.stdout.reconfigure(encoding='utf-8')

def summarize_smt_transactions():
    # 실제 확인된 경로 사용
    target_file = r"C:\Users\hangs\OneDrive\\04. 츮 \츮   Ȳ..xlsx"
    target_date = datetime.datetime(2025, 12, 31)
    
    print(f"파일 검색 중: {target_file}")
    
    try:
        # 이미 열려있는 엑셀 앱 연결
        if xw.apps.count > 0:
            app = xw.apps.active
            print("기존 엑셀 인스턴스에 연결했습니다.")
        else:
            app = xw.App(visible=True)
            print("새로운 엑셀 인스턴스를 시작했습니다.")
            
        # 워크북 찾기 또는 열기
        wb = None
        for b in app.books:
            # 파일명만 비교 (경로에 인코딩 문제가 있을 수 있으므로)
            if "우리집 가계 금융 현황.종합" in b.Name:
                wb = b
                print(f"이미 열려있는 워크북을 찾았습니다: {wb.Name}")
                break
        
        if not wb:
            if os.path.exists(target_file):
                wb = app.books.open(target_file)
                print(f"워크북을 열었습니다: {target_file}")
            else:
                print("파일을 찾을 수 없습니다. 경로를 확인해주세요.")
                return
        
        # '표.거래내역' 테이블 데이터 직접 읽기
        # COM 객체를 통해 컬럼명과 데이터를 안전하게 가져옵니다.
        tbl_trade = None
        for sheet in wb.sheets:
            for tbl in sheet.api.ListObjects:
                if tbl.Name == "표.거래내역":
                    tbl_trade = tbl
                    break
            if tbl_trade: break
            
        if not tbl_trade:
            print("테이블 '표.거래내역'을 찾을 수 없습니다.")
            return

        # 헤더와 데이터 본문 읽기
        headers = [col.Name for col in tbl_trade.ListColumns]
        data_range = tbl_trade.DataBodyRange
        if not data_range:
            print("테이블에 데이터가 없습니다.")
            return
            
        values = data_range.Value
        
        col_map = {name: i for i, name in enumerate(headers)}
        
        # 필요한 컬럼 확인
        try:
            date_idx = col_map["거래일자"]
            type_idx = col_map["자산 분류"]
            qty_idx = col_map["주식수 (매도: 마이너스)"]
            acc_idx = col_map["계좌번호"]
        except KeyError as e:
            print(f"필요한 컬럼을 찾을 수 없습니다: {e}")
            print(f"사용 가능한 컬럼: {headers}")
            return

        # 필터링 및 합산
        summary = {}
        count = 0
        
        for row in values:
            trade_date = row[date_idx]
            asset_type = str(row[type_idx]) if row[type_idx] else ""
            qty = row[qty_idx] if row[qty_idx] is not None else 0
            account = str(row[acc_idx]) if row[acc_idx] else "분류없음"
            
            # 날짜 비교 (엑셀에서 넘어온 datetime 객체 처리)
            if trade_date and isinstance(trade_date, datetime.datetime):
                if trade_date <= target_date and asset_type == "SMT":
                    summary[account] = summary.get(account, 0) + qty
                    count += 1

        # 결과 출력
        print("\n" + "="*50)
        print(f" [SMT 자산 합산 결과]")
        print(f" - 기준: 2025-12-31 이전 거래")
        print(f" - 처리된 행 수: {count}개")
        print("="*50)
        if summary:
            # 정렬하여 출력
            for acc in sorted(summary.keys()):
                total = summary[acc]
                print(f" 계좌번호: {acc:<15} | 합산수량: {total:>12,}")
        else:
            print(" 해당 조건(SMT & ~2025.12.31)에 맞는 데이터가 없습니다.")
        print("="*50)

    except Exception as e:
        import traceback
        print(f"오류 상세: {e}")
        traceback.print_exc()

if __name__ == "__main__":
    summarize_smt_transactions()
