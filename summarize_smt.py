import xlwings as xw
import datetime
import os

def summarize_smt_transactions():
    target_file = r"C:\Users\hangs\OneDrive\\04. 츮 \츮   Ȳ..xlsx"
    target_date = datetime.datetime(2025, 12, 31)
    
    print(f"파일을 여는 중: {target_file}")
    
    try:
        # 이미 열려있는 엑셀 파일 연결 시도, 없으면 새로 열기
        if xw.apps.count > 0:
            app = xw.apps.active
            try:
                wb = app.books[os.path.basename(target_file)]
            except:
                wb = app.books.open(target_file)
        else:
            app = xw.App(visible=True)
            wb = app.books.open(target_file)
        
        # '표.거래내역' 테이블 찾기
        tbl_trade = None
        for sheet in wb.sheets:
            for tbl in sheet.api.ListObjects:
                if tbl.Name == "표.거래내역":
                    # xlwings Table 객체로 변환
                    tbl_trade = sheet.tables["표.거래내역"]
                    break
            if tbl_trade: break
            
        if not tbl_trade:
            print("테이블 '표.거래내역'을 찾을 수 없습니다.")
            return

        # 테이블 데이터 가져오기
        data = tbl_trade.range.value
        headers = data[0]
        rows = data[1:]
        
        col_map = {name: i for i, name in enumerate(headers)}
        
        # 필요한 컬럼 인덱스 확인
        try:
            date_idx = col_map["거래일자"]
            type_idx = col_map["자산 분류"]
            qty_idx = col_map["주식수 (매도: 마이너스)"]
            acc_idx = col_map["계좌번호"] # 또는 계좌분류와 유사한 컬럼 확인
        except KeyError as e:
            print(f"필요한 컬럼을 찾을 수 없습니다: {e}")
            print(f"사용 가능한 컬럼: {headers}")
            return

        # 필터링 및 합산 (SMT & 2025.12.31 이전)
        summary = {}
        
        for row in rows:
            trade_date = row[date_idx]
            asset_type = str(row[type_idx]) if row[type_idx] else ""
            qty = row[qty_idx] if row[qty_idx] is not None else 0
            account = str(row[acc_idx]) if row[acc_idx] else "분류없음"
            
            # 날짜 변환 및 비교
            if isinstance(trade_date, str):
                try:
                    trade_date = datetime.datetime.strptime(trade_date, "%Y-%m-%d")
                except:
                    continue
            
            if trade_date and trade_date <= target_date and asset_type == "SMT":
                summary[account] = summary.get(account, 0) + qty

        # 결과 출력
        print("\n" + "="*40)
        print(f" [SMT 거래 합산 결과 (기준일: {target_date.strftime('%Y-%m-%d')})]")
        print("="*40)
        if summary:
            for acc, total in summary.items():
                print(f" 계좌분류: {acc:<15} | 합산수량: {total:>10,}")
        else:
            print(" 해당 조건에 맞는 거래 내역이 없습니다.")
        print("="*40)

    except Exception as e:
        print(f"오류 발생: {e}")

if __name__ == "__main__":
    summarize_smt_transactions()
