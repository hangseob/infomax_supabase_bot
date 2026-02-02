import xlwings as xw
import os
import time
import sys

# 인코딩 설정
if sys.stdout.encoding != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except:
        pass

def inspect_and_get_kospi200():
    infomax_xlam_path = r"C:\Infomax\bin\excel\infomaxexcel.xlam"
    
    try:
        if xw.apps.count > 0:
            app = xw.apps.active
            print("기존 Excel 인스턴스에 연결합니다.")
        else:
            app = xw.App(visible=True, add_book=False)
            print("새로운 Excel 인스턴스를 실행합니다.")

        # 인포맥스 애드인 로드 확인 및 실행
        is_addin_loaded = False
        for b in app.books:
            if "infomaxexcel" in b.name.lower():
                is_addin_loaded = True
                break
        
        if not is_addin_loaded:
            print(f"인포맥스 애드인을 로드합니다: {infomax_xlam_path}")
            app.books.open(infomax_xlam_path)
            time.sleep(5)

        # 현재 열려 있는 모든 통합 문서에서 KOSPI 200 데이터가 있는지 확인
        print("\n[현재 열린 문서에서 KOSPI 200 데이터 탐색]")
        found_data = None
        for wb in app.books:
            print(f"체크 중: {wb.name}")
            for sheet in wb.sheets:
                # UsedRange 내에서 '삼성전자'가 있는지 확인 (KOSPI 200의 대표 종목)
                try:
                    # 간단하게 A1:Z500 범위 조사
                    vals = sheet.range("A1:Z500").value
                    if vals:
                        for r_idx, row in enumerate(vals):
                            for c_idx, cell in enumerate(row):
                                if cell and "삼성전자" in str(cell):
                                    print(f"  -> '{sheet.name}' 시트 {r_idx+1}행 {c_idx+1}열에서 '삼성전자' 발견!")
                                    # 데이터가 KOSPI 200 리스트인지 확인하기 위해 근처 10행 출력
                                    sample = sheet.range((r_idx + 1, c_idx + 1), (r_idx + 11, c_idx + 2)).value
                                    print(f"  데이터 샘플: {sample[:5]}")
                                    if len(sample) > 5:
                                        found_data = sample
                                        break
                            if found_data: break
                except: continue
                if found_data: break
            if found_data: break

        if found_data:
            # 찾은 데이터를 새 파일로 저장
            new_wb = app.books.add()
            new_wb.sheets[0].range("A1").value = found_data
            save_path = os.path.abspath("kospi200_items_found.xlsx")
            new_wb.save(save_path)
            print(f"\n데이터를 찾아 저장했습니다: {save_path}")
        else:
            print("\n열려 있는 어떤 시트에서도 KOSPI 200 리스트로 보이는 데이터를 찾지 못했습니다.")
            print("인포맥스 로그인이 되어 있는지, 그리고 엑셀에서 '실시간 데이터 수신'이 활성 상태인지 확인이 필요합니다.")

    except Exception as e:
        print(f"오류 발생: {e}")

if __name__ == "__main__":
    inspect_and_get_kospi200()
