import xlwings as xw
import os
import time

def test_excel_ops():
    test_file = os.path.abspath("test_ops.xlsx")
    if os.path.exists(test_file):
        os.remove(test_file)
        
    print(f"1. 엑셀 앱 및 워크북 생성 테스트 시작...")
    try:
        app = xw.App(visible=True, add_book=False)
        wb = app.books.add()
        sheet = wb.sheets[0]
        
        print("2. 데이터 쓰기 테스트...")
        test_data = [["ID", "Name", "Value"], [1, "Test1", 100], [2, "Test2", 200]]
        sheet.range("A1").value = test_data
        
        print(f"3. 파일 저장 테스트: {test_file}")
        wb.save(test_file)
        
        print("4. 닫기 테스트...")
        wb.close()
        app.quit()
        
        print("\n--- 1단계(쓰기 및 저장) 완료 ---\n")
        time.sleep(2)
        
        print("5. 다시 열기 및 읽기 테스트 시작...")
        app2 = xw.App(visible=True, add_book=False)
        wb2 = app2.books.open(test_file)
        sheet2 = wb2.sheets[0]
        
        read_data = sheet2.range("A1:C3").value
        print(f"읽어온 데이터: {read_data}")
        
        success = (read_data[1][1] == "Test1")
        if success:
            print("\n[결과] 모든 엑셀 기본 동작(생성/쓰기/저장/닫기/열기/읽기)이 정상입니다!")
        else:
            print("\n[결과] 데이터 일치 확인 실패.")
            
        wb2.close()
        app2.quit()
        
        # 파일 삭제
        if os.path.exists(test_file):
            os.remove(test_file)
            
    except Exception as e:
        print(f"\n[에러 발생] {e}")
        try:
            app.quit()
        except:
            pass

if __name__ == "__main__":
    test_excel_ops()
