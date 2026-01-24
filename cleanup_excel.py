import xlwings as xw
import os

def cleanup():
    print("기존에 생성된 인포맥스 관련 엑셀 파일 및 앱 인스턴스 정리를 시작합니다...")
    
    # 1. 엑셀 앱 인스턴스 정리
    try:
        for app in xw.apps:
            for wb in app.books:
                # 내가 만든 파일 이름 패턴이거나 저장되지 않은 새 문서인 경우 닫기
                if "infomax" in wb.name.lower() or wb.name.startswith("Book"):
                    print(f"워크북 닫기: {wb.name}")
                    try:
                        wb.close()
                    except:
                        pass
            
            # 앱에 열린 워크북이 없으면 앱 종료
            if len(app.books) == 0:
                print("빈 엑셀 인스턴스 종료")
                try:
                    app.quit()
                except:
                    pass
    except Exception as e:
        print(f"엑셀 앱 정리 중 오류: {e}")

    # 2. 파일 삭제
    files_to_delete = [
        f for f in os.listdir('.') 
        if f.startswith('infomax_') and f.endswith('.xlsx')
    ]
    
    for f in files_to_delete:
        try:
            os.remove(f)
            print(f"파일 삭제 완료: {f}")
        except Exception as e:
            print(f"파일 삭제 실패 ({f}): {e}")

if __name__ == "__main__":
    cleanup()
