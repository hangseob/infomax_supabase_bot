import xlwings as xw

def close_temp_workbooks():
    print("'통합 문서'로 시작하는 모든 엑셀 통합 문서를 종료합니다...")
    closed_count = 0
    
    try:
        # 모든 실행 중인 엑셀 앱 인스턴스 확인
        for app in xw.apps:
            # 각 앱의 워크북 확인 (뒤에서부터 제거해야 인덱스 오류 방지)
            for wb in list(app.books):
                if wb.name.startswith("통합 문서"):
                    print(f"종료 중: {wb.name}")
                    try:
                        wb.close()
                        closed_count += 1
                    except Exception as e:
                        print(f"  - {wb.name} 종료 실패: {e}")
            
            # 앱에 남은 워크북이 없으면 앱 종료
            if len(app.books) == 0:
                try:
                    app.quit()
                except:
                    pass
                    
        print(f"총 {closed_count}개의 통합 문서를 종료했습니다.")
    except Exception as e:
        print(f"처리 중 오류 발생: {e}")

if __name__ == "__main__":
    close_temp_workbooks()
