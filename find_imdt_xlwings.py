import os
import xlwings as xw
import sys

# 인코딩 설정
if sys.stdout.encoding != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except:
        pass

def search_imdt_with_xlwings(root_dir):
    results = []
    
    # 엑셀 앱 실행 (백그라운드)
    app = xw.App(visible=False, add_book=False)
    
    try:
        for root, dirs, files in os.walk(root_dir):
            for file in files:
                if file.startswith('~$'): continue
                if not file.lower().endswith(('.xlsx', '.xlsm', '.xlsb')): continue
                
                file_path = os.path.abspath(os.path.join(root, file))
                print(f"Checking: {file}")
                
                try:
                    # ReadOnly로 열기
                    wb = app.books.open(file_path, read_only=True, update_links=False)
                    
                    for sheet in wb.sheets:
                        # UsedRange에서 IMDT 검색
                        try:
                            # find()는 수식 내 텍스트도 검색 가능
                            # 하지만 xlwings의 find는 조금 까다로울 수 있으니 
                            # UsedRange의 formulas를 직접 확인하는 방식 사용
                            formulas = sheet.used_range.formula
                            
                            # formulas가 단일 값일 수도 있고 2차원 튜플일 수도 있음
                            if isinstance(formulas, str):
                                if 'IMDT' in formulas.upper():
                                    results.append({'file': file_path, 'sheet': sheet.name})
                            elif isinstance(formulas, (list, tuple)):
                                found_in_sheet = False
                                for row in formulas:
                                    for cell_formula in row:
                                        if cell_formula and isinstance(cell_formula, str) and 'IMDT' in cell_formula.upper():
                                            results.append({'file': file_path, 'sheet': sheet.name})
                                            found_in_sheet = True
                                            break
                                    if found_in_sheet: break
                        except Exception as e:
                            # print(f"  Error reading sheet {sheet.name}: {e}")
                            pass
                    
                    wb.close()
                except Exception as e:
                    print(f"  Error opening {file}: {e}")
                    
    finally:
        app.quit()
        
    return results

if __name__ == "__main__":
    target_dir = r"infomax_functions_templetes"
    print(f"Searching for 'IMDT' in {target_dir} using xlwings...")
    
    matches = search_imdt_with_xlwings(target_dir)
    
    if matches:
        print(f"\nFound IMDT functions in {len(matches)} sheets:")
        current_file = ""
        for match in matches:
            if match['file'] != current_file:
                print(f"\n[File: {match['file']}]")
                current_file = match['file']
            print(f"  - Sheet: {match['sheet']}")
    else:
        print("\nNo IMDT functions found.")
