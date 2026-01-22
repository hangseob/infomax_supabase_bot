import os
import win32com.client
import re
import time
import sys

def resolve_formula_simple(sheet, formula):
    if not isinstance(formula, str) or not formula.startswith('='):
        return formula
    
    cell_pattern = r"(\$?[A-Z]+\$?\d+)"
    
    def replace_match(match):
        cell_ref = match.group(1)
        try:
            val = sheet.Range(cell_ref).Value
            if val is None: return "None"
            if isinstance(val, str): return f'"{val}"'
            if isinstance(val, float) and val.is_integer(): return str(int(val))
            return str(val)
        except:
            return cell_ref

    return re.sub(cell_pattern, replace_match, formula)

def find_cells_robust():
    root = os.path.abspath('infomax_functions_templetes')
    output_path = os.path.abspath('imdg_imdi_results.txt')
    
    print("\n[진행 상황] 인포맥스 수식 탐색을 시작합니다 (IMDG/IMDI 집중)...")
    
    # 새로운 독립된 엑셀 인스턴스 사용
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        print("-> 전용 엑셀 인스턴스를 실행했습니다.")
    except Exception as e:
        print(f"-> 엑셀 실행 실패: {e}")
        return
        
    excel.Visible = False
    excel.DisplayAlerts = False
    
    results = []
    files_to_check = []
    for r, d, filenames in os.walk(root):
        for f in filenames:
            if f.endswith(('.xlsx', '.xlsm')) and not f.startswith('~$'):
                files_to_check.append(os.path.join(r, f))

    total_files = len(files_to_check)
    print(f"-> 총 {total_files}개의 파일을 검사합니다.\n")

    for idx, file_path in enumerate(files_to_check, 1):
        filename = os.path.basename(file_path)
        print(f"[{idx}/{total_files}] {filename} 검사 중...", end='\r')
        sys.stdout.flush()
        
        wb = None
        # OLE Busy 에러(0x800ac472) 대응 강화
        max_retries = 5
        for attempt in range(max_retries):
            try:
                wb = excel.Workbooks.Open(file_path, ReadOnly=True, UpdateLinks=False)
                break
            except Exception as e:
                if "0x800ac472" in str(e) or "0x8001010A" in str(e):
                    # 엑셀이 바쁠 때 (수식 계산 중 등)
                    time.sleep(2)
                    if attempt == max_retries - 1:
                        print(f"\n[!] {filename}: 엑셀 바쁨(Busy)으로 열기 실패")
                else:
                    print(f"\n[!] {filename}: 열기 오류 - {e}")
                    break
        
        if wb:
            found_in_file = 0
            try:
                for sheet in wb.Sheets:
                    # Find는 대소문자 구분 없이 검색 가능
                    try:
                        # "IMD"가 포함된 첫 번째 셀 찾기
                        cell = sheet.UsedRange.Find("IMD", LookIn=-4123) # xlFormulas
                        if cell:
                            first_addr = cell.Address
                            while True:
                                formula = str(cell.Formula)
                                # IMDG 혹은 IMDI가 포함된 경우만 수집
                                if "IMDG" in formula.upper() or "IMDI" in formula.upper():
                                    resolved = resolve_formula_simple(sheet, formula)
                                    res = f"File: {file_path}\nSheet: {sheet.Name}\nCell: {cell.Address}\nOriginal: {formula}\nResolved: {resolved}\n" + "-"*50 + "\n"
                                    results.append(res)
                                    found_in_file += 1
                                
                                cell = sheet.UsedRange.FindNext(cell)
                                if not cell or cell.Address == first_addr:
                                    break
                    except:
                        continue
                
                if found_in_file > 0:
                    print(f"\n[*] {filename}: {found_in_file}개 발견! (총 {len(results)}개)")
                    sys.stdout.flush()
                
                wb.Close(False)
            except Exception as e:
                print(f"\n[!] {filename} 처리 중 에러: {e}")
                try: wb.Close(False)
                except: pass

    # 결과 저장
    with open(output_path, 'w', encoding='utf-8') as f_out:
        if results:
            f_out.writelines(results)
        else:
            f_out.write("검사 결과 IMDG 또는 IMDI 수식을 사용하는 셀을 찾지 못했습니다.")
    
    print(f"\n\n[탐색 완료]")
    print(f"- 검사한 파일: {total_files}개")
    print(f"- 발견한 수식: {len(results)}개")
    print(f"- 결과 파일: {output_path}")

    try:
        excel.Quit()
    except:
        pass

if __name__ == "__main__":
    find_cells_robust()
