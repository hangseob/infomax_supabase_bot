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

def find_cells_with_progress():
    root = os.path.abspath('infomax_functions_templetes')
    output_path = os.path.abspath('imdg_imdi_results.txt')
    
    print("\n[진행 상황] 인포맥스 수식(IMDG/IMDI) 탐색을 시작합니다...")
    
    try:
        excel = win32com.client.GetActiveObject("Excel.Application")
        print("-> 실행 중인 엑셀 인스턴스를 사용합니다.")
    except:
        excel = win32com.client.Dispatch("Excel.Application")
        print("-> 새로운 엑셀 인스턴스를 실행했습니다.")
        
    try:
        excel.Visible = False
        excel.DisplayAlerts = False
    except:
        # 이미 열려있는 엑셀의 경우 설정을 변경하지 못할 수 있습니다.
        pass
    
    results = []
    
    # 1. 대상 파일 리스트 확보
    files_to_check = []
    for r, d, filenames in os.walk(root):
        for f in filenames:
            if f.endswith(('.xlsx', '.xlsm')) and not f.startswith('~$'):
                files_to_check.append(os.path.join(r, f))

    total_files = len(files_to_check)
    print(f"-> 총 {total_files}개의 파일을 검사할 예정입니다.\n")
    sys.stdout.flush()

    # 2. 파일 루프 시작
    for idx, file_path in enumerate(files_to_check, 1):
        filename = os.path.basename(file_path)
        print(f"[{idx}/{total_files}] 파일 검사 중: {filename}...", end='\r')
        sys.stdout.flush()
        
        wb = None
        try:
            # OLE Busy 에러 방지를 위한 재시도
            for attempt in range(3):
                try:
                    wb = excel.Workbooks.Open(file_path, ReadOnly=True)
                    break
                except:
                    time.sleep(1)
            
            if not wb:
                print(f"\n[!] 파일 열기 실패 (건너뜀): {filename}")
                continue

            found_in_file = 0
            for sheet in wb.Sheets:
                try:
                    # IMD 문자열 탐색
                    cell = sheet.UsedRange.Find("IMD", LookIn=-4123) # xlFormulas
                    if cell:
                        first_addr = cell.Address
                        while True:
                            formula = str(cell.Formula)
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
                print(f"\n[*] {filename}: {found_in_file}개의 수식을 찾았습니다! (누적: {len(results)}개)")
            
            wb.Close(False)
            
        except Exception as e:
            print(f"\n[!] 에러 발생 ({filename}): {e}")
            if wb: 
                try: wb.Close(False)
                except: pass

    # 3. 최종 결과 저장
    with open(output_path, 'w', encoding='utf-8') as f_out:
        if results:
            f_out.writelines(results)
        else:
            f_out.write("IMDG 혹은 IMDI를 사용하는 수식을 찾지 못했습니다.")
    
    print(f"\n\n[작업 완료]")
    print(f"- 총 검사 파일: {total_files}개")
    print(f"- 총 발견 수식: {len(results)}개")
    print(f"- 결과 저장 경로: {output_path}")

if __name__ == "__main__":
    find_cells_with_progress()
