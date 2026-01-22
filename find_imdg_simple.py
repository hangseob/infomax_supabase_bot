import os
import win32com.client
import re
import time

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

def find_cells_simple():
    root = os.path.abspath('infomax_functions_templetes')
    output_path = os.path.abspath('imdg_imdi_results.txt')
    
    # 엑셀이 열려 있다면 닫아달라고 요청하거나, 아예 새로 띄움
    print("Finding IMDG/IMDI formulas... (Please ensure Excel is not busy)")
    
    try:
        excel = win32com.client.GetActiveObject("Excel.Application")
        print("Using existing Excel instance.")
    except:
        excel = win32com.client.Dispatch("Excel.Application")
        print("Started new Excel instance.")
        
    excel.Visible = False
    excel.DisplayAlerts = False
    
    results = []
    
    # 파일 리스트 먼저 확보 (에러 방지)
    files_to_check = []
    for r, d, filenames in os.walk(root):
        for f in filenames:
            if f.endswith(('.xlsx', '.xlsm')) and not f.startswith('~$'):
                files_to_check.append(os.path.join(r, f))

    for file_path in files_to_check:
        print(f"Checking: {os.path.basename(file_path)}")
        wb = None
        try:
            # 3회 재시도
            for _ in range(3):
                try:
                    wb = excel.Workbooks.Open(file_path, ReadOnly=True)
                    break
                except:
                    time.sleep(2)
            
            if not wb: continue

            for sheet in wb.Sheets:
                try:
                    # Find를 사용해 "IMD"가 포함된 셀 검색
                    cell = sheet.UsedRange.Find("IMD", LookIn=-4123) # xlFormulas
                    if cell:
                        first_addr = cell.Address
                        while True:
                            formula = str(cell.Formula)
                            if "IMDG" in formula.upper() or "IMDI" in formula.upper():
                                resolved = resolve_formula_simple(sheet, formula)
                                res = f"File: {file_path}\nSheet: {sheet.Name}\nCell: {cell.Address}\nOriginal: {formula}\nResolved: {resolved}\n" + "-"*50 + "\n"
                                results.append(res)
                                print(f"  Found in {sheet.Name} {cell.Address}")
                            
                            cell = sheet.UsedRange.FindNext(cell)
                            if not cell or cell.Address == first_addr:
                                break
                except:
                    continue
            wb.Close(False)
        except Exception as e:
            print(f"  Skipped {os.path.basename(file_path)} due to error: {e}")
            if wb: 
                try: wb.Close(False)
                except: pass

    with open(output_path, 'w', encoding='utf-8') as f_out:
        if results:
            f_out.writelines(results)
        else:
            f_out.write("No IMDG/IMDI formulas found.")
    
    print(f"\nDone. Found {len(results)} matches. Results: {output_path}")

if __name__ == "__main__":
    find_cells_simple()
