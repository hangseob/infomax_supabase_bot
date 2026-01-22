import os
import xlwings as xw
import re

def resolve_formula(sheet, formula):
    """
    범용적인 정규표현식을 사용하여 수식 내의 셀 참조(예: $A$1, B2, Sheet1!A1)를 
    해당 셀의 실제 값으로 치환하려고 시도합니다.
    """
    if not isinstance(formula, str) or not formula.startswith('='):
        return formula

    # 셀 참조 패턴 (Sheet1!$A$1, $A1, A$1, A1 등)
    # 1. 시트명 포함 패턴: 'Sheet Name'!A1 or SheetName!A1
    # 2. 일반 셀 패턴: $A$1, A1
    cell_pattern = r"((?:'[^']+'!|[a-zA-Z0-9가-힣]+!)?\$?[A-Z]+\$?\d+)"
    
    def replace_match(match):
        cell_ref = match.group(1)
        try:
            # xlwings를 통해 해당 참조의 값을 가져옴
            val = sheet.range(cell_ref).value
            if val is None:
                return "None"
            if isinstance(val, str):
                return f'"{val}"'
            return str(val)
        except:
            return cell_ref # 변환 실패 시 원래 참조 유지

    resolved = re.sub(cell_pattern, replace_match, formula)
    return resolved

def find_imdg_imdi_cells():
    root = 'infomax_functions_templetes'
    output_file = 'imdg_imdi_results.txt'
    
    # 엑셀 앱 실행 (기존 앱 연결 시도)
    try:
        app = xw.apps.active
    except:
        app = xw.App(visible=False)
    
    results = []
    
    try:
        for r, d, files in os.walk(root):
            for f in files:
                if f.endswith(('.xlsx', '.xlsm')) and not f.startswith('~$'):
                    file_path = os.path.join(r, f)
                    full_path = os.path.abspath(file_path)
                    
                    print(f"Checking: {file_path}")
                    # OLE busy error 대응을 위한 재시도 로직
                    wb = None
                    for retry in range(3):
                        try:
                            wb = app.books.open(full_path, read_only=True, update_links=False)
                            break
                        except:
                            import time
                            time.sleep(1)
                    
                    if not wb:
                        print(f"  Failed to open {file_path} after retries.")
                        continue

                    try:
                        for sheet in wb.sheets:
                            try:
                                used_range = sheet.used_range
                                formulas = used_range.formula
                                
                                if not isinstance(formulas, list):
                                    formulas = [[formulas]]
                                
                                for row_idx, row in enumerate(formulas):
                                    for col_idx, formula in enumerate(row):
                                        if isinstance(formula, str) and any(x in formula.upper() for x in ['IMDG', 'IMDI', 'IMDB', 'IMDH', 'IMDP']):
                                            cell = used_range[row_idx, col_idx]
                                            cell_addr = cell.address
                                            
                                            resolved_formula = resolve_formula(sheet, formula)
                                            
                                            res_str = f"File: {file_path}\n"
                                            res_str += f"Sheet: {sheet.name}\n"
                                            res_str += f"Cell: {cell_addr}\n"
                                            res_str += f"Original: {formula}\n"
                                            res_str += f"Resolved: {resolved_formula}\n"
                                            res_str += "-" * 50 + "\n"
                                            results.append(res_str)
                                            print(f"  Found {formula[:20]}... in {sheet.name} {cell_addr}")
                            except:
                                continue
                        wb.close()
                    except Exception as e:
                        print(f"  Error processing {file_path}: {e}")
                        if wb: wb.close()
                        continue
                        
        # 파일 저장
        with open(output_file, 'w', encoding='utf-8') as f:
            if results:
                f.writelines(results)
            else:
                f.write("IMDG 혹은 IMDI를 사용하는 셀을 찾지 못했습니다.")
                
        print(f"\n작업 완료! 결과가 {output_file}에 저장되었습니다.")

    finally:
        app.quit()

if __name__ == "__main__":
    find_imdg_imdi_cells()
