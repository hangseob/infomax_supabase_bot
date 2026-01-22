import os
import win32com.client
import re

def resolve_formula_simple(sheet, formula):
    if not isinstance(formula, str) or not formula.startswith('='):
        return formula
    
    # Very simple cell reference resolver for COM
    cell_pattern = r"\$?[A-Z]+\$?\d+"
    
    def replace_match(match):
        cell_ref = match.group(0)
        try:
            val = sheet.Range(cell_ref).Value
            if val is None: return "None"
            if isinstance(val, str): return f'"{val}"'
            return str(val)
        except:
            return cell_ref

    return re.sub(cell_pattern, replace_match, formula)

def find_cells_com():
    root = os.path.abspath('infomax_functions_templetes')
    output_path = os.path.abspath('imdg_imdi_results.txt')
    
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    
    results = []
    
    try:
        for r, d, files in os.walk(root):
            for f in files:
                if f.endswith(('.xlsx', '.xlsm')) and not f.startswith('~$'):
                    file_path = os.path.join(r, f)
                    print(f"Checking: {f}")
                    try:
                        wb = excel.Workbooks.Open(file_path, ReadOnly=True)
                        for sheet in wb.Sheets:
                            try:
                                # Search for IMDG or IMDI
                                cells = sheet.UsedRange.Find("IMD", LookIn=-4123) # xlFormulas
                                if cells:
                                    first_addr = cells.Address
                                    curr = cells
                                    while True:
                                        formula = str(curr.Formula)
                                        if "IMDG" in formula.upper() or "IMDI" in formula.upper():
                                            resolved = resolve_formula_simple(sheet, formula)
                                            res = f"File: {file_path}\nSheet: {sheet.Name}\nCell: {curr.Address}\nOriginal: {formula}\nResolved: {resolved}\n" + "-"*50 + "\n"
                                            results.append(res)
                                            print(f"  Found in {sheet.Name} {curr.Address}")
                                        
                                        curr = sheet.UsedRange.FindNext(curr)
                                        if not curr or curr.Address == first_addr:
                                            break
                            except:
                                continue
                        wb.Close(False)
                    except Exception as e:
                        print(f"  Error: {e}")
    finally:
        with open(output_path, 'w', encoding='utf-8') as f_out:
            if results:
                f_out.writelines(results)
            else:
                f_out.write("No IMDG/IMDI formulas found.")
        excel.Quit()
        print(f"Done. Results in {output_path}")

if __name__ == "__main__":
    find_cells_com()
