import os
import win32com.client
import re
import time

def resolve_formula_simple(sheet, formula):
    if not isinstance(formula, str) or not formula.startswith('='):
        return formula
    
    # Simple regex for cell references
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
    
    print("Starting robust scan...")
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
                    
                    wb = None
                    # Retry logic for OLE busy error
                    for attempt in range(5):
                        try:
                            wb = excel.Workbooks.Open(file_path, ReadOnly=True)
                            break
                        except Exception as e:
                            if "0x800ac472" in str(e):
                                print(f"  Excel busy, retrying {attempt+1}/5...")
                                time.sleep(2)
                            else:
                                print(f"  Failed to open: {e}")
                                break
                    
                    if wb:
                        try:
                            for sheet in wb.Sheets:
                                try:
                                    # Use Find to find all cells containing "IMD"
                                    # This is much faster than iterating all cells
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
                                except Exception as e:
                                    print(f"  Error searching sheet {sheet.Name}: {e}")
                        finally:
                            wb.Close(False)
    finally:
        with open(output_path, 'w', encoding='utf-8') as f_out:
            if results:
                f_out.writelines(results)
            else:
                f_out.write("No IMDG/IMDI formulas found.")
        try:
            excel.Quit()
        except:
            pass
        print(f"\nDone. Results saved to: {output_path}")

if __name__ == "__main__":
    find_cells_robust()
