import os
import xlwings as xw

def find_kospi200_template():
    root = 'infomax_functions_templetes'
    mkt_dir = next(d for d in os.listdir(root) if '시장분석' in d)
    stock_dir = next(d for d in os.listdir(os.path.join(root, mkt_dir)) if '주식' in d)
    stock_path = os.path.join(root, mkt_dir, stock_dir)
    
    app = xw.App(visible=False)
    target_files = [f for f in os.listdir(stock_path) if '3206' in f or 'KRX' in f or '200' in f]
    
    for f in target_files:
        path = os.path.abspath(os.path.join(stock_path, f))
        print(f"Checking: {path}")
        try:
            wb = app.books.open(path)
            for sheet in wb.sheets:
                formulas = sheet.range("A1:J50").formula
                for r_idx, row in enumerate(formulas):
                    for c_idx, cell in enumerate(row):
                        if isinstance(cell, str) and ('211' in cell or '200' in cell):
                            print(f"  Match in {sheet.name} at ({r_idx+1}, {c_idx+1}): {cell}")
            wb.close()
        except Exception as e:
            print(f"  Error: {e}")
    app.quit()

if __name__ == "__main__":
    find_kospi200_template()
