import os
import xlwings as xw

def get_3206_formulas():
    root = 'infomax_functions_templetes'
    mkt_dir = next(d for d in os.listdir(root) if '시장분석' in d)
    stock_dir = next(d for d in os.listdir(os.path.join(root, mkt_dir)) if '주식' in d)
    file_name = next(f for f in os.listdir(os.path.join(root, mkt_dir, stock_dir)) if '3206' in f)
    path = os.path.abspath(os.path.join(root, mkt_dir, stock_dir, file_name))
    
    app = xw.App(visible=False)
    try:
        wb = app.books.open(path)
        print(f"Opened: {path}")
        for sheet in wb.sheets:
            print(f"--- Sheet: {sheet.name} ---")
            formulas = sheet.range("A1:M50").formula
            for r_idx, row in enumerate(formulas):
                if any('IMD' in str(c).upper() for c in row):
                    print(f"  Row {r_idx+1}: {row}")
        wb.close()
    finally:
        app.quit()

if __name__ == "__main__":
    get_3206_formulas()
