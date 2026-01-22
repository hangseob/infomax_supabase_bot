import os
import xlwings as xw

def inspect_3206():
    root = 'infomax_functions_templetes'
    mkt_dir = next(d for d in os.listdir(root) if '시장분석' in d)
    stock_dir = next(d for d in os.listdir(os.path.join(root, mkt_dir)) if '주식' in d)
    file_name = next(f for f in os.listdir(os.path.join(root, mkt_dir, stock_dir)) if '3206' in f)
    path = os.path.abspath(os.path.join(root, mkt_dir, stock_dir, file_name))
    
    app = xw.App(visible=False)
    wb = app.books.open(path)
    sheet = wb.sheets['filter']
    print(f"Sheet: {sheet.name}")
    data = sheet.range("A20:F30").value
    for i, row in enumerate(data):
        print(f"Row {i+20}: {row}")
    
    formulas = sheet.range("A20:F30").formula
    for i, row in enumerate(formulas):
        print(f"Formula Row {i+20}: {row}")
        
    wb.close()
    app.quit()

if __name__ == "__main__":
    inspect_3206()
