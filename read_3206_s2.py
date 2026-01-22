import os
import xlwings as xw

def read_3206_sheet2():
    root = 'infomax_functions_templetes'
    mkt_dir = next(d for d in os.listdir(root) if '시장분석' in d)
    stock_dir = next(d for d in os.listdir(os.path.join(root, mkt_dir)) if '주식' in d)
    file_name = next(f for f in os.listdir(os.path.join(root, mkt_dir, stock_dir)) if '3206' in f)
    path = os.path.abspath(os.path.join(root, mkt_dir, stock_dir, file_name))
    
    app = xw.App(visible=False)
    try:
        wb = app.books.open(path)
        sheet = wb.sheets['Sheet2']
        data = sheet.range("A1:C50").value
        for i, row in enumerate(data):
            print(f"{i+1}: {row}")
        wb.close()
    finally:
        app.quit()

if __name__ == "__main__":
    read_3206_sheet2()
