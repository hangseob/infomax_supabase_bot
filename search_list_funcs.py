import os
import xlwings as xw

def search_list_functions():
    root = 'infomax_functions_templetes'
    app = xw.App(visible=False)
    try:
        for r, d, files in os.walk(root):
            for f in files:
                if f.endswith(('.xlsx', '.xlsm')):
                    path = os.path.abspath(os.path.join(r, f))
                    try:
                        wb = app.books.open(path)
                        for sheet in wb.sheets:
                            try:
                                formulas = sheet.range('A1:L50').formula
                                for row in formulas:
                                    for cell in row:
                                        if isinstance(cell, str) and any(func in cell.upper() for func in ['IMDG', 'IMDI', 'IMDB']):
                                            print(f"File: {path}\n  Sheet: {sheet.name}\n  Formula: {cell}")
                            except:
                                continue
                        wb.close()
                    except:
                        continue
    finally:
        app.quit()

if __name__ == "__main__":
    search_list_functions()
