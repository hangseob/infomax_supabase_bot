import os
import xlwings as xw

def check_template():
    root = 'infomax_functions_templetes'
    try:
        # Navigate to the file
        mkt_dir = next(d for d in os.listdir(root) if '시장분석' in d)
        stock_dir = next(d for d in os.listdir(os.path.join(root, mkt_dir)) if '주식' in d)
        file_name = next(f for f in os.listdir(os.path.join(root, mkt_dir, stock_dir)) if '3184' in f)
        full_path = os.path.abspath(os.path.join(root, mkt_dir, stock_dir, file_name))
        
        print(f"Opening: {full_path}")
        app = xw.App(visible=False)
        wb = app.books.open(full_path)
        
        for sheet in wb.sheets:
            if 'hist' in sheet.name.lower():
                print(f"--- Sheet: {sheet.name} ---")
                data = sheet.range("A5:E50").value
                valid = [r for r in data if r[0] is not None]
                print(f"Total rows with data: {len(valid)}")
                if valid:
                    print("First 10 rows:")
                    for i, r in enumerate(valid[:10]):
                        print(f"  {i+1}: {r}")
        
        wb.close()
        app.quit()
    except Exception as e:
        print(f"Error: {e}")
    finally:
        if 'app' in locals():
            try: app.quit()
            except: pass

if __name__ == "__main__":
    check_template()
