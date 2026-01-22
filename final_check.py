import xlwings as xw
import os

def check_file(filename):
    path = os.path.abspath(filename)
    if not os.path.exists(path):
        print(f"File not found: {filename}")
        return
    
    try:
        app = xw.App(visible=False)
        wb = app.books.open(path)
        sheet = wb.sheets[0]
        data = sheet.range('A1:B40').value
        valid = [r for r in data if r is not None and (isinstance(r, list) and r[0] is not None or r is not None)]
        
        # More robust checking for rows with data
        rows_with_data = []
        for r in data:
            if isinstance(r, list):
                if any(cell is not None for cell in r):
                    rows_with_data.append(r)
            elif r is not None:
                rows_with_data.append(r)

        print(f"--- Checking {filename} ---")
        print(f"Rows with any data: {len(rows_with_data)}")
        if rows_with_data:
            print("First 10 rows of data:")
            for i, r in enumerate(rows_with_data[:10]):
                print(f"  Row {i+1}: {r}")
        
        wb.close()
        app.quit()
    except Exception as e:
        print(f"Error checking {filename}: {e}")
    finally:
        if 'app' in locals():
            try: app.quit()
            except: pass

if __name__ == "__main__":
    check_file('samsung_electronics_price.xlsx')
    check_file('samsung_electronics_price_1.xlsx')
    check_file('samsung_electronics_price_2.xlsx')
