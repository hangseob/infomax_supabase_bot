import xlwings as xw
import os
import time

def read_active_or_file(filename):
    print(f"--- Checking {filename} ---")
    app = None
    try:
        # Connect to running Excel
        try:
            app = xw.apps.active
            print("Connected to active Excel app.")
        except:
            print("No active Excel app found, trying to open new one.")
            app = xw.App(visible=False)

        # Find or open workbook
        wb = None
        for book in app.books:
            if filename in book.name:
                wb = book
                print(f"Using open workbook: {wb.name}")
                break
        
        if not wb:
            path = os.path.abspath(filename)
            if os.path.exists(path):
                print(f"Opening file from disk: {path}")
                wb = app.books.open(path)
            else:
                print(f"File not found: {filename}")
                return

        # Read data with retries for OLE error
        for i in range(5):
            try:
                for sheet in wb.sheets:
                    print(f"\n--- Sheet: {sheet.name} ---")
                    data = sheet.range("A1:J100").value
                    rows = []
                    for row_idx, row_data in enumerate(data):
                        if row_data and any(c is not None for c in row_data):
                            rows.append((row_idx + 1, row_data))
                    
                    print(f"Total rows found: {len(rows)}")
                    for line_num, content in rows[:30]:
                        print(f"  Row {line_num}: {content}")
                break # Success
            except Exception as e:
                print(f"Retry {i+1} due to error: {e}")
                time.sleep(1)
                
    except Exception as e:
        print(f"Final Error: {e}")
    finally:
        # We don't close the user's open workbook
        pass

if __name__ == "__main__":
    read_active_or_file('samsung_electronics_price.xlsx')
