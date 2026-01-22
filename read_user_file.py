import xlwings as xw
import os
import time

def read_active_or_file(filename):
    print(f"--- Attempting to read {filename} ---")
    try:
        # Check if Excel is running
        if xw.apps.count > 0:
            print("Excel is running. Checking open workbooks...")
            for app in xw.apps:
                for wb in app.books:
                    if filename in wb.name:
                        print(f"Found open workbook: {wb.name}")
                        read_data(wb)
                        return
        
        # If not open, try to open it
        path = os.path.abspath(filename)
        if os.path.exists(path):
            print(f"Opening file from disk: {path}")
            app = xw.App(visible=False)
            wb = app.books.open(path)
            read_data(wb)
            wb.close()
            app.quit()
        else:
            print(f"File not found on disk: {filename}")
    except Exception as e:
        print(f"Error: {e}")
    finally:
        if 'app' in locals():
            try: app.quit()
            except: pass

def read_data(wb):
    for sheet in wb.sheets:
        print(f"--- Sheet: {sheet.name} ---")
        # Read a larger range including more columns
        data = sheet.range("A1:J100").value
        
        # Filter rows that have any data
        rows_with_data = []
        for i, row in enumerate(data):
            if any(cell is not None for cell in row):
                rows_with_data.append((i + 1, row))
        
        print(f"Total rows with data in {sheet.name}: {len(rows_with_data)}")
        if rows_with_data:
            print("Data preview (first 30 rows):")
            for line_num, content in rows_with_data[:30]:
                # Print only non-None columns to keep it clean
                clean_row = [c for c in content if c is not None]
                print(f"  Line {line_num}: {content}")

if __name__ == "__main__":
    # The user might have named it anything, but let's check the ones we know
    read_active_or_file('samsung_electronics_price.xlsx')
