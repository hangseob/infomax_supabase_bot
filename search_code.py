import xlwings as xw

def search_005930():
    print("--- Searching for '005930' in all open workbooks ---")
    try:
        for app in xw.apps:
            for wb in app.books:
                for sheet in wb.sheets:
                    # Check first few rows and columns for the code
                    try:
                        data = sheet.range("A1:J50").value
                        for r_idx, row in enumerate(data):
                            for c_idx, cell in enumerate(row):
                                if cell and ('005930' in str(cell) or cell == 5930):
                                    print(f"Found '005930' in Workbook: {wb.name}, Sheet: {sheet.name}, Cell: ({r_idx+1}, {c_idx+1})")
                                    # Print the row
                                    print(f"  Row Data: {row}")
                    except:
                        continue
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    search_005930()
