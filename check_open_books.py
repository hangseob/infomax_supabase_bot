import xlwings as xw

def check_all_open_books():
    print("--- Checking all open workbooks ---")
    try:
        if xw.apps.count == 0:
            print("No Excel apps running.")
            return

        for app in xw.apps:
            print(f"App (PID: {app.pid})")
            for wb in app.books:
                print(f"  Workbook: {wb.name}")
                for sheet in wb.sheets:
                    data = sheet.range("A1:C30").value
                    valid_rows = [r for r in data if any(c is not None for c in r)]
                    print(f"    Sheet: {sheet.name} - Rows with data: {len(valid_rows)}")
                    if valid_rows:
                        for row in valid_rows[:10]:
                            print(f"      {row}")
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    check_all_open_books()
