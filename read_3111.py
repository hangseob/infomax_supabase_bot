import xlwings as xw

def read_3111():
    for app in xw.apps:
        for wb in app.books:
            if '3111' in wb.name:
                print(f"Reading Workbook: {wb.name}")
                for sheet in wb.sheets:
                    print(f"--- Sheet: {sheet.name} ---")
                    data = sheet.range("A1:M100").value
                    for i, row in enumerate(data):
                        if any(c is not None for c in row):
                            print(f"  {i+1}: {row}")
                return

if __name__ == "__main__":
    read_3111()
