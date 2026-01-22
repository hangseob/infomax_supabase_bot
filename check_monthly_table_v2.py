import win32com.client

def check_monthly_table_v2():
    try:
        excel = win32com.client.GetActiveObject("Excel.Application")
        target_file = "우리집 가계 금융 현황.종합.xlsx"
        wb = None
        for b in excel.Workbooks:
            if target_file in b.Name:
                wb = b
                break
        
        if wb:
            print(f"Workbook found: {wb.Name}")
            table_name = "표.월중주식"
            found = False
            for sheet in wb.Sheets:
                print(f"Checking sheet: {sheet.Name}")
                for tbl in sheet.ListObjects:
                    print(f"  Found table: {tbl.Name}")
                    if tbl.Name == table_name:
                        columns = [col.Name for col in tbl.ListColumns]
                        print(f"--- MATCH FOUND ---")
                        print(f"TABLE: {tbl.Name}")
                        print(f"SHEET: {sheet.Name}")
                        for col in columns:
                            print(f"COLUMN: {col}")
                        found = True
                        break
                if found: break
            if not found:
                print(f"TABLE_NOT_FOUND: {table_name}")
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    check_monthly_table_v2()
