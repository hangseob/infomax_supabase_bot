import win32com.client

def get_exact_names():
    try:
        excel = win32com.client.GetActiveObject("Excel.Application")
        target_file = "우리집 가계 금융 현황.종합.xlsx"
        wb = None
        for b in excel.Workbooks:
            if target_file in b.Name:
                wb = b
                break
        
        if wb:
            table_name = "표.거래내역"
            for sheet in wb.Sheets:
                for tbl in sheet.ListObjects:
                    if tbl.Name == table_name:
                        columns = [col.Name for col in tbl.ListColumns]
                        print("COLUMNS_START")
                        for col in columns:
                            print(col)
                        print("COLUMNS_END")
                        return
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    get_exact_names()
