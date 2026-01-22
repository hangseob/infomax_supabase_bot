import win32com.client

def get_hex_names():
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
                        for col in tbl.ListColumns:
                            name = col.Name
                            hex_name = ":".join("{:04x}".format(ord(c)) for c in name)
                            print(f"{name} | {hex_name}")
                        return
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    get_hex_names()
