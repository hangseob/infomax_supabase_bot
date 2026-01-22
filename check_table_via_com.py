import win32com.client
import os

def check_table_columns():
    try:
        excel = win32com.client.GetActiveObject("Excel.Application")
        print("Connected to active Excel instance.")
        
        target_file = "우리집 가계 금융 현황.종합.xlsx"
        wb = None
        for b in excel.Workbooks:
            if target_file in b.Name:
                wb = b
                break
        
        if not wb:
            print(f"Workbook '{target_file}' is not open in the active Excel instance.")
            # Let's list open workbooks
            print("Open workbooks:")
            for b in excel.Workbooks:
                print(f"  - {b.Name}")
            return

        print(f"Found workbook: {wb.Name}")
        
        table_name = "표.거래내역"
        found_table = None
        for sheet in wb.Sheets:
            for tbl in sheet.ListObjects:
                if tbl.Name == table_name:
                    found_table = tbl
                    print(f"Found table '{table_name}' on sheet '{sheet.Name}'")
                    break
            if found_table: break
            
        if found_table:
            columns = [col.Name for col in found_table.ListColumns]
            print(f"Columns in '{table_name}':")
            for col in columns:
                print(f"  - {col}")
        else:
            print(f"Table '{table_name}' not found in the workbook.")
            
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    check_table_columns()
