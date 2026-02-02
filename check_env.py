import xlwings as xw
import os
import time

def check_env():
    print("현재 엑셀 환경을 점검합니다...")
    try:
        app = xw.apps.active if xw.apps.count > 0 else xw.App(visible=True)
        print(f"Excel Apps: {xw.apps.count}")
        
        infomax_xlam_path = r"C:\Infomax\bin\excel\infomaxexcel.xlam"
        if os.path.exists(infomax_xlam_path):
            print(f"XLAM file exists: {infomax_xlam_path}")
            # Try to open if not already open
            is_open = False
            for wb in app.books:
                if "infomaxexcel" in wb.name.lower():
                    print(f"Infomax Add-in is already open: {wb.name}")
                    is_open = True
                    break
            if not is_open:
                print("Opening Infomax Add-in...")
                app.books.open(infomax_xlam_path)
        else:
            print(f"XLAM file NOT found: {infomax_xlam_path}")

        # Test a simple IMDP formula first (realtime)
        wb = app.books.add()
        sheet = wb.sheets[0]
        sheet.range("A1").formula = '=IMDP("STK", "005930", "현재가")'
        
        print("Waiting for IMDP value...")
        for i in range(10):
            time.sleep(2)
            val = sheet.range("A1").value
            if val is not None and not (isinstance(val, str) and "#WAITING" in val.upper()):
                print(f"IMDP Success! Value: {val}")
                break
            print(f"IMDP status: {val}")

        # Now try IMDH with date
        sheet.range("A3").formula = '=IMDH("STK", "005930", "날짜,현재가", "20260120", "20260124", 5, "Headers=1")'
        print("Waiting for IMDH value (Date, Price)...")
        for i in range(15):
            time.sleep(2)
            val = sheet.range("A3:B8").value
            if val and val[0][0] is not None and not (isinstance(val[0][0], str) and "#WAITING" in val[0][0].upper()):
                print("IMDH Success!")
                for r in val:
                    print(f"  {r}")
                break
            print(f"IMDH status: {val[0][0] if val else 'None'}")

    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    check_env()
