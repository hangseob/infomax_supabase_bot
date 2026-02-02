import xlwings as xw
import os
import time

def test_ir_date_variations():
    app = xw.apps.active if xw.apps.count > 0 else xw.App(visible=True)
    infomax_xlam_path = r"C:\Infomax\bin\excel\infomaxexcel.xlam"
    if os.path.exists(infomax_xlam_path):
        app.books.open(infomax_xlam_path)
    
    wb = app.books.add()
    sheet = wb.sheets[0]
    
    # Yesterday was 2026-01-23 (Friday)
    target_date = "20260123"
    
    fields_to_try = [
        "날짜,MID종가",
        "일자,MID종가",
        "거래일자,MID종가",
        "DATE,MID종가",
        "TIME,MID종가",
        "MID종가"
    ]
    
    for idx, fields in enumerate(fields_to_try):
        row = (idx * 5) + 1
        formula = f'=IMDH("IR", "CRST25USDCNH06M", "{fields}", "{target_date}", "{target_date}", 1, "Headers=1,Orient=V")'
        sheet.range(row, 1).formula = formula
        print(f"Testing fields: {fields}")
        
    print("Waiting for results...")
    time.sleep(10)
    
    for idx, fields in enumerate(fields_to_try):
        row = (idx * 5) + 1
        data = sheet.range((row, 1), (row + 2, 2)).value
        print(f"\nFields: {fields}")
        if data:
            for r_idx, r in enumerate(data):
                print(f"  Row {r_idx+1}: {r}")

if __name__ == "__main__":
    test_ir_date_variations()
