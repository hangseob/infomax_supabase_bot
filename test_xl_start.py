import xlwings as xw
import time

print("Starting Excel app test...")
try:
    app = xw.App(visible=True)
    print("App created successfully.")
    wb = app.books.add()
    print("Workbook added.")
    time.sleep(2)
    wb.close()
    app.quit()
    print("Excel closed normally.")
except Exception as e:
    print(f"Error: {e}")
