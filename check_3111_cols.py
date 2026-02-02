import pandas as pd
import os

def check_3111_columns():
    file_path = os.path.join("infomax_functions_templetes", "시장분석", "주식", "3111__주식종목정보.xlsx")
    if not os.path.exists(file_path):
        print(f"File not found: {file_path}")
        return

    try:
        df = pd.read_excel(file_path, nrows=5)
        print("Columns found in 3111__주식종목정보.xlsx:")
        for col in df.columns:
            print(f"- {col}")
    except Exception as e:
        print(f"Error reading file: {e}")

if __name__ == "__main__":
    check_3111_columns()
