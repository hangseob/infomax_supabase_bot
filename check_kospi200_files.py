import pandas as pd
import os

def check_kospi200_files():
    files_to_check = [
        "kospi200_test_results.xlsx",
        "infomax_functions_templetes/시장분석/주식/3206__KRX_지수_종합.xlsx",
        "infomax_functions_templetes/시장분석/주식/3111__주식종목정보.xlsx"
    ]
    
    for file in files_to_check:
        if os.path.exists(file):
            print(f"\n[Checking: {file}]")
            try:
                # Read first few rows of each sheet to see what's inside
                xl = pd.ExcelFile(file)
                for sheet_name in xl.sheet_names:
                    print(f"  Sheet: {sheet_name}")
                    df = pd.read_excel(xl, sheet_name=sheet_name, nrows=5)
                    print(df.head())
            except Exception as e:
                print(f"  Error reading {file}: {e}")
        else:
            print(f"\n[File not found: {file}]")

if __name__ == "__main__":
    check_kospi200_files()
