import zipfile
import re
import os
import sys

# 터미널 출력 인코딩을 UTF-8로 강제 설정
if sys.stdout.encoding != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except:
        pass

def search_kospi200_in_3206():
    file_path = "infomax_functions_templetes/시장분석/주식/3206__KRX_지수_종합.xlsx"
    if not os.path.exists(file_path):
        print("File not found.")
        return

    try:
        with zipfile.ZipFile(file_path, 'r') as z:
            for xml_file in z.namelist():
                if xml_file.startswith('xl/worksheets/sheet'):
                    content = z.read(xml_file).decode('utf-8', errors='ignore')
                    if 'KOSPI200' in content.upper() or 'KOSPI 200' in content.upper():
                        print(f"Found KOSPI200 in {xml_file}")
                        idx = content.upper().find('KOSPI')
                        print(f"Context: {content[max(0, idx-50):idx+100]}")
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    search_kospi200_in_3206()
