import os
import zipfile
import re

def search_imdt_raw(root_dir):
    results = []
    
    for root, dirs, files in os.walk(root_dir):
        for file in files:
            if file.startswith('~$'): continue
            if not file.lower().endswith(('.xlsx', '.xlsm')): continue
            
            file_path = os.path.join(root, file)
            
            try:
                with zipfile.ZipFile(file_path, 'r') as z:
                    # 1. 시트 이름 매핑 찾기 (xl/workbook.xml)
                    workbook_xml = z.read('xl/workbook.xml').decode('utf-8')
                    sheet_map = {} # {sheetId: name}
                    # <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
                    sheets_info = re.findall(r'<sheet [^>]*name="([^"]+)" [^>]*sheetId="([^"]+)"', workbook_xml)
                    for name, sid in sheets_info:
                        sheet_map[sid] = name
                    
                    # 2. 각 시트 XML에서 IMDT 검색
                    for xml_file in z.namelist():
                        if xml_file.startswith('xl/worksheets/sheet'):
                            content = z.read(xml_file).decode('utf-8', errors='ignore')
                            if 'IMDT' in content.upper():
                                # 시트 번호 추출 (sheet1.xml -> 1)
                                s_num = re.search(r'sheet(\d+)\.xml', xml_file).group(1)
                                # workbook.xml의 rId와 매칭해야 정확하지만, 
                                # 보통 sheetId와 파일명이 유사하므로 일단 시도
                                # 더 정확하게는 xl/_rels/workbook.xml.rels 를 봐야 함
                                
                                results.append({
                                    'file': file_path,
                                    'xml': xml_file,
                                    'sheet_guess': sheet_map.get(s_num, f"Sheet {s_num}")
                                })
            except Exception as e:
                # print(f"Error reading {file}: {e}")
                pass
                
    return results

if __name__ == "__main__":
    target_dir = r"infomax_functions_templetes"
    print(f"Searching for 'IMDT' in {target_dir} (Raw XML search)...")
    
    matches = search_imdt_raw(target_dir)
    
    if matches:
        print(f"\nFound IMDT occurrences in {len(matches)} sheets:")
        current_file = ""
        for match in matches:
            if match['file'] != current_file:
                print(f"\n[File: {match['file']}]")
                current_file = match['file']
            print(f"  - XML: {match['xml']} (Likely Sheet: {match['sheet_guess']})")
    else:
        print("\nNo IMDT functions found.")
