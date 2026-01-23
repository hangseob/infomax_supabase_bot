import os
import zipfile
import re
import sys

# 인코딩 설정
if sys.stdout.encoding != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except:
        pass

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
                    sheet_map = {} # {rId: name}
                    # <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
                    sheets_info = re.findall(r'<sheet [^>]*name="([^"]+)" [^>]*sheetId="([^"]+)" [^>]*r:id="([^"]+)"', workbook_xml)
                    if not sheets_info:
                         # try another order
                         sheets_info = re.findall(r'<sheet [^>]*name="([^"]+)" [^>]*r:id="([^"]+)" [^>]*sheetId="([^"]+)"', workbook_xml)
                         # rearrange to (name, sid, rid)
                         sheets_info = [(n, s, r) for n, r, s in sheets_info]

                    # 2. 관계 매핑 찾기 (xl/_rels/workbook.xml.rels)
                    rels_xml = z.read('xl/_rels/workbook.xml.rels').decode('utf-8')
                    rel_map = {} # {rId: target_file}
                    rels = re.findall(r'<Relationship [^>]*Id="([^"]+)" [^>]*Target="([^"]+)"', rels_xml)
                    for rid, target in rels:
                        rel_map[rid] = target

                    # 3. 시트 이름과 XML 파일 매칭
                    sheet_to_xml = {} # {xml_path: sheet_name}
                    for name, sid, rid in sheets_info:
                        xml_path = rel_map.get(rid)
                        if xml_path:
                            # Target="worksheets/sheet1.xml" -> "xl/worksheets/sheet1.xml"
                            full_path = 'xl/' + xml_path if not xml_path.startswith('xl/') else xml_path
                            sheet_to_xml[full_path] = name

                    # 4. 각 시트 XML에서 IMDT 검색
                    for xml_file, sheet_name in sheet_to_xml.items():
                        if xml_file in z.namelist():
                            content = z.read(xml_file).decode('utf-8', errors='ignore')
                            if 'IMDT' in content.upper():
                                results.append({
                                    'file': file_path,
                                    'sheet': sheet_name
                                })
            except Exception as e:
                pass
                
    return results

if __name__ == "__main__":
    target_dir = r"infomax_functions_templetes"
    # print(f"Searching for 'IMDT' in {target_dir} (Raw XML search)...")
    
    matches = search_imdt_raw(target_dir)
    
    if matches:
        # print(f"\nFound IMDT occurrences in {len(matches)} sheets:")
        current_file = ""
        for match in matches:
            if match['file'] != current_file:
                print(f"\n[엑셀파일명: {os.path.basename(match['file'])}]")
                print(f"경로: {match['file']}")
                current_file = match['file']
            print(f"  - 시트명: {match['sheet']}")
    else:
        print("\nNo IMDT functions found.")
