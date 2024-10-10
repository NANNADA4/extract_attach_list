"""
폴더를 순회하면서 PDF 파일을 찾고 북마크를 가져와 별도제출자료 추출을 시도합니다
"""

import os
import re
from natsort import natsorted


from module.create_excel import _load_excel, create_excel
from module.create_log import logging
from module.extract_bookmark import extract_bookmark


def process_folder(input_path):
    """폴더를 순회하면서 PDF 파일을 찾아 북마크 추출을 시도합니다"""
    for root, _, files in os.walk(input_path):
        for file in natsorted(files):
            excel_list = []

            if not file.lower().endswith('.pdf'):
                continue
            pdf_path = os.path.join(root, file)

            first_underscore_index = file.find('_')
            second_underscore_index = file.find(
                '_', first_underscore_index + 1)
            if first_underscore_index != -1 and second_underscore_index != -1:
                cmt = file[first_underscore_index +
                           1:second_underscore_index]
            else:
                cmt = ""

            org_matches = re.findall(r'\(([^)]+)\)', file)
            if org_matches:
                org = org_matches[-1]
                if org == '2':
                    org = org_matches[-2]
                if str(org).endswith('(주'):
                    org = str(org).replace('(주', '(주)')
            else:
                org = ""

            for item in extract_bookmark(pdf_path):
                try:
                    if len(item) <= 1 or item['level'] != 4:
                        continue

                    excel_list.append({
                        "cmt": cmt,
                        "org": org,
                        "name": item['parent']['parent']['title'],
                        "question": item['parent']['title'],
                        "realfile_name": item['title'],
                        "real_path": os.path.join(cmt, org, item['parent']['parent']['title'],
                                                  item['title']),
                        "file_name": file
                    })
                except Exception as e:  # pylint: disable=W0703
                    e = "PDF 북마크 추출 오류"
                    logging(e, '', input_path)

            excel_path = os.path.join(input_path, 'SubmitList.xlsx')
            create_excel(_load_excel(excel_path), excel_list, excel_path)
