import openpyxl
import os
import json
from recorder import recorder
from util.configure import column

path_worksheet = os.path.join(os.getcwd(), 'result.xlsx')
path_result_file = os.path.join(os.getcwd(), 'result')

wb = openpyxl.load_workbook(path_worksheet)
st = wb.active

with open(path_result_file, 'r', encoding='utf-8') as f:
    for line in f.readlines():
        item = json.loads(line.strip())
        if len(item) < 10:
            st[column["name"] + item[0]] = item[1]
            st[column["citedby"] + item[0]] = -1
            st[column["hindex"] + item[0]] = -1
            st[column["hindex5y"] + item[0]] = -1
            st[column["i10index"] + item[0]] = -1
            st[column["i10index5y"] + item[0]] = -1
        else:
            st[column["name"] + item[0]] = item[1]
            st[column["affiliation"] + item[0]] = item[2] if st[column["affiliation"] + item[0]].value is not None else st[column["affiliation"] + item[0]].value
            st[column["email"] + item[0]] = item[3]
            st[column["citedby"] + item[0]] = item[4]
            st[column["hindex"] + item[0]] = item[5]
            st[column["hindex5y"] + item[0]] = item[6]
            st[column["i10index"] + item[0]] = item[7]
            st[column["i10index5y"] + item[0]] = item[8]
            st[column["url_picture"] + item[0]] = item[9]

wb.save(path_worksheet)
recorder(path_worksheet)
