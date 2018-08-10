import csv
import openpyxl
import os

'''原本的csv文件转换到excel文件'''

column = {
    "expert": "A",
    "affiliation": "B",
    "interests": "C",
    "name": "D",
    'email':'E',
    "citedby": "F",
    "hindex": "G",
    "hindex5y": "H",
    "i10index": "I",
    "i10index5y": "J",
    "url_picture": "K",
}

path_excel = os.path.join(os.getcwd(), 'result.xlsx')
wb = openpyxl.load_workbook(path_excel)
sheet = wb.active

path_csv = os.path.join(os.getcwd(), 'gene.csv')
with open(path_csv, 'r', encoding='ISO-8859-1') as csv_file:
    reader = csv.reader(csv_file)
    next(reader)
    i = 2
    for r in reader:
        sheet[column["expert"] + str(i)] = r[0]
        sheet[column["affiliation"] + str(i)] = r[1]
        sheet[column["interests"] + str(i)] = r[2]
        i += 1

wb.save(path_excel)

