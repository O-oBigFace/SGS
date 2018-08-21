import csv
import openpyxl
import os
import util.configure

'''原本的csv文件转换到excel文件'''

column = util.configure.column

main_path = os.path.abspath(os.path.join(os.getcwd(), '..'))

path_excel = os.path.join(main_path, 'result.xlsx')
wb = openpyxl.load_workbook(path_excel)
sheet = wb.active

for key,value in column.items():
    sheet[value + "1"] = key

path_csv = os.path.join(main_path, 'export.csv')
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

