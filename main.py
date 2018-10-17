import xlrd
import xlsxwriter
import os

cwd = os.getcwd()

data_file = cwd + '\\Datas\\B3Unit2B.xlsx'

workbook = xlrd.open_workbook(data_file)
print(workbook.sheet_names())
sheet1_name = workbook.sheet_names()[0]
print(sheet1_name)

sheet1 = workbook.sheet_by_index(0)
print(sheet1.name, sheet1.nrows, sheet1.ncols)

a = sheet1.cell(0,0).value
b = sheet1.cell(1,0).value
print(sheet1.cell(0,0).value)
print(sheet1.cell(1,0).value)

data_file = cwd + '\\Datas\\Modify\\B3Unit2B.xlsx'
workbook = xlsxwriter.Workbook(data_file)
worksheet = workbook.add_worksheet()

title = ['单词', '音标', '用户释义','词典释义', '词汇特性', '学习次数', '检测次数', '错误次数', '正确率']
title_len = len(title)

for i in range(title_len):
    worksheet.write(0, i ,title[i])

k = 1
j = 1
for i in range(sheet1.nrows):
    if(i%2 == 1):
        worksheet.write(k, 0, sheet1.cell(i, 0).value)
        k += 1
        continue
    worksheet.write(j, 2, sheet1.cell(i, 0).value)
    j += 1

workbook.close()