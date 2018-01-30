# -*- coding:utf-8 -*-
import xlrd
import codecs
'''
文件路径比较重要，要以这种方式去写文件路径不用
'''
file_path = r"/Users/wujianglong/PycharmProjects/file2dict/门诊医保.xlsx"
# 读取的文件路径
file_path = file_path.decode('utf-8')
# 文件中的中文转码
data = xlrd.open_workbook(file_path)
# 获取数据
table = data.sheet_by_name('Sheet')
# 获取sheet
nrows = table.nrows
# 获取总行数
ncols = table.ncols
# 获取总列数

output_file = "medical_type.dict"
f = codecs.open(output_file, 'a', "utf-8")

for row in range(1, nrows):
    line = []
    for col in [1, 0]:
        line.append(table.cell(row, col).value)
    write_line = "\t".join(line)
    f.write(write_line + '\n')

f.close()

# print table.row_values(1)
# # 获取一行的数值
# print table.col_values(1)
# # 获取一列的数值
#
# # 获取一个单元格的数值
# cell_value = table.cell(1, 1).value
# print cell_value