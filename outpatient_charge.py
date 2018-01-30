# -*- coding:utf-8 -*-
import xlrd
import codecs
'''
文件路径比较重要，要以这种方式去写文件路径不用
'''
file_path = r"/Users/wujianglong/PycharmProjects/file2dict/门诊费用分类 20180130.xlsx"
# 读取的文件路径
file_path = file_path.decode('utf-8')
# 文件中的中文转码
data = xlrd.open_workbook(file_path)
# 获取数据u
table = data.sheet_by_name(u'工作表1')
# 获取sheet
nrows = table.nrows
# 获取总行数
ncols = table.ncols
# 获取总列数

output_file = "outpatient_charge_type.dict"
f = codecs.open(output_file, 'w', "utf-8")

for row in range(1, nrows):
    line = []
    line.append(table.cell(row, 0).value)
    field = table.cell(row, 1).value
    print field
    if field == u"材料费":
        field = u"material_fee"
    elif field == u"挂号费":
        field = u"register_fee"
    elif field == u"检查费":
        field = u"exam_fee"
    elif field == u"其他费":
        field = u"other_fee"
    elif field == u"手术费":
        field = u"oper_fee"
    elif field == u"药费":
        field = u"drug_fee"
    line.append(field)
    write_line = ",".join(line)
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