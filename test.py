# 引入库
import xlrd
import openpyxl
from openpyxl import Workbook
from openpyxl  import load_workbook

# 找到表的所在路径
xpath = r'C:\Users\lxw6-\learngit' # 相对路径目录
xlspath = r'考勤.xlsx' # 相对路径
xxlspath = xpath + "\\" + xlspath # 绝对路径

# 获取表和工作簿
file_kaoqin = openpyxl.load_workbook(xxlspath)# 获取文件
table_kaoqin = file_kaoqin['考勤表']# 获取工作簿
table_huizong = file_kaoqin['汇总表']
# print(table_kaoqin)

# 遍历表中行和列
nrows_kaoqin = table_kaoqin[1]
ncols_kaoqin = table_kaoqin['A']
# print(nrows_kaoqin)
# print(ncols_kaoqin)

# 打印所有单元格
cells_kaoqin = table_kaoqin['C3':'AG17']
# for cell_rows in cells_kaoqin: # 第一层循环取出来是一行（属于元组）
#     print('\n')
#     for data in cell_rows:     # 第二层循环取出来才是一个数
#         if data != 0:
            # print(data.value,end=' ')

#按格式打印单元格
            # print(data.value,end='\t')# 换行模式不对

#转为字典
rows_data = list(table_kaoqin.rows)# 按行获取数据转换成列表
cols_data = list(table_kaoqin.columns)
# print(rows_data,end='\n') # 一个行元组形成的列表
print(cols_data,end='\n')
# titles_kaoqin =[title.value for title in rows_data.pop(2)]# 获取表单的表头信息，也就是列表的第一个元素
# print(titles_kaoqin)

all_row_dict = []
for a_row in rows_data:
    # print(a_row) #一行一个元组
    the_row_data = [cell.value for cell in a_row]# 转换成列表
    # print(the_row_data) # 一行一个列表
    # for cell_rows in cells_kaoqin:
    # row_dict = dict(zip(the_row_data))
    # print(row_dict)
    # print(row_data.value)
#     all_row_dict.append(row_data)

# print(all_row_dict)