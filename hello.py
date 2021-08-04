# import numpy
import xlrd
import xlwt
# import xlwings as xw
#i='internet14975'
# for x in i:
#     if x == 'e':
#         print(x)

# # y=['1',2,'nn','fd','3']
# # print(y[1]*y[2])


# def sumxxxx(l):
#     y=0
#     for i in l:
#         y=y+i
#     return y

# x=[1,2,3,6,10]
# z=[2,4,5,6,7]

# xxx = list('eleven')
# mmm = list('hello')
# yy = set(i)
# zz = 
# yyy = list('eleven')
# aaa = set(yyy)

# ret = xxx + mmm
# a = [1,2,3]
# b = (1,2,3)

# a.append(4)
# print(a)
# b

# dic = dict({
#     "liuxiaomang": {"age": "60", "born": "fujian"},
#     "shuaige": {"age": "18", "born": "earth"}
#     })

# print(dic["liuxiaomang"]["age"])
# print('√')

# b = sorted(list(set(i)))
# c = sorted(list(i))
# # print(b)
# print(c)
# for i,i1 in enumerate(c) :
#     if i!=0:
#         if c[i]==c[i-1]:
#             print(i,i1)
    
# print(xxx)
# print(yyy)

# ret = sumxxxx(x)
# ret1 = sumxxxx(z)
# print(ret1)
    # print(y)

# ceshi = xlrd.open_workbook(r'D:\data\python\pythonProject\test.xls')
# s = ceshi.sheet_by_name('sheet1')
# nrows = s.nrows
# print(nrows)

# xlspath = r'test.xls'
# xpath = r'D:\data\python\pythonProject'
# xxlspath = xpath + "\\" + xlspath
# # print (xxlspath)

# xls = xlrd.open_workbook(xxlspath)# 读取文件
# xls_xlutils = xlutils.copy.copy(xls)
# xlwt_xls = xlwt.Workbook(encoding='utf-8')
# xlwt_sheet_index = xlwt.
# sheet = xls.sheet_by_name("sheet1")
# value = sheet.cell_value(4,3)
# print (value)
# xlwt_xls.write(4,6,"新内容")
# sheet1 = xls.sheet_by_name("sheet1")# 通过sheet名查找
# sheet2 = xls.sheet_by_index(0)# 通过索引查找
#print (sheet1)
#print (sheet2)
# value = sheet1.cell_value(1,2) # 第1行第2列的单元格
# print("第1行第2列的值是：" , value)

# rows = sheet2.row_values(1)
# cols = sheet2.col_values(2)

# for cell in rows:
#     # print ("这1行是：%s"%cell , '')
#     print("这1行是：{}".format(cell))

# xls= xlwt.Workbook(encoding = 'ascii')#创建文件
# worksheet = xls.add_sheet("sheet1")#创建工作簿

# xls.save("new_test.xls")# 保存工作簿

# app = xw.App(visible=True,add_book=False)
# app.display_alerts = False #警告关闭
# app.screen_updating = False # 屏幕更新关闭
# #wb = app.books.add() #新建工作簿

# #wb = app.books.open(xls_path) #打开excel
# wb = xw.Book(xls_path)
# sheet = wb.sheets.active

# B1 = sheet.range('B1').value
# A1_C4 = sheet.range('A1:C4').value
# print (A1_C4)


# # wb.save(path + r'\new_pactice.xlsx')
# wb.save()
# wb.close()
# app.quit()

# name = input('请输入名字')
# age = input ( '请输入年龄')
# message = '姓名:{},年龄:{}'.format(name,age)
# print(message)

a = '@+3'
print(a[1:])
