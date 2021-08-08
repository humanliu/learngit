# 引入库
import xlrd
import re
import openpyxl
from openpyxl import Workbook
from openpyxl  import load_workbook

# 找到表的所在路径
xpath = r'C:\Users\lxw6-\learngit' # 相对路径目录
xlspath = r'考勤.xlsx' # 相对路径
xxlspath = xpath + "\\" + xlspath # 绝对路径

# 获取表和工作簿
file_kaoqin = openpyxl.load_workbook(xxlspath)# 获取文件
table_kaoqin = file_kaoqin['考勤表']# 获取考勤表
table_huizong = file_kaoqin['汇总表']# 获取汇总表
people = 13 # 人员数,这里可以添加input
use_rows = 'AG%d'%(people+4) # 所需要取的行数
#print(use_rows)

# 遍历表中行和列
    # nrows_kaoqin = table_kaoqin[1]
    # ncols_kaoqin = table_kaoqin['A']
    # print(nrows_kaoqin)
    # print(ncols_kaoqin)

# 打印所有单元格
use_kaoqin = table_kaoqin['C3':use_rows]# 只取需要的部分
    # print (use_kaoqin)
    # for cell_rows in cells_kaoqin: # 第一层循环取出来是一行（属于元组）
    #     print('\n')
    #     for data in cell_rows:     # 第二层循环取出来才是一个数
    #         if data != 0:
                # print(data.value,end=' ')

#按格式打印单元格
            # print(data.value,end='\t')# 换行模式不对

# 转为列表
rows_data = list(use_kaoqin)# 按行获取数据转换成列表
# cols_data = list(use_kaoqin)
    # print(rows_data,end='\n') # 一个行元组形成的列表
    # print(cols_data,end='\n')
    # print(len(rows_data),end='\n')# 查看共有几行
    # use_rows = rows_data[2:people+2+2]# 实际取的行要加上标题、部门、日期、星期几，所以从3开始，直到人员数+4
    # 
    # print(len(use_rows),end="")
    # titles_kaoqin =[title.value for title in rows_data.pop(2)]# 获取表单的表头信息，也就是列表的第一个元素
    # print(titles_kaoqin)

# all_row_dict = []
list_rows = [] #以下步骤可选
for a_row in rows_data:
    # print(a_row) # 一行一个元组
    list_rows.append(list(a_row))
# print (list_rows) # 全部为列表 
# the_row_data = [cell.value for cell in a_row]# 转换成列表
    # print(the_row_data) # 一行一个列表
    # print(len(the_row_data))
    # for cell_rows in cells_kaoqin:
    # row_dict = dict(zip(the_row_data))
    # print(row_dict)
    # print(row_data.value)
    # all_row_dict.append(row_data)
    # print(all_row_dict)

# 初始化各个值
chuqin = chuqin_x = chuchai = jiaban = jiaban_x = gongxiu = buxiu = hunjia = sangjia = shijia = shuangxiu = liangxin = jiejia =  0.0
gou = '√'
gou_jiaban = '√+'
c = 'C'
# ban = re.findall("\d+",)
b = 'B'
g = 'G'
h = 'H'
sang = 'S'
shi = '△'
shuangxiufg = 'FFD9D9D9'
sanxinfg = 'FFFFFF00'
for i in list_rows[2]:
    print(i.value,end=" ")
# 汇总每行的√ 出勤
    if i.value == gou :  
        chuqin += 1.0
        # 根据颜色区分是否为公休日、节假日
        if i.fill.fgColor.rgb == shuangxiufg:
            shuangxiu += 1.0
        if i.fill.fgColor.rgb == sanxinfg:
            jiejia += 1.0
    if len(i.value)>2 and (i.value[0:2] == '√/'or i.value[1:3] == '/√' ):
        # print (' 这里增加了:',i.value[0:2])
        chuqin += 0.5
# 汇总每行的C 出差（公休出差是否为双倍、三倍？？？？？？？）
    if i.value == c:
        chuqin += 1.0
        if i.fill.fgColor.rgb == shuangxiufg:
            shuangxiu += 1.0
        if i.fill.fgColor.rgb == sanxinfg:
            jiejia += 1.0
# 汇总每行的+n 加班
    if len(i.value)>2 and i.value[1] == '+' :
        ban = i.value[2:]
        if ban=='夜':# 显示夜班时直接改为8小时
            ban = 8.0
            if i.value[0] == gou:
                chuqin += 1# 这里加入公休日夜间的加班费情况
                if i.fill.fgColor.rgb == shuangxiufg:
                    shuangxiu += 1+ban/8.0
                if i.fill.fgColor.rgb == sanxinfg:
                    jiejia += 1+ban/8.0
                    # print ("jiejia:",jiejia)
                else:
                    jiaban += float(ban)# 是浮点数不是整数
                    print("chuqin:",chuqin)
            # print (n
# 汇总每行的G 公休
    if i.value == g:
        gongxiu += 1.0
    if len(i.value)>2 and (i.value[0:2] == 'G/'or i.value[1:3] == '/G' ):
        gongxiu += 0.5
# 汇总每行的B 补休
    if i.value == b:
        buxiu += 1.0
    if len(i.value)>2 and (i.value[0:2] == 'B/'or i.value[1:3] == '/B' ):
        buxiu += 0.5
# 汇总每行的△ 事假
    if i.value == shi:
        shijia += 1.0
    if len(i.value)>2 and (i.value[0:2] == (shi,'/')or i.value[1:3] == ('/',shi) ):
        shijia += 0.5#  这里有问题
# 汇总每行的H 婚假
    if i.value == h:
        hunjia += 1.0
    if len(i.value)>2 and (i.value[0:2] == 'H/'or i.value[1:3] == '/H' ):
        hunjia += 0.5
# 汇总每行的S 丧假
    if i.value == sang:
        sangjia += 1.0
    if len(i.value)>2 and (i.value[0:2] == 'S/'or i.value[1:3] == '/S' ):
        sangjia += 0.5
# 汇总每行的....
# print ('出差:',chuchai,end=' ')
# 每个月能拿加班费的时间不能大于4.5天
if shuangxiu+jiejia > 4.5:
   shuangxiu = shuangxiu - (4.5-jiejia)
   liangxin = 4.5-jiejia
   buxiu = buxiu + liangxin 
print("加班：",jiaban,end='')
print ('出勤：',chuqin,end='')
print ('公休：',gongxiu,end='')
print ('补休：',buxiu,end='')
print ('事假：',shijia,end='')
# print ('事假：',shi,'/')
print ('婚假：',hunjia,end='')
print ('丧假：',sangjia,end='')
print ('双休加班：',liangxin,end='')
print ('节假日加班：',jiejia,end='')
# print ('颜色：',list_rows[2][0].fill.fgColor.rgb)


