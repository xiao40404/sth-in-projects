import os
import numpy as np
import xlrd
import xlwt

filepath = "F:\\建筑碳排放\\txts"
pathDir =  os.listdir(filepath)

workbook = xlrd.open_workbook(r'F:\建筑碳排放\ed_area.xls')
sheet1 = workbook.sheet_by_index(0)

file = xlwt.Workbook()
table = file.add_sheet('zb')
t = 1
#写入字段名称
l = [0,'平均高度','加权平均高度','高度变异系数','建筑总体积','平均建筑体积','分布均匀度指数','平均覆盖率','容积率','空间拥挤度','建筑体形系数']
for i in range(1,11):
    table.write(0,i,l[i])

for allDir in pathDir:
    filename = os.path.join('%s%s%s' % (filepath,'\\', allDir))
    #filename = filepath + '\\' + allDir
    name = allDir.split("_")

    # .decode('gbk')解决中文显示乱码问题
    print(name[0])

    #读入跳过标题首行
    array0 = np.loadtxt(filename, delimiter=",", skiprows=1)
    if array0.all() == 0:
        print('NB')
    else:
        print("KK")
    #去掉高度为0的建筑物数据
    d1 = np.argwhere(array0[:,3] == 0)
    array1 = np.delete(array0, d1, axis=0)
    d2 = np.argwhere(array1[:,4] == 0)
    array = np.delete(array1, d2, axis=0)
    if array.all() == 0:
        print('NB')
    else:
        print("KK")
