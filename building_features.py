import xlrd
import xlwt
import math

workbook = xlrd.open_workbook(r'F:\LCZ\Data\building_label.xls')
sheet1 = workbook.sheet_by_index(0)

fid = sheet1.col_values(0)
floor = sheet1.col_values(2)
area = sheet1.col_values(3)
label = sheet1.col_values(1)

file = xlwt.Workbook()
table = file.add_sheet('bl')
t=0

for i in range(245):
    n = 0
    area_sum = 0
    floor_area = 0
    h_hm = 0
    active = False
    la = 0

    for j in range(1,sheet1.nrows):
        if fid[j] == i:
            p = floor[j]*area[j]
            floor_area = floor_area+p
            area_sum = area_sum + area[j]
            n = n+1
            id = i
            u =j
            active =True

    while active:
        h_mean = floor_area/area_sum
        for k in range(u-n+1,u+1):
            q = (float(floor[k])-h_mean)**2
            h_hm = h_hm+q
        h_sd = math.sqrt(h_hm/n)
        active = False

    if i == id:
        table.write(t,0,id)
        table.write(t,1,h_mean)
        table.write(t,2,h_sd)
        table.write(t,3,area_sum)
        t = t+1

    if i != id:
        table.write(t, 0, 0)
        table.write(t, 1, 0)
        table.write(t, 2, 0)
        table.write(t, 3, 0)
        t = t + 1
file.save(r'F:\LCZ\Data\building.xls')
