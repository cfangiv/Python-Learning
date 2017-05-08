#---------------------------------
#Author: Wei CAO
#规则文件与脚本文件放于同一目录下，命名为rules.csv。
#待分拣文件放于脚本文件目录下子目录data中。
#Script Version: 1.0
#Python Version: 3.6
#Date: 2017.04.26
#Update:
#---------------------------------
import xlrd
import xlwt
import csv
import os
import sys
# from xlutils.copy import copy

provinces = ['安徽','北京','福建','甘肃','广东','广西','贵州','海南',
            '河北','河南','黑龙江','湖北','湖南','吉林','江苏','江西',
            '辽宁','内蒙古','宁夏','青海','山东','山西','陕西','上海',
            '四川','天津','西藏','新疆','云南','浙江','重庆','政企','物联网',
            '咪咕','互联网','卓望','在线','其它']
data_Dict = {}
summary = {}
log = []
folder_Address = ''
original_File = ''

#Get the path of script
def cur_file_dir():
    path = sys.path[0]
    if os.path.isdir(path):
        return path
    elif os.path.isfile(path):
        return os.path.dirname(path)
folder_Address = cur_file_dir()
print("Got current path:"+cur_file_dir())
if not os.path.exists(folder_Address+"\\result"):
    os.makedirs(cur_file_dir()+"\\result")

#folder_Address = r'C:\Users\Wei CAO\Desktop\sort'

def append_Dict(dict,province,item):
    if province in dict:
        dict[province].append(item)
    else:
        dict[province] = []
        dict[province].append(item)

with open(folder_Address + r'\rules.csv','r',encoding='utf-8') as r:
    reader = csv.reader(r)
    rows = [row for row in reader]
    number_row = [row[0] for row in rows]
    province_row = [row[1] for row in rows]
    print("Have got the rules..")

#Got the file to be sorted
original_File = os.listdir(folder_Address+r'\data\\')[0]
print("Found the original file: "+folder_Address+r'\data\\'+original_File)
workbook = xlrd.open_workbook(folder_Address+r'\data\\'+original_File)
#workbook = xlrd.open_workbook(folder_Address+r'\data\original.xlsx')
original_Sheet = workbook.sheet_by_index(0)
first_row = original_Sheet.row_values(0)
print("Reading the original excel..")

#Achieve the sorting rules
for row_number in range(1,original_Sheet.nrows):
    row = original_Sheet.row_values(row_number)
    for i in number_row:
        if row[4].startswith(i):
            belongs = province_row[number_row.index(i)]
            if belongs == '举报省':
                belongs = row[10]
            break
    else:
        if row[4].startswith('10658'):
            belongs = row[10]
        else:
            belongs = '其它'
    #print(belongs)
    append_Dict(data_Dict,belongs,row)
    log.append("ID: %s NUMBER: %s is inserted into " % (row[0],row[4]) +belongs+ '.')
 #   print("ID: %s NUMBER: %s is inserted into " % (row[0],row[4]) +belongs+ '.')

def addline(sheet,row,linelist):
    for i in range(len(linelist)):
        sheet.write(row,i,linelist[i])

#write rusult dic into excel
print("Have finished the sorting of ",end="")
for prov in provinces:
    f = xlwt.Workbook()
    sheet1 = f.add_sheet(u'sheet1', cell_overwrite_ok=True)
    for i in range(len(first_row)):
        sheet1.write(0,i,first_row[i])
        sheet1.col(i).width = 0x0d00 * 2
    if prov in data_Dict:
        for i in range(len(data_Dict[prov])):
             addline(sheet1,i+1,data_Dict[prov][i])
 #       print(prov + " 添加 %s 条" % len(data_Dict[prov]))
        append_Dict(summary,prov,len(data_Dict[prov]))
    else:
        append_Dict(summary,prov,0)
    f.save(folder_Address + r'\result\\' + prov + '.xls')
    print(prov+" ",end="")
print('')

#create a statistic excel
g = xlwt.Workbook()
sheet1 = g.add_sheet(u'sheet1', cell_overwrite_ok=True)
sheet1.write(0,0,"省份/公司")
sheet1.write(0,1,"统计数")
for i in range(len(provinces)):
    sheet1.write(i+1,0,provinces[i])
    sheet1.write(i+1,1,summary[provinces[i]][0])
sheet1.write(len(provinces)+1,0,"总计")
sheet1.write(len(provinces)+1,1,xlwt.Formula('SUM(B2:B%s)' % str(len(provinces)+1)))
g.save(folder_Address + r'\result\统计数据.xls')
print("Created the statistic excel..")

#save the sorting log
with open(folder_Address+r'\result\log.txt', 'w') as h:
    for i in range(len(log)):
        h.write(log[i]+'\n')
print("Created the log file..")
print("--------------Finish--------------")
#print(summary)
#print(log)

