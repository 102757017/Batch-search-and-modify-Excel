# -*- coding: UTF-8 -*-
import os
import sys
import shutil
import re
import xlwings as xls
import pprint


#将该文件与基准书放置到同一个目录下执行，可以统计本目录及子目录下的所有基准书的页数


if getattr(sys, 'frozen', False):
    bundle_dir = sys._MEIPASS
else:
    bundle_dir = os.path.dirname(os.path.abspath(__file__))

os.chdir(bundle_dir)


partlist = open('重要度统计.csv','w') # 覆盖模式
app = xls.App(visible=False,add_book=True)
print("遍历指定目录下所有的文件和文件夹，包括子目录内的")
list_dirs = os.walk(sys.path[0])
for root, dirs, files in list_dirs:
    for f in files:
        # 分离文件名与扩展名，仅显示txt后缀的文件
        if os.path.splitext(f)[1]=='.xlsx'or os.path.splitext(f)[1]=='.xls':
            file_path=os.path.join(root, f)
            wb=app.books.open(file_path)

            #获取首页
            sht=app.books[f].sheets[0]
            if sht.name!="000000":
                pass
            else:
                sht=app.books[f].sheets[1]
                
            #提取首页的零件号
            address=sht["J8"].api.MergeArea.Address
            text=sht[address].value[0][0]
            Part_No = re.findall(r'([\-0-9A-Z]{10,26})', text)

            #获取首页的零件重要度 
            address=sht["AP6"].api.MergeArea.Address
            imp=sht[address].value[0][0]

            #获取首页的机种
            address=sht["J6"].api.MergeArea.Address
            text=sht[address].value[0][0]
            modle=re.findall(r'(2[A-Z]{2})', text)
            #排序
            modle.sort(reverse=True)
            modle='/'.join(modle)

            #获取各个sheet的页数，再累加得到总页数
            page_num=0
            for sht in wb.sheets:
                page_num=page_num+sht.api.PageSetup.Pages.Count
            

            #写入csv
            for p in Part_No:
                print(p,imp,modle,page_num)
                partlist.write("{},{},{},{}\n".format(p,imp,modle,page_num))
                
            wb.close()
partlist.close()
app.kill()
                
