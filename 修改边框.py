# -*- coding: UTF-8 -*-
import os
import sys
import shutil
import re
import xlwings as xls
import pprint
import pandas as pd

#将该文件与基准书放置到同一个目录下执行，可以统计本目录及子目录下的所有基准书的页数


if getattr(sys, 'frozen', False):
    bundle_dir = sys._MEIPASS
else:
    bundle_dir = os.path.dirname(os.path.abspath(__file__))

os.chdir(bundle_dir)


app = xls.App(visible=True,add_book=True)
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
                
            
            for sht in wb.sheets:
                LeftHeader=sht.api.PageSetup.LeftHeader
                CenterHeader=sht.api.PageSetup.CenterHeader
                RightHeader=sht.api.PageSetup.RightHeader
            
                if LeftHeader =="\nWHTG-QR-8.4-07&G":
                    print("正在处理:",f)
                    rng=sht["A1:BE500"]

                    # 底部边框  LineStyle = 1 直线
                    rng.api.Borders(9).LineStyle = 1
                    rng.api.Borders(9).Weight = 2       # 设置边框粗细。

                    # Borders(7) 左边框，LineStyle = 2 虚线。
                    rng.api.Borders(7).LineStyle = 1
                    rng.api.Borders(7).Weight = 2

                    # Borders(8) 顶部框，LineStyle = 5 双点划线。
                    rng.api.Borders(8).LineStyle = 1
                    rng.api.Borders(8).Weight = 2

                    # Borders(10) 右边框，LineStyle = 4 点划线。
                    rng.api.Borders(10).LineStyle = 1
                    rng.api.Borders(10).Weight = 2
                    
                    # # Borders(11) 内部垂直边线。
                    rng.api.Borders(11).LineStyle = 1
                    rng.api.Borders(11).Weight = 2
 
                    # # Borders(12) 内部水平边线。
                    rng.api.Borders(12).LineStyle = 1
                    rng.api.Borders(12).Weight = 2
                    

                    #app.books[f].save()
                
