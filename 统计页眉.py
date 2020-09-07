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

#创建一个空的dataframe
df = pd.DataFrame(columns=['文件名', '左页眉', '中页眉', '右页眉'])
df=df.set_index(['文件名'])
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
                
            
            for sht in wb.sheets:
                LeftHeader=sht.api.PageSetup.LeftHeader
                CenterHeader=sht.api.PageSetup.CenterHeader
                RightHeader=sht.api.PageSetup.RightHeader
            
                if LeftHeader !="":
                    print("正在处理:",f)
                    #修改页眉
                    #sht.api.PageSetup.LeftHeader="\nWHTG-QR-8.4-07&G"
                    #sht.api.PageSetup.RightHeader="\n\n页码：&P/&N"
                    df.loc[f,:]=[LeftHeader,CenterHeader,RightHeader]
                    app.books[f].save()
            wb.close()
df.to_excel("页眉统计.xls",encoding='utf_8')

app.kill()
                
