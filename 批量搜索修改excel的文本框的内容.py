# -*- coding: UTF-8 -*-
import os
import sys
import shutil
import re
import xlwings as xls
import pprint

if getattr(sys, 'frozen', False):
    bundle_dir = sys._MEIPASS
else:
    bundle_dir = os.path.dirname(os.path.abspath(__file__))

os.chdir(bundle_dir)





app = xls.App(visible=False,add_book=True)
print("遍历指定目录下所有的文件和文件夹，包括子目录内的")
list_dirs = os.walk(sys.path[0])
for root, dirs, files in list_dirs:
    for f in files:
        # 分离文件名与扩展名，仅显示txt后缀的文件
        if os.path.splitext(f)[1]=='.xlsx'or os.path.splitext(f)[1]=='.xls':
            print("正在处理：",f)
            file_path=os.path.join(root, f)
            app.books.open(file_path)
            for sht in app.books[f].sheets:
                for shp in sht.shapes:
                    if shp.type=="text_box":
                        text=shp.api.TextFrame2.TextRange.Text
                        #搜索的目标字符串，如果包含子字符串返回开始的索引值，否则返回-1。
                        result=text.find("测试公差")
                        if result!=-1:
                            print(text)

            #app.books[f].save()
            app.books[f].close()


app.kill()
                
