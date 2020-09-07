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
app.books[0].close()
print("遍历指定目录下所有的文件和文件夹，包括子目录内的")
list_dirs = os.walk(sys.path[0])
for root, dirs, files in list_dirs:
    for f in files:
        # 分离文件名与扩展名，仅显示txt后缀的文件
        if os.path.splitext(f)[1]=='.xls':
            file_path=os.path.join(root, f)
            app.books.open(file_path)
            obj_file=os.path.splitext(f)[0]+".xlsx"
            obj_path=os.path.join(root, obj_file)
            print("保存到:",obj_path)
            #将xls转换为xlsx格式
            app.books[f].api.SaveAs(obj_path,51)
            app.books[obj_file].close()
            #删除旧版
            os.remove(os.path.join(root,f))


app.kill()
                
