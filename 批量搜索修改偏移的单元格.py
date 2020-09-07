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
            file_path=os.path.join(root, f)
            app.books.open(file_path)
            for sht in app.books[f].sheets:
                rng=sht.api.UsedRange.Find("外观")
                if rng==None:
                    pass
                else:
                    #保存第一个搜索结果的地址，防止无限循环
                    firstaddress = rng.Address
                    while True:
                        #目标单元格距离搜索单元格的偏移地址
                        obj_address=sht[rng.Address].offset(row_offset=0,column_offset=21).address
                        #偏移单元格中的内容
                        print(sht[rng.Address].value,sht[obj_address].value,sht,obj_address)

                        #设置替换值
                        sht[obj_address].value="替换文本"
                        #设置字体颜色为紫色
                        sht[obj_address].api.Font.Color=-6279056
                        #缩小字体适应单元格
                        sht[obj_address].api.ShrinkToFit = True
                        #调整字体大小
                        #sht[rng.Address].api.Font.size=8
                        #设置自动换行
                        sht[obj_address].api.WrapText = True
                        rng = sht.api.UsedRange.FindNext(rng)
                        if rng==None:
                            app.books[f].save()
                            break
                        else:
                            if rng.Address == firstaddress:
                                app.books[f].save()
                                break
            app.books[f].close()


app.kill()
                
