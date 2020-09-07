
# coding: utf-8

# In[291]:

#!/usr/bin/python

import xlwings as xls
xls.__version__


# # 打开Excel程序，默认设置：程序可见，并新建一个工作簿

# In[292]:

app = xls.App(visible=True,add_book=True)


# # 再新建一个工作簿

# In[293]:

wb = app.books.add()


# # app.books是可迭代对象，显示目前所有的工作簿

# In[294]:

print(app.books)


# In[295]:

app.books[0]


# In[296]:

app.books["工作簿1"]


# # 关闭工作薄

# In[297]:

app.books["工作簿2"].close()
app.books[0].close()


# # 打开一个现存的工作薄

# In[298]:

wb=app.books.open('基准书1.xlsx')


# # 显示工作薄内所有sheet

# In[299]:

wb.sheets
for sht in wb.sheets:
    print(sht)


# # 显示各个sheet的打印页数

# In[300]:

for sht in wb.sheets:
    print(sht.api.PageSetup.Pages.Count)


# # 显示页眉

# In[301]:

sht=wb.sheets[0]
LeftHeader=sht.api.PageSetup.LeftHeader
CenterHeader=sht.api.PageSetup.CenterHeader
RightHeader=sht.api.PageSetup.RightHeader
print(LeftHeader,CenterHeader,RightHeader)


# # 获取单元格内文本

# In[302]:

print("文本内容：",wb.sheets["正本"]['AD4'].value)
print("文本内容：",wb.sheets["正本"]['$AD$4'].value)
#背景色
print("背景颜色：",wb.sheets["正本"]['AD4'].color)
print("行高：",wb.sheets["正本"]['AD4'].row_height)
print("列宽：",wb.sheets["正本"]['AD4'].column_width)
print("字体:",wb.sheets["正本"]['AD4'].api.Font.Name)
print("字体大小:",wb.sheets["正本"]['AD4'].api.Font.size)
print("字体颜色：",wb.sheets["正本"]['AD4'].api.Font.Color)
print("是否缩小字体适应单元格：",wb.sheets["正本"]['AD4'].api.ShrinkToFit)
print("是否自动换行：",wb.sheets["正本"]['AD4'].api.WrapText)

#设置颜色为(255,0,0)
wb.sheets["正本"]['AD4'].api.Font.Color=0x0000ff

# -4108 水平居中。 -4131 靠左，-4152 靠右。
wb.sheets["正本"]['AD4'].api.HorizontalAlignment = -4108

# -4108 垂直居中（默认）。 -4160 靠上，-4107 靠下， -4130 自动换行对齐。
wb.sheets["正本"]['AD4'].api.VerticalAlignment = -4130


# In[303]:

wb.sheets["正本"]['A4:X4'].value


# # 得到指定单元格所在的合并单元格的Range

# In[304]:

#判断range范围内是否有合并单元格
print(wb.sheets["正本"]['A4:X4'].api.MergeCells)

#合并单元格的range
wb.sheets["正本"]['C4'].api.MergeArea.Address


# # 偏移单元格的地址

# In[305]:

sht["$C$4"].offset(row_offset=0,column_offset=21).address


# # 搜索文本
# 当查找到指定查找区域的末尾时，本方法将环绕至区域的开始继续搜索。发生环绕后，为停止查找，可保存第一次找到的单元格地址，然后测试下一个查找到的单元格地址是否与其相同。

# In[306]:

rng=wb.sheets["正本"].api.UsedRange.Find("尺寸")
if rng==None:
    print("未搜索到该字符串")
else:
    #保存第一个搜索结果的地址，防止无限循环
    firstaddress = rng.Address
    while True:
        print(rng.Address)
        rng = wb.sheets["正本"].api.UsedRange.FindNext(rng)
        if rng==None:
            break
        else:
            if rng.Address == firstaddress:
                break


# # 获取表格中的形状

# In[307]:

wb.sheets["正本"].shapes


# # 文本框

# In[308]:

shape0=wb.sheets["正本"].shapes[0]
print("name:",shape0.name)
#读取shape的value有以下几种方法
print("value:",shape0.api.OLEFormat.Object.Text)
print("value:",shape0.api.TextFrame2.TextRange.Text)
#print("value:",shape0.api.TextFrame.Characters.Text)
print("type:",shape0.type)
print("heigh:",shape0.height)
print("width:",shape0.width)
print("允许文本溢出形状:",shape0.api.TextFrame.VerticalOverflow)


# # 复制一个sheet（此处用到了vba的函数，api后面可以使用任意vba函数）

# In[309]:

wb.sheets["1"].api.Copy()
wb2=app.books[1]
sht2=wb2.sheets[0]


# # 更改sheet的名称（此处用到了vba的函数，api后面可以使用任意vba函数）

# In[310]:

sht2.api.Name="QA2"


# # 图片

# In[311]:

shape1=sht2.shapes[1]
print("name:",shape1.name)
print("type:",shape1.type)
print("heigh:",shape1.height)
print("width:",shape1.width)
print("距离左边的距离:",shape1.left)
print("距离顶端的距离:",shape1.left)

#添加图片
sht2.pictures.add(r'C:\Windows\Web\Wallpaper\Windows\img0.jpg',left=500, top=300, width=10, height=10)

#删除图片
shape1.delete()


# # 修改单元格内文本

# In[312]:

sht2['J11'].value="支架"


# # 进行一组数据的赋值时默认是按行进行赋值

# In[313]:

sht2['A24'].value=[1, 2, 3,4]


# # 按列进行赋值需要添加transpose参数

# In[314]:

sht2['A25'].options(transpose=True).value=[1, 2, 3,4]


# # 保存工作簿

# In[315]:

wb2.save("复制的工作表.xlsx")


# # 转换文件格式
# | 名称                              | 值   | 说明                   | 扩展名 |
# | :-------------------------------- | :--- | :--------------------- | :----- |
# | **xlCSV**                         | 6    | CSV                    | *.csv  |
# | **xlExcel8**                      | 56   | Excel 97-2003 工作簿   | *.xls  |
# | **xlOpenXMLWorkbook**             | 51   | Open XML 工作簿        | *.xlsx |
# | **xlOpenXMLWorkbookMacroEnabled** | 52   | 启用 Open XML 工作簿宏 | *.xlsm |
# | **xlTextMac**                     | 19   | Macintosh 文本         | *.txt  |
# | **xlTextMSDOS**                   | 21   | MSDOS 文本             | *.txt  |
# | **xlTextWindows**                 | 20   | Windows 文本           | *.txt  |
# | **xlXMLSpreadsheet**              | 46   | XML 电子表格           | *.xml  |

# In[316]:

import os
import sys
print(sys.path[0])
#os.chdir(sys.path[0])
wb2.api.SaveAs( "复制的工作表",46)


# # 关闭excel

# In[317]:

app.books[1].close()
app.kill()

