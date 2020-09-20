#!/usr/bin/env python
# coding: utf-8

# In[33]:


#!/usr/bin/python
# -*- coding: UTF-8 -*-
import xlwings as xls
xls.__version__


# # 打开Excel程序，默认设置：程序可见，并新建一个工作簿

# In[34]:


app = xls.App(visible=True,add_book=True)


# # 再新建一个工作簿

# In[35]:


wb = app.books.add()


# # app.books是可迭代对象，显示目前所有的工作簿

# In[36]:


print(app.books)


# In[37]:


app.books[0]


# In[38]:


app.books["工作簿1"]


# # 关闭工作薄

# In[39]:


app.books["工作簿2"].close()
app.books[0].close()


# # 打开一个现存的工作薄

# In[40]:


wb=app.books.open('基准书1.xlsx')


# # 显示工作薄内所有sheet

# In[41]:


wb.sheets
for sht in wb.sheets:
    print(sht)


# # 显示各个sheet的打印页数

# In[42]:


for sht in wb.sheets:
    print(sht.api.PageSetup.Pages.Count)


# # 获取分页符下面的行号

# In[43]:


for sht in wb.sheets:
    for pb in sht.api.HPageBreaks:
        print(sht,pb.Location.Row)


# # 显示页眉

# In[44]:


sht=wb.sheets[0]
LeftHeader=sht.api.PageSetup.LeftHeader
CenterHeader=sht.api.PageSetup.CenterHeader
RightHeader=sht.api.PageSetup.RightHeader
print(LeftHeader,CenterHeader,RightHeader)


# # 获取单元格内文本

# In[45]:


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


# In[46]:


wb.sheets["正本"]['A4:X4'].value


# # 得到指定单元格所在的合并单元格的Range

# In[47]:


#判断range范围内是否有合并单元格
print(wb.sheets["正本"]['A4:X4'].api.MergeCells)

#合并单元格的range
wb.sheets["正本"]['C4'].api.MergeArea.Address


# # 偏移单元格的地址

# In[48]:


sht["$C$4"].offset(row_offset=0,column_offset=21).address


# # 搜索文本
# 当查找到指定查找区域的末尾时，本方法将环绕至区域的开始继续搜索。发生环绕后，为停止查找，可保存第一次找到的单元格地址，然后测试下一个查找到的单元格地址是否与其相同。

# In[49]:


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


# # 删除行，列

# In[50]:


#删除行
wb.sheets["正本"]["34:34"].delete()
#删除列
wb.sheets["正本"]["BB:BB"].delete()


# # 修改边框
# LineStyle = 1 直线  
# LineStyle = 2 虚线  
# LineStyle = 4 点划线  
# LineStyle = 5 双点划线  
# Weight：设置边框线粗细  

# In[51]:


rng=wb.sheets["正本"]['A1:E10']
LineStyle = 5

# Borders(7) 左边框
rng.api.Borders(7).LineStyle = LineStyle
rng.api.Borders(7).Weight = 2

# Borders(8) 顶部框
rng.api.Borders(8).LineStyle = LineStyle
rng.api.Borders(8).Weight = 2

# 底边框
rng.api.Borders(9).LineStyle = LineStyle
rng.api.Borders(9).Weight = 2

# Borders(10) 右边框
rng.api.Borders(10).LineStyle = LineStyle
rng.api.Borders(10).Weight = 2
                    
# # Borders(11) 内部垂直边线。
rng.api.Borders(11).LineStyle = LineStyle
rng.api.Borders(11).Weight = 2
 
# # Borders(12) 内部水平边线。
rng.api.Borders(12).LineStyle = LineStyle
rng.api.Borders(12).Weight = 2


# # 获取表格中的形状

# In[52]:


wb.sheets["正本"].shapes


# # 文本框

# In[53]:


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


# # 自动填充
# Type：根据源范围的内容指定如何填充目标范围  
# 
# 
# | Name               | Value | Description                                                  |
# | :----------------- | :---- | :----------------------------------------------------------- |
# | **xlFillCopy**     | 1     | Copy the values and formats from the source range to the target range, repeating if necessary. |
# | **xlFillDays**     | 5     | Extend the names of the days of the week in the source range into the target range. Formats are copied from the source range to the target range, repeating if necessary. |
# | **xlFillDefault**  | 0     | Excel determines the values and formats used to fill the target range. |
# | **xlFillFormats**  | 3     | Copy only the formats from the source range to the target range, repeating if necessary. |
# | **xlFillMonths**   | 7     | Extend the names of the months in the source range into the target range. Formats are copied from the source range to the target range, repeating if necessary. |
# | **xlFillSeries**   | 2     | Extend the values in the source range into the target range as a series (for example, '1, 2' is extended as '3, 4, 5'). Formats are copied from the source range to the target range, repeating if necessary. |
# | **xlFillValues**   | 4     | Copy only the values from the source range to the target range, repeating if necessary. |
# | **xlFillWeekdays** | 6     | Extend the names of the days of the workweek in the source range into the target range. Formats are copied from the source range to the target range, repeating if necessary. |
# | **xlFillYears**    | 8     | Extend the years in the source range into the target range. Formats are copied from the source range to the target range, repeating if necessary. |
# | **xlGrowthTrend**  | 10    | Extend the numeric values from the source range into the target range, assuming that the relationships between the numbers in the source range are multiplicative (for example, '1, 2,' is extended as '4, 8, 16', assuming that each number is a result of multiplying the previous number by some value). Formats are copied from the source range to the target range, repeating if necessary. |
# | **xlLinearTrend**  | 9     | Extend the numeric values from the source range into the target range, assuming that the relationships between the numbers is additive (for example, '1, 2,' is extended as '3, 4, 5', assuming that each number is a result of adding some value to the previous number). Formats are copied from the source range to the target range, repeating if necessary. |
# | **xlFlashFill**    | 11    | Extend the values from the source range into the target range based on the detected pattern of previous user actions, repeating if necessary. |

# In[54]:


wb.sheets["正本"]['C16:K17'].api.AutoFill(Destination=wb.sheets["正本"]['C16:K33'].api,Type=0)


# # 选择性粘贴
# Paste：指定粘贴的类型
# 
# | 名称                                    | 值    | 描述                             |
# | :-------------------------------------- | :---- | :------------------------------- |
# | **xlPasteAll**                          | -4104 | 一切都会被粘贴。                 |
# | **xlPasteAllExceptBorders**             | 7     | 除边框外的所有内容都会粘贴。     |
# | **xlPasteAllMergingConditionalFormats** | 14    | 将粘贴所有内容，并合并条件格式。 |
# | **xlPasteAllUsingSourceTheme**          | 13    | 一切都将使用源主题进行粘贴。     |
# | **xlPasteColumnWidths**                 | 8     | 粘贴复制的列宽。                 |
# | **xlPasteComments**                     | -4144 | 注释已粘贴。                     |
# | **xlPasteFormats**                      | -4122 | 复制的源格式已粘贴。             |
# | **xlPasteFormulas**                     | -4123 | 粘贴公式。                       |
# | **xlPasteFormulasAndNumberFormats**     | 11    | 粘贴公式和数字格式。             |
# | **xlPasteValidation**                   | 6     | 验证粘贴。                       |
# | **xlPasteValues**                       | -4163 | 值已粘贴。                       |
# | **xlPasteValuesAndNumberFormats**       | 12    | 粘贴值和数字格式。               |
# 
# Operation：指定如何使用工作表上的目标单元格计算数字数据。  
# 
# | 名称                            | 值    | 描述                                   |
# | ------------------------------- | ----- | -------------------------------------- |
# | xlPasteSpecialOperationAdd      | 2     | 复制的数据将添加到目标单元格中的值。   |
# | xlPasteSpecialOperationDivide   | 5     | 复制的数据将在目标单元格中分割该值。   |
# | xlPasteSpecialOperationMultiply | -4    | 复制的数据将乘以目标单元格中的值。     |
# | xlPasteSpecialOperationNone     | -4142 | 粘贴操作将不进行任何计算。             |
# | xlPasteSpecialOperationSubtract | 3     | 复制的数据将从目标单元格中的值中减去。 |
# 
# SkipBlanks：跳过空单元格  
# 
# Transpose：转置  

# In[55]:


wb.sheets["正本"]['C4:K5'].api.Copy()
wb.sheets["正本"]['C6:K17'].api.PasteSpecial(Paste=-4122,Operation=-4142,SkipBlanks=False,Transpose=False)
#情况剪切板，如果不写这句代码会出现提示窗口:是否保存复制的内容到剪贴板,以便下次使用
app.api.CutCopyMode=False


# # 输入公式，相应单元格执行结果

# In[56]:


wb.sheets["正本"]["A1"].formula='=SUM(B6:B7)' 


# # 复制一个sheet（此处用到了vba的函数，api后面可以使用任意vba函数）

# In[57]:


wb.sheets["1"].api.Copy()
wb2=app.books[1]
sht2=wb2.sheets[0]


# # 更改sheet的名称（此处用到了vba的函数，api后面可以使用任意vba函数）

# In[58]:


sht2.api.Name="QA2"


# # 图片

# In[59]:


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

# In[60]:


sht2['J11'].value="支架"


# # 进行一组数据的赋值时默认是按行进行赋值

# In[61]:


sht2['A24'].value=[1, 2, 3,4]


# # 按列进行赋值需要添加transpose参数

# In[62]:


sht2['A25'].options(transpose=True).value=[1, 2, 3,4]


# # 保存工作簿

# In[63]:


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

# In[64]:


import os
import sys
print(sys.path[0])
#os.chdir(sys.path[0])
wb2.api.SaveAs( "复制的工作表",46)


# # 关闭excel

# In[65]:


app.books[1].close()
app.kill()

