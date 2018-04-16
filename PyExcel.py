import xlsxwriter

# 创建文件薄
file_name = "first.xlsx"
workbook = xlsxwriter.Workbook(file_name)

# 创建工作表
worksheet = workbook.add_worksheet("py_sheet")

#写数据,行、列、内容
worksheet.write(0,0,"hello")
worksheet.write("A2","世界")
worksheet.write_row(2,1,[1,2,3,4,5,6])
worksheet.write_column("D2",['a','b','c'])

# 多格式数据写入（url,图片，图表，公式等）

# 插入一个url
worksheet.write_url(0,1,"https://www.baidu.com")

# 插入一个图片
worksheet.insert_image("E4","power.png")
# 插入一个公式
worksheet.write_formula("A4","=SUM(1,2,3)")
# 插入图表




# 关闭
workbook.close()
