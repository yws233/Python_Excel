import xlsxwriter

# 创建文件薄
file_name = "first.xlsx"
workbook = xlsxwriter.Workbook(file_name)

# 关闭
workbook.close()
