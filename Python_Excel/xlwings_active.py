import xlwings as xw

# Mở file có tên 'movies_1.xls' gán cái biến tên là wb1
wb1 = xw.Book('movies_1.xls')

# Mở file có tên 'movies_2.xls' gán cái biến tên là wb2
wb2 = xw.Book('movies_2.xls')
# wb2 = xw.books.active.name

# Mở file có tên 'movies_3.xls' gán cái biến tên là wb3
wb3 = xw.Book('movies_3.xls')

# Tên workbook hiện hành
print("Tên workbook hiện hành:", xw.books.active.name)

# Đóng 3 workbook lại
wb1.close()
wb2.close()
wb3.close()
xw.App().quit()
