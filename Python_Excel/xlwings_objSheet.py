import xlwings as xw

# Khởi tạo file excel bằng modul App
app = xw.App(visible=True, add_book=False) # code này gán modul App với cái tên là app

# Tạo ra file excel mới gán tên là wb_new 
wb = app.books.open('movies_1.xls')  #('movies_1.xls')
# wb_new = app.books.add()





# lưu file mới đến thư mục "D:\ThanhVu\code\python\Data_Science\" 
wb.save(r'D:\ThanhVu\code\python\MyBook\Python_Excel\objSheet.xls')

# Đoạn code này đóng file excel mới toanh đó
wb.close()

# Đoạn code này tắt đối tượng app 
app.quit()


 
    

    
    
    

