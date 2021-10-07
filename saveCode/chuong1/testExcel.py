import xlwings as xw

# Khởi tạo file excel bằng modul App
app = xw.App() # code này gán modul App với cái tên là app

# Tạo ra file excel mới gán tên là wb_new 
wb_new = app.books.add()

""" *Lưu ý: 
 + 'app.books.add()': dòng code này tạo ra 2 workbooks
 + Trong đó: workbook2 là workbook hiện hành (active)
"""

# lưu file mới đến "D:\ThanhVu\code\python\Data_Science\" + "file_moitoanh.xls"
wb_new.save(r'D:\MyHandBook_PyExcel\tem\moitoanh.xls')

# Đoạn code này đóng file excel mới toanh đó
wb_new.close()

# Đoạn code này tắt đối tượng app 
app.quit()

# visible=True, add_book=False


#'code_tinh_camay.xlsm'
 
    

    
    
    

