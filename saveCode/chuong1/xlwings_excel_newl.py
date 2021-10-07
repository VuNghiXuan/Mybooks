import xlwings as xw

# Khởi tạo file excel bằng modul App
app = xw.App(visible=True, add_book=False) # code này gán modul App với cái tên là app

# Tạo ra file excel mới gán tên là wb_new 
wb_new = app.books.add()
""" *Lưu ý tại đoạn code thứ 4, có 2 tham số: 
    + "visible=True" : Nghĩa là cho phép hiện bảng tính mới vừa tạo
    + "add_book=False": Nghĩa là không tạo thêm 01 bảng tính mới nữa
    --> Nếu bỏ tham số "add_book" hoặc cho nó = True, nó sẽ sinh ra 2 workbooks (Trong đó tại dòng thứ 7: 01 workbook đầu tiên sẽ sinh ra do dòng lệnh app.books. Khi bạn thêm add vào nữa sẽ sinh thêm 1 workbook nữa)     
"""
# lưu file mới đến thư mục "D:\ThanhVu\code\python\Data_Science\" 
wb_new.save(r'D:\ThanhVu\code\python\MyBook\Python_Excel\moitoanh.xls')

# Đoạn code này đóng file excel mới toanh đó
wb_new.close()

# Đoạn code này tắt đối tượng app 
app.quit()


 
    

    
    
    

