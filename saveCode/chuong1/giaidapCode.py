import xlwings as xw
import os

# Khởi tạo file excel bằng modul App
app = xw.App(visible= False, add_book= False) # code này gán modul App với cái tên là app

# Dòng này sẽ Tắt các thông báo excel (như update,... xảy ra 1 số file excel)
app.display_alerts=False

""" Giải thích code vòng lặp for bên dưới:
Cho biến book chạy trong phạm vi từ 0->2
Lần chạy đầu tiên book = 0, thêm vào 1 workbook và lưu nó vào tại folder chứa file code.
    + Lần đầu tiên do book =0, nên lần đầu tiên tên file sẽ = book+1, tức là 0+1=1, đổi sang bằng chữ là str(book+1) và thêm đuôi file sẽ là "1.xlsx"     
    + Lần thứ 2, book = 1 tương tự file sẽ là "2.xlsx"
    ->Kết thúc vòng lặp.
"""
for book in range(2):
    book = app.books.add().save(str(book+1) + ".xlsx")

# Tắt tất cả các apps Excel đang hiển thị, kể cả file mở = tay (ko dùng code)
for app in xw.apps:
    app.quit()






# Cho biến n là số đếm từ 0->10 
for n in range(0, 10):
  print("Số n được đếm là:", n)
     

    

    
    
    

