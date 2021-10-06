import xlwings as xw
from xlwings.main import Sheets

# Đoạn code này tắt các file Excel do trong quá trình code bị lỗi mà bạn quên tắt file
for app in xw.apps:
    app.quit()

# Khởi tạo file excel bằng modul App, chọn visible = False: Ko mở file
app = xw.App(visible = True, add_book = True) # code này gán modul App với cái tên là app

# Dòng này sẽ Tắt các thông báo excel (như update,... xảy ra 1 số file excel)
app.display_alerts = False

# Đặt tên cho workbook vừa mới tạo là: wb, 
wb = app.books[0] # Có thể dùng: wb = app.books['Book1']
# books[0]: Nghĩa là book đầu tiên

# Gán biến sh1 cho sheet đầu tiên
sh1= wb.sheets[0] # In ra tên sheet

"Phần code tìm hiểu đối tượng Range"
columns = ["Toán","Lý", "Hóa", "Sinh", "Ngoại Ngữ"]
rows = [[1],[2],[3],[4],[5]]

datas = [[1,5,8],
         [2,7,12],
         [3,9,5],
         [4, 5.9,8.2],
         [5.8, 105,102.5]]

# Cách thực hiện thêm giá trị, dữ liệu vào bảng tính
sh1.range('A1').value = columns # Điền số 1 vào ô A1
sh1.range('A2').value = rows
sh1.range("C4").value = datas 

# # Đoạn code này đóng file excel 
# wb.close()

# # Đoạn code này tắt đối tượng app 
# for app in xw.apps:
#     app.quit()

 
    

    
    
    

