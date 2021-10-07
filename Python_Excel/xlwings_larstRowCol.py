import xlwings as xw
from xlwings.main import Sheets

# Đoạn code này tắt các file Excel do trong quá trình code bị lỗi mà bạn quên tắt file
for app in xw.apps:
    app.quit()

# Đọc dữ liệu từ bảng tính Excel
wb = xw.Book(r'D:\MyBook_PyExcel\xlwings\read_Datas\1.BangDiem.xlsx')
sh = wb.sheets[0]

# Cách 2: Đọc dữ liệu từ ô A1 --> sử dụng phương thức expand mở rộng phạm vi 
datas_2 = sh.range("A1").expand().value

print("datas_2\n", datas_2)
 
    

    
    
    

