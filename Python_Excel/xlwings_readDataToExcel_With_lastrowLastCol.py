import xlwings as xw
from xlwings.main import Sheets

# Đoạn code này tắt các file Excel do trong quá trình code bị lỗi mà bạn quên tắt file
for app in xw.apps:
    app.quit()

# Đọc dữ liệu từ bảng tính Excel
wb = xw.Book(r'D:\MyBook_PyExcel\xlwings\read_Datas\1.BangDiem.xlsx')
sh = wb.sheets[0]

# Sử dụng hàm có sẵn last_cell.row để tìm dòng cuối cùng của bảng tính
lr_table = sh.cells.last_cell.row
print('Dòng cuối cùng bảng tính là:', lr_table)

# Sử dụng hàm có sẵn last_cell.column để tìm cột cuối cùng của bảng tính
lc_table = sh.cells.last_cell.column
print('Cột cuối cùng bảng tính là:', lc_table)

" Thêm code để tìm dòng cuối, cột cuối có chứa dữ liệu";
lr_data = sh.range('A'+ str(lr_table)).end('up').row
print('Dòng cuối cùng có chứa dữ liệu là:', lr_data)

# lcol = sh.range(row_index, col).end("left").column
lc_data = sh.range(1, lc_table).end('left').column
print('Cột cuối cùng có chứa dữ liệu là:', lc_data)

# Cuối cùng ta có bảng dữ liệu cần thu thập như sau:
table_datas = sh.range((1,1), (lr_data,lc_data))
print("Bảng dữ liệu có địa chỉ là:", table_datas.address)

print("Giá trị data là:\n ",table_datas.value)

# lr = sh.range('A' + str(wb.sheets[0].cells.last_cell.row)).end('up').row
# Hàm tự dịnh nghĩa: Tìm dòng cuối dữ liệu
# def lastRow():
#     for ir 


 
    

    
    
    

