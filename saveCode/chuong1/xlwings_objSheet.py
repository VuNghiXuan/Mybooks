import xlwings as xw

# Khởi tạo file excel bằng modul App
app = xw.App(visible=True, add_book=False) # code này gán modul App với cái tên là app

# Đọc file excel có tên là "Covid_VN.xlsx" từ folder "read_Data"
wb = app.books.open(r'D:\ThanhVu\code\python\MyBook\Python_Excel\read_Datas\Covid_VN.xlsx')  #('movies_1.xls')

" -----> Đoạn này là bắt đầu tìm hiểu đối tượng Sheet"

# Đặt tên cho 1 danh sách các tên Sheets là "sh_Names": 
sh_Names = wb.sheets

# In ra terminal đối tượng sh_Names
print ("Đối tượng sh_Names là:", sh_Names)

# Dùng vòng lặp for đi qua từng sheet và in ra tên sheet như sau:
for i_name in range(len(sh_Names)): # Nhớ có các từ khóa "for"; "in" và dấu ":"
    print(f"Tên sheet thứ {i_name+1} là:", sh_Names[i_name].name) # Nhớ dùng phím tab phía đầu hàm print để thụt vào

" Kết thúc tìm hiểu đối tượng Sheet <---------"

# lưu file "ketqua_Covid.xlsx" tại folder "save_Data"
wb.save(r'D:\ThanhVu\code\python\MyBook\Python_Excel\save_Datas\ketqua_Covid.xlsx')

# Đoạn code này đóng file excel 
wb.close()

# Đoạn code này tắt đối tượng app 
for app in xw.apps:
    app.quit()

 
    

    
    
    

