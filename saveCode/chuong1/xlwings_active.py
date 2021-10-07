import os
import xlwings as xw

# Khởi tạo file excel bằng modul App
app = xw.App(visible= False, add_book= False) # code này gán modul App với cái tên là app

app.display_alerts=False

path_folder= "D:\ThanhVu\code\python\MyBook\Python_Excel\save_Datas\\"

for num in range(1):

    path = os.path.normpath(path_folder + str(num+1) + ".xlsx")
    print(path)
    wb = app.books()
    wb.save(path)

for app in xw.App():

    app.quit()
    

    

    
    
    

