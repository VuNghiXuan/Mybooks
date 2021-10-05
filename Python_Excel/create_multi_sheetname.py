from numpy import tile
import pandas as pd
import os
import numpy as np
from pandas.core.frame import DataFrame
from pandas.core.indexes.base import Index

# Import modul ExcelWriter của Pandas
from pandas import ExcelWriter

# Kiểm tra sự tồn tại của file
isTrue = os.path.isfile('./saveExcel/1.xeploai.xlsx')
# Tức là file tồn tại

# Đọc file excel "1.xeploai.xlsx" trong folder saveExcel
if isTrue:
    data = pd.read_excel('./saveExcel/1.xeploai.xlsx')


# Lấy dòng header đưa vào 1 array numpy
header_data = np.array(data.keys()).reshape(1,9)

# Lấy dữ liệu của bảng ko bao gồm header
data = data.values

# Định nghĩa save nhiều sheets bằng hàm 
def save_xlsx(data, path):
    writer = ExcelWriter(path) # Tạo và ghi 1 file Excel
    for i in range(len(data)):
        dt_sh = np.vstack((header_data,data[i])) #Dòng này nối tiêu đề và từng dòng data
        dt_sh = DataFrame(dt_sh) #Chuyển đổi data lại thành 1 DataFram để xuất ra Excel
        dt_sh.to_excel(writer, data[i][1]) # Xuất ra file Excel
    writer.save()

path = './saveExcel/2.xeploai.xlsx' #,list_shs
save_xlsx(data, path)

