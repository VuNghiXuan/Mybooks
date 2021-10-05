import os
import glob
import csv
import pandas as pd
import numpy as np

"Methol 1:";
# Tạo ra 1 list danh sách chứa các file trong folder theo đuôi mở rộng
def fun_get_list_file(path_folder, expention_file):
    #Chọn đuôi file theo chỉ định expention_file : '.xlsx' or '.xlsx'
    list_file = []
    for file in os.listdir(path_folder):
        if file.endswith(expention_file):
            list_file.append(file)
    return list_file



def get_list_files(folder, expention_file):
    cfiles = []
    for root, dirs, files in os.walk(folder): #src
        for file in files:
            if file.endswith(expention_file): #'.c'
                cfiles.append(os.path.join(root, file))
        return cfiles

list_files_1 = fun_get_list_file("./read_datas_txt", ".txt")
list_files_2 = get_list_files("./read_datas_txt", ".txt")


datas = []

"get header từ 1-->2"    
header1 = pd.read_csv(list_files_2[0],sep='|', encoding = "ISO-8859-1",skiprows=0,skipfooter =1 , engine = 'python',header=None,index_col=None)
# engine = 'python', sử dụng python phân tích nhiều tính năng
hd = pd.DataFrame(header1, columns =[header1.iloc[1:2,:]])

# .iloc[:, 1:column_count].values
datas.append(hd)
# print(datas)


for file in list_files_2:    
    "Vì dữ liệu vào excel bị lỗi nên bỏ: skiprows=3: Bỏ dòng 1 -3; skipfooter =1: bỏ 1 cuối"    
    df = pd.read_csv(file, sep='|', encoding = "ISO-8859-1", skiprows=3, skipfooter =1, engine = 'python', header=None, index_col=None) #"ISO-8859-1"
    # engine = 'python', sử dụng python phân tích nhiều tính năng    
    datas.append(df)

df_concat = pd.concat(datas, axis=0, ignore_index = False) 
"""#ignore_index=True :  Đúng 'bỏ qua', nghĩa là không thẳng hàng trên trục kết hợp. nó chỉ đơn giản là dán chúng lại với nhau theo thứ tự mà chúng được truyền, sau đó chỉ định lại một phạm vi cho chỉ mục thực tế (ví dụ: phạm vi (len (chỉ mục)))) để tạo ra sự khác biệt giữa việc kết hợp trên các chỉ mục không chồng chéo (giả sử trục = 1 trong ví dụ) , là với ignore_index = False (mặc định), bạn nhận được kết hợp của các chỉ mục và với ignore_index = True bạn nhận được một phạm vi.""";

df_concat.to_excel("output.xlsx") #, encoding="utf-8"
