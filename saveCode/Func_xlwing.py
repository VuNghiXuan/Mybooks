""" 
Tổng hợp một số hàm thông dụng khi tiếp cận thư viện xlwings.
    *Bao gồm:
    - Lấy tất cả các file trong folder theo đuôi '.xls', '.xlsx', '.xlsm'
    - Đọc hoặc mở file excel
    - Lấy danh sách tên sheets
    - Tìm dòng, cột cuối của dữ liệu trong sheet
    - Đổi tên cột ra số và ngược lại trong sheet

"""
import xlwings as xw
import os
import operator # thư viện toán tử

# Tạo ra 1 list danh sách chứa các file trong folder theo đuôi mở rộng
def fun_get_list_file(path_folder, expention_file):
    #Chọn đuôi file theo chỉ định expention_file : '.xlsx' or '.xlsx'
    list_file = []
    for file in os.listdir(path_folder):
        if file.endswith(expention_file):
            list_file.append(file)
    return list_file

# Đọc và thực hiện mở file excel
def fun_read_excel_byXlwing_open(file): 
    try:
        open_file = xw.Book(file)
    except:
        print("file đã dc mở")
    return open_file

# Đọc file không thực hiện mở file
def fun_read_Excel_byXlwing_Not_open(file):
    excel_app = xw.App(visible=False)
    try:
        excel_book = excel_app.books.open(file)  
    except:
        print("file đã dc mở")
    return excel_book

# Lấy tên sheet
def fun_get_all_sheet_name(excel):
    sheet_Names = [sh.name for sh in excel.sheets]
    return sheet_Names

# Tạo new sheet_name
def fun_create_new_sheet(excel, new_sheet):
    excel.sheets.add(new_sheet)

# Đổi tên sheet
def fun_change_sheet_name(sh, new_sheet): 
    sh.name = new_sheet

# delet sheet
def fun_del_sheet(excel, nameSheet_del):
    for sh in excel.sheets:
        if sh.name == nameSheet_del:
            sh.delete()

#  Xóa toàn bộ dữ liệu sheet
def clear_all_data_for_sheet(sh):
    sh.clear_contents()

# Xóa dòng
def fun_del_row(sh, numRow):
    sh.range(f'{numRow}:{numRow}').delete()

# Xóa cột
def fun_del_col(sh, name_col):
    sh.range(f'{name_col}:{name_col}').delete()


# Đổi số cột sang tên cột
def fun_covert_num_into_nameOfcol(sheets_Name, num):    
    name = sheets_Name.range(1, num).address
    for i, s in enumerate(name):
        if s =="$":
            begin = i+1
            break    
    for j in range(i+1, len(name)):
        if name[j] =="$":
            after = j
            name_col = name[begin:after]
            break
    return name_col

# Đổi tên cột sang số cột
def fun_covert_name_into_numOfcol(sheets_Name, name):
    num_Col = sheets_Name.range(name+str(1)).column
    return num_Col

# Tìm dòng cuối, có chứa dữ liệu: lr
def fun_last_row_xw(sh, col_name):
    lrOfcol = len(sh.range(f'{col_name}:{col_name}'))
    lr = sh.range(col_name+str(lrOfcol)).end('up').row
    return lr

# Tìm cột cuối, có chứa dữ liệu: lcol
def fun_last_col_xw(sh, row_index):
    col = len(sh.range(f'{row_index}:{row_index}'))
    lcol = sh.range(row_index, col).end("left").column
    # sh.range(1, col).address, xác định Ô (ko phải range) cuối cùng, sau đó đếm ngược về
    return lcol

# Lưu file xlsm
def fun_Save_file_xlsm(wb, new_name):
    wb.api.SaveAs(new_name, 51) # Ok lưu file có VBA

# Chuyển range thành list
def fun_covert_range_into_list2D(sh):
    li = sh.value
    return li

# Danh sách không trùng lặp
def fun_no_duplication(input_list):
    list_unique = []
    for x in input_list:
        if x not in list_unique:
            list_unique.append(x)
    return list_unique 

# filter theo cột Danh sách không trùng lặp trong range 2D
def fun_no_duplication_in_range_2D(input_list_2D, col_index_fillter):
    lc = len(input_list_2D[0])
    # lr = len(input_list_2D)
    if col_index_fillter<lc:

        range_unique = []
        out_range =[]

        for x in input_list_2D:
            if x[col_index_fillter] not in range_unique:
                range_unique.append(x[col_index_fillter])
                out_range.append(x)
        return out_range 
    else:
        print("Số cột nhập vào > tổng số cột")

# Xóa nhiều cột trong list 2D: Viết lại kiểu *arg
def fun_del_columns_2D_in_range(input_list, *arg):

    num_id = len(arg) 
    for c in range(num_id, 0, -1):
        col = arg[c-1] 
        # [c-1]: Trả về index list
        # arg[c-1]: là giá trị trong list

        for row in input_list:
            # print(row[col-1])             
            del row[col-1]
            # [col-1]: Trả về index list
            # row[col-1]: là giá trị trong list
            
    return input_list

# Xóa nhiều cột trong list 2D: Viết lại kiểu *arg
""" >>> rng_out = fun_del_columns_2D_in_range(rng_filter_out_sh, 1,2,5,7,8,9,10,11)
    >>> rng_out1 = fun_del_columns_2D_in_range(rng_filter_out_sh, col= [1,2,7,8,9,11])
"""
def fun_del_columns_2D_in_range(input_list, *arg, **kwargs):
    # input_list: nhập vào dạng list = range.value
    #     
    # Xử lý *arg
    num_id = len(arg)     
    for c in range(num_id, 0, -1):
        col = arg[c-1] 
        # [c-1]: Trả về index list
        # arg[c-1]: là giá trị trong list

        for row in input_list:
            # print(row[col-1])             
            del row[col-1]
            # [col-1]: Trả về index list
            # row[col-1]: là giá trị trong list

    # Xử lý **kwarg
    for key, value in kwargs.items():
        num_id = len(value)     
        for c in range(num_id, 0, -1):
            if str(value[c-1]).isdigit():
                col = int(value[c-1]) 
                # [c-1]: Trả về index list
                # arg[c-1]: là giá trị trong list

                for row in input_list:
                    # print(row[col-1])             
                    del row[col-1]
                    # [col-1]: Trả về index list
                    # row[col-1]: là giá trị trong list

    return input_list

# Add thêm tiêu đề cột
def fun_add_column_list(in_list_2D, *col_name):
    # col_name as string
    # in_list_2D =[x + [0] for x in in_list_2D]
    
    
    # Xử lý *col_name
    for c in col_name:
        new_list =[]
        for x in in_list_2D:
            new_list.append(x + [None])
        new_list[0][-1] = c
        in_list_2D = new_list
    return in_list_2D

# Thêm cột và tính toán giá trị cho cột trong list
def fun_add_column_and_return_valueOfThisCol (in_list_2D, **col_name):
    # col_name as string
    # in_list_2D =[x + [0] for x in in_list_2D]
    
    
    # Xử lý **col_name
    for key, c in col_name.key:
        comparison_sign = ['==', '!=', '>','<', '&', '|']
        new_list =[]
        for x in in_list_2D:
            new_list.append(x + [None])
        new_list[0][-1] = c
        in_list_2D = new_list
    return in_list_2D


def get_truth(inp, relate, cut):
    # print(get_truth(1.0, '>', 0.0)) # prints True
    ops = {'>': operator.gt,
           '<': operator.lt,
           '>=': operator.ge,
           '<=': operator.le,
           '==': operator.eq,
           '!=': operator.ne,
           }
    return ops[relate](inp, cut)