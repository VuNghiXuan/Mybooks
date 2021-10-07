from xlwings_VuNghiXuan.File_Exell import Workbook

# Mở file bằng thư viện xlwings_VuNghiXuan
wb = Workbook("D:\\MyBook_PyExcel\\read_Datas","\\Covid_VN.xlsx")

# In ra tên file excel
print("Tên file là:", wb.name)

# In ra đường dẫn đầy đủ của file
print("Tên đường dẫn đầy đủ file là:", wb.pathfull) 

# Mở 1 file excel
wb.open()

