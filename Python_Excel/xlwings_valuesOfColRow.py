datas = [[1,5,8],
         [2,7,12],
         [3,9,5],
         [4, 5.9,8.2],
         [5.8, 105,102.5]]

total_rows = len(datas)
total_columns = len(datas[0]) # Tức là tổng số phần tử hàng đầu tiên 
print(f'Ma trận Datas có {total_rows}(dòng) và {total_columns}(cột)\nCụ thể: ')

for i_row in range(total_rows):
    for i_col in range(total_columns):
        print(f'Dòng {i_row+1}, cột {i_col+1} có giá trị là: {datas[i_row][i_col]}')
    
    
    

# rows = [[1],[2],[3],[4],[5]]
# columns = ["Toán","Lý", "Hóa", "Sinh", "Ngoại Ngữ"]
# datas = [[1,5],[2,7],[3,9],[4, 5.9],[5.8, 105]]