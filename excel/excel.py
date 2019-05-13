import xlrd
import xlwt

# 读取模板加载模板数据到新建的文档之中
data = xlrd.open_workbook(r'模板.xls')
table = data.sheets()[0]
print(table)

nRows = table.nrows
nCols = table.ncols
# print("%d" % nRows)
# print("%d" % nCols)

# 读取Excel表格前七行数据并存储在列表rowsTable
rowsTable = []
for num in range(8):
    rowsTable.append(table.row_values(num))

# print("%s" % rowsTable)
# cell_A1 = table.cell_value(0, 0)
# print("%s" % cell_A1)

# 打开数据文件，将数据加载入上一步新建的文档
print("请输入要处理的BOM文档路径：")
bomURL = input('>')
data1 = xlrd.open_workbook(bomURL)
table1 = data1.sheets()[0]

nRows1 = table1.nrows
nCols1 = table1.ncols
# print("%d" % nRows1)
# print("%d" % nCols1)

colsTable = []
for num1 in range(4):
    colsTable.append(table1.col_values(num1))
# print("%s" % colsTable)


workBook = xlwt.Workbook(encoding='utf-8', style_compression=0)
sheet = workBook.add_sheet('test', cell_overwrite_ok=True)
# sheet.write(0, 0, cell_A1)

# 将模板中的前八行数据填入新建文档
for num1 in range(8):
    ColsNum = 0
    while ColsNum < nCols:
        for cell_value1 in rowsTable[num1]:
            sheet.write(num1, ColsNum, cell_value1)
            if len(rowsTable[num1][ColsNum]) != 0:
                ColsNum += 1
#                print("%d" % ColsNum)
            else:
                ColsNum = 27
                break
#            print("%s" % cell_value1)

# 将第1列的数据填入新建文档
for num2 in range(nRows1):
    sheet.write(num2+8, 2, colsTable[0][num2])

# 将第2列的数据填入新建文档
for num2 in range(nRows1):
    sheet.write(num2+8, 21, colsTable[1][num2])

# 将第3列的数据填入新建文档
for num2 in range(nRows1):
    sheet.write(num2+8, 0, colsTable[2][num2])

# 将第4列的数据填入新建文档
for num2 in range(nRows1):
    sheet.write(num2+8, 1, colsTable[3][num2])

print("请输入保存路径：")
url = input('>')
workBook.save(url)
