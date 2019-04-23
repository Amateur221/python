import xlrd
import xlwt
data = xlrd.open_workbook(r"D:\GitHub\python\TL-441001040_BOM.xls")
table = data.sheets()[0]
print(table)

nRows = table.nrows
nCols = table.ncols
print("%d" % nRows)
print("%d" % nCols)

# 读取Excel表格前七行数据并存储在列表rowsTable
rowsTable = []
for num in range(8):
    rowsTable.append(table.row_values(num))

print("%s" % rowsTable)
# cell_A1 = table.cell_value(0, 0)
# print("%s" % cell_A1)



workBook = xlwt.Workbook(encoding = 'utf-8', style_compression = 0)
sheet = workBook.add_sheet('test', cell_overwrite_ok=True)
# sheet.write(0, 0, cell_A1)

for num1 in range(8):
    ColsNum = 0
    while ColsNum < nCols:
        for cell_value1 in rowsTable[num1]:
            sheet.write(num1, ColsNum, cell_value1)
            if len(rowsTable[num1][ColsNum]) != 0:
                ColsNum += 1
                print("%d" % ColsNum)
            else:
                ColsNum = 27
                break
            print("%s" % cell_value1)

workBook.save(r"D:\GitHub\python\2.xls")
