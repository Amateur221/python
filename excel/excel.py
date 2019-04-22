import xlrd
import xlwt
data = xlrd.open_workbook(r"D:\GitHub\python\TL-441001040_BOM.xls")
table = data.sheets()[0]
print(table)

nRows = table.nrows
nCols = table.ncols
print("%d" % nRows)
print("%d" % nCols)

rows1 = table.row_values(0)
print("%s" % rows1)
#cell_A1 = table.cell_value(0, 0)
#print("%s" % cell_A1)



workbook = xlwt.Workbook(encoding='utf-8', style_compression=0)
sheet = workbook.add_sheet('test', cell_overwrite_ok=True)
#sheet.write(0, 0, cell_A1)
for cell_value1 in rows1:
    sheet.write(0, 0, cell_value1)
    print("%s" % cell_value1)

workbook.save(r"D:\GitHub\python\1.xls")
