import xlrd
import xlwt
data = xlrd.open_workbook(r"C:\Users\Administrator\Desktop\TL-441001040_BOM.xls")
table = data.sheets()[0]
print(table)

nrows = table.nrows
print("%d" % nrows)

cell_A1 = table.cell_value(0, 0)
print("%s" % cell_A1)

workbook = xlwt.Workbook(encoding='utf-8', style_compression=0)
sheet = workbook.add_sheet('test', cell_overwrite_ok=True)
sheet.write(0, 0, cell_A1)
workbook.save(r"C:\Users\Administrator\Desktop\1.xls")
