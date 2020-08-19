# Reading an excel file using Python
import xlrd
import binascii

# Give the location of the file
loc = ("Kontakti.xlsx")

# To open Workbook
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)

# Extracting number of rows
print(sheet.nrows)
print(sheet.ncols)

# For row 0 and column 0
print(int(sheet.cell_value(1, 2)))

for i in range(1, sheet.nrows):
    print("Work street: {}".format(sheet.cell_value(i, 0).encode('utf-8')))
    print("Work city: {}".format(sheet.cell_value(i, 1).encode('utf-8')))
    print("Work code: {}".format(int(sheet.cell_value(i, 2))))
    print("Company: {}".format(sheet.cell_value(i, 3).encode('utf-8')))
    print("Email: {}".format(sheet.cell_value(i, 4).encode('utf-8')))
    print("FN: {}".format(sheet.cell_value(i, 5).encode('utf-8')))
    print("familyName: {}".format(sheet.cell_value(i, 6).encode('utf-8')))
    print("givenName: {}".format(sheet.cell_value(i,7).encode('utf-8')))
    print("ORG: {}".format(sheet.cell_value(i, 8).encode('utf-8')))
    if str(sheet.cell_value(i,9)).startswith('9'):
        print("Cell tel: {}".format(int(sheet.cell_value(i, 9))))
    else:
        print("Cell tel: {}".format(sheet.cell_value(i, 9)))
    if str(sheet.cell_value(i,10)).startswith('9'):
        print("Cell tel: {}".format(int(sheet.cell_value(i, 10))))
    else:
        print("Cell tel: {}".format(sheet.cell_value(i, 10)))
    print("Title: {}".format(sheet.cell_value(i, 11).encode('utf-8')))

file1 = open('slike.txt', 'r')
Lines = file1.readlines()
rows = [r for r in Lines]
print len(Lines)
count = 1
#for line in Lines:
print(rows[3])
count = count + 1
data = rows[3]
data=data.strip()
data=data.replace('0x', '')
data=data.replace(' ', '')
data=data.replace('\n', '')
data=data.rstrip()
print(data)
try:
    data = binascii.a2b_hex(data)
    with open('image.jpg', 'wb') as image_file:
        image_file.write(data)
except TypeError:
    pass
file1.close()