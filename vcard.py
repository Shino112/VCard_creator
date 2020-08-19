# coding=utf-8
import vobject
import base64
import binascii
import xlrd

# Give the location of the file
loc = ("Kontakti.xlsx")

# To open Workbook
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)

file1 = open('slike.txt', 'r')
Lines = file1.readlines()
rows = [r for r in Lines]

for i in range(1, sheet.nrows):

    j = vobject.vCard()

    if (sheet.cell_value(i, 5) != 'NULL'):
        j.add('fn')
        j.fn.value = sheet.cell_value(i, 5).capilize()

    if (sheet.cell_value(i, 6) != 'NULL'):
        if (sheet.cell_value(i, 7) != 'NULL'):
            j.add('n')
            j.n.value = vobject.vcard.Name(family=sheet.cell_value(i, 6), given=sheet.cell_value(i,7))

    if (sheet.cell_value(i, 4) != 'NULL'):
        j.add('email')
        j.email.value = sheet.cell_value(i, 4)
        j.email.type_param = 'INTERNET'

    if (sheet.cell_value(i, 8) != 'NULL'):
        j.add('org')
        j.org.value = [sheet.cell_value(i, 8)]

    if (sheet.cell_value(i, 11) != 'NULL'):
        j.add('title')
        vrsta_posla = sheet.cell_value(i, 11)
        j.title.value = vrsta_posla

    if (sheet.cell_value(i, 9) != 'NULL'):
        o = j.add('tel')
        o.type_param = "cell"
        if str(sheet.cell_value(i, 9)).startswith('9'):
            cell = int(sheet.cell_value(i, 9))
            cell = str(cell)
            o.value = cell
        else:
            o.value = str(sheet.cell_value(i, 9))

    if (sheet.cell_value(i, 10) != 'NULL'):
        o = j.add('tel')
        o.type_param = "work"
        if str(sheet.cell_value(i, 10)).startswith('9'):
            work = int(sheet.cell_value(i, 10))
            work = str(work)
            o.value = work
        else:
            o.value = str(sheet.cell_value(i, 10))

    if (sheet.cell_value(i, 0) != 'NULL'):
        if (sheet.cell_value(i, 1) != 'NULL'):
            if (sheet.cell_value(i, 2) != 'NULL'):
                j.add('adr')
                code = int(sheet.cell_value(i, 2))
                code = str(code)
                j.adr.value = vobject.vcard.Address(street=sheet.cell_value(i, 0), city=sheet.cell_value(i, 1), code= code, country='Hrvatska')
                j.adr.type_param = 'WORK'

    if (sheet.cell_value(i, 3) != 'NULL'):
        j.add('company')
        j.company.value = sheet.cell_value(i, 3)

    data = rows[i-1]
    data = data.strip()
    data = data.replace('0x', '')
    data = data.replace(' ', '')
    data = data.replace('\n', '')
    try:
        data = binascii.a2b_hex(data)
        with open('image.jpg', 'wb') as image_file:
            image_file.write(data)

        photo = open('image.jpg', 'rb')
        j.add('photo')
        j.photo.type_param = 'JPEG'
        j.photo.value = base64.encodestring(photo.read())
    except TypeError:
        pass
    print(j.serialize())
    file = open('kartice/iOSContact' + str(i) + '.vcf', 'w')
    file.write(j.serialize())
    file.close()