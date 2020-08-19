# coding=utf-8
import vobject
import base64
import binascii
import sqlserver

for i in range(0, len(sqlserver.kontakti)):

    j = vobject.vCard()

    if (sqlserver.kontakti[i][5] != 'NULL'):
        j.add('fn')
        j.fn.value = sqlserver.kontakti[i][5]

    if (sqlserver.kontakti[i][6] != 'NULL'):
        if (sqlserver.kontakti[i][7] != 'NULL'):
            j.add('n')
            j.n.value = vobject.vcard.Name(family=sqlserver.kontakti[i][6], given=sqlserver.kontakti[i][7])

    if (sqlserver.kontakti[i][4] != 'NULL'):
        j.add('email')
        j.email.value = sqlserver.kontakti[i][4]
        j.email.type_param = 'INTERNET'

    if (sqlserver.kontakti[i][8] != 'NULL'):
        j.add('org')
        j.org.value = [sqlserver.kontakti[i][8]]

    if (sqlserver.kontakti[i][12] != 'NULL'):
        j.add('title')
        vrsta_posla = sqlserver.kontakti[i][12]
        j.title.value = vrsta_posla

    if (sqlserver.kontakti[i][10] != 'NULL'):
        o = j.add('tel')
        o.type_param = "cell"
        if str(sqlserver.kontakti[i][10]).startswith('9'):
            cell = int(sqlserver.kontakti[i][10])
            cell = str(cell)
            o.value = cell
        else:
            o.value = str(sqlserver.kontakti[i][10])

    if (sqlserver.kontakti[i][11] != 'NULL'):
        o = j.add('tel')
        o.type_param = "work"
        if str(sqlserver.kontakti[i][11]).startswith('9'):
            work = int(sqlserver.kontakti[i][11])
            work = str(work)
            o.value = work
        else:
            o.value = str(sqlserver.kontakti[i][11])

    if (sqlserver.kontakti[i][0] != 'NULL'):
        if (sqlserver.kontakti[i][1] != 'NULL'):
            if (sqlserver.kontakti[i][2] != 'NULL'):
                j.add('adr')
                code = int(sqlserver.kontakti[i][2])
                code = str(code)
                j.adr.value = vobject.vcard.Address(street=sqlserver.kontakti[i][0], city=sqlserver.kontakti[i][1], code= code, country='Hrvatska')
                j.adr.type_param = 'WORK'

    if (sqlserver.kontakti[i][3] != 'NULL'):
        j.add('company')
        j.company.value = sqlserver.kontakti[i][3]

    if (sqlserver.kontakti[i][9] != 'NULL'):
        data = sqlserver.kontakti[i][9]
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
        file = open('kartice/Contacts' + str(i) + '.vcf', 'w')
        file.write(j.serialize())
        file.close()