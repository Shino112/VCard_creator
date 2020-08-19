import pyodbc
import binascii

server = 'SQLCL1INSTANCE6.ri.valamar.local\INSTANCE6'
database = 'HRNetStagingProd'
username = 'MIMuserHRNet'
password = 'iwHb3v7bofu5JAXU'
cnxn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
cursor = cnxn.cursor()

#Select Query
print ('Reading data from table')
tsql = "SELECT WorkLocationAddress + ' ' + WorkLocationStreetNumber AS 'Work street'," \
       "WorkLocationName AS 'Work city', WorkLocationZipCode AS 'Work code'," \
       "CompanyName AS 'Company'," \
       "Email, FName + ' ' + LName AS 'FN'," \
       "LName AS 'familyName'," \
       "FName AS 'givenName'," \
       "OrgUnit1Name AS 'ORG'," \
       "Picture AS 'Photo'," \
       "Mob AS 'Cell tel'," \
       "InternalTel AS 'Work tel'," \
       "JobPositionHR AS 'Title'" \
       "FROM [HRNetStagingProd].[Integration].[Employee]" \
       "WHERE IsActive = 1 AND Mob IS NOT NULL AND Email IS NOT NULL"
with cursor.execute(tsql):
    rows = cursor.fetchall()
    rows = [r for r in rows]
    kontakti = [[[] for x in range(0, 13)] for x in range(0, len(rows))]
    for i in range(0, len(rows)):
        if rows[i][0] is not None:
            print("Work street: {}".format(rows[i][0].encode('utf-8')))
            kontakti[i][0] = rows[i][0].encode('utf-8')
        else:
            print("Work street: NULL")
            kontakti[i][0] = "NULL"
        if rows[i][1] is not None:
            print("Work city: {}".format(rows[i][1].encode('utf-8')))
            kontakti[i][1] = rows[i][1].encode('utf-8')
        else:
            print("Work city: NULL")
            kontakti[i][1] = "NULL"
        if rows[i][2] is not None:
            print("Work code: {}".format(rows[i][2]))
            kontakti[i][2] = rows[i][2].encode('utf-8')
        else:
            print("Work code: NULL")
            kontakti[i][2] = "NULL"
        if rows[i][3] is not None:
            print("Company: {}".format(rows[i][3].encode('utf-8')))
            kontakti[i][3] = rows[i][3].encode('utf-8')
        else:
            print("Company: NULL")
            kontakti[i][3] = "NULL"
        if rows[i][4] is not None:
            print("Email: {}".format(rows[i][4].encode('utf-8')))
            kontakti[i][4] = rows[i][4].encode('utf-8')
        else:
            print("Email: NULL")
            kontakti[i][4] = "NULL"
        if rows[i][5] is not None:
            print("FN: {}".format(rows[i][5].encode('utf-8')))
            kontakti[i][5] = rows[i][5].encode('utf-8')
        else:
            print("FN: NULL")
            kontakti[i][5] = "NULL"
        if rows[i][6] is not None:
            print("familyName: {}".format(rows[i][6].encode('utf-8')))
            kontakti[i][6] = rows[i][6].encode('utf-8')
        else:
            print("familyName: NULL")
            kontakti[i][6] = "NULL"
        if rows[i][7] is not None:
            print("givenName: {}".format(rows[i][7].encode('utf-8')))
            kontakti[i][7] = rows[i][7].encode('utf-8')
        else:
            print("givenName: NULL")
            kontakti[i][7] = "NULL"
        if rows[i][8] is not None:
            print("ORG: {}".format(rows[i][8].encode('utf-8')))
            kontakti[i][8] = rows[i][8].encode('utf-8')
        else:
            print("ORG: NULL")
            kontakti[i][8] = "NULL"
        if rows[i][9] is not None:
            kontakti[i][9] = binascii.hexlify(bytearray(rows[i][9]))
        else:
            kontakti[i][9] = "NULL"
        if rows[i][10] is not None:
            print("Cell tel: {}".format(rows[i][10]))
            kontakti[i][10] = rows[i][10].encode('utf-8')
        else:
            print("Cell tel: NULL")
            kontakti[i][10] = "NULL"
        if rows[i][11] is not None:
            print("Work tel: {}".format(rows[i][11]))
            kontakti[i][11] = rows[i][11].encode('utf-8')
        else:
            print("Work tel: NULL")
            kontakti[i][11] = "NULL"
        if rows[i][12] is not None:
            print("Title: {}".format(rows[i][12].encode('utf-8')))
            kontakti[i][12] = rows[i][12].encode('utf-8')
        else:
            print("Title: NULL")
            kontakti[i][11] = "NULL"