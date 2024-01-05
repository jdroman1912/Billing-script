from ldap3 import Server, Connection, ALL, NTLM, SAFE_SYNC
import os
import xlsxwriter
import re


def getADEntries(ADAddress):
    server = Server('IGLU.LOCAL', get_info=ALL)
    conn = Connection(server, user = 'iglu\jroman.da', password = os.environ['AD_PASSWORD'], client_strategy = SAFE_SYNC, auto_bind = True)
    entries = conn.extend.standard.paged_search(ADAddress, '(objectClass=person)', attributes=['cn', 'givenName', 'sAMAccountName','mobile','extensionAttribute2', 'title', 'pwdLastSet', 'distinguishedName'], paged_size=1)
    result = []
    for entry in entries:
        user = [entry['attributes']['cn'], entry['attributes']['givenName'], entry['attributes']['sAMAccountName'],entry['attributes']['mobile'], entry['attributes']['extensionAttribute2'], entry['attributes']['title'], entry['attributes']['pwdLastSet'], entry['attributes']['distinguishedName']]
        result.append(user)
    
    return result

result = getADEntries('OU=Home Office Staff,OU=Home Office,DC=IGLU,DC=LOCAL') + getADEntries('OU=Central Park Staff,OU=Central Park,DC=IGLU,DC=LOCAL') + getADEntries('OU=Central Staff,OU=Central,DC=IGLU,DC=LOCAL') + getADEntries('OU=Brisbane Staff,OU=Brisbane City,DC=IGLU,DC=LOCAL') + getADEntries('OU=Broadway Staff,OU=Broadway,DC=IGLU,DC=LOCAL') + getADEntries('OU=Chatswood Staff,OU=Chatswood,DC=IGLU,DC=LOCAL') + getADEntries('OU=Flagstaff Staff,OU=Flagstaff,DC=IGLU,DC=LOCAL') + getADEntries('OU=Kelvin Grove Staff,OU=Kelvin Grove,DC=IGLU,DC=LOCAL') + getADEntries('OU=Mascot Staff,OU=Mascot,DC=IGLU,DC=LOCAL') + getADEntries('OU=Melbourne City Staff,OU=Melbourne City,DC=IGLU,DC=LOCAL') + getADEntries('OU=Redfern Staff,OU=Redfern,DC=IGLU,DC=LOCAL') + getADEntries('OU=South Yarra Staff,OU=South Yarra,DC=IGLU,DC=LOCAL') + getADEntries('OU=Summer Hill Staff,OU=Summer Hill,DC=IGLU,DC=LOCAL')
#print(result)
#print(result[0][2])

#Changing the format of the mobile numbers so that it matches my other list
for innerList in result:
    if str(innerList[3]).startswith("+61 "):
        innerList[3] = '0' + innerList[3][4:] #Cambie el 2 por un 3

#Getting rid of the CN on the OU section
for innerList1 in result:
    ou = 'OU='
    pos = innerList1[7].find(ou)
    if pos != -1:
        innerList1[7] = innerList1[7][pos:] #Cambie el 6 por un 7

#Fixing the date format to do the Day Diff in excel
for innerList2 in result:
    dateMatch = re.search(r"\d{4}-\d{2}-\d{2}", str(innerList2[6]))
    if dateMatch:
        innerList2[6] = dateMatch.group(0) #Cambie el 5 por un 6
#---------------------------------------------------------------- 
#This is to create the excel file
row = 1
column = 0
workbook = xlsxwriter.Workbook('ADUsers.xlsx')
format = workbook.add_format({'bold': True,'border':2, 'align':'center', 'bg_color':'#00FF00' })
worksheet = workbook.add_worksheet('AD Users')
worksheet.write('A1', 'CN', format)
worksheet.write('B1', 'Given Name', format)
worksheet.write('C1', 'Login', format)
worksheet.write('D1', 'Mobile Number', format)
worksheet.write('E1', 'Type', format)
worksheet.write('F1', 'Job Title', format)
worksheet.write('G1', 'Last Set Password', format)
worksheet.write('H1', 'OU', format)
worksheet.write('I1', 'Day Diff', format)

cont = 2
for index, user in enumerate(result):

    
    worksheet.write(row, 0, str(user[0]))

    worksheet.write(row, 1, str(user[1]))

    worksheet.write(row, 2, str(user[2]))

    worksheet.write(row, 3, str(user[3]))

    worksheet.write(row, 4, str(user[4]))

    worksheet.write(row, 5, str(user[5]))

    worksheet.write(row, 6, str(user[6]))

    worksheet.write(row, 7, str(user[7]))

    worksheet.write(row, 8, '=DATEDIF(F'+str(index+cont)+',TODAY(),"d")')

    row += 1 
    

workbook.close()
#-----------------------------------------------------------------