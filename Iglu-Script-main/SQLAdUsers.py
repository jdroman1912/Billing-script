import pandas as pd
from sqlalchemy import create_engine, DATE
from ADTest2 import getADEntries
from env import MYSQL_USERNAME, MYSQL_PASSWORD, MYSQL_HOST, MYSQL_DATABASE
import re

#Getting the information from AD by using the function created in ADTest2
result = getADEntries('OU=Home Office Staff,OU=Home Office,DC=IGLU,DC=LOCAL') + getADEntries('OU=Central Park Staff,OU=Central Park,DC=IGLU,DC=LOCAL') + getADEntries('OU=Central Staff,OU=Central,DC=IGLU,DC=LOCAL') + getADEntries('OU=Brisbane Staff,OU=Brisbane City,DC=IGLU,DC=LOCAL') + getADEntries('OU=Broadway Staff,OU=Broadway,DC=IGLU,DC=LOCAL') + getADEntries('OU=Chatswood Staff,OU=Chatswood,DC=IGLU,DC=LOCAL') + getADEntries('OU=Flagstaff Staff,OU=Flagstaff,DC=IGLU,DC=LOCAL') + getADEntries('OU=Kelvin Grove Staff,OU=Kelvin Grove,DC=IGLU,DC=LOCAL') + getADEntries('OU=Mascot Staff,OU=Mascot,DC=IGLU,DC=LOCAL') + getADEntries('OU=Melbourne City Staff,OU=Melbourne City,DC=IGLU,DC=LOCAL') + getADEntries('OU=Redfern Staff,OU=Redfern,DC=IGLU,DC=LOCAL') + getADEntries('OU=South Yarra Staff,OU=South Yarra,DC=IGLU,DC=LOCAL') + getADEntries('OU=Summer Hill Staff,OU=Summer Hill,DC=IGLU,DC=LOCAL')

#Fixing the formatting for the phone numbers
for innerList in result:
    if str(innerList[3]).startswith("+61 "):
        innerList[3] = '0' + innerList[3][4:]

#Fixing the formatting for OU
for innerList1 in result:
    ou = 'OU='
    pos = innerList1[7].find(ou)
    if pos != -1:
        innerList1[7] = innerList1[7][pos:]

#Fixing the formatting for Date
for innerList2 in result:
    dateMatch = re.search(r"\d{4}-\d{2}-\d{2}", str(innerList2[6]))
    if dateMatch:
        innerList2[6] = dateMatch.group(0)

#Escaping the special character '=' for the OU expression
for i in range(len(result)):
    for j in range(len(result[i])):
        if isinstance(result[i][j], list):
            result[i][j] = str(result[i][j]).replace('=', '\=')

#Changing the list into a dictionary
resultDict = {}
for idx, item in enumerate(result):
    resultDict[idx+1]= {
        "CN": item[0],
        "GivenName": item[1],
        "Login": item[2],
        "MobileNumber": item[3],
        "Type": item[4],
        "JobTitle": item[5],
        "LastSetPassword": item[6],
        "OU": item[7]
    }

#Creating the pandas dataframe
df = pd.DataFrame(resultDict, columns=["CN", "GivenName", "Login","MobileNumber", "Type", "JobTitle", "LastSetPassword", "OU"])

#Getting the values of the dictionary and adding them to the dataframe
for x in resultDict.values():
    df = df.append(x, ignore_index = True)

#Connecting to MySQL
engine = create_engine('mysql+mysqldb://' + MYSQL_USERNAME + ':' + MYSQL_PASSWORD + '@' + MYSQL_HOST + '/' + MYSQL_DATABASE)

#Assign type
dtypes = {'LastSetPassword': DATE}

#Uploading the table to MySQL
df.to_sql(name='ADUser', con=engine, if_exists='replace', chunksize = 1000, index=False, dtype=dtypes)

