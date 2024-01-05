import xlsxwriter
from PyPDF2 import PdfReader
import PyPDF2
import itertools
import re

#-----------------------------
#Attributes to edit
file = "April23.pdf"
initialPage = 4
lastPage = 24
excelFileName = 'TelstraBillsApril.xlsx'
month = "April 2023"
#------------------------------
with open(file, "rb") as f:
    reader = PdfReader(f)
    number_of_pages = len(reader.pages)
    matches = []
    # Iterate through all the pages
    for i in range(initialPage, lastPage):
        page = reader.pages[i]
        text = page.extract_text() 
        firstMatch = re.finditer(r'\d{4} \d{3} \d{3}', text)
        for match in firstMatch: #finditer is used to find all occurrences of the numbers in the text variable. The function returns a callable iterator that yields match objects
            start = match.start() #Each match object has a start() and end() method that can be used to find the position of the match in the text
            end = match.end()
            nextMatch = re.search(r'\d{4} \d{3} \d{3}', text[end:])
            if nextMatch:
                nextStart = nextMatch.start() + end
                lines = text[start:nextStart] 

                matches.append([lines])
            else:
                lines = text[start:]
                matches.append([lines])


#print(matches)
#print(matches[4][0])
#print(matches[0][0].split("\n"))
#firstItem = matches[0][0].split("\n")

matchesSplit = [match1[0].split("\n") for match1 in matches] 
#print(matchesSplit)
for innerList in matchesSplit:
    if innerList[0].endswith(" continued"):
        innerList[0] = innerList[0].replace(" continued", "") #Getting rid of the continued


mergedList =[]
for i in range(len(matchesSplit)):
    if i==0:
        mergedList.append(matchesSplit[i])
    elif matchesSplit[i][0] == matchesSplit[i-1][0]:
        mergedList[-1].extend(matchesSplit[i][1:])
    else:
        mergedList.append(matchesSplit[i])
#print(mergedList)

number = '\d{4} \d{3} \d{3}'
nationalCalls = '(?<=National\s)\d+(?=\scalls)'
nationalTelstraCalls = '(?<=Mobiles\s)\d+(?=\scalls)'
forwardedCalls = '(?<=service\s)\d+(?=\scalls)'
dataUsage = '(\d+.\d+ GB)'
dataUsageMB = '(\d+.\d+ MB)'
plan = "Business Mobile Plan -.*?-"
planXS = 'Business Data Plan XS'
planBasic = 'Business Mobile Plan Basic'
total = '(?<=)(\$\d+\.\d+)(?=\',\s\'excl)'
totalCr = '(\$\d+\.\d+)(?=cr\',\s\'excl)'

#---------------------------------------------------------------- 
# #This is to create the excel file
row = 1
column = 0
workbook = xlsxwriter.Workbook(excelFileName)
format = workbook.add_format({'bold': True,'border':2, 'align':'center', 'bg_color':'#00FF00' })
crFormat = workbook.add_format({'font_color': 'red'})
monthFormat = workbook.add_format({'bg_color':'#C0C0C0'})
worksheet = workbook.add_worksheet('August Bill')
worksheet.write('A1', 'Month', format)
worksheet.write('B1', 'Mobile Number', format)
worksheet.write('C1', 'National', format)
worksheet.write('D1', 'National to Telstra Mobile', format)
worksheet.write('E1', 'Forwarded Calls', format)
worksheet.write('F1', 'Data Usage', format)
worksheet.write('G1', 'Plan', format)
worksheet.write('H1', 'Total', format)
cont = 0

for mobile in mergedList:
    numberMatch = re.findall(number, str(mobile))
    worksheet.write(row, 0, month, monthFormat)
    if numberMatch:
        for x in numberMatch:
            worksheet.write(row, 1, x) 
        
    else:
        worksheet.write(row, 1, 'No number')


    nationalMatch = re.search(nationalCalls, str(mobile))
    if nationalMatch:
        worksheet.write(row, 2, nationalMatch.group())
    else:
        worksheet.write(row, 2, '0')

    nationalTelstraCallsMatch = re.search(nationalTelstraCalls, str(mobile))
    if nationalTelstraCallsMatch:
        worksheet.write(row, 3, nationalTelstraCallsMatch.group())
    else:
        worksheet.write(row, 3, '0')


    forwardedCallsMatch = re.search(forwardedCalls, str(mobile))
    if forwardedCallsMatch:
        worksheet.write(row, 4, forwardedCallsMatch.group())
    else:
        worksheet.write(row, 4, '0')
    
    dataUsageMatch = re.search(dataUsage, str(mobile))
    dataUsageMatchMB = re.search(dataUsageMB, str(mobile))
    if dataUsageMatch:
        worksheet.write(row, 5, dataUsageMatch.group())
    
    elif dataUsageMatchMB:
        worksheet.write(row, 5, dataUsageMatchMB.group())
    else:
        worksheet.write(row, 5, 'No Data Usage')#
    
    planMatch = re.search(plan, str(mobile))
    planXSMatch = re.search(planXS, str(mobile))
    planBasicMatch = re.search(planBasic, str(mobile))
    if planMatch:
        worksheet.write(row, 6, planMatch.group())
    
    elif planXSMatch:
        worksheet.write(row, 6, planXSMatch.group())
    
    elif planBasicMatch:
        worksheet.write(row, 6, planBasicMatch.group())
    else:
        worksheet.write(row, 6, 'No Business Plan Detected')

    totalMatch = re.search(total, str(mobile))
    totalMatchCr = re.search(totalCr, str(mobile))
    if totalMatch:
        totalValue = totalMatch.group()
        worksheet.write(row, 7, totalValue)
    elif totalMatchCr:
        worksheet.write(row, 7, '-'+totalMatchCr.group(), crFormat)
    else:
        worksheet.write(row, 7, 'Not found')

    row += 1 
    

workbook.close()
#-----------------------------------------------------------------
#XlsxWriter colors
#https://xlsxwriter.readthedocs.io/working_with_colors.html#colors
