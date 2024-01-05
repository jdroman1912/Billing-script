import pandas as pd
import xlsxwriter

file1 = pd.read_excel('TelstraBills.xlsx')
file2 = pd.read_excel('ADUsers.xlsx')



phoneNumbers1 = file2['Mobile Number'].tolist()

nonMatchingNumbers = []

for index, row in file1.iterrows():
    if row['MMonth'] == '2023-04-01' and row['Mobile Number'] not in phoneNumbers1:
        nonMatchingNumbers.append(row['Mobile Number'])

nonMatchingNumbers = list(set(nonMatchingNumbers))
#-----------------------------------------------------

row = 1
workbook = xlsxwriter.Workbook('MissingNumbersApril.xlsx')
format = workbook.add_format({'bold': True,'border':2, 'align':'center', 'bg_color':'#00FF00' })
worksheet = workbook.add_worksheet('Missing Numbers')
worksheet.write('A1', 'Mobile Numbers', format)

for number in nonMatchingNumbers:
    worksheet.write(row, 0, number)
    row += 1 


workbook.close()
#-----------------------------------------------------
