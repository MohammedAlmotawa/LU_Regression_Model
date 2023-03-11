
import openpyxl

#Open a file in py and import CO data without the NULL DATA 

#The source Data
workbook = openpyxl.load_workbook('CO_Step2_DeleteNULLEmptyRows.py')
sheet = workbook.active       

#Because there is still some empty rows, I needed to apply the method for the new file
#Delete empty rows
for row in range (2,6415): 
    cell = sheet['J' + str (row)].value
    if cell == None:
        sheet.delete_rows(row)
    else:
        pass 
for row in range (2,6415): 
    cell = sheet['J' + str (row)].value
    if cell == None:
        sheet.delete_rows(row)
    else:
        pass 
for row in range (2,6415): 
    cell = sheet['J' + str (row)].value
    if cell == None:
        sheet.delete_rows(row)
    else:
        pass 
for row in range (2,6415): 
    cell = sheet['J' + str (row)].value
    if cell == None:
        sheet.delete_rows(row)
    else:
        pass 
for row in range (2,6415): 
    cell = sheet['J' + str (row)].value
    if cell == None:
        sheet.delete_rows(row)
    else:
        pass 
for row in range (2,6415): 
    cell = sheet['J' + str (row)].value
    if cell == None:
        sheet.delete_rows(row)
    else:
        pass

workbook.save('CO_Step3_DeletedEmptyRows.xlsx')






