
import openpyxl

#Open a file in py and import CO data without the NULL DATA 

#The source Data
workbook = openpyxl.load_workbook('CO_Step1_DATA_WithoutNULL.xlsx')
sheet = workbook.active         



#Deleting empty rows
#The method was repeated becaues the method didn't delete 2 or more empty rows next to each other
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

workbook.save('CO_Step2_DeletedEmptyRows.xlsx')






