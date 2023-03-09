#Defining (back to (Excel_CO2019) file for more details
import openpyxl

#Open a file in py and import CO data without the NULL DATA 

#The source Data
workbook = openpyxl.load_workbook('CO_Step1_DATA_WithoutNULL.xlsx')
sheet = workbook.active         



#Delete empty rows
#The method is repeated becaues the method doesn't delete 2 empty rows in order
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






