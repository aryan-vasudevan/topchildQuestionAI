import openpyxl

path = "questions.xlsx" # path to the excel file
wb = openpyxl.load_workbook(path) # load the workbook

sheet = wb.active  
  
x1 = sheet['A1']  
x2 = sheet['A2']  
  
print("The first cell value:",x1.value)  
print("The second cell value:",x2.value)  


