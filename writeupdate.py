import openpyxl

wb=openpyxl.load_workbook("students.xlsx")
ws=wb["Sheet"]
print(ws['B5'].value)   # used to read the values in sheet cell
ws['B5'].value=471      # used to update the value of particular cell
ws['A9']="meera"        # helps to write in a particular cell
ws['B9']=453
wb.save("students.xlsx")
