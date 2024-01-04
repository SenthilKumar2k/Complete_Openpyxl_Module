import openpyxl

workbook = openpyxl.Workbook()
sheet = workbook.active

# Add data to the sheet
sheet['A1'] = 'Hello'
sheet['B1'] = 'World!'

# Save the workbook
workbook.save('example.xlsx')

try:
    workbook=openpyxl.load_workbook("students.xlsx")
    print(workbook.sheetnames)
except Exception as e:
    print("error:{}".format(e))
