import openpyxl
from openpyxl.worksheet.datavalidation import DataValidation

wb=openpyxl.load_workbook("mark.xlsx")
ws=wb['Sheet']

update='"in complete, in progress, not started "'

rule=DataValidation(type='list', formula1=update, allow_blank=True)
rule.error="your entry is not valid"
rule.errorTitle="Invalid Entry"
rule.prompt="please select from list"
rule.promptTitle="select option"
rule.add("C1:C5")
ws.add_data_validation(rule)
wb.save("mark.xlsx")