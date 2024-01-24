import openpyxl

master_report=openpyxl.load_workbook("lifetime_report.xlsx")
daily_report=openpyxl.load_workbook("daily_report.xlsx")
master_data=master_report['Data']
daily_data=daily_report['Sheet1']

#get row count for daily_report
data_daily=True
daily_row_count=0
while data_daily:
    daily_row_count+=1
    data=daily_data.cell(row=daily_row_count, column=1).value
    if data==None:
        data_daily=False

print("row count for daily report : ",daily_row_count)

#get row count for master_report
data_master=True
master_row_count=0
while data_master:
    master_row_count+=1
    data=master_data.cell(row=master_row_count,column=1).value
    if data==None:
        data_master=False

print("row count for master report : ", master_row_count)
