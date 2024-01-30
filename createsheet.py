import openpyxl

wb=openpyxl.Workbook()
ws=wb.active

result={"senthil":99, "kumar":98, "raji":100, "priya":95, "abi":93}

for i,j in result.items():
    ws.append([i,j])

wb.save("mark.xlsx")

