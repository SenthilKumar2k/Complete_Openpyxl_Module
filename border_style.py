import openpyxl
from openpyxl.styles import Border, Side

wb=openpyxl.load_workbook("students.xlsx")
ws=wb['Sheet2']

top=Side(border_style="dashed", color="CD5C5C" )
bot=Side(border_style="dotted", color="DE3163")
right=Side(border_style="double", color="800000")
left=Side(border_style="dashDot", color="0000FF")

border=Border(top=top, bottom=bot, right=right, left=left)
ws['A6'].border=border
wb.save("students.xlsx")

# the border style.
#  hair, dashed, medium, slantDashDot, thick, dotted, dashDotDo,
# dashDot, double, mediumDashDot, mediumDashed, thin, mediumDashDotDot