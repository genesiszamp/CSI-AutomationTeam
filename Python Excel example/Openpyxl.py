from openpyxl import Workbook, load_workbook
 
wb = load_workbook('SLMS.xlsx')

ws = wb.active
sheet_names = ws.title

 
list = []
 
for row in ws.iter_rows():
    list.append(row[1].value)
    list.remove("Overdue Trainings")
    list = [int(i) for i in list]
    Overdue_Trainings = sum(list)


sum=wb.create_sheet('Sum',0)
sum.append([ws.title,Overdue_Trainings])

#for sheet in wb.worksheets:
#    print(sheet)



print(Overdue_Trainings)
print(sheet_names)
wb.save("SLMS.xlsx")