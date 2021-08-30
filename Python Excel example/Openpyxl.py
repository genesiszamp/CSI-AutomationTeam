from openpyxl import Workbook, load_workbook
 
wb = load_workbook('SLMS.xlsx')
write = wb.create_sheet('Total')
ws = wb.active
sheets = wb.sheetnames
n = 0

 
total = []
list = []
Overdue = 0
for ws in wb.worksheets:
     #Move along worksheets
     ws = wb[sheets[n]]
     n = n + 1

      #Create a list, convert to int  
     for row in ws.iter_rows(min_row=2,min_col=1):
        list.append(row[1].value)
        list = [int(i) for i in list]
        #list = list.clear()
#total.append(list)       
           
      
        
print(list)
        