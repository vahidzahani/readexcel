from openpyxl import load_workbook
import re

book = load_workbook('main.xlsx')
sheet=book['Sheet1']
rows=sheet.rows
head=[cell.value for cell in next(rows)]

extract_phone_number_pattern = "\\+?[1-9][0-9]{7,14}"
# a=re.findall(extract_phone_number_pattern, 'You can reach me out at +989989978798 and +56667778888') # returns ['+12223334444', '+56667778888']

# print(a[0])
print("start")
f=open("main.csv","w")
for now in rows:
    data={}
    for title,cell in zip(head,now):
        data[title] = cell.value
    a=str(data['from'])
    b=str(data['to'])
    x=re.findall(extract_phone_number_pattern, a) # returns ['+12223334444', '+56667778888']
    y=re.findall(extract_phone_number_pattern, b) # returns ['+12223334444', '+56667778888']
    #print(x[0] + "," + y[0])
    f.writelines(x[0] + "," + y[0] + "\n")
f.close()
print("end")


    

        # data[title]=cell.value
        # print(data['from'])
        #print(data)
        #print(cell.value + "," cell.value )