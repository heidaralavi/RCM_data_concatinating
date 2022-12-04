import pandas as pd


#read sheet names as list

#xl = pd.ExcelFile('Origin.xlsx')
#sheet_names=xl.sheet_names
#print(sheet_names[0])


main_dict = pd.read_excel("Origin.xlsx", sheet_name=None)
sheet_names=list(main_dict.keys())


print(type(sheet_names))
position_ID = main_dict[sheet_names[0]].to_dict(orient='records')
system_ID = main_dict[sheet_names[1]].to_dict(orient='records')
trade_ID = main_dict[sheet_names[2]].to_dict(orient='records')
anzardaghigh = main_dict[sheet_names[3]].to_dict(orient='records')
automasion = main_dict[sheet_names[4]].to_dict(orient='records')
labratory = main_dict[sheet_names[5]].to_dict(orient='records')
IT = main_dict[sheet_names[6]].to_dict(orient='records')
transport = main_dict[sheet_names[7]].to_dict(orient='records')
nasouz = main_dict[sheet_names[8]].to_dict(orient='records')




 
#print(trade_ID)

for items in nasouz:
    print(items['نام کارشناس دفتر فنی'])
    #line_text='{} {}\n'.format(items['Name'],items['ID'])

'''
for sheet in sheet_names:
    df=pd.read_excel("Origin.xlsx",dtype=str, sheet_name=sheet)
    dic=df.to_dict(orient='records')
    















df=pd.read_excel("a.xlsx",dtype=str,, sheet_name='Employees')
f = open("demofile3.txt", "w")


dic=df.to_dict(orient='records')

for items in dic:
    print(items['name'],items['id'])
    line_text='{} {}\n'.format(items['name'],items['id'])
    f.write(line_text)
    
f.close()
'''