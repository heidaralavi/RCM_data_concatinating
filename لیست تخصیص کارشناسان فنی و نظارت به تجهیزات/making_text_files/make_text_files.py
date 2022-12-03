import pandas as pd


#read sheet names as list

#xl = pd.ExcelFile('Origin.xlsx')
#sheet_names=xl.sheet_names
#print(sheet_names[0])


main_dict = pd.read_excel("Origin.xlsx", sheet_name=None)
sheet_names=main_dict.keys()
print(type(sheet_names))

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