import pandas as pd




# read origin file
main_dict = pd.read_excel("Origin.xlsx", sheet_name=None)
sheet_names=list(main_dict.keys())


# make sheet data as dictionary
position_ID = main_dict[sheet_names[0]].to_dict(orient='records')
system_ID = main_dict[sheet_names[1]].to_dict(orient='records')
trade_ID = main_dict[sheet_names[2]].to_dict(orient='records')
abzardaghigh = main_dict[sheet_names[3]].to_dict(orient='records')
automasion = main_dict[sheet_names[4]].to_dict(orient='records')
labratory = main_dict[sheet_names[5]].to_dict(orient='records')
IT = main_dict[sheet_names[6]].to_dict(orient='records')
transport = main_dict[sheet_names[7]].to_dict(orient='records')
nasouz = main_dict[sheet_names[8]].to_dict(orient='records')


#make daftarfanni text file

f = open("daftar_fanni.txt", "w",encoding='utf-16')
f.write('SELECT        FQ.PositionID\n')
f.write('FROM            (\n')

for position in position_ID:
    position_name = position['Employee']
    for trade in trade_ID:
        trade_name = trade['Name']
        for system in system_ID:
            code_system = system['کد سیستم']
            
            for items in abzardaghigh:
                if items['کد سیستم'] == code_system and items['نوع کار']==trade_name and items['نام کارشناس دفتر فنی']==position_name:
                    text_line='UNION ALL SELECT \'{}\' as ParentSystemID, \'{}\' as WoTradeID, \'{}\' as PositionID --{}-{}-{}\n'.format(system['ID'],trade['ID'],position['Position ID'],code_system,position_name,trade_name)
                    f.write(text_line)
                            
            for items in automasion:
                if items['کد سیستم'] == code_system and items['نوع کار']==trade_name and items['نام کارشناس دفتر فنی']==position_name:
                    text_line='UNION ALL SELECT \'{}\' as ParentSystemID, \'{}\' as WoTradeID, \'{}\' as PositionID --{}-{}-{}\n'.format(system['ID'],trade['ID'],position['Position ID'],code_system,position_name,trade_name)
                    f.write(text_line)
  
            for items in labratory:
                if items['کد سیستم'] == code_system and items['نوع کار']==trade_name and items['نام کارشناس دفتر فنی']==position_name:
                    text_line='UNION ALL SELECT \'{}\' as ParentSystemID, \'{}\' as WoTradeID, \'{}\' as PositionID --{}-{}-{}\n'.format(system['ID'],trade['ID'],position['Position ID'],code_system,position_name,trade_name)
                    f.write(text_line)
  
            for items in transport:
                if items['کد سیستم'] == code_system and items['نوع کار']==trade_name and items['نام کارشناس دفتر فنی']==position_name:
                    text_line='UNION ALL SELECT \'{}\' as ParentSystemID, \'{}\' as WoTradeID, \'{}\' as PositionID --{}-{}-{}\n'.format(system['ID'],trade['ID'],position['Position ID'],code_system,position_name,trade_name)
                    f.write(text_line)
  
            for items in transport:
                if items['کد سیستم'] == code_system and items['نوع کار']==trade_name and items['نام کارشناس دفتر فنی']==position_name:
                    text_line='UNION ALL SELECT \'{}\' as ParentSystemID, \'{}\' as WoTradeID, \'{}\' as PositionID --{}-{}-{}\n'.format(system['ID'],trade['ID'],position['Position ID'],code_system,position_name,trade_name)
                    f.write(text_line)
  

f.write(') AS FQ RIGHT OUTER JOIN\n')
f.write('dbo.WorkOrder ON FQ.ParentSystemID =\n')
f.write('dbo.WorkOrder.ParentSystemID AND FQ.WoTradeID =\n')
f.write('dbo.WorkOrder.WOTradeID\n')
f.write('WHERE (WorkOrder.ID LIKE \'{0}\')\n')
f.close()



'''
for items in nasouz:
    print(items['نام کارشناس دفتر فنی'])
    #line_text='{} {}\n'.format(items['Name'],items['ID'])


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