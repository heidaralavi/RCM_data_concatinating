import pandas as pd
import numpy as np

# read origin file
main_dict = pd.read_excel("Origin.xlsx", sheet_name=None,keep_default_na=True)
sheet_names=list(main_dict.keys())


# make sheet data as dictionary
position_ID = main_dict[sheet_names[0]].to_dict(orient='records')
system_ID = main_dict[sheet_names[1]].to_dict(orient='records')
trade_ID = main_dict[sheet_names[2]].to_dict(orient='records')
abzardaghigh = main_dict[sheet_names[3]].to_dict(orient='records')
automasion = main_dict[sheet_names[4]].to_dict(orient='records')
labratory = main_dict[sheet_names[5]].to_dict(orient='records')
transport = main_dict[sheet_names[6]].to_dict(orient='records')
nasouz = main_dict[sheet_names[7]].to_dict(orient='records')
hydrolic = main_dict[sheet_names[8]].to_dict(orient='records')
mechanic = main_dict[sheet_names[9]].to_dict(orient='records')
tasisat = main_dict[sheet_names[10]].to_dict(orient='records')
bargh = main_dict[sheet_names[11]].to_dict(orient='records')



#ID returner functions
def position_id_returner(text):
    for item in position_ID:
        if item['Employee'] == text:
            return item['Position ID']

def system_id_returner(text):
    for item in system_ID:
        if item['کد سیستم'] == text:
            return item['ID']

def trade_id_returner(text):
    for item in trade_ID:
        if item['Name'] == text:
            return item['ID']
        

#body text generator
def make_text_body(body,naghsh):
    for items in body:
        code_system = items['کد سیستم']
        trade_name = items['نوع کار']
        position_name = str(items[naghsh])
        if position_name != "nan" :
            text_line='UNION ALL SELECT \'{}\' as ParentSystemID, \'{}\' as WoTradeID, \'{}\' as PositionID --{}-{}-{}\n'.format(system_id_returner(code_system),trade_id_returner(trade_name),position_id_returner(position_name),code_system,position_name,trade_name)
            f.write(text_line)


#header and footer
def make_file(fname,vahed,naghsh):
    global f
    f = open(fname, "w",encoding='utf-16')
    f.write('SELECT        FQ.PositionID\n')
    f.write('FROM            (\n')
    for i in range(len(vahed)):
        make_text_body(vahed[i],naghsh)
    f.write(') AS FQ RIGHT OUTER JOIN\n')
    f.write('dbo.WorkOrder ON FQ.ParentSystemID =\n')
    f.write('dbo.WorkOrder.ParentSystemID AND FQ.WoTradeID =\n')
    f.write('dbo.WorkOrder.WOTradeID\n')
    f.write('WHERE (WorkOrder.ID LIKE \'{0}\')\n')
    f.close()



#make summery file from all 
vahed_ejraii = [mechanic,abzardaghigh,automasion,labratory,transport,nasouz,hydrolic,tasisat,bargh]
#vahed_ejraii = [mechanic]

make_file('pamidco_summery.txt',vahed_ejraii,"نام کارشناس دفتر فنی")


make_file('nezarat_summery.txt',vahed_ejraii,"نام شخص کارشناس نظارت")







