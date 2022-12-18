import pandas as pd
import numpy as np



# make sheet data as dataframe
position_ID = pd.read_excel("Origin.xlsx", sheet_name='Position_ID')
system_ID = pd.read_excel("Origin.xlsx", sheet_name='System_ID')
trade_ID = pd.read_excel("Origin.xlsx", sheet_name='Trade_ID')
ejraii_ID = pd.read_excel("Origin.xlsx", sheet_name='واحد اجرایی')
abzardaghigh = pd.read_excel("Origin.xlsx", sheet_name='ابزاردقیق')
automasion = pd.read_excel("Origin.xlsx", sheet_name='اتوماسیون')
labratory = pd.read_excel("Origin.xlsx", sheet_name='آزمایشگاه و کنترل فرایند')
transport = pd.read_excel("Origin.xlsx", sheet_name='عمرانی خدماتی ترانسپورت')
nasouz = pd.read_excel("Origin.xlsx", sheet_name='نسوز')
hydrolic = pd.read_excel("Origin.xlsx", sheet_name='هیدرولیک و روانکاری')
mechanic = pd.read_excel("Origin.xlsx", sheet_name='مکانیک')
tasisat = pd.read_excel("Origin.xlsx", sheet_name='تاسیسات آبرسانی')
bargh = pd.read_excel("Origin.xlsx", sheet_name='برق')



vahed_ejraii = [mechanic,abzardaghigh,automasion,labratory,transport,nasouz,hydrolic,tasisat,bargh]
#vahed_ejraii = [mechanic]


for radif,vahed in enumerate(vahed_ejraii):
    for index,items in vahed.iterrows():
        code_system_ID = system_ID['ID'].loc[system_ID['کد سیستم']==items['کد سیستم']].values[0]
        trade_name_ID = trade_ID['ID'].loc[trade_ID['Name'] == items['نوع کار']].values[0]
        position_name_ID = position_ID['Position ID'].loc[position_ID['Employee'] == items['نام کارشناس دفتر فنی']].values
        text='hello {}'.format(position_name_ID[0])
        print(radif,text)
        













'''















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

def make_ejraii_file(fname,column):
    global f
    f = open(fname, "w",encoding='utf-16')
    f.write('select case\n')
    for item in ejraii_ID:
        text_line='when DepartmentID = \'{}\' then \'{}\''.format(item['ID'],position_id_returner(item[column]))
        f.write(text_line)
        f.write('\n')
    f.write('else \'7c70790f-81e3-4efd-ae03-700d677984bd\'\n')
    f.write('end\n')
    f.write('from dbo.WorkOrder\n')
    f.write('WHERE (WorkOrder.ID LIKE \'{0}\')\n')
    f.close()

#make summery file from all 
vahed_ejraii = [mechanic,abzardaghigh,automasion,labratory,transport,nasouz,hydrolic,tasisat,bargh]
#vahed_ejraii = [mechanic]

make_file('pamidco_summery.txt',vahed_ejraii,"نام کارشناس دفتر فنی")


make_file('nezarat_summery.txt',vahed_ejraii,"نام شخص کارشناس نظارت")

make_ejraii_file('raiis_ejraii_summery.txt','رئیس اجرایی')
make_ejraii_file('sarparast_ejraii_summery.txt','سرپرست واحد اجرایی')


'''




