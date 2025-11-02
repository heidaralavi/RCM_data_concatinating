import pandas as pd
import math



# make sheet data as dataframe
position_ID = pd.read_excel("Origin.xlsx", sheet_name='Position_ID',dtype=str).to_dict(orient='records')
system_ID = pd.read_excel("Origin.xlsx", sheet_name='System_ID',dtype=str).to_dict(orient='records')
trade_ID = pd.read_excel("Origin.xlsx", sheet_name='Trade_ID',dtype=str).to_dict(orient='records')
ejraii_ID = pd.read_excel("Origin.xlsx", sheet_name='واحد اجرایی',dtype=str).to_dict(orient='records')
abzardaghigh = pd.read_excel("Origin.xlsx", sheet_name='ابزاردقیق',dtype=str).to_dict(orient='records')
automasion = pd.read_excel("Origin.xlsx", sheet_name='اتوماسیون',dtype=str).to_dict(orient='records')
labratory = pd.read_excel("Origin.xlsx", sheet_name='کنترل کیفی',dtype=str).to_dict(orient='records')
transport1 = pd.read_excel("Origin.xlsx", sheet_name='عمرانی خدماتی',dtype=str).to_dict(orient='records')
transport2 = pd.read_excel("Origin.xlsx", sheet_name='ترانسپورت',dtype=str).to_dict(orient='records')
nasouz = pd.read_excel("Origin.xlsx", sheet_name='نسوز',dtype=str).to_dict(orient='records')
hydrolic = pd.read_excel("Origin.xlsx", sheet_name='هیدرولیک و روانکاری',dtype=str).to_dict(orient='records')
mechanic = pd.read_excel("Origin.xlsx", sheet_name='مکانیک',dtype=str).to_dict(orient='records')
tasisat = pd.read_excel("Origin.xlsx", sheet_name='تاسیسات آبرسانی',dtype=str).to_dict(orient='records')
bargh = pd.read_excel("Origin.xlsx", sheet_name='برق',dtype=str).to_dict(orient='records')
tolid = pd.read_excel("Origin.xlsx", sheet_name='نقش بهره بردار تولید',dtype=str).to_dict(orient='records')
    


#ID returner functions
def position_id_returner(text):
    for item in position_ID:
        if item['Position'] == text:
            return item['Position ID']

def name_returner(text):
    for item in position_ID:
        if item['Position'] == text:
            return item['Employee']

def system_id_returner(text):
    for item in system_ID:
        if item['کد سیستم'] == text:
            return item['ID']

def trade_id_returner(text):
    for item in trade_ID:
        if item['Name'] == text:
            return item['ID']
        


def make_naghsh(fname,naghsh='سمت کارشناس دفتر فنی'):
    vahed_ejraii = [mechanic,abzardaghigh,automasion,labratory,transport1,transport2,nasouz,hydrolic,tasisat,bargh]
    #vahed_ejraii = [mechanic]
    f = open(fname, "w",encoding='utf-16')
    f.write('SELECT        FQ.PositionID\n')
    f.write('FROM            (\n')
        
    for radif,vahed in enumerate(vahed_ejraii):
        i=0
        for items in vahed:
            code_system_ID =system_id_returner(items['کد سیستم'])
            trade_name_ID = trade_id_returner(items['نوع کار'])
            #position_name_ID = items[naghsh]
            position_name_ID = position_id_returner(items[naghsh])
            employee_name = name_returner(items[naghsh])
                        
            if radif == 0 and i ==0:
                text_line='SELECT \'{}\' as ParentSystemID, \'{}\' as WoTradeID, \'{}\' as PositionID --{}-{}-{}\n'.format(code_system_ID,trade_name_ID,position_name_ID,items['کد سیستم'],employee_name,items['نوع کار'])
                f.write(text_line)
                
            i=i+1
            text_line='UNION ALL SELECT \'{}\' as ParentSystemID, \'{}\' as WoTradeID, \'{}\' as PositionID --{}-{}-{}\n'.format(code_system_ID,trade_name_ID,position_name_ID,items['کد سیستم'],employee_name,items['نوع کار'])
            f.write(text_line)
            
    
    f.write(') AS FQ RIGHT OUTER JOIN\n')
    f.write('dbo.WorkOrder ON FQ.ParentSystemID =\n')
    f.write('dbo.WorkOrder.ParentSystemID AND FQ.WoTradeID =\n')
    f.write('dbo.WorkOrder.WOTradeID\n')
    f.write('WHERE (WorkOrder.ID LIKE \'{0}\')\n')
    f.close()
    


make_naghsh('نقش دفترفنی.txt',naghsh='سمت کارشناس دفتر فنی')
make_naghsh('نقش نظارت.txt',naghsh='سمت کارشناس نظارت')



    

def make_ejraii_file(fname,column):
    global f
    f = open(fname, "w",encoding='utf-16')
    f.write('select case\n')
    for item in ejraii_ID:
        text_line='when DepartmentID = \'{}\' then \'{}\' --{}-{}'.format(item['ID'],position_id_returner(item[column]),name_returner(item[column]),item[column])
        f.write(text_line)
        f.write('\n')
    f.write('else \'7c70790f-81e3-4efd-ae03-700d677984bd\'\n')
    f.write('end\n')
    f.write('from dbo.WorkOrder\n')
    f.write('WHERE (WorkOrder.ID LIKE \'{0}\')\n')
    f.close()

#make summery file from all 


make_ejraii_file('نقش رییس اجرایی.txt','سمت رئیس اجرایی')
make_ejraii_file('نقش سرپرست اجرایی.txt','سمت سرپرست اجرایی')
make_ejraii_file('نقش سرپرست اجرایی پی ام.txt','سمت سرپرست اجرایی پی ام')


def make_naghsh_tolid(fname):
    f = open(fname, "w",encoding='utf-16')
    f.write('select case\n')
    for item in tolid:
        persion_counter = 0
        position_name = []
        if str(item['سمت بهره بردار تولید 1']) != "nan" :
            persion_counter += 1
            position_name.append(item['سمت بهره بردار تولید 1'])
        if str(item['سمت بهره بردار تولید 2']) != "nan" :
            persion_counter += 1
            position_name.append(item['سمت بهره بردار تولید 2'])
        if str(item['سمت بهره بردار تولید 3']) != "nan" :
            persion_counter += 1 
            position_name.append(item['سمت بهره بردار تولید 3'])
        if str(item['سمت بهره بردار تولید 4']) != "nan" :
            persion_counter += 1
            position_name.append(item['سمت بهره بردار تولید 4'])
    
        match persion_counter:
            case 1:
                text_line = f'when ParentSystemID = \'{system_id_returner(item['کد سیستم'])}\' then \'{position_id_returner(position_name[0])}\' --{item['کد سیستم']} / {name_returner(position_name[0])}'
            case 2:
                text_line = f'when ParentSystemID = \'{system_id_returner(item['کد سیستم'])}\' then \'{position_id_returner(position_name[0])}\' + \',\' + \'{position_id_returner(position_name[1])}\'  --{item['کد سیستم']} / {name_returner(position_name[0])}+{name_returner(position_name[1])}'
            case 3:
                text_line = f'when ParentSystemID = \'{system_id_returner(item['کد سیستم'])}\' then \'{position_id_returner(position_name[0])}\' + \',\' + \'{position_id_returner(position_name[1])}\' + \',\' + \'{position_id_returner(position_name[2])}\'  --{item['کد سیستم']} / {name_returner(position_name[0])}+{name_returner(position_name[1])}+{name_returner(position_name[2])}'
            case 4:
                text_line = f'when ParentSystemID = \'{system_id_returner(item['کد سیستم'])}\' then \'{position_id_returner(position_name[0])}\' + \',\' + \'{position_id_returner(position_name[1])}\' + \',\' + \'{position_id_returner(position_name[2])}\' + \',\' + \'{position_id_returner(position_name[3])}\'  --{item['کد سیستم']} / {name_returner(position_name[0])}+{name_returner(position_name[1])}+{name_returner(position_name[2])}+{name_returner(position_name[3])}'
    
        if persion_counter > 0:
            f.write(text_line)
            f.write('\n')
    f.write('else OrganizationPos.ID --درخواست کننده\n')
    f.write('end as ID \n')
    f.write('from WorkOrder Inner Join\n')
    f.write('     Employee On WorkOrder.ReportByID = Employee.ID Inner Join\n')
    f.write('	      OrganizationPos On Employee.ID = OrganizationPos.EmployeeID\n')
    f.write('where (workorder.ID LIKE \'{0}\')\n')
    f.close()

make_naghsh_tolid('نقش بهره بردار دستگاه.txt')





