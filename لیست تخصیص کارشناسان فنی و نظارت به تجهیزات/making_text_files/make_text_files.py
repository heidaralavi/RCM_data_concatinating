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
transport = main_dict[sheet_names[6]].to_dict(orient='records')
nasouz = main_dict[sheet_names[7]].to_dict(orient='records')
#IT = main_dict[sheet_names[8]].to_dict(orient='records')


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
        position_name = items[naghsh]
        text_line='UNION ALL SELECT \'{}\' as ParentSystemID, \'{}\' as WoTradeID, \'{}\' as PositionID --{}-{}-{}\n'.format(system_id_returner(code_system),trade_id_returner(trade_name),position_id_returner(position_name),code_system,position_name,trade_name)
        f.write(text_line)


#header and footer
def make_file(fname,vahed,naghsh):
    global f
    f = open(fname, "w",encoding='utf-16')
    f.write('SELECT        FQ.PositionID\n')
    f.write('FROM            (\n')
    make_text_body(vahed,naghsh)
    f.write(') AS FQ RIGHT OUTER JOIN\n')
    f.write('dbo.WorkOrder ON FQ.ParentSystemID =\n')
    f.write('dbo.WorkOrder.ParentSystemID AND FQ.WoTradeID =\n')
    f.write('dbo.WorkOrder.WOTradeID\n')
    f.write('WHERE (WorkOrder.ID LIKE \'{0}\')\n')
    f.close()


make_file('پامیدکو آزمایشگاه و کنترل فرآیند.txt',labratory,"نام کارشناس دفتر فنی")
make_file('پامیدکو ابزاردقیق.txt',abzardaghigh,"نام کارشناس دفتر فنی")
make_file('پامیدکو اتوماسیون.txt',automasion,"نام کارشناس دفتر فنی")
make_file('پامیدکو ترانسپورت.txt',transport,"نام کارشناس دفتر فنی")
make_file('پامیدکو نسوز.txt',nasouz,"نام کارشناس دفتر فنی")


make_file('نظارت آزمایشگاه و فرآیند.txt',labratory,"نام شخص کارشناس نظارت")
make_file('نظارت ابزاردقیق.txt',abzardaghigh,"نام شخص کارشناس نظارت")
make_file('نظارت اتوماسیون.txt',automasion,"نام شخص کارشناس نظارت")
make_file('نظارت ترانسپورت.txt',transport,"نام شخص کارشناس نظارت")
make_file('نظارت نسوز.txt',transport,"نام شخص کارشناس نظارت")





