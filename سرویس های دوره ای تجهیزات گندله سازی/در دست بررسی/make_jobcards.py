import pandas as pd
import os
from xlsxwriter import Workbook

category_name = pd.read_excel('class-group-vlockup.xlsx',usecols=['class-group','FullCategoryName','App_Tag']).to_dict(orient='records')
code_faaliat_mojri = pd.read_excel('vlookup_code_faaliat_mijri.xlsx').to_dict(orient='records')
def class_group(text):
    for item in category_name:
        if item['class-group'] == text:
            return item['FullCategoryName']
def faaliat_mojri(text):
    for item in code_faaliat_mojri:
        if item['code_faaliat'] == text:
            return item['mojri']
        
def app_tag(text):
    for item in category_name:
        if item['class-group'] == text:
            return item['App_Tag']


def trans_txt(text):
    arr = text.split('.')
    if arr[5][:2] =='SR':
        code_tajhiz='ME-Slip Ring'
    elif arr[5][:2] =='GM':
        code_tajhiz='ME-Grease'
    elif arr[5][:2] =='LQ':
        code_tajhiz='ME-Liquid Starter'
    elif arr[5][:2] =='BR':
        code_tajhiz='ME-Brake'
    elif arr[5][:3] =='BCG':
        code_tajhiz='BC-Gearbox'        
    elif arr[5][:3] =='WFB':
        code_tajhiz='WF-Belt'
    elif arr[5][:3] =='WFS':
        code_tajhiz='WF-Screw'        
        
    else:
        code_tajhiz=arr[5][:2]
    output_txt = '{} {} {} تجهیزات ({}) - {}'.format(vocablary[arr[0]],vocablary[arr[1]][0],vocablary[arr[2]],code_tajhiz,vocablary[arr[3]])
    return output_txt

def make_row(text):
    arr = text.split('.')
    row = []
    row.append(trans_txt(text))
    row.append(text)
    row.append(vocablary[arr[2][1]])
    row.append(arr[2][0])
    for i in range(7):
        row.append('')
    row.append('Active')
    for i in range(2):
        row.append('')
    row.append('سرویس دوره ای - PM')
    row.append(vocablary[arr[1]][1])
    row.append(class_group(arr[4]))
    row.append(faaliat_mojri(text))
    print(faaliat_mojri(text))
    row.append('Hour')
    row.append('')
    row.append('')
    return(row)

def file_list(mypath):
    f_list=[]
    for root, dirs, files in os.walk(mypath):
        for f in files:
            if f.endswith('.xlsx') and f.startswith('PM'):
                f_list.append(f[:-5])
    return f_list
    



vocablary = {}
vocablary['PM']='سرویس دوره ای نت'
vocablary['ME']=('مکانیکی','مکانیک')
vocablary['EL']=('برقی','برق')
vocablary['IN']=('ابزاردقیقی','ابزار دقیق')     
vocablary['UT']=('تاسیساتی','تاسیسات/آبرسانی')
vocablary['HY']=('مکانیکی','هیدرولیک')
vocablary['LB']=('آزمایشگاه','آزمایشگاه')
vocablary['TR']=('ترانسپورت','ترانسپورت')      
vocablary['1W']='هفتگی'  
vocablary['2W']='دوهفته یکبار'
vocablary['1M']='ماهیانه'
vocablary['2M']='دو ماه یکبار'
vocablary['3M']='سه ماه یکبار'
vocablary['6M']='شش ماه یکبار'
vocablary['1Y']='سالیانه'
vocablary['RUN']='بدون نیاز به توقف'
vocablary['STP']='نیاز به توقف'
vocablary['W']='Week'
vocablary['M']='Month'
vocablary['Y']='Year'

print(os.getcwd().replace('\\','/'))
#mypath = 'C:/Users/heidaralavi/Documents/GitHub/tasvinshodeha-pm'
mypath = os.getcwd().replace('\\','/')

f_list_txt = file_list(mypath)

rows = []
for item in f_list_txt:
    rows.append(make_row(item))

#col=pd.read_excel("jobcard.xlsx",dtype=str)

col_names = ['Name نام کارت فعالیت','Code کد کارت فعالت','CalendarUnit واحد تقویمی','CalendarPeriod دوره تقویمی','CalendarPeriodPlus شناوری مثبت تقویمی','calendarPeriodMinus شناوری منفی تقویمی','MeterUnit واحد کارکردی','MeterPeriod دوره کارکردی','MeterPeriodPlus شناوری کارکردی مثبت','MeterPeriodMinus شناوری کارکردی منفی','SafetyInstruction نکته ایمنی','PlanningStatus وضعیت کارت فعالیت','CalendarCoverageLimit حد همپوشانی تقویمی','MeterCoverageLimit حد همپوشانی کارکردی','JCType نوع کارت فعالیت','WorkTrade نوع کار','AssetClass کلاس دستگاه','Department مجری','DurationUnit واحد انجام','Duration مدت انجام','Problem کد ایراد و مشکل'] #col.columns.values

df=pd.DataFrame(rows,columns=col_names)
writer = pd.ExcelWriter('jobcard.xlsx', engine='xlsxwriter')
df.to_excel(writer, startrow=1 ,sheet_name = 'sheet1',header=False,index=False)
workbook = writer.book
worksheet = writer.sheets['sheet1']
(max_row, max_col) = df.shape
column_settings = []
for header in df.columns:
    column_settings.append({'header': header})
worksheet.add_table(0, 0, max_row, max_col - 1, {'columns': column_settings})
worksheet.set_column(0, max_col - 1, 12)
worksheet.set_column('E:E', 40)
writer.close()



