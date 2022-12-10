import pandas as pd
import numpy as np
import os


def find_dore(dore):
    for items in dore_service_dict:
        #print(dore)
        if items['Column1'] == dore:
            return items['Column2'],items['Column3'],items['Column4'],items['Column5']
    
def find_groups(App_groups):
    for items in groups_dict:
        if items['app_machine_group'] == App_groups:
            return items['Column1']
        


df=pd.read_excel("service_list_group_by_use_power_query.xlsx",dtype=str)

col_names = df.columns.values

dore_service_dict=pd.read_excel("dore_service.xlsx").to_dict(orient='records')

groups_dict=pd.read_excel("App_groups_Vs_Parseh.xlsx").to_dict(orient='records')


jobcard_test_dict={'Name نام کارت فعالیت':[],'Code کد کارت فعالت':[],'CalendarUnit واحد تقویمی':[],
                   'CalendarPeriod دوره تقویمی ':[],'MeterUnit واحد کارکردی':[],'MeterPeriod دوره کارکردی':[],
                   'WorkTrade نوع کار':[],'AssetClass کلاس دستگاه':[],'Department مجری':[]}

gb_category = df.groupby('app_machine_group')  

for item in gb_category.groups:
    sub_df1=gb_category.get_group(item)
    gb_priod=sub_df1.groupby('priod')
    for item2 in gb_priod.groups:
        sub_df2=gb_priod.get_group(item2)
        gb_noekar=sub_df2.groupby('noekar.ID')
        for item3 in gb_noekar.groups:
            
            temp=gb_noekar.get_group(item3).reset_index()
            temp=temp.drop_duplicates(subset=['service_no_taxonomy'],keep='first').sort_values(by=['tozihat_taxonomy'],ascending=True)
            #print(temp['noekar.نوع کار'])
            i=1
            text='پی ام های {} - {} تجهیزات ({}) '.format(temp['noekar.نوع کار'][0],find_dore(int(item2))[0],item)
            #print(text)
            cart_faaliat_code='PM.{}.{}.{}.{}'.format(item3,find_dore(int(item2))[1],find_groups(item),str(i).zfill(4))
            i=i+1
            jobcard_test_dict['Name نام کارت فعالیت'].append(text)
            jobcard_test_dict['Code کد کارت فعالت'].append(cart_faaliat_code)
            if find_dore(int(item2))[2] == 'Hours':
                jobcard_test_dict['MeterUnit واحد کارکردی'].append(find_dore(int(item2))[2])
                jobcard_test_dict['MeterPeriod دوره کارکردی'].append(find_dore(int(item2))[3])
                jobcard_test_dict['CalendarPeriod دوره تقویمی '].append('')
                jobcard_test_dict['CalendarUnit واحد تقویمی'].append('')
            else:
                jobcard_test_dict['CalendarUnit واحد تقویمی'].append(find_dore(int(item2))[2])
                jobcard_test_dict['CalendarPeriod دوره تقویمی '].append(find_dore(int(item2))[3])
                jobcard_test_dict['MeterUnit واحد کارکردی'].append('')
                jobcard_test_dict['MeterPeriod دوره کارکردی'].append('')
 
            jobcard_test_dict['WorkTrade نوع کار'].append(temp['noekar.نوع کار'][0])
            jobcard_test_dict['AssetClass کلاس دستگاه'].append(temp['Table3.category-class'][0])
            jobcard_test_dict['Department مجری'].append(temp['mojri.Department مجری'][0])


jobcard_test=pd.DataFrame(jobcard_test_dict).drop_duplicates(subset=['Name نام کارت فعالیت'],keep='first')

jobcard_test.insert(loc=4, column='CalendarPeriodPlus شناوری مثبت تقویمی', value='')
jobcard_test.insert(loc=5, column='calendarPeriodMinus شناوری منفی تقویمی', value='')
jobcard_test.insert(loc=8, column='MeterPeriodPlus شناوری کارکردی مثبت' , value='')
jobcard_test.insert(loc=9, column='MeterPeriodMinus شناوری کارکردی منفی' , value='')
jobcard_test.insert(loc=10, column='SafetyInstruction نکته ایمنی' , value='')
jobcard_test.insert(loc=11, column='PlanningStatus وضعیت کارت فعالیت' , value='Active')
jobcard_test.insert(loc=12, column='CalendarCoverageLimit حد همپوشانی تقویمی' , value='')
jobcard_test.insert(loc=13, column='MeterCoverageLimit حد همپوشانی کارکردی' , value='')
jobcard_test.insert(loc=14, column='JCType نوع کارت فعالیت' , value='سرویس دوره ای - PM')
jobcard_test.insert(loc=18, column='DurationUnit واحد انجام' , value='Hour')
jobcard_test.insert(loc=19, column='Duration مدت انجام' , value=1)
jobcard_test.insert(loc=20, column='Problem کد ایراد و مشکل' , value='')


jobcard_test.to_excel('jobcard_test.xlsx',index=False)

