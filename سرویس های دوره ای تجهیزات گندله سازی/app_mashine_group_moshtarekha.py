import pandas as pd
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
        

dore_service_dict=pd.read_excel("dore_service.xlsx").to_dict(orient='records')

groups_dict=pd.read_excel("App_groups_Vs_Parseh.xlsx").to_dict(orient='records')

df=pd.read_excel("service_list_group_by_use_power_query.xlsx",dtype=str)
#print(df.head())

#group by calegory-class

col_names = df.columns.values


gb_category = df.groupby('app_machine_group')  

for item in gb_category.groups:
    #print(item)
    #print(gb_category.get_group(item))
    sub_df1=gb_category.get_group(item)
    gb_priod=sub_df1.groupby('priod')
    
    for item2 in gb_priod.groups:
        sub_df2=gb_priod.get_group(item2)
        gb_noekar=sub_df2.groupby('noekar.ID')
        
        for item3 in gb_noekar.groups:
            
            i=1
            cart_faaliat_code='PM.{}.{}.{}.{}'.format(item3,find_dore(int(item2))[1],find_groups(item),str(i).zfill(4))
            directory = ".\\moshtarak\\{}\\".format(item) 
            file_name = ".\\moshtarak\\{}\\{}-{}.xlsx".format(item,cart_faaliat_code,item2)
            
            if not os.path.exists(directory):
                os.makedirs(directory)
            temp=gb_noekar.get_group(item3).reset_index()
            temp=temp.drop_duplicates(subset=['service_no_taxonomy'],keep='first').sort_values(by=['tozihat_taxonomy'],ascending=True)
            text='پی ام های {} - {} تجهیزات ({}) '.format(temp['noekar.نوع کار'][0],find_dore(int(item2))[0],item)
            temp.reset_index(inplace=True)
            temp.drop(['index'],axis=1,inplace=True)
            output_col=['tozihat_taxonomy','active']
            temp=temp[output_col]
            temp.rename(columns={'tozihat_taxonomy': 'شرح و دستورالعمل'}, inplace=True)
            temp.insert(loc=0, column='کد کارت فعالیت', value=cart_faaliat_code)
            temp.insert(loc=1, column='ترتیب', value= temp.index+1)
            temp.insert(loc=2, column='شرح کار یا فعالیت', value= text)
            temp.insert(loc=3, column='زمان انجام دقیقه', value= '')
            temp.insert(loc=4, column='زمان انجام ساعت', value= '')
            temp.insert(loc=6, column='نکات ایمنی', value= '')
            temp.insert(loc=7, column='نرخ انجام', value= '')

            temp.to_excel(file_name,index=False)
            
            
        
        
     
        
        
        
        
        
        

