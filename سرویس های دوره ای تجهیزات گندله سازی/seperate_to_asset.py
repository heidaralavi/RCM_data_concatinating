import pandas as pd
import numpy as np
import os

col_names = ['location','machine_code','joze_machine','sharhe_service_fa',
             'service','tozihat','zamane_anjam','zamane_standard','priod',
             'noe_service','tarikh_anjam','maharat','vahede_ejraii','active']
df=pd.read_excel("service_list_row.xlsx",dtype=str,names=col_names)


#data cleaning
df = df.replace('ك', 'ک', regex=True)
df = df.replace('ي', 'ی', regex=True)
df = df.replace(chr(10),' ',regex=True) #Two Line replace by one Line

for n in range(6):
    df = df.replace('  ',' ',regex=True)

for item in col_names:
    df[item]=df[item].astype(str).str.strip()

#seperate data 
first_char=df['location'].str[0]
map = first_char.str.isdigit()
df[map].to_excel('service_list.xlsx',index=False)
df[~map].to_excel('not_procced.xlsx',index=False)
del df

#seperate EM 

df=pd.read_excel("service_list.xlsx",dtype=str)
filter_data = df['location'].str.contains('ME')
df['map']=filter_data




df.to_excel('temp.xlsx')

'''
df['area']=df['location'][map].str.split('.', expand=True)[0]

df['area1']=df['area'][map].str.split('(\d+)', expand=True)[1]
df['area1']=df['area1'][map].astype(str).str.zfill(4)
df['tajhiz']=df['area'][map].str.split('(\d+)', expand=True)[2]
df['tajhiz']=df['tajhiz'].str[0:2]
df['asset']=df['area1']+df['tajhiz']
'''




'''
df['category-class'].fillna(0, inplace=True)
df['app_asset_code'].fillna(0, inplace=True)

col_names = df.columns.values


gb_category = df.groupby('category-class')  

for item in gb_category.groups:
    #print(item)
    #print(gb_category.get_group(item))
    sub_df1=gb_category.get_group(item)
    gb_priod=sub_df1.groupby('priod')
    for item2 in gb_priod.groups:
        sub_df2=gb_priod.get_group(item2)
        gb_app_asset=sub_df2.groupby('app_asset_code')
        for item3 in gb_app_asset.groups:
            directory = ".\\test\\{}\\".format(item) 
            file_name = ".\\test\\{}\\{}-{}-{}.xlsx".format(item,item,item3,item2)
            if not os.path.exists(directory):
                os.makedirs(directory)
            temp=gb_app_asset.get_group(item3).reset_index()
            temp=temp.drop_duplicates(subset=['service_no_taxonomy'],keep='first').sort_values(by=['tozihat_taxonomy'],ascending=True)
            temp.drop(['index'],axis=1,inplace=True)
            temp[['asset','Asset Number','category-class','tozihat_taxonomy','priod','vahed_ejraii','active']].to_excel(file_name,index=False)
    
'''    
