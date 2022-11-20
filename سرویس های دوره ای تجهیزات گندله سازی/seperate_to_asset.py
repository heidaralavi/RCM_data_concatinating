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

#seperate ME 

df=pd.read_excel("service_list.xlsx",dtype=str)
filter_by_ME = df['location'].str.contains('ME')
df['filter']=filter_by_ME
df[~df['filter']].to_excel('non_electromotors_services.xlsx',index=False)
df[df['filter']].to_excel('electromotors_services.xlsx',index=False)
del df


#seperate GB 

df=pd.read_excel("non_electromotors_services.xlsx",dtype=str)
filter_by_GB = df['location'].str.contains('GB')
df['filter_gb']=filter_by_GB
df[~df['filter_gb']].to_excel('non_gearbox_services.xlsx',index=False)
df[df['filter_gb']].to_excel('gearbox_services.xlsx',index=False)
del df



#make app asset_no for non_electromotor_geabox elements

df=pd.read_excel("non_gearbox_services.xlsx",dtype=str)
df['temp1']=df['location'].str.split('.', expand=True)[0]
df['area']=df['temp1'].str.split('(\d+)', expand=True)[1]
df['area']=df['area'].astype(str).str.zfill(4)
df['tajhiz']=df['temp1'].str.split('(\d+)', expand=True)[2]
df['tajhiz']=df['tajhiz'].str[0:2]
df['no']=df['temp1'].str.split('(\d+)', expand=True)[3]
df['no']=df['no'].astype(str).str.zfill(2)
df['app_asset_no']=df['area']+df['tajhiz']+df['no']
df.to_excel('regenerate_app_asset_no.xlsx',index=False)
del df

#vlookup for category-class
df1=pd.read_excel("regenerate_app_asset_no.xlsx",dtype=str)
df2=pd.read_excel("for_vlookup.xlsx",dtype=str)
inner_join = pd.merge(df1, 
                      df2, 
                      on ='app_asset_no', 
                      how ='inner')
inner_join['1']=' ('
inner_join['2']=')'
inner_join['tozihat_taxonomy']=inner_join['tozihat']+inner_join['1']+inner_join['joze_machine']+inner_join['2']

inner_join['service_no_taxonomy']=inner_join['joze_machine']+inner_join['service']

inner_join.to_excel('final.xlsx',index=False)
del df1,df2,inner_join

#group by calegory-class

df=pd.read_excel("final.xlsx",dtype=str)
#df['category-class'].fillna(0, inplace=True)
#df['app_asset_code'].fillna(0, inplace=True)

col_names = df.columns.values


gb_category = df.groupby('category-class')  

for item in gb_category.groups:
    #print(item)
    #print(gb_category.get_group(item))
    sub_df1=gb_category.get_group(item)
    gb_priod=sub_df1.groupby('priod')
    for item2 in gb_priod.groups:
        sub_df2=gb_priod.get_group(item2)
        gb_app_asset=sub_df2.groupby('app_asset_no')
        for item3 in gb_app_asset.groups:
            directory = ".\\test\\{}\\".format(item) 
            file_name = ".\\test\\{}\\{}-{}-{}.xlsx".format(item,item,item3,item2)
            if not os.path.exists(directory):
                os.makedirs(directory)
            temp=gb_app_asset.get_group(item3).reset_index()
            temp=temp.drop_duplicates(subset=['service_no_taxonomy'],keep='first').sort_values(by=['tozihat_taxonomy'],ascending=True)
            temp.drop(['index'],axis=1,inplace=True)
            temp[['Asset Number','machine_code','app_asset_no','priod',
                  'tozihat_taxonomy','vahede_ejraii','active']].to_excel(file_name,index=False)
    

