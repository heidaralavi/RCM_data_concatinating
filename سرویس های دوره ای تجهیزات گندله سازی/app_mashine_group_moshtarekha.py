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
        i=1
        cart_faaliat_code='PM.{}.{}.{}'.format(find_dore(int(item2))[1],find_groups(item),str(i).zfill(4))
        
        directory = ".\\moshtarak\\{}\\".format(item) 
        file_name = ".\\moshtarak\\{}\\{}-{}.xlsx".format(item,cart_faaliat_code,item2)
        if not os.path.exists(directory):
            os.makedirs(directory)
        temp=gb_priod.get_group(item2).reset_index()
        temp=temp.drop_duplicates(subset=['service_no_taxonomy'],keep='first').sort_values(by=['tozihat_taxonomy'],ascending=True)

        output_col=['location','machine_code','tozihat_taxonomy','priod','noe_service',
                    'vahede_ejraii','active','Table3.category-class']
        temp[output_col].to_excel(file_name,index=False)

        
     
        
        
        
        
        
        
        
        
'''        
        
        
        gb_app_asset=sub_df2.groupby('machine_code - Copy.1.1.1')
        for item3 in gb_app_asset.groups:
            directory = ".\\moshtarek\\{}\\".format(item) 
            file_name = ".\\moshtarek\\{}\\{}-{}-{}.xlsx".format(item,item,item3,item2)
            if not os.path.exists(directory):
                os.makedirs(directory)
            temp=gb_app_asset.get_group(item3).reset_index()
            temp=temp.drop_duplicates(subset=['service_no_taxonomy'],keep='first').sort_values(by=['tozihat_taxonomy'],ascending=True)
            temp.drop(['index'],axis=1,inplace=True)
            output_col=['location','machine_code','tozihat_taxonomy','priod','noe_service',
                        'vahede_ejraii','active','Table3.category-class']
            temp[output_col].to_excel(file_name,index=False)
    
'''
