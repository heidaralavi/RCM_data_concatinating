import pandas as pd
import numpy as np
import os


df=pd.read_excel("service_list_group_by_use_power_query.xlsx",dtype=str)
print(df.head())

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
        gb_app_asset=sub_df2.groupby('machine_code - Copy.1.1.1')
        for item3 in gb_app_asset.groups:
            directory = ".\\test\\{}\\".format(item) 
            file_name = ".\\test\\{}\\{}-{}-{}.xlsx".format(item,item,item3,item2)
            if not os.path.exists(directory):
                os.makedirs(directory)
            temp=gb_app_asset.get_group(item3).reset_index()
            temp=temp.drop_duplicates(subset=['service_no_taxonomy'],keep='first').sort_values(by=['tozihat_taxonomy'],ascending=True)
            temp.drop(['index'],axis=1,inplace=True)
            output_col=['location','machine_code','tozihat_taxonomy','priod','noe_service',
                        'vahede_ejraii','active','Table3.category-class']
            temp[output_col].to_excel(file_name,index=False)
    

