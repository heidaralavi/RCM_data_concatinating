import pandas as pd
import numpy as np


df=pd.read_excel("service_list.xlsx")
df['category-class'].fillna(0, inplace=True)

gb = df.groupby('category-class')    

for item in gb.groups:
    print(item)
    print(gb.get_group(item))
    sub_df=gb.get_group(item)
    sub_gb=sub_df.groupby('priod')
    for item2 in sub_gb.groups:
        file_name = ".\\test\{}-{}.xlsx".format(item,item2)
        sub_gb.get_group(item2).to_excel(file_name,index=False)
    print(file_name)
    
    
