import pandas as pd
import numpy as np
import os


df=pd.read_excel("for_vlookup.xlsx",dtype=str)

used_col=['app_machine_group','Column1','category-class']
df=df[used_col]
#print(df.head())
#group by calegory-class



gb_category = df.groupby('app_machine_group')  
summery=pd.DataFrame()

for item in gb_category.groups:
    sub_df1=gb_category.get_group(item)
    sub_df1=sub_df1.drop_duplicates(subset=['category-class'],keep='first')
    summery = pd.concat([sub_df1, summery], axis=0)


summery.to_excel('temp.xlsx',index=False)
    
