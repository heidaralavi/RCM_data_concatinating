import pandas as pd
import numpy as np


df=pd.read_excel("service_list.xlsx")

gb = df.groupby('asset')    

for item in gb.groups:
    print(item)
    print(gb.get_group(item))
    file_name = ".\\test\{}.xlsx".format(item)
    gb.get_group(item).to_excel(file_name,index=False)
    print(file_name)
    
    
