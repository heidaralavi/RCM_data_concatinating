import pandas as pd
import numpy as np
import os
from treelib import Node,Tree

fileslist=os.listdir('./')
file_lists=[]
for filename in fileslist:
    if filename.endswith('.xlsx'):
        file_lists.append(filename)
        
#print(file_lists)


data_items=np.empty(0)

for f in file_lists:
    df=pd.read_excel(f)
    #print(f,df.shape)
    data_items=np.append(data_items,df.values)
    del df

data_items=data_items.reshape(-1,11)
#print(data_items[-30:])
col_names=['AssetNumber','TagNo','نام جز','نوع جز','سطح تجهیز','پارت نامبر PN','سریال نامبر SN','جز بالاتر (سریال نامبر بالاسری)','مشخصه فنی','توضیحات','نوع فنی']
df = pd.DataFrame(data_items, columns = col_names)
#print(df)

df.to_excel('TOTAL-14010809.xlsx',sheet_name='asset',index=False)





