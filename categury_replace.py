import pandas as pd
import numpy as np

file_name='RCM Origenal Date - 1401-02-19 (Recovered).xlsx'
sheet_name='Class & Category'

df=pd.read_excel(file_name,sheet_name=sheet_name,header=None)
df.drop(index=[0,1], inplace=True)



vals = df.values

dic={}
for val in vals:
    string = '{}{}{}'.format(val[1],val[4],val[7])
    dic[string]=[val[0],val[3],val[6]]
    #print(string)
#print(dic)

df1=pd.DataFrame(dic)

df1.T.to_excel('aa.xlsx')


#if __name__ == "__main__":
    