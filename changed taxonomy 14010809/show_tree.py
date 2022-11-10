#!pip install treelib
from treelib import Node,Tree
import pandas as pd

fname='710TG01-DEHESTANI0.xlsx'
df=pd.read_excel(fname)

for index , row in df.iterrows():
    print(row[2]) #name joze
    print(row[6]) #serial nomber 
    print(row[7]) #sreial balasari
    

df=df.fillna(0)    
tree = Tree()

i=0
tree.create_node(fname,'top')
for index , row in df.iterrows():
    print(i)
    if row[7] == 0:
        tree.create_node(row[2],row[6],parent='top')
    else:
        tree.create_node(row[2],row[6],parent=row[7])
    i += 1


tree.show()

fname=format(fname+'.txt')
print(fname)
tree.save2file(fname)

    
    