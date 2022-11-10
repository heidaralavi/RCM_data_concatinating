#!pip install treelib

from treelib import Node,Tree
import pandas as pd
import os

fileslist=os.listdir('./')
file_lists=[]
for filename in fileslist:
    if filename.endswith('.xlsx'):
        file_lists.append(filename)
        
#print(file_lists)


for fname in file_lists:
    
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
        print(fname,i)
        if row[7] == 0:
            tree.create_node(row[2],row[6],parent='top')
        else:
            tree.create_node(row[2],row[6],parent=row[7])
        i += 1


    tree.show()

    textname=format(fname[:-5]+'.txt')
    print(textname)
    tree.save2file(textname)
    del df
    del tree
    
    