import os
import pandas as pd
import numpy as np



def list_of_files(path):
    myfiles=[]
    for mypath in path:
        for root, dirs, files in os.walk(mypath):
            for f in files:
                myfiles.append(root+'/'+f)
                #break #only one loop needed
    return myfiles

def sheets_name(exel_path):
    xl = pd.ExcelFile(exel_path)
    return np.array(xl.sheet_names).reshape(-1)


def asset_list(path):
    asset_list = np.empty(0)
    for f in list_of_files(path):
        asset_list=np.append(asset_list,sheets_name(f))
        
    return np.unique(asset_list)


if __name__ == "__main__":
    
    
    mydirectories= ['abdi','azizi','fahimi','ghafari','hakimzadeh',
                    'hossaini','kargar','mansourian','mehdizadeh',
                    'nazemizadeh','tabibi','voshtani']
    
    ass_list = asset_list(mydirectories)
    
    file_list=list_of_files(mydirectories)
    
    for f in file_list:
        names=sheets_name(f)
        print(names)
        if ass_list[0] in names:
            df= pd.read_excel(f, asset_list[0])
            print(df)
        
        
    
    
    