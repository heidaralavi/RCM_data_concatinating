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


def find_files(asset,path):
    res=[]
    for f in list_of_files(path):
        sh_names=sheets_name(f)
        if asset in sh_names:
            res.append(f)
            
    return res
    
    
def make_dataframe(asset,files):
    col_names=['AssetNumber','TagNo','نام جز','نوع جز','سطح تجهیز','پارت نامبر PN','سریال نامبر SN','جز بالاتر (سریال نامبر بالاسری)','مشخصه فنی','توضیحات','نوع فنی']
    res_df=pd.DataFrame(columns=col_names)
    for f in files:
        print(f,asset)
        df=pd.read_excel(f,sheet_name=asset,header=0,names=col_names)
        res_df=pd.concat([res_df,df])
    
    f_name='./results/{}.xlsx'.format(asset)
    res_df.to_excel(f_name,sheet_name=asset,index=False)



if __name__ == "__main__":
    
    
    mydirectories= ['abdi','azizi','fahimi','ghafari','hakimzadeh',
                    'hossaini','kargar','mansourian','mehdizadeh',
                    'nazemizadeh','tabibi','voshtani']
    
    ass_list = asset_list(mydirectories)
    
    #file_list=list_of_files(mydirectories)
    
    for asset in ass_list:
        files=find_files(asset, mydirectories)
        make_dataframe(asset, files)
        
    