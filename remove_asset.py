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


def assetno_tagno():
    df = pd.read_excel('RCM Origenal Date.xlsx',sheet_name='Asset',header=None,usecols=[0,1,2])
    df = df.iloc[1: , :]
    dic=df.set_index(1).to_dict()[0]
    return dic


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
    
    
    mydirectories= ['C:/Users/heidar/Documents/GitHub/RCM_data_concatinating/results']
    #mydirectories= ['C:/Users/heidar/Desktop/temp/1']
    l_o_f=list_of_files(mydirectories)
    #print(l_o_f)
    #dict_of_asset_tag=assetno_tagno()
    #print(dict_of_asset_tag)
    #error_list=[]
    col_names=['AssetNumber','TagNo','نام جز','نوع جز','سطح تجهیز','پارت نامبر PN','سریال نامبر SN','جز بالاتر (سریال نامبر بالاسری)','مشخصه فنی','توضیحات','نوع فنی']
    must_be_asset_df=pd.DataFrame(columns=col_names)
    asset_tax_df = pd.DataFrame(columns=col_names)
    for file in l_o_f:
        xl = pd.ExcelFile(file)
        sh=xl.sheet_names
        df=pd.read_excel(file)
        #print(df)
        col_name='نوع جز'
        asset_df=df[df[col_name] == 'ASSET']
        must_be_asset_df=pd.concat([must_be_asset_df,asset_df])
        
        df.drop(df.index[df[col_name] =='ASSET'],inplace=True)
        col='سریال نامبر SN'
        df.drop_duplicates(subset=[col],keep='first',inplace=True)
        asset_tax_df= pd.concat([asset_tax_df,df])
        
    
    must_be_asset_df.to_excel('must_be_subasset.xlsx',index=False)
    asset_tax_df.to_excel('asset_taxonumi.xlsx',index=False)
    
    #    try:
    #        df['AssetNumber']=dict_of_asset_tag[sh[0]]
    #        f_name='./results/{}_new.xlsx'.format(sh[0])
    #        print(file)
    #        df.to_excel(f_name,sheet_name=sh[0],index=False)
    #    except:
    #        error_list.append(file)
        
    #print(error_list)
    
    #ass_list = asset_list(mydirectories)
    
    #file_list=list_of_files(mydirectories)
    
    #for asset in ass_list:
    #    files=find_files(asset, mydirectories)
    #    make_dataframe(asset, files)
        
    