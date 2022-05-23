import os
import pandas as pd



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
    print(xl.sheet_names)





if __name__ == "__main__":
    
    mydirectories= ['azizi']
    list_of_files(mydirectories)
    for f in list_of_files(mydirectories):
        sheets_name(f)
        
    
    
    