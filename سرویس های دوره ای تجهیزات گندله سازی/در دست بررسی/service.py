import pandas as pd
import os

df=pd.read_excel("service-by-group-latest.xlsx",sheet_name="Table1",dtype=str)
#print(df.info())

gb_category_class = df.groupby('Table3.category-class')
#col_name = ['کد کارت فعالیت','ترتیب','شرح کار یا فعالیت','زمان انجام دقیقه','زمان انجام ساعت','شرح و دستورالعمل','نکات ایمنی','نرخ انجام','active']
#final_df = pd.DataFrame(columns=col_name)

for item1 in gb_category_class.groups:
    i=1
    sub_def1 = gb_category_class.get_group(item1)
    group_class_id=sub_def1['Table3.Column1'].to_list()[0]
    app_id = sub_def1['app_machine_group'].to_list()[0]
    gb_noekar = sub_def1.groupby('noekar.ID')
    for item2 in gb_noekar.groups:
        sub_def2 = gb_noekar.get_group(item2)
        gb_priod = sub_def2.groupby('parseh_priod')  
        for item3 in gb_priod.groups:
            sub_def3 = gb_priod.get_group(item3)
            gb_status = sub_def3.groupby('status')
            for item4 in gb_status.groups:
                print(group_class_id,item1,item2,item3,item4)
                path = ".\mp\{}\\".format(app_id)
                isExist = os.path.exists(path)
                if not isExist:
                    os.makedirs(path)
                f_name="{}({})PM.{}.{}.{}.{}.{}.xlsx".format(path,app_id,item2,item3,item4,group_class_id,str(i).zfill(4))
                code_faaliat="PM.{}.{}.{}.{}.{}".format(item2,item3,item4,group_class_id,str(i).zfill(4))
                print(f_name)
                sub_def4 = gb_status.get_group(item4) #.to_dict() # (orient='records')
                final_df = sub_def4[['sharhe_service_fa','zamane_anjam1','Count']].rename(columns={'sharhe_service_fa':'شرح کار یا فعالیت','zamane_anjam1': 'زمان انجام دقیقه','Count':'شرح و دستورالعمل'})
                final_df.insert(0,'کد کارت فعالیت',code_faaliat)
                final_df.to_excel(f_name,index=False)
                
        



    
    #group_class_id=sub_df['Table3.Column1'].to_list()[0]
    #noe_kar_id=sub_df['noekar.ID'].to_list()[0]
    #priod_id=sub_df['parseh_priod'].to_list()[0]
    #status_id=sub_df['status'].to_list()[0]

    
    #f_name="PM.{}.{}.{}.{}.{}".format(noe_kar_id,priod_id,status_id,group_class_id,str(i).zfill(4))
    #print(f_name)
    #sub_df.to_excel("file_name.xlsx",index=False)
    

