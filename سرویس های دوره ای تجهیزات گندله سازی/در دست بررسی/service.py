import pandas as pd
import os

df=pd.read_excel("service-by-group-latest.xlsx",sheet_name="Table1",dtype=str)
#print(df.info())

gb_category_class = df.groupby('Table3.category-class')

for item1 in gb_category_class.groups:
    i=1
    sub_def1 = gb_category_class.get_group(item1)
    group_class_id=sub_def1['Table3.Column1'].to_list()[0]
    app_id = sub_def1['app_machine_group'].to_list()[0]
    gb_noekar = sub_def1.groupby('noekar.ID')
    for item2 in gb_noekar.groups:
        sub_def2 = gb_noekar.get_group(item2)
        gb_status = sub_def2.groupby('status')
        for item3 in gb_status.groups:
            sub_def3 = gb_status.get_group(item3)
            gb_priod = sub_def3.groupby('parseh_priod')
            for item4 in gb_priod.groups:
                #print(group_class_id,item1,item2,item3,item4)
                f_name="({})PM.{}.{}.{}.{}.{}".format(app_id,item2,item4,item3,group_class_id,str(i).zfill(4))
                print(f_name)

    
    #group_class_id=sub_df['Table3.Column1'].to_list()[0]
    #noe_kar_id=sub_df['noekar.ID'].to_list()[0]
    #priod_id=sub_df['parseh_priod'].to_list()[0]
    #status_id=sub_df['status'].to_list()[0]

    
    #f_name="PM.{}.{}.{}.{}.{}".format(noe_kar_id,priod_id,status_id,group_class_id,str(i).zfill(4))
    #print(f_name)
    #sub_df.to_excel("file_name.xlsx",index=False)
    

