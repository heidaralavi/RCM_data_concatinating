import pandas as pd
import os
from xlsxwriter import Workbook

df=pd.read_excel("service-by-group-latest.xlsx",sheet_name="Table1",dtype=str)
#print(df.info())

gb_category_class = df.groupby('Table3.category-class')
#col_name = ['کد کارت فعالیت','ترتیب','شرح کار یا فعالیت','زمان انجام دقیقه','زمان انجام ساعت','شرح و دستورالعمل','نکات ایمنی','نرخ انجام','active']
#final_df = pd.DataFrame(columns=col_name)
code_faaliat_mojri = []
for item1 in gb_category_class.groups:
    i=1
    sub_def1 = gb_category_class.get_group(item1)
    group_class_id=sub_def1['Table3.Column1'].to_list()[0]
    gb_app_group = sub_def1.groupby('app_machine_group')
    for item5 in gb_app_group.groups:
        sub_def5 = gb_app_group.get_group(item5)
        app_id = sub_def5['app_machine_group'].to_list()[0]
        gb_noekar = sub_def5.groupby('noekar.ID')
        for item2 in gb_noekar.groups:
            sub_def2 = gb_noekar.get_group(item2)
            gb_priod = sub_def2.groupby('parseh_priod')  
            for item3 in gb_priod.groups:
                sub_def3 = gb_priod.get_group(item3)
                gb_status = sub_def3.groupby('status')
                for item4 in gb_status.groups:
                    print(group_class_id,item1,item2,item3,item4)
                    path = ".\mp\{}\\".format(item5)
                    isExist = os.path.exists(path)
                    if not isExist:
                        os.makedirs(path)
                    f_name="{}PM.{}.{}.{}.{}.{}{}.xlsx".format(path,item2,item3,item4,group_class_id,item5,str(i).zfill(2))
                    code_faaliat="PM.{}.{}.{}.{}.{}{}".format(item2,item3,item4,group_class_id,item5,str(i).zfill(2))
                    print(f_name)
                    sub_def4 = gb_status.get_group(item4) #.to_dict() # (orient='records')
                    final_df = sub_def4[['sharhe_service_fa','zamane_anjam1','Count','mojri.Department مجری']].rename(columns={'sharhe_service_fa':'شرح کار یا فعالیت','zamane_anjam1': 'زمان انجام دقیقه','Count':'شرح و دستورالعمل','mojri.Department مجری':'مجری'})
                    final_df.insert(0,'کد کارت فعالیت',code_faaliat)
                    final_df.insert(1,'ترتیب',range(1,len(final_df)+1))
                    final_df.insert(4,'زمان انجام ساعت',"")
                    final_df.insert(7,'نکات ایمنی',"")
                    final_df.insert(8,'نرخ انجام',"")
                    final_df.insert(9,'active',"YES")
                    code_faaliat_mojri.append((code_faaliat,final_df['مجری'].to_list()[0]))
                    sh_name = "({}){}".format(app_id[:3],code_faaliat[:20])
                    
                    writer = pd.ExcelWriter(f_name, engine='xlsxwriter')
                  
                    final_df.to_excel(writer,sheet_name = sh_name , startrow=1 ,header=False,index=False)
                    workbook = writer.book
                    worksheet = writer.sheets[sh_name]
                    (max_row, max_col) = final_df.shape
                    column_settings = []
                    for header in final_df.columns:
                        column_settings.append({'header': header})
                    worksheet.add_table(0, 0, max_row, max_col - 1, {'columns': column_settings})
                    cell_format = workbook.add_format({'text_wrap': True})
                    cell_format.set_align('vcenter')
                    cell_format.set_align('center')
                    cell_format1 = workbook.add_format({'text_wrap': True})
                    cell_format1.set_align('vcenter')
                    cell_format1.set_align('right')
                    worksheet.set_column('A:A', 24,cell_format)
                    worksheet.set_column('B:B', 5,cell_format)
                    worksheet.set_column('C:C', 20,cell_format)
                    worksheet.set_column('D:D', 5,cell_format)
                    worksheet.set_column('E:E', 20,cell_format)
                    worksheet.set_column('F:F', 50,cell_format1)
                    worksheet.set_column('G:G', 30,cell_format)
                    worksheet.set_column('H:J', 5,cell_format)
                    writer.close()
                    

code_faaliat_mojri_df = pd.DataFrame(code_faaliat_mojri,columns = ['code_faaliat','mojri'])                
code_faaliat_mojri_df.to_excel('vlookup_code_faaliat_mijri.xlsx',index=False)        



    
    #group_class_id=sub_df['Table3.Column1'].to_list()[0]
    #noe_kar_id=sub_df['noekar.ID'].to_list()[0]
    #priod_id=sub_df['parseh_priod'].to_list()[0]
    #status_id=sub_df['status'].to_list()[0]

    
    #f_name="PM.{}.{}.{}.{}.{}".format(noe_kar_id,priod_id,status_id,group_class_id,str(i).zfill(4))
    #print(f_name)
    #sub_df.to_excel("file_name.xlsx",index=False)
    

