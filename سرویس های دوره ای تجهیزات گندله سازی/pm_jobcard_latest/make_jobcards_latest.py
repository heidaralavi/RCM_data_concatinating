import pandas as pd
import os
from xlsxwriter import Workbook


jobcards_dict = {'Code':[],'Name':[],'JCTypesName':[],'WOTradesName':[],'AssetClassName':[],'category_class':[],'DepartmentName':[],'CalendarPeriod':[],'CalenderUnitName':[],'MeterUnitName':[],'NeedSafety':[],'ShutDownTypes':[],'Duration':[],'DurationUnitName':[]}
JobcardActivity_dict = {'Sequence':[],'ActivityTitle':[],'DurationHRS':[],'DurationMIN':[],'Descroption':[],'DepartmentName':[],'JobCardCode':[]}

sorce_df = pd.read_excel('Asset_Services.xlsx',dtype=str)

gb_category_class = sorce_df.groupby('category-class')

for category_class_item in gb_category_class.groups:
    sub_def1 = gb_category_class.get_group(category_class_item)
    
    category_class_id=sub_def1['category-class-id'].to_list()[0]
    AssetClassName_txt=sub_def1['AssetClassName'].to_list()[0]
    
    gb_namad_group_code = sub_def1.groupby('namad_group_code')
    
    for namad_group_code_item in gb_namad_group_code.groups:
        sub_def2 = gb_namad_group_code.get_group(namad_group_code_item)
        
        gb_WOTradesName = sub_def2.groupby('WOTradesName')
        
        for WOTradesName_item in gb_WOTradesName.groups:
            sub_def3 = gb_WOTradesName.get_group(WOTradesName_item)
            
            WOTradesName_id=sub_def3['work_type_id'].to_list()[0]
            WOTradesName_txt=sub_def3['WOTradesName_Text'].to_list()[0]
            
            gb_Priod = sub_def3.groupby('Priod')
            
            for Priod_item in gb_Priod.groups:
                sub_def4 = gb_Priod.get_group(Priod_item)
                
                priod_txt = sub_def4['Priod_text'].to_list()[0]
                CalendarPeriod_text = sub_def4['CalendarPeriod'].to_list()[0]
                CalenderUnitName_text = sub_def4['CalenderUnitName'].to_list()[0]

                gb_run_stop = sub_def4.groupby('Run_Stop')
               
                for run_stop_item in gb_run_stop.groups:
                    sub_def5 = gb_run_stop.get_group(run_stop_item)
                                  
                    run_stop__txt = sub_def5['Run_Stop_Text'].to_list()[0]
                    ShutDownTypes_txt = sub_def5['ShutDownTypes'].to_list()[0]

                    gb_DepartmentName = sub_def5.groupby('DepartmentName')
                    
                    i=1
                    for DepartmentName_item in gb_DepartmentName.groups:
                        sub_def6 = gb_DepartmentName.get_group(DepartmentName_item)
                        sub_def6 = sub_def6. drop_duplicates()
                        JCTypesName_txt = sub_def5['JCTypesName'].to_list()[0]
                        Duration_txt = sub_def5['zamane_anjam'].astype(int).sum()

                 
                        jobcard_name = "سرویس دوره ای نت"+" {} ".format(WOTradesName_txt)+" {} ".format(priod_txt)+"تجهیزات"+" ({})- ".format(namad_group_code_item)+"<مجری <{}>> ".format(DepartmentName_item)+"{}".format(run_stop__txt)
                        code_jabcard = "PM.{}.{}.{}.{}.{}{}".format(WOTradesName_id,Priod_item,run_stop_item,category_class_id,namad_group_code_item,str(i).zfill(2))
                        jobcards_dict['Code'].append(code_jabcard)
                        jobcards_dict['Name'].append(jobcard_name)
                        jobcards_dict['JCTypesName'].append(JCTypesName_txt)
                        jobcards_dict['WOTradesName'].append(WOTradesName_item)
                        jobcards_dict['AssetClassName'].append(AssetClassName_txt)
                        jobcards_dict['category_class'].append(category_class_item)
                        jobcards_dict['DepartmentName'].append(DepartmentName_item)
                        jobcards_dict['CalendarPeriod'].append(CalendarPeriod_text)
                        jobcards_dict['CalenderUnitName'].append(CalenderUnitName_text)
                        jobcards_dict['MeterUnitName'].append('ساعت')
                        jobcards_dict['NeedSafety'].append('False')
                        jobcards_dict['ShutDownTypes'].append(ShutDownTypes_txt)
                        jobcards_dict['Duration'].append(Duration_txt)
                        jobcards_dict['DurationUnitName'].append('Minute')

                        
                        for (item1,item2,item3) in zip(sub_def6['service_title'].to_list(),sub_def6['zamane_anjam'].to_list(),sub_def6['Description'].to_list()):
                            JobcardActivity_dict['ActivityTitle'].append(item1)
                            JobcardActivity_dict['DurationMIN'].append(int(item2) % 60)
                            JobcardActivity_dict['DurationHRS'].append(int(item2) // 60)
                            JobcardActivity_dict['Descroption'].append(item3)

                        for j in range(len(sub_def6)):
                            JobcardActivity_dict['JobCardCode'].append(code_jabcard)
                            JobcardActivity_dict['Sequence'].append(j+1)
                            JobcardActivity_dict['DepartmentName'].append(DepartmentName_item)
                       
                        i=i+1
 


jobcard =pd.DataFrame(jobcards_dict)

writer_jobcard = pd.ExcelWriter('jobcard_new.xlsx', engine='xlsxwriter')
jobcard.to_excel(writer_jobcard, startrow=1 ,sheet_name = 'JobCard',header=False,index=False)
workbook = writer_jobcard.book
worksheet = writer_jobcard.sheets['JobCard']
(max_row, max_col) = jobcard.shape
column_settings = []
for header in jobcard.columns:
    column_settings.append({'header': header})
worksheet.add_table(0, 0, max_row, max_col - 1, {'columns': column_settings} )
text_format = workbook.add_format({'text_wrap' : True})
worksheet.set_column(0, max_col - 1, 20 ,text_format)
worksheet.set_column('A:A',25)
worksheet.set_column('B:B',50,text_format)
writer_jobcard.close()



jobcardActivity =pd.DataFrame(JobcardActivity_dict)

writer_jobcardActivity = pd.ExcelWriter('jobcardActivity_new.xlsx', engine='xlsxwriter')
jobcardActivity.to_excel(writer_jobcardActivity, startrow=1 ,sheet_name = 'JobCardActivity',header=False,index=False)
workbook = writer_jobcardActivity.book
worksheet = writer_jobcardActivity.sheets['JobCardActivity']
(max_row, max_col) = jobcardActivity.shape
column_settings = []
for header in jobcardActivity.columns:
    column_settings.append({'header': header})
worksheet.add_table(0, 0, max_row, max_col - 1, {'columns': column_settings} )
text_format = workbook.add_format({'text_wrap' : True})
worksheet.set_column(0, max_col - 1, 15 ,text_format)
worksheet.set_column('E:E',50)
worksheet.set_column('G:G',30,text_format)
writer_jobcardActivity.close()
