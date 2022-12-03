import pandas as pd



col_names = ['location','machine_code','joze_machine','sharhe_service_fa',
             'service','tozihat','zamane_anjam','zamane_standard','priod',
             'noe_service','tarikh_anjam','maharat','vahede_ejraii','active']
df=pd.read_excel("service_list_row.xlsx",dtype=str,names=col_names)


#data cleaning
df = df.replace('ك', 'ک', regex=True)
df = df.replace('ي', 'ی', regex=True)
df = df.replace(chr(10),' ',regex=True) #Two Line replace by one Line

for n in range(6):
    df = df.replace('  ',' ',regex=True)

for item in col_names:
    df[item]=df[item].astype(str).str.strip()

df.to_excel('clean_data.xlsx',index=False)
del df
