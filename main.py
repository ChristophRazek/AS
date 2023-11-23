import pandas as pd
import pyodbc
import Email
import SQL as s
import SchmiedUpdate as schmied
import warnings
import pandasql as ps
from datetime import datetime
from tkinter import messagebox

warnings.filterwarnings('ignore')

connx_string = r'DRIVER={SQL Server}; server=172.19.128.2\emeadb; database=emea_enventa_live; UID=usr_razek; PWD=wB382^%H3INJ'
conx = pyodbc.connect(connx_string)

#Abfrage offene Bestellungen
df_bestellpos = pd.read_sql_query(s.bestellpos, conx)
df_bestellpos[['BELEGNR', 'Reference']]= df_bestellpos[['BELEGNR', 'Reference']].astype('int64')

#Abfrage Schmied Update
df_total = schmied.update()
df_total.rename(columns={'Ref.':'Reference'}, inplace=True)
df_total['Reference']= df_total['Reference'].astype('int64')



#Verbinde offene Bestellungen mit Schmied Update
df_final = pd.merge(df_bestellpos, df_total, on='Reference', how='left')
df_final['Versanddatum'] = df_final['Versanddatum'].fillna(value='01.01.1900')

#Wenn Überschriebenes Lieferdatum (Versanddatum) vorhanden, nimm dieses, sonst berechnetes Lieferdatum (Delivery Date)
df_final['Datum_Final'] = df_final.apply(lambda x: x['Versanddatum'] if x['Versanddatum'] != '01.01.1900' else x['DeliveryDate'], axis=1 )
df_final.to_csv(r'L:\PM\AndreasSchmied.csv',index = False, sep=';')

########################################################################################################################
#Feedback für Schmied

feedback = {}
###Kein ATD
######Konvertiere als String gespeichertes Datum zu Datumsformat
df_final[['ETD','ATD','ETA','ATA', 'DeliveryDate']] = df_final[['ETD','ATD','ETA','ATA','DeliveryDate']].apply(pd.to_datetime,format='%d.%m.%Y' )
d = datetime.today()

no_atd_query = f"""select * from df_final where ATD is null and ETD < '{d}'"""
df_no_atd = ps.sqldf(no_atd_query)
df_no_atd = df_no_atd[['Reference','ETD','ATD','ETA','ATA','DeliveryDate']].drop_duplicates()

feedback['Missing ATD'] = df_no_atd

###Kein ATA
no_ata_query = f"""select * from df_final where ATA is null and ETA < '{d}'"""
df_no_ata = ps.sqldf(no_ata_query)
df_no_ata = df_no_ata[['Reference','ETD','ATD','ETA','ATA','DeliveryDate']].drop_duplicates()

feedback['Missing ATA'] = df_no_ata

'''###Zollproblem
no_customs_info_query = f"""select * from df_final where Zollproblem is null and DeliveryDate < '{d}'"""
df_no_customs_info = ps.sqldf(no_customs_info_query)
df_no_customs_info = df_no_customs_info[['Reference', 'DeliveryDate', 'Zollproblem']].drop_duplicates()
if df_no_customs_info.shape[0] > 0:
    feedback['Missing Customs'] = df_no_customs_info
#print(df_no_customs_info.to_markdown())'''

###Doppelte Referenz Nummer
df_duplicates = pd.read_excel(r'S:\EMEA\Kontrollabfragen\AndreasSchmied_Doppelte.xlsx')
df_duplicates.rename(columns={'Ref.':'Reference'}, inplace=True)
old_duplicates_query = """select * from df_duplicates where Reference not in (2212017,2304226) """
df_duplicates = ps.sqldf(old_duplicates_query)

feedback['Duplicates'] = df_duplicates


with pd.ExcelWriter(r'S:\Schmid\AS_Feedback.xlsx') as writer:
    for key in feedback:
        feedback[key].to_excel(writer, sheet_name=key, index=False)

Email.send_mail()

with open(r'S:\EMEA\Kontrollabfragen\AS_Update.txt', 'w') as f:
    f.write(f'Last MPS copied at: {d}')
    f.close()

messagebox.showinfo('Update Erfolgreich!', 'Das Schmied Update wurde erfolgreich durchgeführt.')

