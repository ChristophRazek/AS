import pandas as pd
import pyodbc
import SQL as s
import SchmiedUpdate as schmied
import warnings
import pandasql as ps
from datetime import datetime
from tkinter import messagebox

warnings.filterwarnings('ignore')

connx_string = r'DRIVER={SQL Server}; server=172.19.128.2\emeadb; database=emea_enventa_live; UID=usr_razek; PWD=wB382^%H3INJ'
conx = pyodbc.connect(connx_string)

#Query Relevant Orders from Enventa
df_bestellpos = pd.read_sql_query(s.bestellpos, conx)
df_bestellpos[['BELEGNR', 'Reference']]= df_bestellpos[['BELEGNR', 'Reference']].astype('int64')

#Get Latest and cleaned Schmied data from SchmiedUpdate
df_total = schmied.update()
df_total.rename(columns={'Ref.':'Reference'}, inplace=True)
df_total['Reference']= df_total['Reference'].astype('int64')

#Merge and Save Update
df_final = pd.merge(df_bestellpos, df_total, on='Reference', how='left')
df_final.to_csv(r'L:\PM\AndreasSchmied.csv',index = False, sep=';')

#Prepare Feedback for Schmied
###Missing ATD
######Convert strings to date
df_final[['ETD','ATD','ETA','ATA', 'DeliveryDate']] = df_final[['ETD','ATD','ETA','ATA','DeliveryDate']].apply(pd.to_datetime,format='%d.%m.%Y' )
d = datetime.today()

no_atd_query = f"""select * from df_final where ATD is null and ETD < '{d}'"""
df_no_atd = ps.sqldf(no_atd_query)
df_no_atd = df_no_atd[['Reference','ETD','ATD','ETA','ATA','DeliveryDate']].drop_duplicates()


###Missing ATA
no_ata_query = f"""select * from df_final where ATA is null and ETA < '{d}'"""
df_no_ata = ps.sqldf(no_ata_query)
df_no_ata = df_no_ata[['Reference','ETD','ATD','ETA','ATA','DeliveryDate']].drop_duplicates()

###Zollproblem
no_customs_info_query = f"""select * from df_final where Zollproblem is null and DeliveryDate < '{d}'"""
df_no_customs_info = ps.sqldf(no_customs_info_query)
df_no_customs_info = df_no_customs_info[['Reference', 'DeliveryDate', 'Zollproblem']].drop_duplicates()
#print(df_no_customs_info.to_markdown())

#messagebox.showinfo('Update Erfolgreich!', 'Das Schmied Update wurde erfolgreich durchgeführt.')