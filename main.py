import pandas as pd
import pyodbc
import SQL as s
import SchmiedUpdate as schmied
import warnings
from tkinter import messagebox

warnings.filterwarnings('ignore')

connx_string = r'DRIVER={SQL Server}; server=172.19.128.2\emeadb; database=emea_enventa_live; UID=usr_razek; PWD=wB382^%H3INJ'
conx = pyodbc.connect(connx_string)


df_bestellpos = pd.read_sql_query(s.bestellpos, conx)
df_bestellpos[['BELEGNR', 'Reference']]= df_bestellpos[['BELEGNR', 'Reference']].astype('int64')

df_total = schmied.update()

df_total.rename(columns={'Ref.':'Reference'}, inplace=True)
df_total['Reference']= df_total['Reference'].astype('int64')


df_final = pd.merge(df_bestellpos, df_total, on='Reference', how='left')
df_final.to_csv(r'L:\PM\AndreasSchmied.csv',index = False, sep=';')

messagebox.showinfo('Update Erfolgreich!', 'Das Schmied Update wurde erfolgreich durchgeführt.')