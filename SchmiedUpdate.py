import pandas as pd
import pyodbc
from ast import literal_eval
import pandasql as ps
import warnings
import Email


#Feedback

#IMPORT UND DATENBEREINIGUNG DER ANDREAS SCHMIED INFORMATIONEN
#ERSTELLUNG DER FEHLERPROTOKOLLE

warnings.filterwarnings('ignore')

#Datenbankverbindung
connx_string = r'DRIVER={SQL Server}; server=172.19.128.2\emeadb; database=emea_enventa_live; UID=usr_razek; PWD=wB382^%H3INJ'
conx = pyodbc.connect(connx_string)

def update():
    #Email.get_mail()

    #Import CSV
    df = pd.read_csv(r'S:\Schmid\AS_Updates\Schmid_Update_neu.csv')
    df = df[['Ref.', 'ETD', 'ATD', 'ETA', 'ATA', 'DeliveryDate', 'Zollproblem']]
    df['Ref.'] = df['Ref.'].str.strip()
    df['Laenge'] = df['Ref.'].str.len()

    #Aufteilen der Daten in jene mit 1 Referenznummer und jenen mit mehreren
    df_clean = df[df['Laenge'] == 7].drop('Laenge', axis = 'columns')
    df_mult_ref = df[df['Laenge'] != 7]


    #Erstellen des Fehlerprotokolls für falsch Formatierte Referenzen
    query_multiple_lines_errors = """select * from df_mult_ref where Laenge not in (15,23)"""
    df_errors = ps.sqldf(query_multiple_lines_errors)


    #Vorbereiten für Zeilen mit mehreren
    query_multiple_lines = """select * from df_mult_ref where Laenge in (15,23)"""
    df_mult_ref = ps.sqldf(query_multiple_lines)



    #Separate Multiple Ref Numbers
        #Creating a List-Like Format
    df_mult_ref['Ref.'] = df_mult_ref['Ref.'].apply(lambda x: '[' + x + ']')
    df_mult_ref['Ref.'] = df_mult_ref['Ref.'].replace('/', ',', regex=True)
    df_mult_ref['Ref.'] = df_mult_ref['Ref.'].replace(',', ',', regex=True)
    df_mult_ref['Ref.'] = df_mult_ref['Ref.'].replace(';', ',', regex=True)
    df_mult_ref = df_mult_ref.drop('Laenge', axis ='columns')

        #Pandas Function "Explode" to create rows of List Objects
    df_mult_ref['Ref.'] = df_mult_ref['Ref.'].apply(literal_eval)
    df_mult_ref=df_mult_ref.explode('Ref.')

    #Combine all valid References
    query_union = """select * from df_clean union select * from df_mult_ref """
    df_total = ps.sqldf(query_union)

    #Find Duplicates
    df_duplicates = df_total[df_total['Ref.'].duplicated()==True]
    df_total = df_total.drop_duplicates(subset='Ref.')


    df_duplicates.to_excel(r'S:\EMEA\Kontrollabfragen\AndreasSchmied_Doppelte.xlsx',index = False)
    df_errors.to_excel(r'S:\EMEA\Kontrollabfragen\AndreasSchmied_Fehler.xlsx',index = False)

    return df_total


