import win32com.client as win32
import os
from datetime import date, datetime


def get_mail():

    outlook = win32.Dispatch('Outlook.Application').GetNameSpace('MAPI')
    #os.startfile("outlook")


    inbox = outlook.GetDefaultFolder(6)
    pyfolder = inbox.Folders['Schmied']

    #Filter auf Schmied-Update Emails
    messages = inbox.Items.Restrict("[SenderEmailAddress]='as-automate@andreas-schmid.de'")
    for message in messages:

        #ermittle die Datums-Information der Email
        received_date = str(message)[-10:]
        #Wandle die Info in eine Datum um
        dateobject = datetime.strptime(received_date, '%d.%m.%Y').date()

        #wenn email von heute, dann...
        if dateobject == date.today():
            #print(dateobject)
            attachment = message.Attachments
            for a in attachment:
                #print(a)
                a_name = str(a)
                a.SaveAsFile(fr'S:\Schmid\AS_Updates\Archiv\{a_name}')
                a.SaveAsFile(fr'S:\Schmid\AS_Updates\Schmid_Update_neu.csv')
            message.Move(pyfolder)


def send_mail():

    receivers = ['michael.pacher@emea-cosmetics.com','rudolf.swerak@emea-cosmetics.com','christoph.razek@emea-cosmetics.com','operation-airsea@andreas-schmid.de']

    today =date.today()
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = ";".join(receivers)
    mail.Subject = f'Feedback zu AS Update'
    mail.Body = f'Guten Morgen\n \nAnbei findest du die Anmerkungen zu den übermittelten Daten.\nBei Fragen bitte um Rückmeldung.\n \nMit freundlichen Grüßen\n \nChristoph Razek, M.Sc.\nERP&Prozessmanager\nemea Handelsgesellschaft mbH\nWallnerstraße 3/22\nA-1010 Wien\nTel.:    +43 1 535 10 01 - 280\nFax:    +43 1 535 10 01 - 900'
    mail.Attachments.Add(r'S:\Schmid\AS_Feedback.xlsx')

    mail.Display()
    mail.Save()
    #mail.Send()






