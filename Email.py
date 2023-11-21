import win32com.client as win32
import os
import pandas as pd

def get_mail():

    outlook = win32.Dispatch('Outlook.Application').GetNameSpace('MAPI')
    os.startfile("outlook")

    inbox = outlook.GetDefaultFolder(6)
    pyfolder = inbox.Folders['Schmied']

    messages = inbox.Items.Restrict("[SenderEmailAddress]='as-automate@andreas-schmid.de'")



    for message in messages:
        attachment = message.Attachments
        for a in attachment:
            a_name = str(a)
            a.SaveAsFile(fr'S:\Schmid\AS_Updates\Archiv\{a_name}')
            a.SaveAsFile(fr'S:\Schmid\AS_Updates\Schmid_Update_neu.csv')
        message.Move(pyfolder)







