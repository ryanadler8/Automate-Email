import win32com.client
import os
import datetime

date = datetime.datetime.today().strftime ('%b-%d-%Y') # pulling in the date and format we want

outlook = win32com.client.Dispatch( "Outlook.Application")
inbox = outlook.GetNamespace("MAPI").GetDefaultFolder(6).Folders(1)

for message in inbox.Items:
    if message.UnRead == True:
        print(message.Subject + '--------------Email Subject') #or whatever command you want to do
        attachment = message.Attachments(1)
        print(attachment)
        email = inbox.Items
        #message = email.GetNext()
        time = message.ReceivedTime
        message.UnRead = False
        updtime = '%.10s' % time
        print(updtime)
        attachment.SaveAsFile('S:\\Shared-Financial-Data-Governance\\Ryan\\Requester Query' + '\\' + updtime + ' ' + str(attachment))  # saving the attachment
