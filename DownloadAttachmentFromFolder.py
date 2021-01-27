import win32com.client
import os
import datetime

date = datetime.datetime.today().strftime ('%m-%d-%Y') 

outlook = win32com.client.Dispatch( "Outlook.Application")
inbox = outlook.GetNamespace("MAPI").GetDefaultFolder(6).Folders(1)

for message in inbox.Items:
    if message.UnRead == True:
        print(message.Subject + '--------------Email Subject') 
        attachment = message.Attachments(1)
        print(attachment)
       
        time = message.ReceivedTime
        message.UnRead = False
        updtime = '%.10s' % time
        print(updtime)
        attachment.SaveAsFile(filepath + '\\' + updtime + ' ' + str(attachment)) 
