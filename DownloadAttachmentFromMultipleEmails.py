import win32com.client
import datetime

outlook = win32com.client.Dispatch( "Outlook.Application")
inbox = outlook.GetNamespace("MAPI").GetDefaultFolder(6).Folders(2)

for message in inbox.Items:
    if message.UnRead == True and message.Subject == 'SUBJECT_1':
        attachment = message.Attachments(1)
        email = inbox.Items
        time = message.ReceivedTime
        message.UnRead = False
        updtime =  time.strftime('%m-%d-%Y') 

        attachment.SaveAsFile(SAVE_LOCATION_1)  # saving the attachment

    elif message.UnRead == True and message.Subject == 'SUBJECT_2':
        attachment = message.Attachments(1)
        email = inbox.Items
        time = message.ReceivedTime
        message.UnRead = False
        updtime =  time.strftime('%m-%d-%Y')

        attachment.SaveAsFile(SAVE_LOCATION_2) 


    elif message.UnRead == True and message.Subject == 'SUBJECT_3':
        attachment = message.Attachments(1)
        email = inbox.Items
        time = message.ReceivedTime
        message.UnRead = False
        updtime =  time.strftime('%m-%d-%Y')

        attachment.SaveAsFile(SAVE_LOCATION_3) 


    elif message.UnRead == True and message.Subject == 'SUBJECT_4':
        attachment = message.Attachments(1)
        email = inbox.Items
        time = message.ReceivedTime
        message.UnRead = False
        updtime =  time.strftime('%m-%d-%Y')

        attachment.SaveAsFile(SAVE_LOCATION_4) 


    elif message.UnRead == True and message.Subject == 'SUBJECT_5':
        attachment = message.Attachments(1)
        email = inbox.Items
        time = message.ReceivedTime
        message.UnRead = False
        updtime =  time.strftime('%m-%d-%Y')

        attachment.SaveAsFile(SAVE_LOCATION_5) 


    else:

        pass
