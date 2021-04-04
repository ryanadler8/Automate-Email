# Automate Email
Script to iterate through Outlook Inbox and save attachments. Looks through a folder for unread emails and if the email is unread, it downloads and saves the attachment. I created a rule in outlook that would send the certain email to the folder. Using the Task Scheduler I have this script run everyday at 10:20 AM.

## Setup - Libraries
  - **pywin32** (https://pypi.org/project/pywin32/)
  - **os** (https://docs.python.org/3/library/os.html)
  - **Time** (https://docs.python.org/3/library/time.html)
  
## Explanation of Code
(DownloadAttachmentFromFolder.py)

Import libraries

![image](https://user-images.githubusercontent.com/55520621/105937919-7ef4e580-6024-11eb-8e69-630a28855267.png)

Created a variable to format the date.

![image](https://user-images.githubusercontent.com/55520621/105939162-a187fe00-6026-11eb-806f-a37d86429781.png)

Create a variable to call on the Outlook Application and another for the folder of where you will be downloading the email from. I used a print function to get the name of the folder. 

![image](https://user-images.githubusercontent.com/55520621/105939235-cb412500-6026-11eb-83fb-c197e305fb9f.png)

![image](https://user-images.githubusercontent.com/55520621/105939394-1ce9af80-6027-11eb-915a-459bb4a7ab24.png)

Next, I created a for loop to iterate through my folder. I created a for loop so on Monday mornings it would download the past three emails from Saturday, Sunday and Monday. Using an IF function after that to narrow down the search to only unread emails. Created a variable for the email attachment and a variable for the date. Lastly, I saved the file and used the variables I created to alter the name of the attachment. Throughout the script I used a few print statments to see what I am downloading, its date and the attachment name. 

![image](https://user-images.githubusercontent.com/55520621/105941151-c7af9d00-602a-11eb-9715-99c3faec7c51.png)

## Alternative Code
  - **.Subject == 'Daily Report'** - If there is an email that comes over with the same subject daily we can identify the subject and then perform an action, such as saving email/attachment. 
  - **.ReceivedTime** - Emails by date recieved you can use this to also perform actions. I have used if statements with *.Subject and .ReceivedTime* to identify emails and saved them. 
