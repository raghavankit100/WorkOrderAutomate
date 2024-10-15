import win32com.client
import re

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

inbox = outlook.GetDefaultFolder(6)

messages = inbox.Items
messages = messages.Restrict("[Subject] = 'PresDate'")  

messages.Sort("[ReceivedTime]", True)

if messages.Count > 0:
    latest_message = messages[0]  
    email_body = latest_message.Body
    match = re.search(r'presDate=(\d+)', email_body)
    if match:
        pres_date_data = match.group(1)
        print(pres_date_data)
    else:
        print('presDate not found in the email body.')
else:
    print('No emails found with the subject "PresDate".')
