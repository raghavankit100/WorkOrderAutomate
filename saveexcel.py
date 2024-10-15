import os
import win32com.client

# Folder to save attachments
save_folder = r"C:\Users\araghav6\OneDrive - DXC Production\Desktop\workorderfile"

# Create the save folder if it doesn't exist
if not os.path.exists(save_folder):
    os.makedirs(save_folder)

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# Access the inbox (folder index 6 is the inbox)
inbox = outlook.GetDefaultFolder(6)
print(inbox)
# Get all items in the inbox
messages = inbox.Items

# Function to save attachments from a message
def save_attachments(message, save_folder):
    attachments = message.Attachments
    for attachment in attachments:
        attachment.SaveAsFile(os.path.join(save_folder, attachment.FileName))
        print(f"Saved {attachment.FileName} to {save_folder}")

# Process each message
for message in messages:
    
    try:
        subject = message.Subject
        if "Work order report" in subject:
            print(f"Processing email with subject: {subject}")
            save_attachments(message, save_folder)
    except Exception as e:
        print(f"Error processing message: {e}")

