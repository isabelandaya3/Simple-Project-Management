import win32com.client
import os

outlook = win32com.client.Dispatch("Outlook.Application")
ns = outlook.GetNamespace("MAPI")
inbox = ns.GetDefaultFolder(6)

messages = inbox.Items.Restrict("@SQL=\"urn:schemas:httpmail:subject\" LIKE '%RFI #155%'")
print(f"Found {messages.Count} messages matching 'RFI #155'")

for msg in messages:
    print(f"\nSubject: {msg.Subject}")
    print(f"Attachments count: {msg.Attachments.Count}")
    for i in range(1, msg.Attachments.Count + 1):
        att = msg.Attachments.Item(i)
        print(f"  Attachment {i}: {att.FileName} ({att.Size} bytes)")
        
        # Save attachment to RFI 155 folder
        dest = r"\\sac-filsrv1\Projects\Structural-028\Projects\LEB\9.0_Const_Svcs\Mortenson\RFIs\RFI - 155 - LEB50 - Zone Of Influence Clairification"
        dest_path = os.path.join(dest, att.FileName)
        att.SaveAsFile(dest_path)
        print(f"  Saved to: {dest_path}")
