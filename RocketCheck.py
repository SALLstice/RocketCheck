import win32com.client
import os
import pythoncom

class Handler_Class(object):
  def OnNewMailEx(self, receivedItemsIDs):
    for ID in receivedItemsIDs.split(","):
        message = mapi.Session.GetItemFromID(ID)
        if message.UnRead:
            print(f"New Message Found: {message.Subject}")
            for attachment in message.Attachments:
                attachment.SaveAsFile(os.path.join(path, str(attachment.FileName)))
                message.UnRead = False

                if attachment.FileName[-3:] == "txt":
                    with open(os.path.join(path, str(attachment.FileName))) as file:
                        for line in file:
                            if "TODO:" in line:
                                newtask = outlook.CreateItem(win32com.client.constants.olTaskItem)
                                newtask.Subject = line
                                pdfpath = os.path.join(path, str(attachment.FileName)[:-23]+".pdf")
                                newtask.Attachments.Add(pdfpath)
                                newtask.Save()

outlook = win32com.client.DispatchWithEvents("Outlook.Application", Handler_Class)
mapi = outlook.GetNamespace("MAPI")

path = r'L:\Rocketbook Scans'
inbox = mapi.GetDefaultFolder(6)
rocketbookFolder = inbox.Folders.Item("Scanned Docs").Folders.Item("Rocketbook")

print("Running...")
pythoncom.PumpMessages()