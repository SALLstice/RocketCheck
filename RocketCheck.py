import win32com.client
import os
import pythoncom
from win10toast import ToastNotifier

class Handler_Class(object):
  def OnNewMailEx(self, receivedItemsIDs):
    for ID in receivedItemsIDs.split(","):
        message = mapi.Session.GetItemFromID(ID)
        if message.SenderEmailAddress == "notes@email.getrocketbook.com":
            print(f"New Message Found: {message.Subject}")
            for attachment in message.Attachments:
                attachment.SaveAsFile(os.path.join(path, str(attachment.FileName)))
                message.UnRead = False

                if attachment.FileName[-3:] == "txt":
                    with open(os.path.join(path, str(attachment.FileName))) as file:
                        for line in file:
                            if "TODO" in line:
                                newtask = outlook.CreateItem(win32com.client.constants.olTaskItem)
                                newtask.Subject = line
                                pdfpath = os.path.join(path, str(attachment.FileName)[:-23]+".pdf")
                                newtask.Attachments.Add(pdfpath)
                                newtask.Save()
                                toaster.show_toast("New Rocketbook TODO", line)

outlook = win32com.client.DispatchWithEvents("Outlook.Application", Handler_Class)
mapi = outlook.GetNamespace("MAPI")
path = r'L:\Rocketbook Scans'
toaster = ToastNotifier()

print("Running...")
pythoncom.PumpMessages()