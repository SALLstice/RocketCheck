import win32com.client
import os
import pythoncom
from win10toast import ToastNotifier
from datetime import date, timedelta
import pywintypes

class Handler_Class(object):
  def OnNewMailEx(self, receivedItemsIDs):
    for ID in receivedItemsIDs.split(","):
        message = mapi.Session.GetItemFromID(ID)
        if message.SenderEmailAddress == "notes@email.getrocketbook.com":
            print(f"New Message Found: {message.Subject}")
            toaster.show_toast(f"New Rocketbook Message",message.Subject)
            for attachment in message.Attachments:
                attachment.SaveAsFile(os.path.join(path, str(attachment.FileName)))
                message.UnRead = False

                if attachment.FileName[-3:] == "txt":
                    with open(os.path.join(path, str(attachment.FileName))) as file:
                        for line in file:
                            if "TODO" in line:
                                newtask = outlook.CreateItem(win32com.client.constants.olTaskItem)
                                newtask.StartDate = pywintypes.Time(date.today())

                                #sets due date to yesterday if todo is high priority, to highlight red
                                if "!TODO" in line:
                                    newtask.DueDate = pywintypes.Time(date.today() - timedelta(days=1))
                                else:
                                    newtask.DueDate = pywintypes.Time(date.today())

                                if "xTODO" in line or "XTODO" in line:
                                    newtask.DateCompleted = pywintypes.Time(date.today())

                                #trims from front of string until finds TODO
                                TODOCount = 0
                                while TODOCount < 4:
                                    if line[0] in ["T", "O", "D"]:
                                        TODOCount += 1
                                    line = line[1:]

                                # keeps trimming from front until an alphanumeric character is found
                                while not (line[0].isalnum()):
                                    line = line[1:]

                                # sets subject to TODO line and attaches pdf of RB scan
                                newtask.Subject = line
                                pdfpath = os.path.join(path, str(attachment.FileName)[:-23] + ".pdf")
                                newtask.Attachments.Add(pdfpath)
                                newtask.Save()

outlook = win32com.client.DispatchWithEvents("Outlook.Application", Handler_Class)
mapi = outlook.GetNamespace("MAPI")
path = r'L:\Rocketbook Scans'
toaster = ToastNotifier()

print("Running...")
pythoncom.PumpMessages()