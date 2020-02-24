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
        doTheThing(message)

def doTheThing(message):
    if message.SenderEmailAddress == "notes@email.getrocketbook.com":
        print(f"New Rocketbook Message: {message.Subject}")
        toaster.show_toast(f"New Rocketbook Message",message.Subject)
        for attachment in message.Attachments:
            attachment.SaveAsFile(os.path.join(path, str(attachment.FileName)))
            message.UnRead = False

            if attachment.FileName[-3:] == "txt":
                with open(os.path.join(path, str(attachment.FileName))) as file:
                    for line in file:

                        #adds TODOs but ignores xTODOs
                        if "TODO" in line and not ("xTODO" in line or "XTODO" in line):
                            newtask = outlook.CreateItem(win32com.client.constants.olTaskItem)
                            newtask.StartDate = pywintypes.Time(date.today())
                            newtask.DueDate = pywintypes.Time(date.today())

                            #sets priority to high of !TODOs
                            if "!TODO" in line:
                                newtask.Importance = 2

                            #trims from front of string until finds TODO
                            TODOCount = 0
                            while TODOCount < 4:
                                if line[0] in ["T", "O", "D"]:
                                    TODOCount += 1
                                line = line[1:]

                            # keeps trimming from front until an alphanumeric character is found
                            while not (line[0].isalnum()):
                                line = line[1:]

                            # sets subject to TODOs line and attaches pdf of RB scan
                            newtask.Subject = line
                            pdfpath = os.path.join(path, str(attachment.FileName)[:-23] + ".pdf")
                            newtask.Attachments.Add(pdfpath)
                            newtask.Save()
                            print(f"New TODO Created: {newtask.Subject}")

outlook = win32com.client.DispatchWithEvents("Outlook.Application", Handler_Class)
mapi = outlook.GetNamespace("MAPI")
path = r'L:\Rocketbook Scans'
toaster = ToastNotifier()

print("Running...")

inbox = mapi.GetDefaultFolder(6)
messages = inbox.Folders["Scanned Docs"].Folders["Rocketbook"].Items

for message in messages:
    if message.UnRead:
        doTheThing(message)

while True:
    pythoncom.PumpMessages()