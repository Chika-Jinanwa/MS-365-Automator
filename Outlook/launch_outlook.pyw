from tkinter import Tk
from time import sleep
from tkinter.messagebox import showwarning
import win32com.client as win32

warn = lambda app: showwarning(app, 'Exit?')
def outlook():
    app = "Outlook"
    outlook = win32.gencache.EnsureDispatch('%s.Application' % app)
    mail =outlook.CreateItem(win32.constants.olMailItem)
    recip = mail.Recipients.Add('harryegege@gmail.com')
    subj = mail.Subject = 'Chika Testing His Automated Python Mail Script'
    body =  "Hey Harry!;). This is Chika's Automated Python Script sending you a message!"
    mail.Body = body
    mail.Send() 

    #open outbox folder to view sent mail
    ns = outlook.GetNamespace("MAPI")
    obox = ns.GetDefaultFolder(win32.constants.olFolderOutbox)
    obox.Display()
    obox.Items.Item(1).Display()

    warn(app)
    outlook.Quit()
if __name__ == '__main__':
    Tk().withdraw()
    outlook()

