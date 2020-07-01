#pip install pywin32
from tkinter import Tk
from time import sleep
from tkinter.messagebox import showwarning
import win32com.client as win32


warn = lambda app: showwarning(app, "Exit?")
def excel():
    app = 'Excel'
    x1 = win32.gencache.EnsureDispatch('%s.Application' % app)
    ss = x1.Workbooks.Add()
    sh = ss.ActiveSheet
    x1.Visible = True
    sleep(2)
    sh.Cells(1,1).Value = 'Python-to-%s Demo'% app
    sleep(2)
    for i in range(3,8):
        sh.Cells(i,1).Value = 'Line %d' % i
        sleep(2)
    sh.Cells(i+2, 1).Value = "Hello World!"
    warn(app)
    ss.Close(False)
    x1.Application.Quit()

if __name__ == '__main__':
    Tk().withdraw
    excel()
