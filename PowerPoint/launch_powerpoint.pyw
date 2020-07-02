from tkinter import Tk
from time import sleep
from tkinter.messagebox import showwarning
import win32com.client as win32

warn = lambda app: showwarning(app, 'Exit?')
def ppoint():
    app = 'PowerPoint'
    ppoint = win32.gencache.EnsureDispatch('%s.Application'% app)
    pres = ppoint.Presentations.Add()
    ppoint.Visible = True
    sl = pres.Slides.Add(1, win32.constants.ppLayoutText)
    sleep(2)
    sla = sl.Shapes(1).TextFrame.TextRange 
    sla.Text = 'Python to- %s Demo' %app 
    sleep(2)
    slb = sl.Shapes(2).TextFrame.TextRange
    for i in range(3, 6):
        slb.InsertAfter('Line %d\r\n'% i)
        sleep(2)
    slb.InsertAfter("\r\n Hello World! with Powerpoint")
    warn(app)
    pres.Close()
    ppoint.Quit()
if __name__ == '__main__':
    Tk().withdraw()
    ppoint()