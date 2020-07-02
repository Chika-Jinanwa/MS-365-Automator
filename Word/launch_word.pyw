#pip install pywin32

from tkinter import Tk
from time import sleep
from tkinter.messagebox import showwarning
import win32com.client as win32

warn = lambda app: showwarning(app, 'Exit?')

def word():
    app = "Word"
    word = win32.gencache.EnsureDispatch('%s.Application' % app)
    doc = word.Documents.Add()
    word.Visible = True
    sleep(2)

    rng = doc.Range(0,0)
    rng.InsertAfter ('Python-to %s Test \r\n\r\n' % app)
    sleep(2)
    for i in range(4,9):
        rng.InsertAfter('Line %d\r\n' %i)
        sleep(2)
        rng.InsertAfter("\r\n Hello World! in MS Word")
        warn(app)
        doc.Close(False)
        word.Application.Quit()
if __name__ == '__main__':
    Tk().withdraw()
    word()