from tkinter import Tk, Label, Entry, Button
from time import sleep
from tkinter.messagebox import showwarning
import win32com.client as win32

demo ='''
PRESENTATION TITLE
    test title
slide 1 title
    slide 1 bullet 1
    slide 1 bullet 2
slide 2 title
    slide 2 bullet 1 
    slide 2 bullet 2
'''


def txtToPPt(lines):
    speaker = win32.Dispatch("SAPI.SpVoice") 
    speak = "Hi, this is Chika, your dedicated Microsoft Office 365 AI assistant. I will launch a demo automated powerpoint presentation. Enjoy!"
    speaker.Speak(speak)
    ppt = win32.gencache.EnsureDispatch('Powerpoint.Application')
    pres = ppt.Presentations.Add()
    ppt.Visible = True
    sleep(2)
    nslide = 1
    for line in lines:
        if not line:
            continue
        line_data = line.split('    ')
        if len(line_data) ==1:
            title = (line ==line.upper()) #treats capitalized words as title
            if title:
                stype = win32.constants.ppLayoutTitle
            else:
                stype = win32.constants.ppLayoutText
            s = pres.Slides.Add(nslide, stype)
            ppt.ActiveWindow.View.GotoSlide(nslide)
            s.Shapes(1).TextFrame.TextRange.Text = line.title()
            body = s.Shapes(2).TextFrame.TextRange
            nline = 1
            nslide+=1
            sleep((nslide <4) and 0.25 or 0.01)
        else:
            line = '%s\r\n' % line.lstrip()
            body.InsertAfter(line)
            para = body.Paragraphs(nline)
            para.IndentLevel = len(line_data)-1
            nline+=1
            sleep((nslide<4) and 0.25 or 0.01)
    s = pres.Slides.Add(nslide,win32.constants.ppLayoutTitle)
    ppt.ActiveWindow.View.GotoSlide(nslide)
    sla= s.Shapes(1).TextFrame.TextRange
    sla.Text = 'Loading SlideShow!'.upper()
    sleep(2)
    for i in range(3, 0, -1):
        s.Shapes(1).TextFrame.TextRange.Text = str(i)
        sleep(1)
    pres.SlideShowSettings.ShowType=win32.constants.ppShowTypeSpeaker
    ss = pres.SlideShowSettings.Run()
    pres.ApplyTemplate(r'C:\Users\Chika Jinanwa\Documents\GitHub\MS-365-Automator\PowerPoint\Conference.potx')
    s.Shapes(1).TextFrame.TextRange.Text = 'END OF PRESENTATION!'
    s.Shapes(2).TextFrame.TextRange.Text = ''
def _start(ev=None):
    fn = en.get().strip()
    try:
        f = open(fn)
    except IOError:
        from io import StringIO
        f = StringIO(demo)
        en.delete(0, 'end')
        if fn.lower() == 'demo':
            en.insert(0, fn)
        else:
            import os
            en.insert(0, r"Demo (can't open %s: %s)"%(os.path.join(os.getcwd(), fn),str(e)))
            en.update_idletasks()
        txtToPPt(line.rstrip() for line in f)
        f.close()
if __name__ =='__main__':
    tk = Tk()
    lb = Label(tk, text= 'Enter File Path or Demo to begin Presentation')
    lb.pack()
    en = Entry(tk)
    en.bind('<Return>', _start)
    en.pack()
    en.focus_set()
    quit = Button(tk, text = 'Quit', command = tk.quit, fg= 'white', bg = 'red')
    quit.pack(fill= 'x', expand = True)
    tk.mainloop()