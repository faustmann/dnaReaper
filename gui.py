from tkinter import *
from tkinter import filedialog
from tkinter import messagebox

import doc


# build with command pyinstaller -Fw gui.py

def on_closing(zop):
    if messagebox.askokcancel("Quit", "Do you want to quit?"):
        top.destroy()

def storeCallBack(top,T1, T2):
    print(stepSizeVar.get())
    filename = filedialog.asksaveasfilename(initialdir=".", title="Select file",
                                           filetypes=(("doc files", "*.doc"), ("all files", "*.*")))

    stepSize = stepSizeVar.get()

    text1 = T1.get("1.0","end-1c")
    text2 = T2.get("1.0","end-1c")
    texts = [text1,text2]
    print(text1)

    name1=TName1.get("1.0","end-1c")
    name2=TName2.get("1.0","end-1c")
    names = [name1, name2]

    doc.generateAndPersistDocument(filename, names, texts, stepSize)
    messagebox.showinfo("sucess", "File %s generated"%(filename))


top = Tk()
top.title("DNA reaper 3000")
top.protocol("WM_DELETE_WINDOW", lambda: on_closing(top))

Label(top,
		 text="DNA reaper 3000",
		 fg = "blue",
		 font = "Verdana 10 bold").pack()


Label(top, text="Name1:").pack()
TName1 = Text(top, height=1, width=20)
TName1.pack()

Label(top, text="Name2:").pack()
TName2 = Text(top, height=1, width=20)
TName2.pack()

Label(top, text="Step Size:").pack()
selectOptions = tuple(range(6,61,3))
stepSizeVar = IntVar(top)
stepSizeVar.set(selectOptions[-1])
stepSizeMenu = OptionMenu(top, stepSizeVar, *selectOptions)
stepSizeMenu.pack()

B = Button(top, text ="Generate Doc File", command = lambda: storeCallBack(top,T1, T2))
B.pack(side=BOTTOM)

combFrame = Frame(top)

frame1 = Frame(combFrame)
Label(frame1, text="Insert the main sequence:").pack()
S1 = Scrollbar(frame1)
T1 = Text(frame1, height=20, width=40)
S1.pack(side=RIGHT, fill=Y)
T1.pack(side=BOTTOM, fill=Y)
S1.config(command=T1.yview)
T1.config(yscrollcommand=S1.set)

frame2 = Frame(combFrame)
Label(frame2, text="Insert the 2nd sequence:").pack()
S2 = Scrollbar(frame2)
T2 = Text(frame2, height=20, width=40)
S2.pack(side=RIGHT, fill=Y)
T2.pack(side=BOTTOM, fill=Y)
S2.config(command=T2.yview)
T2.config(yscrollcommand=S2.set)

frame1.pack(side=TOP, expand=True)
frame2.pack(side=BOTTOM, expand=True)
combFrame.pack(side=BOTTOM, expand=True)

top.mainloop()
