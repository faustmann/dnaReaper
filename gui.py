from tkinter import *
from tkinter import filedialog
from tkinter import messagebox

import doc


# build with command pyinstaller -Fw gui.py

def on_closing(zop):
    if messagebox.askokcancel("Quit", "Do you want to quit?"):
        top.destroy()

def storeCallBack(top,T):
    filename = filedialog.asksaveasfilename(initialdir=".", title="Select file",
                                           filetypes=(("doc files", "*.doc"), ("all files", "*.*")))

    text = T.get("1.0","end-1c")
    print(text)

    name1=TName1.get("1.0","end-1c")
    name2=TName2.get("1.0","end-1c")
    names = [name1, name2]

    doc.generateAndPersistDocument(filename, names, text)
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

B = Button(top, text ="Generate Doc File", command = lambda: storeCallBack(top,T))
B.pack(side=BOTTOM)


Label(top, text="Insert the sequence:").pack()
S = Scrollbar(top)
T = Text(top, height=40, width=40)
S.pack(side=RIGHT, fill=Y)
T.pack(side=BOTTOM, fill=Y)
S.config(command=T.yview)
T.config(yscrollcommand=S.set)

top.mainloop()