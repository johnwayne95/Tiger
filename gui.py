import tkinter
from tkinter import *
import time
import os

root = tkinter.Tk()
text = Text(root, spacing1 = 1, font=("Courier New", 14), width = 70, background = 'gray10', fg = 'gray80') 

def setup():
    global root, text
    text.pack(fill=X)

    # scroll = Scrollbar(root)
    # scroll.pack(side=RIGHT, fill=Y)
    # scroll.config(command=text.yview)

    trey1 = os.path.join(sys.path[0], 'trey_1.ico')
    root.title("Reports Automation")
    root.iconbitmap(trey1)

    # text.config(yscrollcommand=scroll.set)

def print(textinput):
    global root, text
    text.insert(END, "\n" + textinput + "\n")
    root.update()
    text.see("end")

def dots():
    global root, text
    text.insert(END, ".")
    root.update()
    text.see("end")

def loadingprint(textinput):
    global root, text
    text.insert(END, "\n" + textinput)
    root.update()
    text.see("end")

def update():
    root.mainloop()

def changeicon(file):
    root.iconbitmap(file)

def inputgoals():
    goalEntry = Entry(root, font=("Courier New", 16), width = 70, background = 'gray80', fg = 'gray10')
    goalEntry.pack(fill = X)

    
