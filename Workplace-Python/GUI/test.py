from tkinter import ttk
import tkinter 

root = tkinter.Tk()
root.overrideredirect(True)

def btn():
    global btn
    print("123")


btn = ttk.Button(text="最短",command = btn).pack()

root.mainloop()