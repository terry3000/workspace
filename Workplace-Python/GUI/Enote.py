from tkinter import ttk
import tkinter 
import pyautogui

root = tkinter.Tk()
root.overrideredirect(True)
root.wm_attributes('-topmost',1)

def btn():
    global btn
    pyautogui.keyDown('ctrl')
    pyautogui.keyDown('alt')
    pyautogui.keyDown('n')
    pyautogui.keyUp('ctrl')
    pyautogui.keyUp('alt')
    pyautogui.keyUp('n')


btn = ttk.Button(text="新建印象笔记",command = btn).pack()

root.mainloop()