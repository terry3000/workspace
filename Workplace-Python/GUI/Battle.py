from tkinter import ttk
import tkinter 
import pyautogui
import os

root = tkinter.Tk()
root.overrideredirect(True)
root.wm_attributes('-topmost',1)

p = 'Workplace-Python\img\\'

def btn():
    global btn

    # 点击win；打开battle；输入用户名，密码；点击进入游戏；
    # 点击win；打开runasdate；点击提升权限；点击运行；点击开始；
    pyautogui.PAUSE = 1.5
    pyautogui.press('win')  
    pyautogui.PAUSE = 1.5
    pyautogui.click(700,600)

    #等待battle程序出现，如果没找到，继续等待2.5，出现了跳出
    while pyautogui.locateOnScreen(p +'logo.png') == None :
        print('等待battle登入口')
        pyautogui.PAUSE = 2.5
        if pyautogui.locateOnScreen(p +'logo.png') != None:
            print('登入窗口出现')
            break

    pyautogui.hotkey('shift')
    pyautogui.PAUSE = 0.5          
    pyautogui.press(['3', '5', '2','8','5','5','5','@','q','q','.','c','o','m'])
    pyautogui.PAUSE = 0.5
    pyautogui.hotkey('tab')
    pyautogui.PAUSE = 0.5
    pyautogui.press(['1', '2', '3','4','5','6'])
    pyautogui.PAUSE = 0.5
    pyautogui.hotkey('enter')

    while pyautogui.locateOnScreen(p +'enter.png') == None :
        print('等待进入游戏的界面')
        pyautogui.PAUSE = 1
        if pyautogui.locateOnScreen(p +'enter.png') != None:
            print('进入游戏窗口出现')
            break

    pyautogui.click(p +'enter.png')       

    
    # 处理runasdate
    os.startfile(r'E:\迅雷下载\runasdate\RunAsDate.exe')
    pyautogui.PAUSE = 1.5
    pyautogui.click(1000,680,1.1)
    pyautogui.PAUSE = 1.5
    pyautogui.click(1024,680,1.1)
    pyautogui.PAUSE = 1.5
    pyautogui.click(1712,1012)



btn = ttk.Button(text="韩承熹",command = btn).pack()

root.mainloop()
