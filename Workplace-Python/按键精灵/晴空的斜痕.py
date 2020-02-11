
import pyautogui
import os

p = 'D:\workspace\Workplace-Python\img\\'

# 打开battle；输入用户名，密码；点击进入游戏；
# 打开runasdate；点击提升权限；点击运行；点击开始；

os.startfile(r'D:\Program Files (x86)\Battle.net\Battle.net.exe')
while pyautogui.locateOnScreen(p +'logo.png') == None :
    print('等待battle登入口')
    pyautogui.PAUSE = 2.5
    if pyautogui.locateOnScreen(p +'logo.png') != None:
        print('登入窗口出现')
        break

pyautogui.hotkey('shift')
pyautogui.PAUSE = 0.2          
pyautogui.press(['7', '2', '2','2','5','9','3','@','q','q','.','c','o','m'])
pyautogui.PAUSE = 0.2
pyautogui.hotkey('tab')
pyautogui.PAUSE = 0.2
pyautogui.press(['1', '2', '3','4','5','6','m','m'])
pyautogui.PAUSE = 0.2
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

# 点击界面
pyautogui.PAUSE = 3.5
pyautogui.click(1900,900)

# 提升权限
pyautogui.PAUSE = 2.5
pyautogui.click(1700,1013)

# 点击运行
pyautogui.PAUSE = 2.5
pyautogui.click(1043,968)
