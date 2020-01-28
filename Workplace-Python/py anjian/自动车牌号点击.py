#自动车牌号点击
import pyautogui

#定义字母
A = 738,550
B = 810,550
Num0 = 730,447
Num1 = 808.446

pyautogui.sleep = 1000

def C1(a,b,c,d,e):
    
    pyautogui.click(a)
    pyautogui.click(b)
    pyautogui.click(c)
    pyautogui.click(d)
    pyautogui.click(e)

#AB0A01
pyautogui.click(A)  #A
C1(B,Num0,A,Num0,Num1)

