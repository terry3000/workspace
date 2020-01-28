import tkinter as tk  # 使用Tkinter前需要先导入

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



window = tk.Tk()
window.title('自动输入车牌号')
window.geometry('300x400') # 这里的乘是小x
window.wm_attributes('-topmost',1)

t = tk.Entry(window)

t.insert("insert", "B0A01、、、、")
print(t.get())

on_hit = False
def hit_me():
    global on_hit
    if on_hit == False:
        on_hit = True
        #AB0A01  第1组数
        pyautogui.click(A)  #A
        C1(B,Num0,A,Num0,Num1)

        #AB0A01  第2组数
        pyautogui.click(A)  #A
        C1(B,Num0,A,Num0,Num1)


        #AB0A01  第3组数
        pyautogui.click(A)  #A
        C1(B,Num0,A,Num0,Num1)


        #AB0A01  第4组数
        pyautogui.click(A)  #A
        C1(B,Num0,A,Num0,Num1)


        #AB0A01  第5组数
        pyautogui.click(A)  #A
        C1(B,Num0,A,Num0,Num1)



    else:
        on_hit = False
        
b = tk.Button(window, text='hit me', font=('Arial', 12), width=10, height=1, command=hit_me)
b.pack()


t.pack()# 第8步，主窗口循环显示window.mainloop()

window.mainloop()




