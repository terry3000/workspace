import pyautogui


# move mouse to XY coordinates over num_second seconds
# 将鼠标移到XY坐标,花费x秒
pyautogui.moveTo(x, y, duration=num_seconds)  

# move mouse relative to its current position
# 相对于当前位置移动鼠标,花费x秒
pyautogui.moveRel(xOffset, yOffset, duration=num_seconds)  

# 将鼠标拖动到XY
pyautogui.dragTo(x, y, duration=num_seconds)  

# 相对于当前位置拖动鼠标
pyautogui.dragRel(xOffset, yOffset, duration=num_seconds)  

# Calling click() just clicks the mouse once with the left button at the mouse’s current location, but the keyword arguments can change that:
# 调用click（）只需在鼠标的当前位置用左键单击鼠标一次，但关键字参数可以更改：
# The button keyword argument can be 'left', 'middle', or 'right'.
pyautogui.click(x=moveToX, y=moveToY, clicks=num_of_clicks, interval=secs_between_clicks, button='left')


# All clicks can be done with click(), but these functions exist for readability. Keyword args are optional:
# 所有点击都可以用Clice()完成，但是这些函数存在可读性。关键字参数是可选的：
pyautogui.rightClick(x=moveToX, y=moveToY)
pyautogui.middleClick(x=moveToX, y=moveToY)
pyautogui.doubleClick(x=moveToX, y=moveToY)
pyautogui.tripleClick(x=moveToX, y=moveToY)

pyautogui.scroll(amount_to_scroll, x=moveToX, y=moveToY)

# Positive scrolling will scroll up, negative scrolling will scroll down:
# 正滚动向上滚动，负滚动向下滚动：
pyautogui.scroll(amount_to_scroll, x=moveToX, y=moveToY)

# Individual button down and up events can be called separately:
# 单独的按钮向下和向上事件可以单独调用：
pyautogui.mouseDown(x=moveToX, y=moveToY, button='left')
pyautogui.mouseUp(x=moveToX, y=moveToY, button='left')



######  键盘操作 ###### 

# useful for entering text, newline is Enter
# 用于文本输入，用回车下一行。
pyautogui.typewrite('Hello world!\n', interval=secs_between_keys)  

# A list of key names can be passed too:
# 也可以传递密钥名称列表：
pyautogui.typewrite(['a', 'b', 'c', 'left', 'backspace', 'enter', 'f1'], interval=secs_between_keys)

# Keyboard hotkeys like Ctrl-S or Ctrl-Shift-1 can be done by passing a list of key names to hotkey():
# 键盘热键（如Ctrl-S或Ctrl-Shift-1）可以通过向 hotkey() 传递键名列表来完成：
pyautogui.hotkey('ctrl', 'c')  # ctrl-c to copy
pyautogui.hotkey('ctrl', 'v')  # ctrl-v to paste

# Individual button down and up events can be called separately:
# 单独的按钮向下和向上事件可以单独调用：
pyautogui.keyDown(key_name)
pyautogui.keyUp(key_name)

pyautogui.press()


######  Message Box ###### 

pyautogui.alert('This displays some text with an OK button.')
pyautogui.confirm('This displays text and has an OK and Cancel button.')
# >>>'OK'
pyautogui.prompt('This lets the user type in a string and press OK.')
# >>>'This is what I typed in.'







######  截图操作 ###### 

#定义路径
LnPath = 'Workplace-Python\img\\'

im1 = pyautogui.screenshot()
im2 = pyautogui.screenshot('my_screenshot.png')
im3 = pyautogui.screenshot(region=(431,431,1741 ,981))


# 返回一个Pillow/PIL图像对象  >>> <PIL.Image.Image image mode=RGB size=1920x1080 at 0x24C3EF0>
# im1 = pyautogui.screenshot()

# 返回一个Pillow/PIL图像对象,并存储到文件  >>> <PIL.Image.Image image mode=RGB size=1920x1080 at 0x31AA198>
# im2 = pyautogui.screenshot(LnPath +'1.png')

# 如果你不想要整个屏幕的截图。可以传递要捕获的区域的左、上、宽和高的四个整数元组：
# im3 = pyautogui.screenshot(region=(0,0, 300, 400))

# 返回第一次出现的位置坐标 (left, top, width, height)   
# pyautogui.locateOnScreen(LnPath +'4.png')
# print(pyautogui.locateOnScreen(LnPath +'1.png'))
# >>> (863, 417, 70, 13)

# 返回在屏幕上找到的所有位置的坐标
# for i in pyautogui.locateAllOnScreen(LnPath +'1.png')
# >>> ...
# >>> (863, 117, 70, 13)
# >>> (623, 137, 70, 13)
# >>> (853, 577, 70, 13)
# >>> (883, 617, 70, 13)
# list(pyautogui.locateAllOnScreen(LnPath +'1.png'))
print(list(pyautogui.locateAllOnScreen(LnPath +'1.png')))
pyautogui.PAUSE = 2.5
pyautogui.PAUSE = 2.5
print(list(pyautogui.locateAllOnScreen(LnPath +'1.png')))



# returns center x and y
# pyautogui.locateCenterOnScreen('looksLikeThis.png')  
# >>> (898, 423)