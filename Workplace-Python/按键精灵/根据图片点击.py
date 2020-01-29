import pyautogui

#定义路径
LnPath = 'Workplace-Python\py anjian\ln122'

#定义字母
A = LnPath+'\A.png'
B = LnPath+'\B.png'
Num0 = LnPath+'\0.png'

#im1 = pyautogui.screenshot()
#im2 = pyautogui.screenshot('my_screenshot.png')
#im3 = pyautogui.screenshot(region=(431,431,1741 ,981))


#定义字母位置
pyautogui.click(A)
pyautogui.screenshot()
pyautogui.PAUSE = 2.5

#pyautogui.screenshot('my_screenshot.png')
#pyautogui.locateOnScreen('my_screenshot.png')
pyautogui.click(B)



