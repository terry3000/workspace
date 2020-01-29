from selenium import webdriver

browser = webdriver.Chrome()
browser.get('https://m.12123.com/')
browser.maximize_window()

browser.implicitly_wait(15)

#用户名密码
browser.find_element_by_xpath('//*[@id="carNo"]').send_keys('M002K')
browser.find_element_by_xpath('//*[@id="frameNo"]').send_keys('082394')
browser.find_element_by_xpath('//*[@id="engineNo"]').send_keys('393919')

#点击提交登录
browser.find_element_by_xpath('//*[@id="btnSubmit"]').click()


