from selenium import webdriver
import time

browser = webdriver.Chrome()
browser.get('https://ln.122.gov.cn/views/member/')
#browser.maximize_window()

#等待浏览器
browser.implicitly_wait(15)
#点击登录




