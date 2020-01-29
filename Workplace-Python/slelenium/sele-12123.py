from selenium import webdriver
import time

browser = webdriver.Chrome()
browser.get('https://ln.122.gov.cn/m/login')
#browser.maximize_window()

#等待浏览器
browser.implicitly_wait(15)

#用户名密码
browser.find_element_by_xpath('//*[@id="inputId"]').send_keys('210381198705280816')
browser.find_element_by_xpath('//*[@id="inputPassword"]').send_keys('123456mm')
browser.find_element_by_xpath('//*[@id="csessionid"]').click()
time.sleep(3)
#点击提交登录
browser.find_element_by_xpath('//*[@id="btnGryhdl"]').click()
#点击机动车业务
browser.find_element_by_xpath('//*[@id="sidebar_menu_11"]/a').click()

#点击模拟选号 
browser.find_element_by_xpath('//*[@id="mem-content"]/div/div[2]/ul/li[1]/p[2]/a[2]').click()

#选号第一步，选择车管所     
browser.find_element_by_xpath('//*[@id="glbm"]/option[2]').click()  
time.sleep(0.5)
browser.find_element_by_xpath('//*[@id="next"]').click()  

#点击同意阅读  等待6秒
time.sleep(6)
browser.find_element_by_xpath('//*[@id="agree"]').click()


#第二步-1  号牌种类     
browser.find_element_by_xpath('//*[@id="hpzl"]/option[2]').click()

#第二步-2  合格证号    
browser.find_element_by_xpath('//*[@id="zsbh"] ').send_keys('123456mm')
#第二步-3  品牌型号  等待10秒    1V2UR2CA
time.sleep(15)

#第二步-4    车辆识别代号   //*[@id="clsbdh"]    .send_keys('123456123451234512')
browser.find_element_by_xpath('//*[@id="zsbh"] ').send_keys('123456123451234512')

#   提交   
browser.find_element_by_xpath('//*[@id="submit"]').click()

time.sleep(2)

#点击模拟选号   //*[@id="startMnxh"]



#点击验证手机    //*[@id="submit"]
