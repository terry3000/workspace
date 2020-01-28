from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import time

chrome_options = Options()
chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
chrome_driver = "D:\chrome\chromedriver.exe"
browser = webdriver.Chrome(chrome_driver, chrome_options=chrome_options)


def openLink():
    browser.get('https://ln.122.gov.cn/m/login')
    browser.implicitly_wait(15)


def inputId():
    # 用户名密码
    browser.find_element_by_xpath(
        '//*[@id="inputId"]').send_keys('210381198705280816')
    browser.find_element_by_xpath(
        '//*[@id="inputPassword"]').send_keys('123456mm')
    browser.find_element_by_xpath('//*[@id="csessionid"]').click()
    time.sleep(3)
    # 点击提交登录
    browser.find_element_by_xpath('//*[@id="btnGryhdl"]').click()


def clickJdc():
    # 点击机动车业务
    browser.find_element_by_xpath('//*[@id="sidebar_menu_11"]/a').click()

    # 点击模拟选号
    browser.find_element_by_xpath(
        '//*[@id="mem-content"]/div/div[2]/ul/li[1]/p[2]/a[2]').click()

    # 选号第一步，选择车管所
    browser.find_element_by_xpath('//*[@id="glbm"]/option[2]').click()
    time.sleep(0.5)
    browser.find_element_by_xpath('//*[@id="next"]').click()

    # 点击同意阅读  等待6秒
    time.sleep(5.5)
    browser.find_element_by_xpath('//*[@id="agree"]').click()


def queren():
    # 第二步-1  号牌种类
    browser.find_element_by_xpath('//*[@id="hpzl"]/option[2]').click()

    # 第二步-2  合格证号
    browser.find_element_by_xpath('//*[@id="zsbh"] ').send_keys('123456mm')
    # 第二步-3  品牌型号  等待10秒    1V2UR2CA
    time.sleep(15)

    # 第二步-4    车辆识别代号   //*[@id="clsbdh"]    .send_keys('123456123451234512')
    browser.find_element_by_xpath('//*[@id="clsbdh"]').send_keys('123456123451234512')

    #   提交
    browser.find_element_by_xpath('//*[@id="submit"]').click()

    time.sleep(2)

    # 点击模拟选号
    browser.find_element_by_xpath('//*[@id="startMnxh"]').click()
    time.sleep(0.5)

    # 点击验证手机
    browser.find_element_by_xpath('//*[@id="submit"]').click()
    time.sleep(0.5)

    # 点击自编
    browser.find_element_by_xpath('//*[@id="tab_zb"]').click()


def text_create(lst):
   
    file = open('C:\\Users\\hyt\\Desktop\\chepai.txt','w')
    for i in range(len(lst)):
        s = str(lst[i]).replace('[','').replace(']','')#去除[],这两行按数据不同，可以选择
        s = s.replace("'",'').replace(',','') +'\n'   #去除单引号，逗号，每行末尾追加换行符
        file.write(s)
    file.close()
    print("保存文件成功") 

def chaxun():

    # 辽A
    browser.switch_to_frame('zbxh')
    s1 = browser.find_elements_by_class_name('active')
    # print("第1位选择了--辽"+s1[0].text)
    s1[0].click()
    
    #形成退格键
    dele = browser.find_elements_by_class_name('zfx')

    #创建列表
    lst = []

    # 第2位
    s2 = browser.find_elements_by_class_name('active')
    for i2 in range(len(s2)-2):
        s2[i2].click()
        print('第2位的个数是：'+str(len(s2)-2))
        print('当前点击的是：'+(s2[i2].text))
   
        #第3位
        s3 = browser.find_elements_by_class_name('active')
        for i3 in range(len(s3)-2):            
            s3[i3].click()
                       
            # 第4位
            s4 = browser.find_elements_by_class_name('active')            
            for i4 in range(len(s4)-2):                
                s4[i4].click()
            #     str1 = '辽A'+s2[i2].text+s3[i3].text+s4[i4].text
            #     lst.append(str1) 
            #     print(str1)
            # dele[0].click()   
              
                # 第5位
                s5 = browser.find_elements_by_class_name('active')
                for i5 in range(len(s5)-2):
                    s5[i5].click()
                    
                    # 第6位
                    s6 = browser.find_elements_by_class_name('active')
                    for i6 in range(len(s6)-2):
                        str1 = '辽A'+s2[i2].text+s3[i3].text+s4[i4].text+s5[i5].text+s5[i6].text
                        lst.append(str1)                     
                        print(str1)
                    dele[0].click()

                dele[0].click()
                print(lst)
                text_create(lst)  
            dele[0].click()
            text_create(lst) 
            print(lst) 
        dele[0].click()
        text_create(lst)  
        print(lst)
    dele[0].click()
    print(lst) 
    text_create(lst)     
   

#    , '辽AA0P25', '辽AA0P26', '辽AA0P27', '辽AA0P28', '辽AA0P29', '辽AA0P30', '辽AA0P31', '辽AA0P32', '辽AA0P33'
# 辽AA9AAA
# 辽AA9AAB
# 辽AA9AAx
# 辽AA9AA确认选号


# openLink()
# inputId()
# clickJdc()
# queren()
chaxun()




