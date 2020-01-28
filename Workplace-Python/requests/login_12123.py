# coding:utf-8
import requests
# 登录请求地址
url = 'https://ln.122.gov.cn/user/m/login'
# 请求头
headers = {
        "User-Agent":"Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/78.0.3904.108 Safari/537.36"
}
# body数据
data = {
        'usertype': "1",
        'systemid': 'main',
        'username': "210381198705280816",
        'password': "MTIzNDU2bW0=",
        'csessionid': "10"
}
# 发送请求
r = requests.post(url,headers=headers,data=data)
# 判断是否登录成功
if '成功' in r.text:
    print('登录成功')
    print(r.text)
else:
    print('登录失败')