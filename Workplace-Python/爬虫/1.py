import requests

#无参数
r = requests.get('http://www.baidu.com')
print(r.url)

#有参数
payload = {
    'key1': 'value1',
    'key2': 'value2'
}
r = requests.get('http://www.baidu.com', params=payload)
print(r.url)