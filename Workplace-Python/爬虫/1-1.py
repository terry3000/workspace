import requests
url="https://ln.122.gov.cn/"
r = requests.get(url)
r.encoding = r.apparent_encoding
print(r.text)