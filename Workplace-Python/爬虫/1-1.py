import requests
url="http://1soccer.com/odds/detail/id/1122653"
r = requests.get(url)
r.encoding = r.apparent_encoding
print(r.text)