import requests
from bs4 import BeautifulSoup

# url="http://1soccer.com/odds/detail/id/1122653"
url="http://1soccer.com/Gameschedule/score/type/fb/gametime/over/"


r = requests.get(url)
r.encoding = r.apparent_encoding


demo = r.text
soup = BeautifulSoup(demo, "html.parser")  # 给出解析器


# print(soup.prettify())

print(soup.title)
# print(soup.a)
# print(soup.a.name)
# print(soup.a.parent.parent.name)

# tag = soup.a
# print(tag.attrs)    #属性
