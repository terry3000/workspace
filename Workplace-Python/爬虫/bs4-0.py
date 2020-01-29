
import requests
from bs4 import BeautifulSoup

r = requests.get("https://python123.io/ws/demo.html")
# print(r.text)

demo = r.text
soup = BeautifulSoup(demo, "html.parser")  # 给出解析器
print(soup.prettify())

print(soup.title)
print(soup.a)
print(soup.a.name)
print(soup.a.parent.parent.name)

tag = soup.a
print(tag.attrs)    #属性
print(tag.attrs['class']) 
print(tag.attrs['href'])

print(type(tag.attrs))   #<class 'dict'>   字典类型

print(tag.string)  #标签表达的内容