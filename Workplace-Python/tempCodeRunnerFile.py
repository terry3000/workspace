
import requests
form  bs4 import beautifulSoup

url = "https://www.baidu.com"
getHTMLText(url)


def getHTMLText(url):
    try:
        r.requests.get(url, timeout=30)
        r.raise_for_status()
        r.encoding = r.apparent_encoding
        return r.text
    except:
        return "产生异常"
