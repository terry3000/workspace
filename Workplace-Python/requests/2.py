import requests

def getHTMLText(url):
    try:
        r.requests.get(url,timeout=30)
        r.raise_for_status()
        r.encoding =r.apparent_encoding
        print(r.text)
        return r.text
    except:
        return "产生异常"

getHTMLText('http://www.baidu.com')