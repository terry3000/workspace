
import requests

def getHTMLText(url):
    try:
        r.requests.get(url)
        r.raise_for_status()
        r.encoding = r.apparent_encoding
        print(r.text)
        return r.text
    except:
        return "产生异常"

if __name__ == "__main__":
   url = "http://www.baidu.com"
   getHTMLText(url)