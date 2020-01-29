import requests
from bs4 import BeautifulSoup
import bs4


def getHTMLText(url):
    try:
        r = requests.get(url)
        print(url)
        r.raise_for_status()
        r.encoding = r.apparent_encoding
        
        return r.text
    except:
        return ""


def fillUnivList(ulist, html):
    soup = BeautifulSoup(html, "html.parser") #使用html解析器
    for tr in soup.find('tbody').children:    #解析body
        if isinstance(tr,bs4.element.Tag):     #
            tds=tr('td')
            print("1233")
            
            ulist.append(tds[0].string,tds[1].string,tds[2].string)
            
    pass


def printUnivList(ulist, num):
    print("{:^10}\t{:^6}\t{:^10}".format("排名","学校名称","总分"))
    for i in range(num):
        u=ulist[i]
        print("{:^10}\t{:^6}\t{:^10}".format(u[0],u[1],u[2]))
    print("Suc"+str(num))


if __name__ == "__main__":
    uinfo = []
    
    url = "http://www.zuihaodaxue.com/zuihaodaxuepaiming2016.html"
    html = getHTMLText(url)
    fillUnivList(uinfo, html)
    printUnivList(uinfo, 20)  # 20所学校
