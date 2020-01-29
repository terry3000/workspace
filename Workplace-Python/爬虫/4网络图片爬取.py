import requests
import os

url = "http://img0.dili360.com/pic/2019/11/27/5dde402bb8b5b6w77086630.jpg"

root = "C://Users//Administrator//Desktop"
path = root+url.split('/')[-1]

try:
    if not os.path.exists(root):
        os.mkdir(root)
    if not os.path.exists(path):
        r=requests.get(url)    
    with open(path, 'wb') as f:
        f.write(r.content)
        f.close
        print("图片保存成功")
except:
    print("爬取失败")
