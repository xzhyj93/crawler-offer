#coding=utf-8
# 这个爬虫是为了爬一个大神写的2017年offer内容的信息，顺便学一下写爬虫
# 原信息网址： http://www.offershow.online:8000/index/

import urllib
import re
#首先下载安装以下三个包
import xlrd                         # http://pypi.python.org/pypi/xlrd
import xlwt                         # http://pypi.python.org/pypi/xlwt
from xlutils.copy import copy       # http://pypi.python.org/pypi/xlutils
import os

# 爬取html文档
def getHtml(url):
    page = urllib.urlopen(url)
    html = page.read().decode('utf-8')
    return html

# 这是公司列表，目的在于获取所有公司
def getHref(html):
    reg = re.compile(r'href="\/offerdetail\/\d+"')
    match = reg.findall(html);

    return match

# 定义一个全局变量，为了记录excel文档里存到了第几行
global nr
nr=0
#把数据保存进excel中
def saveToExcel(data):
    #定义文档路径
    #读取xls文档内容
    #复制到wb中
    #xlwt拷贝了原文件内容，然后添加新的内容，删除源文件，保存新文件
    xls = "offer.xls"
    wd = xlrd.open_workbook(xls,formatting_info=True)
    global nr
    nr = nr+1
    wb = copy(wd)
    ws = wb.get_sheet(0)
    ws.write(1,1,"opffer")

    for i in range(0,8):
        ws.write(nr,i,data[i])

    os.remove(xls)
    wb.save("offer.xls")

# 解析每条信息页面的内容，保存到数据中，存储进excel中
def getItem(item):
    html = getHtml("http://www.offershow.online:8000/"+item)
    reg = re.compile(r'class="ui-block-b"[\s\S]*?\<\/p');
    match = reg.findall(html)
    data = ['']*8
    i = 0
    data[0] = item[12:-1]       #序号
    print(item + " begin...")
    for i in range(1,8):
        it = match[i-1]
        index = it.find('data-theme')
        data[i] = (it[index+15:-3])
        #print("i=%d; data[i]=%s"%(i,data[i]))

    # reg = re.compile(r'textarea([\s\S]*)?textarea')
    # match = reg.findall(html)
    # data[i] = match[0][28:-2]
    saveToExcel(data)
    print(item + " down.")

if __name__ == "__main__":
    #
    html = getHtml("http://www.offershow.online:8000/sort/1")
    html = getHref(html)

    for item in html:
       getItem(item[7:-1]+"/")


