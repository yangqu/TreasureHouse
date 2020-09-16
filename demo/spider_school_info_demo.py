# -*- coding: utf-8 -*-
from bs4 import BeautifulSoup
import requests
import re
import sys
import xlwt
import xlrd
from xlutils.copy import copy
from xml.etree import ElementTree as ET
schoolDictP = {}
# 获取html
def getHtmlText(url, code="UTF-8"):
    try:
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/51.0.2704.63 Safari/537.36'
            ,'content-type': 'text/html'
        }
        r = requests.get(url, headers=headers, timeout=30)
        r.raise_for_status()
        r.encoding = code
        return r.text
    except:
        return "获取html异常"


# 解析地区，返回地区清单
'''
def getGroundList(htext):
    try:
        grounddict = {}
        soup = BeautifulSoup(htext, "html.parser")
        gdname = soup.find('dl', attrs={'class':'nobackground'})
        keyList = gdname.find_all('a')
        for i in range(1,len(keyList)):
            key = keyList[i].text
            val = keyList[i].get('href')
            grounddict[key] = val
        return grounddict
    except:
        print("getGroundList异常")
'''


# 解析页码
def getPageCode(htext, typeitem):
    try:
        soup = BeautifulSoup(htext, "html.parser")
        myz = soup.find(text='末页')
        print(myz)

    except:
        print("getPageCode异常")


# 解析学校信息，返回学校名称、地址、电话、网址
def getSchoolList(htext, fileAddress, cityitem, typeitem):

        schoolDict = {}
        soup = BeautifulSoup(htext, "html.parser")

        list = soup.find_all('dl')[1:]
        for item in list:
            try:
                print(item.find_all('span')[0].text)
                print(item.find_all('span')[1].text)
                print(item.find('p').text)
                schoolDict['城市'] = cityitem
                schoolDict['类型'] = typeitem
                schoolDict['地址'] = item.find_all('span')[0].text
                schoolDict['电话'] = item.find_all('span')[1].text
                schoolDict['学校名称'] = item.find('p').text
                savefile(schoolDict, fileAddress)
            except:
                print("getPageCode异常")
                continue

        """
        for item in sclist:
            print(item)
            print(cityitem)
            schoolDict['城市'] = cityitem
            schoolDict['类型'] = typeitem
            schoolDict['学习名称'] = item.find('p').text
            sl = item.find_all('li')
            schoolDict['地址'] = sl[0].text
            schoolDict['电话'] = sl[1].text
            schoolDict['网址'] = sl[2].text
            # f = open(fileAddress, 'a', encoding='utf-8')
            # f.write(str(schoolDict)  + '\n' )
            savefile(schoolDict, fileAddress)
        sclist1 = soup.find_all('dl', attrs={'class': 'left'})
        sclist2 = soup.find_all('dl', attrs={'class': 'right'})
        sclist = sclist1 + sclist2"""



# 保存到excel
def savefile(schoolDict, fileAddress):
    workbook = xlrd.open_workbook(fileAddress, 'w+b')
    sheet = workbook.sheet_by_index(0)
    wb = copy(workbook)
    ws = wb.get_sheet(0)
    rowNum = sheet.nrows
    ws.write(rowNum, 0, schoolDict['城市'])
    ws.write(rowNum, 1, schoolDict['类型'])
    ws.write(rowNum, 2, schoolDict['地址'])
    ws.write(rowNum, 3, schoolDict['电话'])
    ws.write(rowNum, 4, schoolDict['学校名称'])
    wb.save(fileAddress)


# 获取城市列表,城市由EXCEL文件存储
def getCityList():
    try:
        """
        cityFileAddress = r'D:\中国省市名称拼音.xlsx'
        file = xlrd.open_workbook(cityFileAddress)
        sheet = file.sheet_by_name('city')
        cityDic = {}
        for i in range(sheet.nrows):
            key = sheet.col_values(0)[i]
            value = sheet.col_values(1)[i].lower()
            cityDic[key] = value"""
        cityDic = {'北京':'beijing','上海':'shanghai','昆明':'kunmingshi',
                   '广州':'guangzhou','深圳':'shenzhenshi','成都':'chengdushi',
                   '杭州':'hangzhoushi','南京':'nanjingshi','重庆':'chongqing',
                   '西安':'xianshi','武汉':'wuhanshi','大连':'dalianshi','青岛':'qingdaoshi',
                   '厦门':'xiamenshi'}
        return cityDic
    except:
        print("getCityList失败")


def main():
    cityList = getCityList()
    typeList = {'小学': '/xiaoxue/', '初中': '/zhongxue/', '高中': '/gaozhong/'}
    for cityitem in cityList:
        for typeitem in typeList:
            searchUrl = 'https://www.mingxiaow.com/' + cityList[cityitem]
            fileAddress = 'D:/school.xls'
            htext = getHtmlText(searchUrl + typeList[typeitem])
            getSchoolList(htext, fileAddress, cityitem, typeitem)
            pagecode = 50
            if pagecode != 0:
                for i in range(1, pagecode):
                    h1text = getHtmlText(searchUrl + typeList[typeitem] + 'list_' + str(i) + '.html')
                    getSchoolList(h1text, fileAddress, cityitem, typeitem)


main()
